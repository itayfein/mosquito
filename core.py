"""
MosquiOrder - לוגיקת סידור העמודות (ללא ממשק).

הקובץ הזה לא פותח רשת, לא מריץ תהליכים ולא כותב לשום מקום חוץ מהקובץ המסודר
שנשמר ליד קובץ המקור. כל קלט (קובץ אקסל, קובץ הגדרות) נבדק לפני שימוש.
"""

from __future__ import annotations

import json
import os
import re
import tempfile
import zipfile
from copy import copy
from dataclasses import dataclass, field

import openpyxl
import openpyxl.xml
from defusedxml import DefusedXmlException
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.cell_range import CellRange

if not openpyxl.xml.DEFUSEDXML:  # בלי defusedxml אין הגנה מ-XXE – לא רצים בכלל
    raise ImportError("defusedxml חייב להיות מותקן (pip install defusedxml)")

APP_NAME = "MosquiOrder"
APP_VERSION = "2.1.0"

# ---------------------------------------------------------------- סדר העמודות (ברירת מחדל, מ-Report Run)

DEFAULT_COLUMNS = [
    {"name": "מיקום מבחנה בקופסה", "width": 19.1},
    {"name": "WNV No.", "width": 9.2},
    {"name": "מזהה ביקור", "width": 11.2},
    {"name": "מזהה מבחנה", "width": 28.2},
    {"name": "שם נקודת ניטור", "width": 58.0},
    {"name": "שם אתר", "width": 13.4},
    {"name": "רשות שיפוט", "width": 11.0},
    {"name": "תאריך איסוף", "width": 11.1},
    {"name": "שם דוגם", "width": 11.2},
    {"name": "מחוז", "width": 11.5},
    {"name": "Point X ITM", "width": 11.1},
    {"name": "Point Y ITM", "width": 13.0},
    {"name": "בדיקות נגיף - מין", "width": 16.4},
    {"name": "בדיקות נגיף - כמות נקבות", "width": 21.9},
    {"name": "Sent/ Received at Sheba", "width": 22.9},
    {"name": "Date mosquitoes homogenized", "width": 29.1},
    {"name": "Date RNA extracted", "width": 18.6},
    {"name": "Date of TaqMan", "width": 15.1},
    {"name": "Result RT-PCR (Lin1+2)", "width": 21.0},
    {"name": "TaqMan Ct ENV", "width": 14.0},
    {"name": "TaqMan Ct NS3", "width": 14.0},
    {"name": "TaqMan Copies/ml NS3", "width": 20.0},
    {"name": "TaqMan Ct NS5", "width": 14.0},
    {"name": "TaqMan Copies/ml NS5", "width": 20.0},
    {"name": "Result Culture", "width": 14.0},
    {"name": "Date result out", "width": 15.0},
    {"name": "Primer & probe sets", "width": 18.0},
    {"name": "signature", "width": 12.0},
]

CONFIG_FILE = "columns.json"
OUTPUT_SUFFIX = " - מסודר"
HEADER_SCAN_ROWS = 15

# ---------------------------------------------------------------- מגבלות בטיחות לקלט

MAX_FILE_BYTES = 50 * 1024 * 1024            # קובץ אקסל עד 50MB
MAX_UNCOMPRESSED_BYTES = 400 * 1024 * 1024   # הגנה מ-zip bomb
MAX_COMPRESSION_RATIO = 200
MAX_ZIP_ENTRIES = 5000
MAX_CELLS = 5_000_000                        # הגנה מגיליון ענק שיתקע את המחשב
MAX_CONFIG_BYTES = 64 * 1024
MAX_COLUMNS = 200
ALLOWED_EXTENSIONS = (".xlsx", ".xlsm")


class InputError(Exception):
    """קלט לא תקין או חשוד. ההודעה מוצגת למשתמש כמו שהיא."""


# ---------------------------------------------------------------- הגדרות

@dataclass(frozen=True)
class Column:
    name: str
    width: float | None = None
    aliases: tuple[str, ...] = ()

    def keys(self):
        return {norm(self.name), *(norm(a) for a in self.aliases)} - {""}


@dataclass(frozen=True)
class Config:
    columns: tuple[Column, ...]
    freeze_header: bool = True
    min_header_matches: int = 2
    source: str = "מובנה"


def _parse_columns(raw):
    if not isinstance(raw, list) or not raw:
        raise InputError("columns חייב להיות רשימה לא ריקה")
    if len(raw) > MAX_COLUMNS:
        raise InputError(f"יותר מדי עמודות (מקסימום {MAX_COLUMNS})")
    cols, seen = [], set()
    for i, item in enumerate(raw, start=1):
        if isinstance(item, str):
            item = {"name": item}
        if not isinstance(item, dict) or set(item) - {"name", "width", "aliases"}:
            raise InputError(f"עמודה {i}: מותרים רק השדות name, width, aliases")
        name = item.get("name")
        if not isinstance(name, str) or not name.strip() or len(name) > 200:
            raise InputError(f"עמודה {i}: name חסר או ארוך מדי")
        width = item.get("width")
        if width is not None and (isinstance(width, bool) or not isinstance(width, (int, float))
                                  or not 1 <= width <= 255):
            raise InputError(f"עמודה {i}: width חייב להיות מספר בין 1 ל-255")
        aliases = item.get("aliases", [])
        if not isinstance(aliases, list) or not all(isinstance(a, str) and len(a) <= 200 for a in aliases):
            raise InputError(f"עמודה {i}: aliases חייב להיות רשימת מחרוזות")
        col = Column(name.strip(), float(width) if width is not None else None, tuple(aliases))
        dup = col.keys() & seen
        if dup:
            raise InputError(f"עמודה {i} ({name}): כותרת כפולה בהגדרות")
        seen |= col.keys()
        cols.append(col)
    return tuple(cols)


def default_config():
    return Config(_parse_columns(DEFAULT_COLUMNS))


def load_config(folder):
    """טוען columns.json מתיקיית התוכנה אם קיים. קובץ פגום = שגיאה ברורה, לא נפילה שקטה."""
    path = os.path.join(folder, CONFIG_FILE)
    if not os.path.isfile(path):
        return default_config()
    if os.path.getsize(path) > MAX_CONFIG_BYTES:
        raise InputError(f"{CONFIG_FILE} גדול מדי")
    try:
        with open(path, encoding="utf-8-sig") as f:
            data = json.load(f)
    except (OSError, ValueError) as e:
        raise InputError(f"לא ניתן לקרוא את {CONFIG_FILE}: {e}") from None
    if not isinstance(data, dict) or set(data) - {"columns", "freeze_header", "min_header_matches"}:
        raise InputError(f"{CONFIG_FILE}: מותרים רק המפתחות columns, freeze_header, min_header_matches")
    freeze = data.get("freeze_header", True)
    min_matches = data.get("min_header_matches", 2)
    if not isinstance(freeze, bool):
        raise InputError(f"{CONFIG_FILE}: freeze_header חייב להיות true/false")
    if isinstance(min_matches, bool) or not isinstance(min_matches, int) or not 1 <= min_matches <= 50:
        raise InputError(f"{CONFIG_FILE}: min_header_matches חייב להיות מספר שלם 1-50")
    return Config(_parse_columns(data.get("columns")), freeze, min_matches, source=CONFIG_FILE)


# ---------------------------------------------------------------- עזרים

_NON_WORD = re.compile(r"[\W_]+", re.UNICODE)


def norm(value):
    """מנרמל כותרת להשוואה: מתעלם מרווחים, סימני פיסוק ואותיות גדולות/קטנות.
    כך 'WNV No.' = 'wnv no' = 'WNV  No'."""
    if value is None:
        return ""
    return _NON_WORD.sub("", str(value)).casefold()


def output_path_for(path):
    base, _ = os.path.splitext(path)
    if base.endswith(OUTPUT_SUFFIX):
        base = base[: -len(OUTPUT_SUFFIX)]
    # תמיד xlsx: מאקרו (VBA) לא עוברים לקובץ המסודר
    return f"{base}{OUTPUT_SUFFIX}.xlsx"


def check_input_file(path):
    """בדיקות בטיחות לפני שפותחים קובץ. מחזיר רשימת הערות (למשל: הוסרו מאקרו)."""
    notes = []
    if not path.lower().endswith(ALLOWED_EXTENSIONS):
        raise InputError("סוג קובץ לא נתמך (רק xlsx / xlsm)")
    if not os.path.isfile(path):
        raise InputError("הקובץ לא נמצא")
    if os.path.islink(path):
        raise InputError("קיצורי דרך / קישורים סימבוליים לא נתמכים")
    size = os.path.getsize(path)
    if size == 0:
        raise InputError("הקובץ ריק")
    if size > MAX_FILE_BYTES:
        raise InputError(f"הקובץ גדול מדי (מעל {MAX_FILE_BYTES // 2**20}MB)")
    try:
        with zipfile.ZipFile(path) as z:
            infos = z.infolist()
            names = {i.filename for i in infos}
    except zipfile.BadZipFile:
        raise InputError("הקובץ אינו קובץ Excel תקין (ייתכן שהוא מוצפן בסיסמה או בפורמט xls ישן)") from None
    if len(infos) > MAX_ZIP_ENTRIES:
        raise InputError("מבנה קובץ חשוד (יותר מדי רכיבים)")
    total = sum(i.file_size for i in infos)
    packed = sum(i.compress_size for i in infos) or 1
    if total > MAX_UNCOMPRESSED_BYTES or total / packed > MAX_COMPRESSION_RATIO:
        raise InputError("מבנה קובץ חשוד (יחס דחיסה חריג)")
    if any(n.startswith("/") or ".." in n.split("/") for n in names):
        raise InputError("מבנה קובץ חשוד (נתיבים לא תקינים)")
    if "xl/workbook.xml" not in names:
        raise InputError("הקובץ אינו חוברת Excel תקינה")
    if any(n.lower().endswith("vbaproject.bin") for n in names):
        notes.append("הקובץ הכיל מאקרו (VBA) – הם לא הועתקו לקובץ המסודר")
    if any(n.startswith("xl/externalLinks/") for n in names):
        notes.append("קישורים לקבצים חיצוניים הוסרו")
    if any(n.startswith("xl/embeddings/") or n.startswith("xl/activeX/") for n in names):
        notes.append("אובייקטים מוטמעים (OLE/ActiveX) לא הועתקו")
    return notes


# ---------------------------------------------------------------- סידור גיליון

def find_header_row(ws, config):
    """שורת הכותרות = השורה (מתוך הראשונות) עם הכי הרבה כותרות מוכרות."""
    keys = set().union(*(c.keys() for c in config.columns))
    best_row, best_count = None, 0
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, HEADER_SCAN_ROWS),
                                         values_only=True), start=1):
        count = sum(1 for v in row if norm(v) in keys)
        if count > best_count:
            best_row, best_count = r, count
    return best_row, best_count


def _remap_range(rng, col_map, hr):
    """ממפה טווח תאים לעמודות החדשות. מחזיר מחרוזת טווח, או None אם אי אפשר לשמור אותו."""
    min_col, min_row, max_col, max_row = range_boundaries(str(rng))
    if max_row < hr:
        return str(rng)                                       # מעל הכותרת – לא זז
    new_cols = [col_map.get(c) for c in range(min_col, max_col + 1)]
    if None in new_cols:
        return None
    if new_cols != list(range(new_cols[0], new_cols[0] + len(new_cols))):
        return None                                           # הטווח התפצל
    if min_row < hr and new_cols[0] != min_col:
        return None                                           # חוצה את שורת הכותרת וזז
    return CellRange(min_col=new_cols[0], min_row=min_row, max_col=new_cols[-1], max_row=max_row).coord


@dataclass
class SheetReport:
    sheet: str
    header_row: int
    moved: int = 0
    missing: list = field(default_factory=list)
    extra: list = field(default_factory=list)
    duplicates: list = field(default_factory=list)
    dropped: list = field(default_factory=list)


def reorder_sheet(ws, config):
    """מסדר את העמודות בגיליון. מחזיר SheetReport, או None אם הגיליון לא רלוונטי."""
    hr, matches = find_header_row(ws, config)
    if hr is None or matches < config.min_header_matches:
        return None
    max_col, max_row = ws.max_column, ws.max_row
    if max_col * max_row > MAX_CELLS:
        raise InputError(f"הגיליון '{ws.title}' גדול מדי ({max_row}×{max_col} תאים)")

    report = SheetReport(ws.title, hr)
    src_headers = {c: norm(ws.cell(row=hr, column=c).value) for c in range(1, max_col + 1)}

    # בניית הסדר החדש: (עמודת מקור או None, כותרת לעמודה חסרה, רוחב מההגדרות)
    used, plan = set(), []
    for col in config.columns:
        keys = col.keys()
        hits = [c for c in range(1, max_col + 1) if c not in used and src_headers[c] in keys]
        if hits:
            used.add(hits[0])
            plan.append((hits[0], None, col.width))
            if len(hits) > 1:
                report.duplicates.append(col.name)
        else:
            plan.append((None, col.name, col.width))
            report.missing.append(col.name)

    # עמודות שלא מופיעות בהגדרות – נשמרות בסוף
    for c in range(1, max_col + 1):
        if c in used:
            continue
        if any(ws.cell(row=r, column=c).value is not None for r in range(hr, max_row + 1)):
            plan.append((c, None, None))
            report.extra.append(str(ws.cell(row=hr, column=c).value or get_column_letter(c)))

    report.moved = sum(1 for i, (src, _, _) in enumerate(plan, start=1) if src is not None and src != i)
    col_map = {src: new for new, (src, _, _) in enumerate(plan, start=1) if src is not None}
    unchanged = all(col_map.get(c, c) == c for c in range(1, max_col + 1)) and not report.missing

    # טווחים מאוחדים, אימות נתונים ועיצוב מותנה – ממופים למיקום החדש
    merged = []
    for rng in list(ws.merged_cells.ranges):
        if rng.max_row >= hr:
            new = _remap_range(rng, col_map, hr)
            ws.unmerge_cells(str(rng))
            if new:
                merged.append(new)
            else:
                report.dropped.append(f"איחוד תאים {rng.coord}")

    if not unchanged:
        for dv in ws.data_validations.dataValidation:
            parts = [_remap_range(r, col_map, hr) for r in dv.sqref.ranges]
            if None in parts:
                report.dropped.append(f"אימות נתונים {dv.sqref}")
            dv.sqref = openpyxl.worksheet.cell_range.MultiCellRange(" ".join(p for p in parts if p))
        ws.data_validations.dataValidation = [dv for dv in ws.data_validations.dataValidation
                                              if str(dv.sqref).strip()]

        # עיצוב מותנה עם נוסחאות תלוי במיקום העמודות – נשמר רק אם הוא לא זז
        from openpyxl.formatting.formatting import ConditionalFormattingList
        new_cf = ConditionalFormattingList()
        for cf in ws.conditional_formatting:
            parts = [_remap_range(r, col_map, hr) for r in cf.sqref.ranges]
            has_formula = any(rule.formula for rule in cf.rules)
            if None in parts or (has_formula and [str(r) for r in cf.sqref.ranges] != parts):
                report.dropped.append(f"עיצוב מותנה {cf.sqref}")
                continue
            for rule in cf.rules:
                new_cf.add(" ".join(parts), rule)
        ws.conditional_formatting = new_cf

    # צילום מצב של כל התאים מהכותרת ומטה
    snapshot = {}
    for c in range(1, max_col + 1):
        snapshot[c] = [(cell.value, copy(cell._style), copy(cell.hyperlink) if cell.hyperlink else None,
                        copy(cell.comment) if cell.comment else None)
                       for cell in (ws.cell(row=r, column=c) for r in range(hr, max_row + 1))]
    widths = {c: ws.column_dimensions[get_column_letter(c)].width for c in range(1, max_col + 1)}
    hidden = {c: ws.column_dimensions[get_column_letter(c)].hidden for c in range(1, max_col + 1)}
    header_style = copy(ws.cell(row=hr, column=min(used) if used else 1)._style)

    # ניקוי האזור
    for c in range(1, max(max_col, len(plan)) + 1):
        ws.column_dimensions[get_column_letter(c)].hidden = False
        for r in range(hr, max_row + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.hyperlink = None
            cell.comment = None
            cell.style = "Normal"

    # כתיבה בסדר החדש
    for new_c, (src, new_header, width) in enumerate(plan, start=1):
        letter = get_column_letter(new_c)
        dim = ws.column_dimensions[letter]
        if src is not None:
            for i, (value, style, link, comment) in enumerate(snapshot[src]):
                cell = ws.cell(row=hr + i, column=new_c)
                cell.value = value
                cell._style = copy(style)
                if link:
                    cell.hyperlink = link
                if comment:
                    cell.comment = comment
            if widths[src]:
                dim.width = widths[src]
            dim.hidden = bool(hidden[src])
        else:
            cell = ws.cell(row=hr, column=new_c)
            cell.value = new_header
            cell._style = copy(header_style)
            if width:
                dim.width = width

    for rng in merged:
        ws.merge_cells(rng)

    last = get_column_letter(len(plan))
    if ws.auto_filter.ref:
        ws.auto_filter.ref = f"A{hr}:{last}{max_row}"
    if config.freeze_header and not ws.freeze_panes:
        ws.freeze_panes = f"A{hr + 1}"
    return report


# ---------------------------------------------------------------- קובץ שלם

@dataclass
class FileResult:
    status: str            # ok / warn / err
    message: str
    output: str | None = None
    sheets: list = field(default_factory=list)


def _atomic_save(wb, target):
    """שומר לקובץ זמני באותה תיקייה ואז מחליף – אין קובץ חצי-כתוב אם משהו נכשל."""
    folder = os.path.dirname(target) or "."
    fd, tmp = tempfile.mkstemp(prefix="~$mosq_", suffix=".xlsx", dir=folder)
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, target)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def process_file(path, config=None):
    """מסדר קובץ אחד ושומר עותק. לא זורק חריגות – מחזיר FileResult."""
    config = config or default_config()
    path = os.path.abspath(path)
    out = output_path_for(path)
    try:
        if os.path.normcase(out) == os.path.normcase(path):
            raise InputError("שם הקובץ זהה לשם הפלט – שנה את שם הקובץ")
        notes = check_input_file(path)
        # keep_vba=False + keep_links=False: מאקרו וקישורים חיצוניים לא עוברים לפלט.
        # openpyxl משתמש אוטומטית ב-defusedxml (מותקן) להגנה מ-XML bomb / XXE.
        wb = openpyxl.load_workbook(path, keep_vba=False, keep_links=False)
        try:
            reports = [r for r in (reorder_sheet(ws, config) for ws in wb.worksheets) if r]
            if not reports:
                raise InputError("לא נמצאו כותרות מתאימות באף גיליון")
            _atomic_save(wb, out)
        finally:
            wb.close()
    except InputError as e:
        return FileResult("err", str(e))
    except PermissionError:
        return FileResult("err", "אין הרשאה לשמור – אם הקובץ המסודר פתוח באקסל, סגור אותו ונסה שוב")
    except DefusedXmlException:
        return FileResult("err", "הקובץ נחסם: מבנה XML חשוד (DTD / ישויות חיצוניות)")
    except MemoryError:
        return FileResult("err", "הקובץ גדול מדי לזיכרון")
    except Exception as e:  # קובץ פגום מסוג שלא צפינו – מוצג למשתמש ולא מפיל את התוכנה
        cause = e
        while cause is not None:                 # openpyxl עוטף את חסימת defusedxml ב-ValueError
            if isinstance(cause, DefusedXmlException):
                return FileResult("err", "הקובץ נחסם: מבנה XML חשוד (DTD / ישויות חיצוניות)")
            cause = cause.__cause__ or cause.__context__
        return FileResult("err", f"שגיאה בקריאת הקובץ ({type(e).__name__})")

    moved = sum(r.moved for r in reports)
    lines = []
    if moved == 0 and not any(r.missing for r in reports):
        lines.append("העמודות כבר היו מסודרות")
    else:
        lines.append(f"הוזזו {moved} עמודות למקומן" + (f" ב-{len(reports)} גיליונות" if len(reports) > 1 else ""))
    for label, attr in (("נוספו ריקות", "missing"), ("הועברו לסוף", "extra"),
                        ("כותרת כפולה (נלקחה הראשונה)", "duplicates"), ("לא הועבר", "dropped")):
        items = [x for r in reports for x in getattr(r, attr)]
        if items:
            lines.append(f"{label}: " + ", ".join(dict.fromkeys(items)))
    lines.extend(notes)
    status = "warn" if len(lines) > 1 else "ok"
    return FileResult(status, "\n".join(lines), out, reports)
