import json
import os
import random
import sys
import zipfile

import openpyxl
import pytest
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import core  # noqa: E402

NAMES = [c["name"] for c in core.DEFAULT_COLUMNS]


def make_book(path, headers, rows=3, title_rows=1, ext_cols=()):
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in range(1, title_rows + 1):
        ws.cell(row=r, column=1, value="דוח ריצה")
    hr = title_rows + 1
    for c, h in enumerate(list(headers) + list(ext_cols), start=1):
        ws.cell(row=hr, column=c, value=h)
        for r in range(rows):
            ws.cell(row=hr + 1 + r, column=c, value=f"{h}|{r}")
    ws.auto_filter.ref = f"A{hr}:{openpyxl.utils.get_column_letter(len(headers) + len(ext_cols))}{hr + rows}"
    wb.save(path)
    return hr


def read_header(path, hr):
    ws = openpyxl.load_workbook(path).active
    return [ws.cell(row=hr, column=c).value for c in range(1, ws.max_column + 1)], ws


def test_shuffled_columns_are_reordered(tmp_path):
    src = tmp_path / "run.xlsx"
    shuffled = NAMES[:]
    random.Random(1).shuffle(shuffled)
    hr = make_book(src, shuffled)
    res = core.process_file(str(src))
    assert res.status == "ok", res.message
    assert res.output.endswith(" - מסודר.xlsx")
    headers, ws = read_header(res.output, hr)
    assert headers == NAMES
    for c, name in enumerate(NAMES, start=1):          # הנתונים זזו יחד עם הכותרת
        assert ws.cell(row=hr + 1, column=c).value == f"{name}|0"
    assert ws.freeze_panes == f"A{hr + 1}"
    assert ws.auto_filter.ref.startswith(f"A{hr}:")
    assert open(src, "rb").read()[:2] == b"PK"           # המקור לא נפגע


def test_missing_and_extra_and_fuzzy_headers(tmp_path):
    src = tmp_path / "run.xlsx"
    headers = [h.upper().replace(".", "") for h in NAMES[:5][::-1]]   # 'WNV NO' בלי נקודה
    hr = make_book(src, headers, ext_cols=["הערות"])
    res = core.process_file(str(src))
    assert res.status == "warn"
    assert "נוספו ריקות" in res.message and "הערות" in res.message
    out, ws = read_header(res.output, hr)
    assert out[:len(NAMES)] == [headers[::-1][i] if i < 5 else NAMES[i] for i in range(len(NAMES))]
    assert out[-1] == "הערות"


def test_no_matching_headers(tmp_path):
    src = tmp_path / "x.xlsx"
    make_book(src, ["a", "b", "c"])
    res = core.process_file(str(src))
    assert res.status == "err"
    assert not os.path.exists(core.output_path_for(str(src)))


def test_merged_and_validation_are_remapped(tmp_path):
    src = tmp_path / "run.xlsx"
    headers = NAMES[::-1]
    hr = make_book(src, headers, rows=4)
    wb = openpyxl.load_workbook(src)
    ws = wb.active
    ws.merge_cells(start_row=hr + 1, start_column=1, end_row=hr + 2, end_column=1)  # עמודה אחרונה בסדר
    dv = DataValidation(type="list", formula1='"+,-"')
    dv.add(f"A{hr + 1}:A{hr + 4}")
    ws.add_data_validation(dv)
    wb.save(src)
    res = core.process_file(str(src))
    assert res.status in ("ok", "warn"), res.message
    ws = openpyxl.load_workbook(res.output).active
    last = openpyxl.utils.get_column_letter(len(NAMES))
    assert f"{last}{hr + 1}:{last}{hr + 2}" in [r.coord for r in ws.merged_cells.ranges]
    assert str(ws.data_validations.dataValidation[0].sqref) == f"{last}{hr + 1}:{last}{hr + 4}"


def test_xlsm_output_is_xlsx(tmp_path):
    src = tmp_path / "run.xlsm"
    make_book(tmp_path / "tmp.xlsx", NAMES[::-1])
    os.replace(tmp_path / "tmp.xlsx", src)
    res = core.process_file(str(src))
    assert res.status in ("ok", "warn"), res.message
    assert res.output.endswith(".xlsx")


def test_rerun_on_output_does_not_overwrite_itself(tmp_path):
    src = tmp_path / "run - מסודר.xlsx"
    make_book(src, NAMES)
    res = core.process_file(str(src))
    assert res.status == "err"


@pytest.mark.parametrize("content,msg", [
    (b"not a zip", "אינו קובץ Excel"),
    (b"", "ריק"),
])
def test_bad_files(tmp_path, content, msg):
    p = tmp_path / "bad.xlsx"
    p.write_bytes(content)
    res = core.process_file(str(p))
    assert res.status == "err" and msg in res.message


def test_zip_bomb_rejected(tmp_path):
    p = tmp_path / "bomb.xlsx"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("xl/workbook.xml", "<x/>")
        z.writestr("xl/big.xml", b"0" * (50 * 1024 * 1024))
    res = core.process_file(str(p))
    assert res.status == "err" and "חשוד" in res.message


def test_path_traversal_entries_rejected(tmp_path):
    p = tmp_path / "trav.xlsx"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("xl/workbook.xml", "<x/>")
        z.writestr("../evil.txt", "x")
    res = core.process_file(str(p))
    assert res.status == "err" and "חשוד" in res.message


def test_xxe_not_resolved(tmp_path):
    secret = tmp_path / "secret.txt"
    secret.write_text("TOPSECRET")
    src = tmp_path / "run.xlsx"
    make_book(src, NAMES)
    evil = tmp_path / "evil.xlsx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(evil, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                doctype = f'<!DOCTYPE worksheet [<!ENTITY e SYSTEM "file:///{secret.as_posix()}">]>'
                data = doctype.encode() + data.replace("דוח ריצה".encode(), b"&e;", 1)
            zout.writestr(item, data)
    res = core.process_file(str(evil))
    assert res.status == "err" and "XML חשוד" in res.message   # defusedxml חוסם DTD


def test_config_validation(tmp_path):
    (tmp_path / core.CONFIG_FILE).write_text(json.dumps({
        "columns": [{"name": "B", "width": 10}, "A", {"name": "C", "aliases": ["סי"]}],
        "freeze_header": False}), encoding="utf-8")
    cfg = core.load_config(str(tmp_path))
    assert [c.name for c in cfg.columns] == ["B", "A", "C"] and cfg.freeze_header is False

    for bad in ({"columns": []}, {"columns": [{"name": "A", "cmd": "x"}]}, {"columns": ["A", "a"]},
                {"columns": [{"name": "A", "width": 9999}]}, {"columns": ["A"], "evil": 1}, [1, 2]):
        (tmp_path / core.CONFIG_FILE).write_text(json.dumps(bad), encoding="utf-8")
        with pytest.raises(core.InputError):
            core.load_config(str(tmp_path))


def test_default_config_without_file(tmp_path):
    assert [c.name for c in core.load_config(str(tmp_path)).columns] == NAMES
