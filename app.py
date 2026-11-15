"""
MosquiOrder - סידור עמודות אקסל לניטור יתושים (ממשק משתמש).

בוחרים או גוררים קובץ אחד או יותר, והתוכנה שומרת לצד כל קובץ
עותק מסודר בשם "<שם הקובץ> - מסודר.xlsx". הקובץ המקורי לא משתנה.

התוכנה: ללא רשת, ללא הרשאות מנהל, ללא כתיבה לרג'יסטרי, ללא קיצורי דרך
וללא הרצת תהליכים חיצוניים. פתיחת קובץ/תיקייה נעשית רק על פי לחיצת המשתמש.
"""

import math
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import core
except ImportError as e:
    tk.Tk().withdraw()
    messagebox.showerror("חסרה ספרייה", f"{e}\n\npip install -r requirements.txt")
    raise SystemExit(1)

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

FROZEN = getattr(sys, "frozen", False)
APP_DIR = os.path.dirname(sys.executable if FROZEN else os.path.abspath(__file__))
RESOURCE_DIR = getattr(sys, "_MEIPASS", APP_DIR)

# ---------------------------------------------------------------- עיצוב

BG = "#f6f7fb"
SURFACE = "#ffffff"
SURFACE_2 = "#f0f2f7"
BORDER = "#e3e6ee"
TEXT = "#1b1e28"
MUTED = "#6b7385"
FAINT = "#a3a9b8"
ACCENT = "#6c4cf5"
ACCENT_HOVER = "#5a3be3"
ACCENT_SOFT = "#efebff"
ACCENT_2 = "#4f8cff"
FONT = "Segoe UI"

SCALE = 1.0  # מתעדכן לפי הגדלת התצוגה של Windows


def px(v):
    return int(round(v * SCALE))


BADGES = {
    "wait": ("ממתין", "#eef0f5", "#5f6779"),
    "run":  ("מעבד…", "#e3f1ff", "#1f74c9"),
    "ok":   ("סודר ✓", "#e3f8ec", "#16834a"),
    "warn": ("סודר ✓", "#fff4dc", "#b26a00"),
    "err":  ("שגיאה", "#fde8e8", "#c62828"),
}


def round_rect(canvas, x1, y1, x2, y2, r, **kw):
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return canvas.create_polygon(pts, smooth=True, **kw)


def hex_mix(a, b, t):
    a = [int(a[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(b[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * t):02x}" for x, y in zip(a, b))


class RoundButton(tk.Canvas):
    """כפתור מעוגל. style: primary / ghost."""

    def __init__(self, parent, text, command, width=150, height=36, style="primary",
                 parent_bg=BG, font_size=10):
        width, height = px(width), px(height)
        super().__init__(parent, width=width, height=height, bg=parent_bg,
                         highlightthickness=0, cursor="hand2", takefocus=1)
        if style == "primary":
            self.colors = {"bg": ACCENT, "hover": ACCENT_HOVER, "fg": "white", "outline": ""}
        else:
            self.colors = {"bg": parent_bg, "hover": SURFACE_2, "fg": TEXT, "outline": BORDER}
        self.command = command
        self.enabled = True
        self.shape = round_rect(self, 1, 1, width - 1, height - 1, px(12), fill=self.colors["bg"],
                                outline=self.colors["outline"], width=1)
        self.label = self.create_text(width / 2, height / 2, text=text, fill=self.colors["fg"],
                                      font=(FONT, font_size, "bold"))
        self.bind("<Enter>", lambda e: self.enabled and self.itemconfig(self.shape, fill=self.colors["hover"]))
        self.bind("<Leave>", lambda e: self.enabled and self.itemconfig(self.shape, fill=self.colors["bg"]))
        self.bind("<Button-1>", lambda e: self._fire())
        self.bind("<Return>", lambda e: self._fire())
        self.bind("<space>", lambda e: self._fire())

    def _fire(self):
        if self.enabled:
            self.command()

    def set_text(self, text):
        self.itemconfig(self.label, text=text)

    def set_enabled(self, enabled):
        self.enabled = enabled
        self.itemconfig(self.shape, fill=self.colors["bg"] if enabled else SURFACE_2)
        self.itemconfig(self.label, fill=self.colors["fg"] if enabled else FAINT)
        self.config(cursor="hand2" if enabled else "arrow")


def _ellipse(cx, cy, rx, ry, angle=0.0, n=40):
    a = math.radians(angle)
    pts = []
    for i in range(n):
        t = 2 * math.pi * i / n
        x, y = rx * math.cos(t), ry * math.sin(t)
        pts.append((cx + x * math.cos(a) - y * math.sin(a), cy + x * math.sin(a) + y * math.cos(a)))
    return pts


def mosquito_shapes():
    """יתוש במבט מלמעלה, ביחידות שבהן הגובה הכולל הוא בערך 1 והמרכז ב-(0,0)."""
    shapes = []
    for s in (-1, 1):
        shapes.append(("poly", _ellipse(s * 0.13, 0.03, 0.065, 0.21, angle=-s * 28), "wing"))
    legs = [
        [(0.04, -0.19), (0.17, -0.27), (0.27, -0.42)],
        [(0.05, -0.15), (0.24, -0.13), (0.40, -0.02)],
        [(0.04, -0.11), (0.17, 0.04), (0.29, 0.40)],
    ]
    for s in (-1, 1):
        for leg in legs:
            shapes.append(("line", [(s * x, y) for x, y in leg], 0.022, "body"))
        shapes.append(("line", [(s * 0.015, -0.31), (s * 0.08, -0.40)], 0.018, "body"))  # מחוש
    shapes.append(("line", [(0, -0.30), (0, -0.50)], 0.02, "body"))      # חדק
    shapes.append(("poly", _ellipse(0, 0.12, 0.045, 0.22), "body"))      # בטן
    shapes.append(("poly", _ellipse(0, -0.16, 0.072, 0.085), "body"))    # חזה
    shapes.append(("poly", _ellipse(0, -0.275, 0.042, 0.042), "body"))   # ראש
    return shapes


def draw_mosquito(c, cx, cy, size, body, wing):
    for kind, pts, *rest in mosquito_shapes():
        flat = [v for x, y in pts for v in (cx + x * size, cy + y * size)]
        if kind == "poly":
            c.create_polygon(flat, fill=wing if rest[-1] == "wing" else body, outline="")
        else:
            c.create_line(flat, fill=body, width=max(1, rest[0] * size), capstyle="round",
                          joinstyle="round")


def logo(parent, size=34, bg=BG):
    s = px(size)
    c = tk.Canvas(parent, width=s, height=s, bg=bg, highlightthickness=0)
    for i in range(s):
        c.create_line(i, 0, i, s, fill=hex_mix(ACCENT_2, ACCENT, i / s))
    r = px(9)
    for (x0, y0, start) in ((0, 0, 90), (s - 2 * r, 0, 0), (0, s - 2 * r, 180), (s - 2 * r, s - 2 * r, 270)):
        cx, cy = x0 + r, y0 + r
        sx = -1 if start in (90, 180) else 1
        sy = -1 if start in (0, 90) else 1
        c.create_rectangle(cx, cy, cx + sx * r, cy + sy * r, fill=bg, outline="")
        c.create_arc(x0, y0, x0 + 2 * r, y0 + 2 * r, start=start, extent=90, style="pieslice",
                     fill=hex_mix(ACCENT_2, ACCENT, cx / s), outline="")
    draw_mosquito(c, s / 2, s * 0.53, s * 0.8, "white", hex_mix("#ffffff", ACCENT, 0.35))
    return c


def open_path(path):
    """פותח קובץ/תיקייה שהתוכנה עצמה יצרה, באפליקציית ברירת המחדל. רק בלחיצת משתמש."""
    if path and os.path.exists(path) and (os.path.isdir(path) or path.lower().endswith(".xlsx")):
        os.startfile(path)  # nosec B606 - נתיב שנוצר ע"י התוכנה, ללא shell


# ---------------------------------------------------------------- ממשק

BaseTk = TkinterDnD.Tk if HAS_DND else tk.Tk


class App(BaseTk):
    def __init__(self):
        super().__init__()
        global SCALE
        SCALE = self.winfo_fpixels("1i") / 96
        self.title(f"סידור עמודות – {core.APP_NAME} {core.APP_VERSION}")
        w, h = px(660), px(500)
        self.geometry(f"{w}x{h}+{(self.winfo_screenwidth() - w) // 2}+{(self.winfo_screenheight() - h) // 3}")
        self.minsize(px(560), px(440))
        self.configure(bg=BG)
        self._set_icon()

        self.files = []        # נתיבים לפי סדר הוספה
        self.rows = {}         # נתיב -> רכיבי השורה ברשימה
        self.results = {}      # נתיב -> FileResult
        self.popup = None
        self.busy = False
        self.events = queue.Queue()
        self.config_error = None
        try:
            self.cfg = core.load_config(APP_DIR)
        except core.InputError as e:
            self.cfg, self.config_error = core.default_config(), str(e)

        wrap = tk.Frame(self, bg=BG)
        wrap.pack(fill="both", expand=True, padx=px(18), pady=px(14))
        self._build_topbar(wrap)
        self._build_dropzone(wrap)
        self._build_actions(wrap)
        self._build_progress(wrap)
        self._build_file_list(wrap)
        self._refresh()
        self.bind("<Configure>", lambda e: e.widget is self and self._close_popup(), add="+")
        self.bind("<Escape>", lambda e: self._close_popup())
        self.bind("<Control-o>", lambda e: self.pick_files())
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        if self.config_error:
            self.after(300, lambda: messagebox.showwarning(
                "קובץ הגדרות לא תקין",
                f"{self.config_error}\n\nהתוכנה תשתמש בסדר העמודות המובנה."))

    def _set_icon(self):
        ico = os.path.join(RESOURCE_DIR, "mosquito.ico")
        if os.path.exists(ico):
            try:
                self.iconbitmap(default=ico)
            except tk.TclError:
                pass

    def _on_close(self):
        if self.busy and not messagebox.askyesno("עיבוד פעיל", "יש קבצים בעיבוד. לצאת בכל זאת?"):
            return
        self.destroy()

    # --- שורה עליונה
    def _build_topbar(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(fill="x", pady=(0, px(12)))
        logo(bar).pack(side="right")
        titles = tk.Frame(bar, bg=BG)
        titles.pack(side="right", padx=px(10))
        tk.Label(titles, text="סידור עמודות", bg=BG, fg=TEXT,
                 font=(FONT, 13, "bold")).pack(anchor="e")
        src = "" if self.cfg.source == "מובנה" else "  ·  הגדרות מ-columns.json"
        tk.Label(titles, text=f"ניטור יתושים  ·  גרסה {core.APP_VERSION}{src}", bg=BG, fg=MUTED,
                 font=(FONT, 8)).pack(anchor="e")
        self.order_btn = RoundButton(bar, f"סדר העמודות ({len(self.cfg.columns)})", self.toggle_order,
                                     width=130, height=30, style="ghost", font_size=9)
        self.order_btn.pack(side="left")

    # --- חלון קופץ עם סדר העמודות
    def _close_popup(self):
        if self.popup is not None and self.popup.winfo_exists():
            self.popup.destroy()
        self.popup = None

    def toggle_order(self):
        if self.popup is not None:
            self._close_popup()
            return
        pop = tk.Toplevel(self, bg=BORDER)
        pop.overrideredirect(True)
        inner = tk.Frame(pop, bg=SURFACE)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(inner, text="סדר העמודות הקבוע", bg=SURFACE, fg=TEXT,
                 font=(FONT, 10, "bold")).pack(anchor="e", padx=px(14), pady=(px(10), px(2)))
        tk.Label(inner, text=f"מקור: {self.cfg.source}", bg=SURFACE, fg=MUTED,
                 font=(FONT, 7)).pack(anchor="e", padx=px(14), pady=(0, px(6)))
        grid = tk.Frame(inner, bg=SURFACE)
        grid.pack(padx=px(10), pady=(0, px(10)))
        cols = self.cfg.columns
        per_col = 12
        ncols = (len(cols) + per_col - 1) // per_col
        per_col = (len(cols) + ncols - 1) // ncols
        for i, col in enumerate(cols):
            gc, gr = ncols - 1 - i // per_col, i % per_col   # הטור הימני קודם
            cell = tk.Frame(grid, bg=SURFACE)
            cell.grid(row=gr, column=gc, sticky="e", padx=px(6), pady=px(1))
            tk.Label(cell, text=str(i + 1), bg=ACCENT_SOFT, fg=ACCENT_HOVER, width=3,
                     font=(FONT, 7, "bold")).pack(side="right")
            tk.Label(cell, text=col.name, bg=SURFACE, fg=TEXT, font=(FONT, 8)).pack(side="right", padx=(0, px(6)))
        pop.update_idletasks()
        x = self.order_btn.winfo_rootx()
        y = self.order_btn.winfo_rooty() + self.order_btn.winfo_height() + px(6)
        pop.geometry(f"+{x}+{y}")
        pop.bind("<Button-1>", lambda e: self._close_popup())
        self.popup = pop

    # --- אזור גרירה
    def _build_dropzone(self, parent):
        self.drop = tk.Canvas(parent, height=px(84), bg=BG, highlightthickness=0, cursor="hand2")
        self.drop.pack(fill="x")
        self.drop_hover = False
        self.drop.bind("<Configure>", lambda e: self._draw_drop())
        self.drop.bind("<Enter>", lambda e: self._set_drop_hover(True))
        self.drop.bind("<Leave>", lambda e: self._set_drop_hover(False))
        self.drop.bind("<Button-1>", lambda e: self.pick_files())
        if HAS_DND:
            for w in (self.drop, self):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)
            self.drop.dnd_bind("<<DropEnter>>", lambda e: self._set_drop_hover(True))
            self.drop.dnd_bind("<<DropLeave>>", lambda e: self._set_drop_hover(False))

    def _set_drop_hover(self, value):
        self.drop_hover = value
        self._draw_drop()
        return "copy"

    def _draw_drop(self):
        c = self.drop
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        hover = self.drop_hover
        round_rect(c, 1, 1, w - 1, h - 1, px(16), fill=ACCENT_SOFT if hover else SURFACE,
                   outline=ACCENT if hover else BORDER, dash=(5, 4), width=max(1, px(1)))
        cx, cy, r = w - px(46), h / 2, px(20)
        c.create_oval(cx - r, cy - r, cx + r, cy + r, fill=ACCENT if hover else ACCENT_SOFT, outline="")
        col = "white" if hover else ACCENT_HOVER
        c.create_polygon(cx, cy - px(10), cx - px(7.5), cy - px(2), cx - px(3), cy - px(2),
                         cx - px(3), cy + px(6), cx + px(3), cy + px(6), cx + px(3), cy - px(2),
                         cx + px(7.5), cy - px(2), fill=col, outline="")
        c.create_line(cx - px(8), cy + px(10), cx + px(8), cy + px(10), fill=col,
                      width=px(2), capstyle="round")
        main = "גרור לכאן קבצי אקסל או תיקייה" if HAS_DND else "בחר קבצי אקסל"
        c.create_text(w - px(80), cy - px(9), text=main, anchor="e", fill=TEXT, font=(FONT, 11, "bold"))
        sub = "או לחץ לבחירה (Ctrl+O)  ·  אפשר כמה קבצים יחד" if HAS_DND else "לחץ כאן (Ctrl+O)  ·  אפשר כמה קבצים יחד"
        c.create_text(w - px(80), cy + px(11), text=sub, anchor="e", fill=MUTED, font=(FONT, 8))

    # --- פס התקדמות
    def _build_progress(self, parent):
        self.progress = tk.Canvas(parent, height=px(4), bg=BG, highlightthickness=0)
        self.progress.pack(side="bottom", fill="x", pady=(px(8), 0))

    def _draw_progress(self, done, total):
        c = self.progress
        c.delete("all")
        if not total:
            return
        w = c.winfo_width()
        c.create_rectangle(0, 0, w, px(4), fill=SURFACE_2, outline="")
        fill_w = w * done / total
        c.create_rectangle(w - fill_w, 0, w, px(4), fill=ACCENT, outline="")   # מימין לשמאל

    # --- רשימת קבצים
    def _build_file_list(self, parent):
        head = tk.Frame(parent, bg=BG)
        head.pack(fill="x", pady=(px(14), px(6)))
        tk.Label(head, text="קבצים", bg=BG, fg=MUTED, font=(FONT, 9, "bold")).pack(side="right")
        self.count_lbl = tk.Label(head, bg=SURFACE_2, fg=TEXT, font=(FONT, 7, "bold"), padx=px(6))
        self.count_lbl.pack(side="right", padx=px(6))
        self.summary_lbl = tk.Label(head, bg=BG, fg=MUTED, font=(FONT, 8))
        self.summary_lbl.pack(side="right", padx=px(6))
        self.clear_lbl = tk.Label(head, text="נקה הכל", bg=BG, fg=MUTED, font=(FONT, 8), cursor="hand2")
        self.clear_lbl.pack(side="left")
        self.clear_lbl.bind("<Button-1>", lambda e: self.clear())
        self.clear_lbl.bind("<Enter>", lambda e: self.clear_lbl.config(fg=TEXT))
        self.clear_lbl.bind("<Leave>", lambda e: self.clear_lbl.config(fg=MUTED))

        self.list_canvas = tk.Canvas(parent, bg=BG, highlightthickness=0)
        self.list_canvas.pack(fill="both", expand=True)
        self.list_inner = tk.Frame(self.list_canvas, bg=BG)
        win = self.list_canvas.create_window(0, 0, window=self.list_inner, anchor="nw")
        self.list_inner.bind("<Configure>", lambda e: self.list_canvas.configure(
            scrollregion=self.list_canvas.bbox("all")))
        self.list_canvas.bind("<Configure>", lambda e: self.list_canvas.itemconfig(win, width=e.width))
        self.list_canvas.bind("<Enter>", lambda e: self.bind_all("<MouseWheel>", self._on_wheel))
        self.list_canvas.bind("<Leave>", lambda e: self.unbind_all("<MouseWheel>"))

        self.empty_lbl = tk.Frame(self.list_inner, bg=BG)
        art = tk.Canvas(self.empty_lbl, width=px(56), height=px(56), bg=BG, highlightthickness=0)
        draw_mosquito(art, px(28), px(30), px(52), "#cdd1dc", "#e6e8f0")
        art.pack()
        tk.Label(self.empty_lbl, text="עדיין לא נבחרו קבצים", bg=BG, fg=FAINT,
                 font=(FONT, 9)).pack(pady=(px(4), 0))

    def _on_wheel(self, e):
        if self.list_inner.winfo_height() > self.list_canvas.winfo_height():
            self.list_canvas.yview_scroll(int(-e.delta / 120), "units")

    def _add_row(self, path):
        row = tk.Frame(self.list_inner, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        row.pack(fill="x", pady=(0, px(6)))

        icon = tk.Canvas(row, width=px(26), height=px(26), bg=SURFACE, highlightthickness=0)
        round_rect(icon, 0, 0, px(26), px(26), px(7), fill="#1f7a4d", outline="")
        icon.create_text(px(13), px(13), text="X", fill="white", font=(FONT, 9, "bold"))
        icon.pack(side="right", padx=(0, px(10)), pady=px(9), anchor="n")

        texts = tk.Frame(row, bg=SURFACE)
        texts.pack(side="right", fill="x", expand=True, padx=px(10), pady=px(7))
        name = tk.Label(texts, text=os.path.basename(path), bg=SURFACE, fg=TEXT,
                        font=(FONT, 9, "bold"), anchor="e")
        name.pack(fill="x")
        detail = tk.Label(texts, text=os.path.dirname(path), bg=SURFACE, fg=MUTED,
                          font=(FONT, 8), anchor="e", justify="right")
        detail.pack(fill="x")

        remove = tk.Label(row, text="✕", bg=SURFACE, fg=FAINT, font=(FONT, 10), cursor="hand2")
        remove.pack(side="left", padx=(px(10), px(2)))
        remove.bind("<Button-1>", lambda e: self.remove(path))
        remove.bind("<Enter>", lambda e: remove.config(fg="#f87171"))
        remove.bind("<Leave>", lambda e: remove.config(fg=FAINT))
        badge = tk.Label(row, font=(FONT, 8, "bold"), padx=px(8), pady=px(1))
        badge.pack(side="left", padx=px(4))

        for w in (row, texts, name, detail, icon):     # לחיצה כפולה פותחת את הקובץ המסודר
            w.bind("<Double-Button-1>", lambda e: self._open_result(path))
        texts.bind("<Configure>", lambda e: detail.config(wraplength=max(e.width - 4, px(100))))
        self.rows[path] = {"frame": row, "badge": badge, "detail": detail, "name": name}
        self._set_status(path, "wait")

    def _open_result(self, path):
        res = self.results.get(path)
        if res and res.output:
            open_path(res.output)

    def _set_status(self, path, status, detail=None):
        text, bg, fg = BADGES[status]
        r = self.rows.get(path)
        if not r:
            return
        r["badge"].config(text=text, bg=bg, fg=fg)
        if detail is not None:
            r["detail"].config(text=detail, fg=fg if status in ("err", "warn") else MUTED)
        r["name"].config(cursor="hand2" if status in ("ok", "warn") else "")

    # --- כפתורים (נבנים לפני הרשימה כדי שיישארו בתחתית)
    def _build_actions(self, parent):
        bar = tk.Frame(parent, bg=BG)
        bar.pack(side="bottom", fill="x", pady=(px(10), 0))
        self.run_btn = RoundButton(bar, "סדר את הקבצים", self.run, width=150, height=38)
        self.run_btn.pack(side="right")
        RoundButton(bar, "פתח תיקייה", self.open_folder, width=104, height=38,
                    style="ghost", font_size=9).pack(side="right", padx=px(8))
        tk.Label(bar, text="הקובץ המסודר נשמר ליד המקור  ·  לחיצה כפולה פותחת אותו", bg=BG, fg=FAINT,
                 font=(FONT, 8)).pack(side="left")

    # --- פעולות
    def _on_drop(self, event):
        self._set_drop_hover(False)
        self.add_files(self.tk.splitlist(event.data))
        return "copy"

    def pick_files(self):
        if self.busy:
            return
        start = os.path.dirname(self.files[-1]) if self.files else APP_DIR
        paths = filedialog.askopenfilenames(title="בחר קבצי אקסל לסידור", initialdir=start,
                                            filetypes=[("Excel", "*.xlsx *.xlsm")])
        self.add_files(paths)

    @staticmethod
    def _expand(paths):
        """תיקייה שנגררה = כל קבצי האקסל שבה (לא כולל תתי-תיקיות ופלטים קודמים)."""
        for p in paths:
            p = os.path.normpath(p)
            if os.path.isdir(p):
                try:
                    names = sorted(os.listdir(p))
                except OSError:
                    continue
                for n in names:
                    if not os.path.splitext(n)[0].endswith(core.OUTPUT_SUFFIX):
                        yield os.path.join(p, n)
            else:
                yield p

    def add_files(self, paths):
        if self.busy:
            return
        for p in self._expand(paths):
            name = os.path.basename(p)
            if (not p.lower().endswith(core.ALLOWED_EXTENSIONS) or name.startswith("~$")
                    or p in self.files or not os.path.isfile(p)):
                continue
            self.files.append(p)
            self._add_row(p)
        self._refresh()

    def remove(self, path):
        if self.busy:
            return
        self.files.remove(path)
        self.results.pop(path, None)
        self.rows.pop(path)["frame"].destroy()
        self._refresh()

    def clear(self):
        for p in list(self.files):
            self.remove(p)
        self._draw_progress(0, 0)
        self.summary_lbl.config(text="")

    def _refresh(self):
        self.count_lbl.config(text=str(len(self.files)))
        if self.files:
            self.empty_lbl.pack_forget()
        else:
            self.empty_lbl.pack(pady=px(18))
        self.run_btn.set_enabled(bool(self.files) and not self.busy)

    def open_folder(self):
        open_path(os.path.dirname(self.files[-1]) if self.files else APP_DIR)

    # --- עיבוד ברקע: הממשק לא נתקע גם בקבצים גדולים
    def run(self):
        if self.busy or not self.files:
            return
        self.busy = True
        self.run_btn.set_text("מעבד…")
        self._refresh()
        self.summary_lbl.config(text="")
        files, cfg = list(self.files), self.cfg
        for p in files:
            self._set_status(p, "wait", os.path.dirname(p))
        self._draw_progress(0, len(files))

        def worker():
            for i, path in enumerate(files):
                self.events.put(("start", path, i))
                self.events.put(("done", path, core.process_file(path, cfg)))
            self.events.put(("finish", None, None))

        threading.Thread(target=worker, daemon=True).start()
        self.after(40, self._poll)

    def _poll(self):
        try:
            while True:
                kind, path, data = self.events.get_nowait()
                if kind == "start":
                    self._set_status(path, "run")
                elif kind == "done":
                    self.results[path] = data
                    msg = data.message
                    if data.output:
                        msg += f"\n← {os.path.basename(data.output)}"
                    self._set_status(path, data.status, msg)
                    self._draw_progress(len(self.results), len(self.files))
                else:
                    self._finish()
                    return
        except queue.Empty:
            pass
        self.after(40, self._poll)

    def _finish(self):
        self.busy = False
        self.run_btn.set_text("סדר את הקבצים")
        self._refresh()
        res = [self.results[p] for p in self.files if p in self.results]
        ok = sum(r.status != "err" for r in res)
        warn = sum(r.status == "warn" for r in res)
        failed = len(res) - ok
        parts = [f"{ok} סודרו"] + ([f"{warn} עם הערות"] if warn else []) + ([f"{failed} נכשלו"] if failed else [])
        self.summary_lbl.config(text="  ·  ".join(parts), fg="#c62828" if failed else MUTED)
        if failed:
            messagebox.showwarning("הסתיים", f"סודרו {ok} קבצים, ונכשלו {failed}.\nפרטים מופיעים ברשימה.")
        elif len(res) == 1:
            if messagebox.askyesno("הסתיים", "הקובץ סודר ונשמר בהצלחה.\nלפתוח אותו עכשיו?"):
                open_path(res[0].output)
        elif messagebox.askyesno("הסתיים", f"כל {ok} הקבצים סודרו בהצלחה.\nלפתוח את התיקייה?"):
            self.open_folder()


def selftest():
    """בדיקה עצמית ללא ממשק (לאחר בנייה / לצוות IT): קוד יציאה 0 = תקין."""
    import tempfile
    import openpyxl
    names = [c.name for c in core.default_config().columns]
    with tempfile.TemporaryDirectory() as d:
        src = os.path.join(d, "selftest.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        for c, h in enumerate(reversed(names), start=1):
            ws.cell(row=1, column=c, value=h)
            ws.cell(row=2, column=c, value=h + "_v")
        wb.save(src)
        res = core.process_file(src)
        out = openpyxl.load_workbook(res.output).active if res.output else None
        good = (res.status == "ok" and out is not None
                and [out.cell(row=1, column=c).value for c in range(1, len(names) + 1)] == names
                and out.cell(row=2, column=1).value == names[0] + "_v")
    return 0 if good else 1


def main():
    if "--selftest" in sys.argv[1:]:
        raise SystemExit(selftest())
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)  # טקסט חד במסכים ברזולוציה גבוהה
    except (ImportError, AttributeError, OSError):
        pass
    App().mainloop()


if __name__ == "__main__":
    main()
