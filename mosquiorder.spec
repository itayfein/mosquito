# -*- mode: python ; coding: utf-8 -*-
# בנייה: build.ps1  (או: pyinstaller --clean --noconfirm mosquiorder.spec)
#
# החלטות אבטחה בבנייה:
#  * onedir ולא onefile – התוכנה לא נפרסת לתיקיית TEMP בכל הפעלה
#    (AppLocker / EDR בבתי חולים חוסמים הרצה מ-TEMP, וזו הסיבה הנפוצה לחסימה).
#  * בלי UPX – קבצים דחוסים ב-UPX מסומנים כחשודים ע"י אנטי-וירוס.
#  * רק ה-DLL של tkdnd ל-Windows x64 – בלי קבצי Linux/macOS/ARM מיותרים.
#  * משאב גרסה (VERSIONINFO) עם שם חברה/מוצר – exe "אנונימי" מסומן כחשוד.
#  * מודולי רשת שהתוכנה לא צריכה מוחרגים מהחבילה.
import os
import tkinterdnd2

tkdnd_dir = os.path.join(os.path.dirname(tkinterdnd2.__file__), "tkdnd", "win-x64")
tkdnd_files = [(os.path.join(tkdnd_dir, f), "tkinterdnd2/tkdnd/win-x64")
               for f in os.listdir(tkdnd_dir) if not f.endswith(".lib")]

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=[],
    datas=[("mosquito.ico", ".")] + tkdnd_files,
    hiddenimports=["defusedxml", "defusedxml.ElementTree"],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        "ftplib", "smtplib", "poplib", "imaplib", "telnetlib", "xmlrpc", "http.server",
        "socketserver", "asyncio", "ssl", "_ssl", "sqlite3", "unittest", "pydoc",
        "lib2to3", "multiprocessing", "numpy", "pandas", "PIL", "pytest",
        # בלי OpenSSL (libcrypto 1.1 שהגיע לסוף חייו), רשת ו-lxml – לא נדרשים לתוכנה
        "_hashlib", "_ssl", "_socket", "select", "lxml", "_bz2", "_lzma", "bz2", "lzma",
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="MosquiOrder",
    debug=False,
    strip=False,
    upx=False,
    console=False,
    icon="mosquito.ico",
    version="version_info.txt",
    uac_admin=False,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="MosquiOrder",
)
