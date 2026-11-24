# -*- mode: python ; coding: utf-8 -*-
# גרסה ניידת – קובץ exe יחיד, למחשבים רגילים שלא דורשים הלבנה.
# בנייה: build.ps1  (או: pyinstaller --clean --noconfirm mosquiorder-portable.spec)
#
# שימו לב: onefile נפרס ל-TEMP בכל הפעלה, ולכן הוא נחסם במערכות סגורות (AppLocker/EDR).
# לבתי חולים / הלבנה – להשתמש ב-mosquiorder.spec (גרסת התיקייה).
# שאר החלטות האבטחה (בלי UPX, בלי רשת/OpenSSL, VERSIONINFO) זהות לגרסת התיקייה.
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
    a.binaries,
    a.datas,
    [],
    name="MosquiOrder-portable",
    debug=False,
    strip=False,
    upx=False,
    console=False,
    icon="mosquito.ico",
    version="version_info.txt",
    uac_admin=False,
    runtime_tmpdir=None,
)
