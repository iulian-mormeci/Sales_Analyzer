# -*- mode: python ; coding: utf-8 -*-
# PyInstaller build spec for Sales Analyzer
# Build: pyinstaller build.spec

import sys
from pathlib import Path

block_cipher = None

# Collect all data files that need to be bundled
added_files = [
    ("VERSION", "."),
]

# Try to include icon if it exists
icon_path = Path("icon.ico")
icon_arg = str(icon_path) if icon_path.exists() else None

a = Analysis(
    ["main.py"],
    pathex=["."],
    binaries=[],
    datas=added_files,
    hiddenimports=[
        # Matplotlib backends
        "matplotlib.backends.backend_tkagg",
        "matplotlib.backends.backend_agg",
        "matplotlib.backends._backend_tk",
        "matplotlib.figure",
        "matplotlib.pyplot",
        # Pandas/numpy internals
        "pandas",
        "pandas._libs.tslibs.base",
        "pandas._libs.tslibs.nattype",
        "pandas._libs.tslibs.timedeltas",
        "pandas._libs.tslibs.timestamps",
        "pandas._libs.tslibs.np_datetime",
        "pandas._libs.tslibs.offsets",
        "pandas._libs.tslibs.strptime",
        "pandas._libs.tslibs.period",
        "pandas._libs.tslibs.parsing",
        "pandas._libs.hashtable",
        "pandas._libs.index",
        "pandas._libs.lib",
        "numpy",
        "numpy.core",
        "numpy.core._multiarray_umath",
        # Excel
        "openpyxl",
        "openpyxl.workbook",
        "openpyxl.worksheet",
        # Networking (updater)
        "requests",
        "requests.adapters",
        "requests.auth",
        "requests.packages",
        "urllib3",
        "certifi",
        "charset_normalizer",
        "idna",
        # Packaging
        "packaging",
        "packaging.version",
        # Optional calendar
        "tkcalendar",
        # tkinter
        "tkinter",
        "tkinter.ttk",
        "tkinter.filedialog",
        "tkinter.messagebox",
        # Pillow (optional, for image handling)
        "PIL",
        "PIL.Image",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "PyQt5", "PyQt6", "wx", "gi",
        "IPython", "notebook", "jupyter",
        "scipy", "sklearn", "tensorflow",
        "cv2", "pygame",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="SalesAnalyzer",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_arg,
)
