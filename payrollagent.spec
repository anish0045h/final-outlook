# -*- mode: python ; coding: utf-8 -*-
#
# PayrollAgent.spec — PyInstaller build specification
#
# Security hardening applied:
#   - upx=False  → prevents UPX-packer false positives in AV/EDR engines
#   - icon set   → reduces SmartScreen suspicion score for generic PyInstaller icon
#   - noarchive=False retained (default) — no change to bytecode storage
#   - codesign_identity left as None — EXE must be signed post-build via
#     signtool.exe before enterprise distribution (see build_exe.bat notes)
#   - console=True retained — agent is an operator-facing CLI tool
#   - No runtime_tmpdir set — PyInstaller default (_MEIxxxxxx) is fine;
#     setting a fixed tmpdir can create predictable extraction paths

block_cipher = None

a = Analysis(
    ["agent.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        # Project modules
        "config",
        "outlook_reader",
        "excel_parser",
        "master_sheet",
        # openpyxl
        "openpyxl",
        "openpyxl.cell",
        "openpyxl.styles",
        "openpyxl.utils",
        "openpyxl.workbook",
        "openpyxl.worksheet",
        # pywin32 / COM  — required for Outlook automation
        "win32com",
        "win32com.client",
        "win32com.server",
        "win32api",
        "win32timezone",
        "pythoncom",
        "pywintypes",
        # stdlib
        "sqlite3",
        "email",
        "io",
        "logging",
        "logging.handlers",   # ADDED — RotatingFileHandler lives here
        "datetime",
        "argparse",           # ADDED — replaces input() prompts
        "pathlib",            # ADDED — used for safe filename handling
        "tempfile",           # ADDED — used for secure temp file creation
        "os",
        "sys",
        "re",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name="PayrollAgent",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,              # CHANGED: was True — UPX triggers AV/EDR false positives
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None, # Sign post-build with signtool.exe + EV cert
    entitlements_file=None,
    icon=None,
)
