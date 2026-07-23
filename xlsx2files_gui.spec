# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.utils.hooks import collect_all

project = Path(SPECPATH)
browser_dir = project / ".playwright-browsers"

if not browser_dir.exists():
    raise SystemExit(
        "缺少 .playwright-browsers。请先设置 PLAYWRIGHT_BROWSERS_PATH 后执行："
        "python -m playwright install --only-shell chromium"
    )

playwright_datas, playwright_binaries, playwright_hiddenimports = collect_all("playwright")
datas = playwright_datas + [
    (str(browser_dir), ".playwright-browsers"),
    (str(project / "assets"), "assets"),
]
binaries = playwright_binaries
hiddenimports = playwright_hiddenimports + [
    "playwright.sync_api",
    "gdown",
    "py7zr",
    "PIL._tkinter_finder",
]

a = Analysis(
    [str(project / "xlsx2files_gui.py")],
    pathex=[str(project)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name="Excel订单素材下载工具",
    icon=str(project / "assets" / "lingtu-logo.ico"),
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version=str(project / "version_info.txt"),
)
