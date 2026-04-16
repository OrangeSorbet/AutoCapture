# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all

# Collect everything properly (fixes your crash)
datas = [("icon.ico", ".")]
binaries = []
hiddenimports = []

for pkg in ["reportlab", "pynput", "PIL", "pypdf"]:
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h


a = Analysis(
    ['AutoCapture.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PySide6'],
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
    name='AutoCapture',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    icon='icon.ico',
)