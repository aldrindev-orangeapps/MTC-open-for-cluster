# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['mtc_open_order.py'],
    pathex=[],
    binaries=[],
    datas=[('assets/MTC Template.xlsx', 'assets')],
    hiddenimports=['win32com', 'win32com.client', 'pythoncom', 'pywintypes'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas,
    [],
    name='MTC Open Order',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)
