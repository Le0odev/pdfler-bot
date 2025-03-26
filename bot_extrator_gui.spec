# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['bot_extrator_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\pc\\AppData\\Local\\Programs\\Python\\Python313\\Lib\\site-packages\\fitz', 'fitz'), ('C:\\Users\\pc\\AppData\\Local\\Programs\\Python\\Python313\\Lib\\site-packages\\customtkinter', 'customtkinter'), ('C:\\Users\\pc\\AppData\\Local\\Programs\\Python\\Python313\\Lib\\site-packages\\mysql', 'mysql')],
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name='bot_extrator_gui',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='bot_extrator_gui',
)
