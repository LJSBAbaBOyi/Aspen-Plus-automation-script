# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src\\ui_app.py'],
    pathex=['src'],
    binaries=[],
    datas=[('src', 'src'), ('ico', 'ico')],
    hiddenimports=['win32com', 'pythoncom', 'openpyxl', 'aspen_interface'],
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
    name='Aspen自动化调参工具',
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
    codesign_identity=None,
    entitlements_file=None,
    icon='ico\\128x128.ico',
    version='version_info.txt',
)
