# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('config.py', '.'), ('research_core.py', '.'), ('research_gui.py', '.'), ('test', 'test')],
    hiddenimports=['pandas', 'openpyxl', 'tkinter', 'numpy', 'colorlog', 'psutil', 'pillow', 'xlsxwriter', 'xlwings', 'xlwings.server', 'xlwings._xlwindows', 'win32com', 'pythoncom', 'pywintypes'],
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
    name='연구비처리도우미_포터블',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
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
    name='연구비처리도우미_포터블',
)
