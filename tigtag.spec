# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for TigTag.
#
# Build with:
#   pyinstaller tigtag.spec
#
# Produces a single-file windowed executable at dist\TigTag.exe with
# tigtag.png bundled inside it (extracted to a temp folder at runtime,
# handled automatically by the sys._MEIPASS logic in tigtag.py).

block_cipher = None

a = Analysis(
    ['tigtag.py'],
    pathex=[],
    binaries=[],
    datas=[('tigtag.png', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='TigTag',
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
    icon='tigtag.ico',
)
