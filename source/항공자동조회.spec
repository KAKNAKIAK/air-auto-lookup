# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['topas_live_collector', 'update_client', 'fare_store']
hiddenimports += collect_submodules('selenium')


a = Analysis(
    ['air_auto_lookup_mvp.py'],
    pathex=[],
    binaries=[],
    datas=[('flight-master.mjs', '.'), ('fares_snapshot.seed.json', '.'), ('assets', 'assets')],
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
    name='항공자동조회',
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
    icon=['assets\\air_auto_lookup_icon.ico'],
)
