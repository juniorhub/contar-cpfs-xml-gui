# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['contar_cpfs_xml_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['pandas', 'numpy', 'tkinter'],
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
    name='contar_cpfs_xml_gui',
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
    name='contar_cpfs_xml_gui',
)
app = BUNDLE(
    coll,
    name='contar_cpfs_xml_gui.app',
    icon=None,
    bundle_identifier=None,
)
