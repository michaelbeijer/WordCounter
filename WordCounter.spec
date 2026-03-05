# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['WordCounter.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('bundle/jre', 'jre'),
        ('bundle/tika-server-standard-3.1.0.jar', 'tika'),
        ('bundle/tika-server-standard-3.1.0.jar.md5', 'tika'),
    ],
    hiddenimports=['tika', 'requests', 'urllib3', 'certifi', 'idna', 'charset_normalizer'],
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
    name='WordCounter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='windows_version_info.txt',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='WordCounter',
)
