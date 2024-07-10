# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['https://github.com/Jars18/Sistema_Lector_de_Facturas/blob/54581a6acd8d780e78dc10a1b56ef7a5313ba15f/Sistema_Lector_de_Facturas_Windows.py'],
    pathex=[],
    binaries=[],
    datas=[('/Users/jurgenalejandrorocasalvosanchez/Documents/Programa_PDG/Lector_de_Facturas/lib/python3.11/site-packages/ultralytics', 'ultralytics/'), ('https://github.com/Jars18/Sistema_Lector_de_Facturas/blob/2ca8c9c3fbc34fa4b4b443627ea97acb31b279d8/background.png', '.'), ('https://github.com/Jars18/Sistema_Lector_de_Facturas/blob/4d07e517878959f4cc6bb41c08d22cab2075bf5e/best_ultimo.pt', '.')],
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
    a.binaries,
    a.datas,
    [],
    name='Sistema_Lector_de_Facturas_Windows',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
