# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['/Users/jurgenalejandrorocasalvosanchez/Documents/Programa_PDG/Lector_de_Facturas/Sistema_Lector_de_Facturas_Windows.py'],
    pathex=[],
    binaries=[],
    datas=[('/Users/jurgenalejandrorocasalvosanchez/Documents/Programa_PDG/Lector_de_Facturas/lib/python3.11/site-packages/ultralytics', 'ultralytics/'), ('/Users/jurgenalejandrorocasalvosanchez/Documents/Programa_PDG/Lector_de_Facturas/background.png', '.'), ('/Users/jurgenalejandrorocasalvosanchez/Documents/Programa_PDG/Lector_de_Facturas/best_ultimo.pt', '.')],
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
