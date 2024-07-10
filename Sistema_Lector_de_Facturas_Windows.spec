# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['Sistema_Lector_de_Facturas.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('background.png', '.'),  # Incluye el archivo de imagen
        ('best_ultimo.pt', '.')   # Incluye el archivo de modelo
    ],
    hiddenimports=[
        'fitz',
        'ultralytics',
        'transformers',
        'PIL',
        'easyocr',
        'cv2',
        'numpy',
        'openpyxl',
        'torch'
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Sistema_Lector_de_Facturas_Windows',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Sistema_Lector_de_Facturas_Windows'
)
