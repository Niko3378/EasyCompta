# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['pdf_extractor.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'pdfminer',
        'pdfminer.high_level',
        'pdfminer.layout',
        'pdfminer.pdfinterp',
        'pdfminer.converter',
        'pdfminer.pdfpage',
        'pdfminer.pdfdocument',
        'pdfminer.pdfparser',
        'pdfminer.pdftypes',
        'charset_normalizer',
    ],
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
    name='pdf_extractor',
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
