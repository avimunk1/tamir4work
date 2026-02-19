# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for Excel Extractor

import sys

# Add src to path so excel_extractor can be found
sys.path.insert(0, 'src')

a = Analysis(
    ['src/excel_extractor/main.py'],
    pathex=['src'],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.cell._writer',
        'openpyxl.styles',
        'openpyxl.worksheet',
        'openpyxl.utils',
        'numpy',
        'et_xmlfile',
    ],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ExcelExtractor',
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
