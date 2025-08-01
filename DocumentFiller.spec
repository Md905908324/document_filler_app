# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# Collect specific data files (Excel files, templates, pdftk.exe, and libiconv2.dll)
datas = [
    ('assets/data/DFT All Input Data - DO NOT USE.xlsx', 'assets/data'),
    ('assets/data/DFT RP Change Inputs.xlsx', 'assets/data'),
    ('assets/data/DFT TN Dissolution Inputs.xlsx', 'assets/data'),
    ('assets/templates/Use This For RP Change/*', 'assets/templates/Use This For RP Change'),
    ('assets/templates/Use This For TN Dissolution/*', 'assets/templates/Use This For TN Dissolution'),
    ('tools/pdftk.exe', 'tools'),
    ('tools/libiconv2.dll', 'tools')
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['pandas', 'docxtpl', 'pypdf', 'fdfgen', 'docx2pdf'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='DocumentFiller',
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
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DocumentFiller',
)