# main.spec

# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

import os
from PyInstaller.utils.hooks import collect_data_files

# Paths
tess_path = os.path.join('resource', 'tess')
datas = [
    ('resource/wamy.ico', 'resource'),  # Icon
    (os.path.join(tess_path, 'tesseract.exe'), os.path.join('resource', 'tess')),
]

# Add tessdata files dynamically
tessdata_dir = os.path.join(tess_path, 'tessdata')
if os.path.isdir(tessdata_dir):
    for file in os.listdir(tessdata_dir):
        datas.append((os.path.join(tessdata_dir, file), os.path.join('resource', 'tess', 'tessdata')))

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='OCR_Renamer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='resource/wamy.ico'
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='OCR_Renamer'
)
