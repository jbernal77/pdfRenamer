# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for PDF Renamer Tool v2,
now embedding version and metadata via version.txt
"""
import sys
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# Collect PyQt5 resources
pyqt5_binaries, pyqt5_datas, pyqt5_hiddenimports = collect_all('PyQt5')

# Analysis: include icon and version resource
a = Analysis(
    ['pdf_renamer_tool_v2.py'],
    pathex=[],
    binaries=pyqt5_binaries,
    datas=[('pdf_renamer_icon.ico', '.'), ('version.txt', '.')]+pyqt5_datas,
    hiddenimports=pyqt5_hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Build the Python archive
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# Build the executable, embedding version info
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='pdf_renamer_tool_v2',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='pdf_renamer_icon.ico',
    version='version.txt'
)

# Collate all into one folder
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='pdf_renamer_tool_v2'
)
