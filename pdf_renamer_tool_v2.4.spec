# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for PDF Renamer Tool v2.4,
onefile build with embedded version, icon, and PyQt5 resources.
"""
import sys
from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

# Collect PyQt5 binaries, datas, and hidden imports
pyqt5_binaries, pyqt5_datas, pyqt5_hiddenimports = collect_all('PyQt5')

# Analysis: include main script, icon, version, and all PyQt5 data
a = Analysis(
    ['pdf_renamer_tool_v2.4.py'],
    pathex=[],
    binaries=pyqt5_binaries,
    datas=[('pdf_renamer_icon.ico', '.'), ('version.txt', '.')] + pyqt5_datas,
    hiddenimports=(
        pyqt5_hiddenimports
        + collect_submodules('azure')
        + collect_submodules('opentelemetry')
    ),
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

# Build the Python bytecode archive
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# Create a onefile executable
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='pdf_renamer_tool_v2.4',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='pdf_renamer_icon.ico',
    version='version.txt'
)
