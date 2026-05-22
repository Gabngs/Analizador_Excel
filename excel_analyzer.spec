# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file para Excel Analyzer
Compila la aplicación en un único ejecutable
"""

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Recopilar archivos de datos necesarios
datas = []
datas += collect_data_files('flask')
datas += collect_data_files('openpyxl')

# Recopilar submódulos
hiddenimports = []
hiddenimports += collect_submodules('flask')
hiddenimports += collect_submodules('pandas')
hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('werkzeug')
hiddenimports += ['engineio.async_drivers.threading']

block_cipher = None

a = Analysis(
    ['excel_analyzer.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExcelAnalyzer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # True para ver output de consola
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None  # Puedes agregar un icono aquí si tienes uno
)
