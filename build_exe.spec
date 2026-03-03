# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file untuk aplikasi Formulir Ritual.
Bundel semua dependensi + assets dalam satu folder.
"""

import os
import sys
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# Path project
PROJECT_DIR = os.path.abspath('.')

# Collect customtkinter data files (themes, assets)
ctk_datas = collect_data_files('customtkinter')

a = Analysis(
    ['main.py'],
    pathex=[PROJECT_DIR],
    binaries=[],
    datas=[
        # Assets aplikasi
        (os.path.join('assets', 'label.pdf'), 'assets'),
        (os.path.join('assets', 'fonts', 'HanyiSentyPagoda.ttf'), os.path.join('assets', 'fonts')),
        # Modul internal
        (os.path.join('database', '__init__.py'), 'database'),
        (os.path.join('database', 'database.py'), 'database'),
        (os.path.join('modules', '__init__.py'), 'modules'),
        (os.path.join('modules', 'calibration_preview.py'), 'modules'),
        (os.path.join('modules', 'excel_template.py'), 'modules'),
        (os.path.join('modules', 'pdf_engine.py'), 'modules'),
        (os.path.join('modules', 'pdf_preview.py'), 'modules'),
        (os.path.join('modules', 'updater.py'), 'modules'),
    ] + ctk_datas,
    hiddenimports=[
        'customtkinter',
        'packaging',
        'packaging.version',
        'reportlab',
        'reportlab.lib.pagesizes',
        'reportlab.pdfgen.canvas',
        'reportlab.pdfbase.ttfonts',
        'reportlab.pdfbase.pdfmetrics',
        'reportlab.lib.units',
        'PyPDF2',
        'fitz',
        'pypinyin',
        'openpyxl',
        'sqlite3',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
    ],
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
    [],
    exclude_binaries=True,
    name='FormulirRitual',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # No console window (GUI app)
    disable_windowed_traceback=False,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='FormulirRitual',
)
