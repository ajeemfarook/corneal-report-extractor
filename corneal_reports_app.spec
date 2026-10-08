# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for Corneal Reports PDF Extractor.
# Run from this folder:
#     pyinstaller --noconfirm corneal_reports_app.spec
# Resulting executable:  dist\CornealReportsApp.exe

from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

a = Analysis(
    ['corneal_reports_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'corneal_reports_core',
        'pymupdf',
        'pandas',
        'openpyxl',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Trim common scientific stacks we don't use to keep the .exe smaller.
        'matplotlib',
        'numpy.tests',
        'scipy',
        'PIL',
        'PyQt5',
        'PySide2',
        'IPython',
    ],
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
    name='CornealReportsApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,            # GUI app, no extra console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,                # Drop an .ico path here if you have one
)
