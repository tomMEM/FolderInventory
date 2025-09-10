# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

a = Analysis(
    ['fileinventory_cgp.py'],
    pathex=[],
    binaries=[],
    datas=[
        *collect_data_files('gradio', include_py_files=True),
        *collect_data_files('gradio_client'),
        *collect_data_files('safehttpx'),
        *collect_data_files('groovy'),
        *collect_data_files('markdown2'),
        *collect_data_files('ffmpy'),
        *collect_data_files('pandas'),
        *collect_data_files('openpyxl')
    ],
    hiddenimports=[
        'gradio',
        'markdown2',
        'pandas',
        'openpyxl',
        'ffmpy',
        'docx',
        'webbrowser',
    ] + collect_submodules('gradio'),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='InventoryDashboard',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None
)