# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_submodules
from PyInstaller.building.api import *
from PyInstaller.building.build_main import *
from PyInstaller.archive.pyz_crypto import PyiBlockCipher

# Generate a secure key - this should be kept secret
import secrets
key = secrets.token_hex(16)  # 32 character hex string
print(f"Generated key: {key}")  # Save this key somewhere secure!

block_cipher = PyiBlockCipher(key=key)

a = Analysis(
    ['fileinventory.py'],
    pathex=[],
    binaries=[],
    datas=[
        *collect_data_files('gradio', include_py_files=True),
        *collect_data_files('gradio_client'),
        *collect_data_files('safehttpx'),
        *collect_data_files('groovy')
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
    cipher=block_cipher,  # Add cipher here
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(
    a.pure, 
    a.zipped_data,
    cipher=block_cipher  # Add cipher here too
)

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
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='app.ico'  # Optional: Add this if you have an icon file
)