 
# -*- mode: python ; coding: utf-8 -*-

import os
from PyInstaller.utils.hooks import collect_submodules

# Specify any hidden imports (if needed) for your project.
hidden_imports = collect_submodules('customtkinter')  # for example, if needed

block_cipher = None

a = Analysis(
    ['myapp.py'],  # Replace with the name of your main Python file
    pathex=[os.path.abspath(".")],
    binaries=[],
    datas=[
        # Include extra image files and any other necessary resources:
        ('header.png', '.'), 
        ('footer.png', '.'),
        ('signeture.jpg', '.'),
        ('ss.jpg', '.'),
        ('seal.png', '.'),
    ],
    hiddenimports=hidden_imports,
    hookspath=[],
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
    name='InvoiceSOAGenerator',  # The name of the EXE
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False  # Use False for a GUI application (no console window)
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='InvoiceSOAGenerator'
)
