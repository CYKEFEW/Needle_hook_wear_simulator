# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

datas = []
datas += collect_data_files("matplotlib")
datas += [("defaultData.ini", ".")]
datas += [("app.ico", ".")]

hiddenimports = []
hiddenimports += collect_submodules("matplotlib")

a = Analysis(
    ["needle_hook_wear_sim_gui_app.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name="needle_hook_wear_sim",
    console=False,
    icon="app.ico",
    onefile=True,
    strip=False,
    upx=True,
)
