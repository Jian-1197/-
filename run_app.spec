# -*- mode: python ; coding: utf-8 -*-


from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import copy_metadata
 
datas = [(r"e:\miniconda\envs\office\Lib\site-packages\streamlit\runtime","./streamlit/runtime")]
datas += collect_data_files("streamlit")
datas += copy_metadata("streamlit")


block_cipher = None

a = Analysis(
    ['run_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['openpyxl.utils.dataframe','openpyxl.styles','openpyxl.utils','docx','docx.shared','docx.enum.text',
    'docx.oxml.ns','docx2pdf','tqdm','tqdm.auto'],
    hookspath=['./hooks'],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='上课啦表格制作',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    icon='12.ico',
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
