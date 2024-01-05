# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata

hiddenimports = []
datas = []

with open('requirements.txt', 'r') as f:
    for line in f:
        package = line.strip().split('==')[0]
        datas += copy_metadata(package) 
        hiddenimports += [package]

datas += [('ner_bert_base_msra_20211227_114712/', 'ner_bert_base_msra_20211227_114712')]
datas += [('huggingface/', 'huggingface')]

hiddenimports += collect_submodules('hanlp')

a = Analysis(
    ['windows_ver.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='文档脱敏工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=r'C:\Users\echooff\Desktop\pyproject\tuomin\tuomin\a0jzs-2pnyt-001.ico',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='文档脱敏工具',
)
