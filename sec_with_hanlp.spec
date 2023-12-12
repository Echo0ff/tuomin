# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata

block_cipher = None

datas = []
datas += collect_data_files('hanlp')
datas += copy_metadata('tqdm')
datas += copy_metadata('regex')
datas += copy_metadata('requests')
datas += copy_metadata('packaging')
datas += copy_metadata('filelock')
datas += copy_metadata('numpy')
datas += copy_metadata('tokenizers')
datas += copy_metadata('tensorflow')
datas += copy_metadata('huggingface-hub')
datas += copy_metadata('safetensors')
datas += copy_metadata('pyyaml')

a = Analysis(
    ['sec_with_hanlp.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=collect_submodules('hanlp')+['tqdm']+['huggingface-hub']+['safetensors']+['pyyaml'],
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
    name='with_hanlp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='with_hanlp',
)
app = BUNDLE(
    coll,
    name='desensitization_hanlp.app',
    icon=None,
    bundle_identifier=None,
)
