# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(['price-tag-generator-for-almaz-store.py'],
             pathex=['C:\\Users\\ROG STRIX\\Projects\\Python\\price-tag-generator-for-almaz-store'],
             binaries=[
                ('C:/Users/ROG STRIX/miniconda3/envs/price-tag-generator-for-almaz-store/Library/bin/mkl_intel_thread.1.dll', '.'),
                ('C:/Users/ROG STRIX/miniconda3/envs/price-tag-generator-for-almaz-store/Library/bin/mkl_core.1.dll', '.'),
                ('C:/Users/ROG STRIX/miniconda3/envs/price-tag-generator-for-almaz-store/Library/bin/mkl_def.1.dll', '.'),
                ('C:/Users/ROG STRIX/miniconda3/envs/price-tag-generator-for-almaz-store/Library/bin/libiomp5md.dll', '.')
             ],
             datas=[
                ('almaz.ico', '.'),
                ('arial_bold.ttf', '.')
             ],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='Генератор ценников для магазина Алмаз',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          icon='almaz.ico')
