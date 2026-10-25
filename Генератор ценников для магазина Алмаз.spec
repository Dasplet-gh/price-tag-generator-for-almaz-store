# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

project_directory = Path(SPECPATH).resolve()

a = Analysis(
    [str(project_directory / 'price-tag-generator-for-almaz-store.py')],
    pathex=[str(project_directory)],
    binaries=[],
    datas=[
        (str(project_directory / 'almaz.ico'), '.'),
        (str(project_directory / 'arial_bold.ttf'), '.'),
    ],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
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
    icon=str(project_directory / 'almaz.ico'),
)
