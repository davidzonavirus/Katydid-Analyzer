# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for Data Analyzer
# Mac: creates Data Analyzer.app | Windows: creates Data Analyzer.exe

import sys

block_cipher = None

a = Analysis(
    ['Data Analyzer.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'numpy', 'scipy', 'scipy.io', 'scipy.io.wavfile',
        'matplotlib', 'matplotlib.pyplot', 'matplotlib.backends.backend_qt5agg',
        'matplotlib.figure', 'PyQt5', 'PyQt5.QtCore', 'PyQt5.QtGui',
        'PyQt5.QtWidgets', 'pandas', 'openpyxl',
        'openpyxl.styles', 'openpyxl.utils.dataframe',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

if sys.platform == 'darwin':
    # Mac: onedir creates proper .app bundle
    exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='Data Analyzer',
              debug=False, strip=True, upx=True, console=False)
    coll = COLLECT(exe, a.binaries, a.zipfiles, a.datas, strip=True, upx=True, upx_exclude=[], name='Data Analyzer')
    app = BUNDLE(coll, name='Data Analyzer.app', icon=None)
else:
    # Windows: onefile creates single .exe
    exe = EXE(pyz, a.scripts, a.binaries, a.zipfiles, a.datas, [],
              name='Data Analyzer', debug=False, strip=False, upx=True,
              upx_exclude=[], runtime_tmpdir=None, console=False)
