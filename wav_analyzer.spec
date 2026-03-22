# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for Wav Analyzer
# Mac: creates Wav Analyzer.app | Windows: creates Wav Analyzer.exe

import sys

block_cipher = None

a = Analysis(
    ['Wav Analyzer.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'numpy', 'scipy', 'scipy.io', 'scipy.io.wavfile',
        'matplotlib', 'matplotlib.pyplot', 'matplotlib.backends.backend_qt5agg',
        'matplotlib.figure', 'PyQt5', 'PyQt5.QtCore', 'PyQt5.QtGui',
        'PyQt5.QtWidgets', 'PyQt5.QtMultimedia',
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
    exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name='Wav Analyzer',
              debug=False, strip=True, upx=True, console=False)
    coll = COLLECT(exe, a.binaries, a.zipfiles, a.datas, strip=True, upx=True, upx_exclude=[], name='Wav Analyzer')
    app = BUNDLE(coll, name='Wav Analyzer.app', icon=None)
else:
    # Windows: onefile creates single .exe
    exe = EXE(pyz, a.scripts, a.binaries, a.zipfiles, a.datas, [],
              name='Wav Analyzer', debug=False, strip=False, upx=True,
              upx_exclude=[], runtime_tmpdir=None, console=False)
