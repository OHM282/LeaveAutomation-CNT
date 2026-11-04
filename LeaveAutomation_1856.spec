# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
import glob, os

datas = [('clock.ico', '.'), ('cnt_logo.png', '.')]
binaries = []
hiddenimports = ['windnd']
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# รวบรวม Visual C++ Runtime และ UCRT DLLs ทั้งหมด เพื่อป้องกันปัญหาบนเครื่องลูกข่ายที่ไม่มี VC++ Redistributable
system32 = os.path.join(os.environ.get('SystemRoot', r'C:\Windows'), 'System32')
for pat in ['vcruntime140*.dll', 'msvcp140*.dll', 'vcomp140*.dll', 'concrt140*.dll', 'ucrtbase.dll']:
    for f in glob.glob(os.path.join(system32, pat)):
        binaries.append((f, '.'))

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
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
    name='LeaveAutomation_1856',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['clock.ico'],
)
