# -*- mode: python ; coding: utf-8 -*-
import os

block_cipher = None

# Make sure to include the sample CRM data file
sample_data = [('Kapil-Dutta-Open-Deals.xlsx', '.')]

a = Analysis(
    ['time-tracker.py'],
    pathex=[],
    binaries=[],
    datas=sample_data,
    hiddenimports=['pandas', 'openpyxl'],
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

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='TimeTracker',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,
)

# macOS app bundle
app = BUNDLE(
    exe,
    name='TimeTracker.app',
    icon='icon.icns' if os.path.exists('icon.icns') else None,
    bundle_identifier='com.timetracker.app',
    info_plist={
        'NSHighResolutionCapable': 'True',
        'CFBundleShortVersionString': '1.0.0',
    },
) 