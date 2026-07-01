# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for macOS .app (run: python3 -m PyInstaller interview_coding_mac.spec)

from PyInstaller.utils.hooks import collect_all

block_cipher = None

tk_datas, tk_binaries, tk_hidden = collect_all("tkinter")
hiddenimports = ["openpyxl", "_tkinter", *tk_hidden]

a = Analysis(
    ["export_segments_gui.py"],
    pathex=[],
    binaries=tk_binaries,
    datas=tk_datas,
    hiddenimports=hiddenimports,
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
    [],
    exclude_binaries=True,
    name="InterviewCodingToExcel",
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
    name="InterviewCodingToExcel",
)

app = BUNDLE(
    coll,
    name="InterviewCodingToExcel.app",
    icon=None,
    bundle_identifier="com.transcripttocorpus.interviewcoding",
)
