# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('src', 'src'), ('config/email_accounts.json', 'config'), ('docs/README_OPTIMIZATION.md', 'docs'), ('data', 'data')],
    hiddenimports=[
        'json', 'os', 'sys', 'time', 'tkinter', 'tkinter.messagebox',
        'smtplib', 'email.mime.multipart', 'email.mime.text',
        'email.mime.application', 'email.header', 'logging',
        'tkinter.ttk', 'tkinter.filedialog', 'datetime', '_tkinter',
        're', 'threading', 'queue', 'plyer', 'subprocess', 'uuid', 'multiprocessing'
    ],
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
    [],
    exclude_binaries=True,
    name='Casiow邮件发送系统',
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
    icon=['resources/Casiow.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Casiow邮件发送系统',
)
