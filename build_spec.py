from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# 收集所有需要的数据文件
datas = [
    ('templates', 'templates'),  # 包含模板文件
    ('static', 'static'),       # 包含静态文件
]

# 配置信息
a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'webview',
        'pandas',
        'openpyxl',
        'numpy',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Excel处理工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 如果不需要控制台窗口，设置为False
    icon='static/icon.ico'  # 如果你有图标文件的话
)