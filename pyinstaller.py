import PyInstaller.__main__

PyInstaller.__main__.run([
    'ESheetSearchMaster.py',
    '--noconfirm',
    '--onefile',
    '--windowed',  # 或者使用 '--noconsole'
])
