import PyInstaller.__main__

PyInstaller.__main__.run([
    'main.py',
    '--noconfirm',
    '--onefile',
    '--windowed'  # 或者使用 '--noconsole'
])