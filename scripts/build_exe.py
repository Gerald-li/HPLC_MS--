# -*- coding: utf-8 -*-
import PyInstaller.__main__
import os
import shutil

def build_executable():
    # 清理之前的构建文件
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    
    # 使用更简单的PyInstaller配置
    PyInstaller.__main__.run([
        'extract_excel_gui.py',
        '--name=Excel数据提取工具',
        '--onefile',
        '--windowed',
        '--clean',
        '--noconfirm'
    ])

if __name__ == '__main__':
    build_executable()