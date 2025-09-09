# build_exe.py
# 用于将 Word 转 HTML 转换器打包成 EXE 的脚本

import subprocess
import sys
import os
import shutil

def check_pyinstaller():
    """检查 PyInstaller 是否已安装"""
    try:
        import PyInstaller
        print("✓ PyInstaller 已安装")
        return True
    except ImportError:
        print("✗ PyInstaller 未安装")
        return False

def install_pyinstaller():
    """安装 PyInstaller"""
    print("正在安装 PyInstaller...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller 安装成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ PyInstaller 安装失败: {e}")
        return False

def check_dependencies():
    """检查所需依赖是否已安装"""
    required_packages = [
        'python-docx',
        'mammoth', 
        'beautifulsoup4',
        'lxml'
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            if package == 'python-docx':
                import docx
            elif package == 'mammoth':
                import mammoth
            elif package == 'beautifulsoup4':
                from bs4 import BeautifulSoup
            elif package == 'lxml':
                import lxml
            print(f"✓ {package} 已安装")
        except ImportError:
            print(f"✗ {package} 未安装")
            missing_packages.append(package)
    
    if missing_packages:
        print(f"\n需要安装以下依赖包: {', '.join(missing_packages)}")
        print("正在安装...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing_packages)
            print("✓ 依赖包安装成功")
            return True
        except subprocess.CalledProcessError as e:
            print(f"✗ 依赖包安装失败: {e}")
            return False
    
    return True

def create_spec_file():
    """创建 PyInstaller 规格文件"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['word_to_html_converter.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['docx', 'mammoth', 'bs4', 'lxml', 'tkinter', 'tkinter.ttk', 'tkinter.filedialog', 'tkinter.scrolledtext', 'tkinter.messagebox'],
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
    name='Word转HTML转换器v24',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 设为 False 隐藏控制台窗口
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 如果有 .ico 文件可以在这里指定
)
'''
    
    with open('word_converter.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    print("✓ 规格文件创建成功")

def build_exe():
    """使用 PyInstaller 构建 EXE"""
    print("\n开始打包 EXE...")
    try:
        # 使用规格文件构建
        subprocess.check_call([
            'pyinstaller', 
            'word_converter.spec',
            '--clean'
        ])
        print("✓ EXE 打包成功！")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ EXE 打包失败: {e}")
        return False
    except FileNotFoundError:
        print("✗ PyInstaller 命令未找到，请确保已正确安装")
        return False

def cleanup():
    """清理临时文件"""
    print("\n清理临时文件...")
    cleanup_items = ['build', '__pycache__', 'word_converter.spec']
    
    for item in cleanup_items:
        if os.path.exists(item):
            if os.path.isdir(item):
                shutil.rmtree(item)
                print(f"✓ 删除目录: {item}")
            else:
                os.remove(item)
                print(f"✓ 删除文件: {item}")

def main():
    print("=" * 60)
    print("Word 转 HTML 转换器 - EXE 打包工具")
    print("=" * 60)
    
    # 检查主程序文件是否存在
    if not os.path.exists('word_to_html_converter.py'):
        print("✗ 找不到主程序文件 'word_to_html_converter.py'")
        print("请确保该文件存在于当前目录中")
        return False
    
    print("✓ 找到主程序文件")
    
    # 检查并安装 PyInstaller
    if not check_pyinstaller():
        if not install_pyinstaller():
            return False
    
    # 检查依赖包
    if not check_dependencies():
        return False
    
    # 创建规格文件
    create_spec_file()
    
    # 构建 EXE
    if not build_exe():
        return False
    
    # 检查输出文件
    exe_path = os.path.join('dist', 'Word转HTML转换器v24.exe')
    if os.path.exists(exe_path):
        file_size = os.path.getsize(exe_path) / (1024 * 1024)  # MB
        print(f"✓ EXE 文件已生成: {exe_path}")
        print(f"  文件大小: {file_size:.1f} MB")
    else:
        print("✗ 未找到生成的 EXE 文件")
        return False
    
    # 询问是否清理临时文件
    response = input("\n是否删除临时文件? (y/n): ").lower()
    if response in ['y', 'yes', '是']:
        cleanup()
    
    print("\n" + "=" * 60)
    print("打包完成！")
    print(f"EXE 文件位置: {os.path.abspath(exe_path)}")
    print("=" * 60)
    
    return True

if __name__ == "__main__":
    try:
        success = main()
        if not success:
            input("\n按任意键退出...")
            sys.exit(1)
        else:
            input("\n按任意键退出...")
    except KeyboardInterrupt:
        print("\n\n操作被用户取消")
        sys.exit(1)
    except Exception as e:
        print(f"\n发生未预期的错误: {e}")
        input("\n按任意键退出...")
        sys.exit(1)