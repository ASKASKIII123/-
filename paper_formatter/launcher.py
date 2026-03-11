#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
import subprocess

def print_header():
    print("=" * 50)
    print("   论文格式自动修改软件 - 启动器")
    print("=" * 50)
    print()

def check_python_version():
    print("[1/4] 检查Python版本...")
    version = sys.version_info
    print(f"Python版本: {version.major}.{version.minor}.{version.micro}")
    
    if version.major < 3 or (version.major == 3 and version.minor < 7):
        print("[错误] Python版本过低！需要Python 3.7或更高版本")
        return False
    
    print("[成功] Python版本符合要求")
    print()
    return True

def check_dependencies():
    print("[2/4] 检查依赖包...")
    
    required_packages = [
        ("PyQt5", "PyQt5"),
        ("python-docx", "docx"),
        ("PyMuPDF", "fitz"),
        ("PyPDF2", "PyPDF2"),
        ("reportlab", "reportlab")
    ]
    
    missing_packages = []
    
    for package_name, import_name in required_packages:
        try:
            __import__(import_name)
            print(f"  ✓ {package_name}")
        except ImportError:
            print(f"  ✗ {package_name} (未安装)")
            missing_packages.append(package_name)
    
    print()
    
    if missing_packages:
        print(f"[提示] 发现 {len(missing_packages)} 个缺失的依赖包")
        print("正在安装依赖包（使用清华镜像源加速）...")
        print()
        
        try:
            subprocess.check_call([
                sys.executable, "-m", "pip", "install", "-r", "requirements.txt",
                "-i", "https://pypi.tuna.tsinghua.edu.cn/simple",
                "--trusted-host", "pypi.tuna.tsinghua.edu.cn"
            ])
            print()
            print("[成功] 依赖包安装完成")
        except subprocess.CalledProcessError:
            print()
            print("[错误] 依赖包安装失败")
            print("请尝试手动运行: python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple")
            return False
    
    print("[成功] 所有依赖包已就绪")
    print()
    return True

def check_main_file():
    print("[3/4] 检查程序文件...")
    
    if not os.path.exists("main.py"):
        print("[错误] 未找到 main.py 文件")
        return False
    
    print("[成功] 程序文件检查通过")
    print()
    return True

def launch_program():
    print("[4/4] 启动程序...")
    print()
    print("=" * 50)
    print()
    
    try:
        subprocess.run([sys.executable, "main.py"], check=True)
        return True
    except subprocess.CalledProcessError as e:
        print()
        print(f"[错误] 程序运行失败，退出码: {e.returncode}")
        return False
    except KeyboardInterrupt:
        print()
        print("[提示] 程序被用户中断")
        return True
    except Exception as e:
        print()
        print(f"[错误] 发生未知错误: {str(e)}")
        return False

def main():
    print_header()
    
    if not check_python_version():
        print("\n请安装Python 3.7或更高版本: https://www.python.org/downloads/")
        input("\n按回车键退出...")
        return 1
    
    if not check_dependencies():
        input("\n按回车键退出...")
        return 1
    
    if not check_main_file():
        input("\n按回车键退出...")
        return 1
    
    if not launch_program():
        input("\n按回车键退出...")
        return 1
    
    print()
    print("=" * 50)
    print("程序已正常关闭")
    print("=" * 50)
    input("\n按回车键退出...")
    return 0

if __name__ == "__main__":
    sys.exit(main())