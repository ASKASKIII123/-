#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys

def check_dependencies():
    print("正在检查依赖包...")
    print("=" * 50)
    
    dependencies = [
        ("PyQt5", "PyQt5"),
        ("python-docx", "docx"),
        ("PyMuPDF", "fitz"),
        ("PyPDF2", "PyPDF2"),
        ("reportlab", "reportlab")
    ]
    
    all_ok = True
    
    for package_name, import_name in dependencies:
        try:
            __import__(import_name)
            print(f"✓ {package_name:20s} - 已安装")
        except ImportError:
            print(f"✗ {package_name:20s} - 未安装")
            all_ok = False
    
    print("=" * 50)
    
    if all_ok:
        print("\n所有依赖包已正确安装！")
        print("可以运行 python main.py 启动程序")
        return 0
    else:
        print("\n部分依赖包未安装，请运行以下命令安装：")
        print("python -m pip install -r requirements.txt")
        return 1

if __name__ == "__main__":
    sys.exit(check_dependencies())