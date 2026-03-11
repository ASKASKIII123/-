#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import subprocess
import os

def print_header():
    print("=" * 60)
    print("   论文格式软件 - 依赖包详细安装工具")
    print("=" * 60)
    print()

def check_python():
    print("[1/6] 检查Python环境...")
    print(f"Python路径: {sys.executable}")
    print(f"Python版本: {sys.version}")
    print()
    return True

def check_pip():
    print("[2/6] 检查pip...")
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "--version"],
            capture_output=True,
            text=True,
            timeout=30
        )
        if result.returncode == 0:
            print(f"pip版本: {result.stdout.strip()}")
            print()
            return True
        else:
            print(f"pip错误: {result.stderr}")
            print()
            return False
    except Exception as e:
        print(f"pip检查失败: {str(e)}")
        print()
        return False

def upgrade_pip():
    print("[3/6] 升级pip（使用清华镜像源）...")
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "--upgrade", "pip",
             "-i", "https://pypi.tuna.tsinghua.edu.cn/simple",
             "--trusted-host", "pypi.tuna.tsinghua.edu.cn"],
            capture_output=True,
            text=True,
            timeout=300
        )
        
        if result.returncode == 0:
            print("✓ pip升级成功")
            print()
            return True
        else:
            print("✗ pip升级失败")
            print(f"错误信息: {result.stderr}")
            print()
            return False
    except subprocess.TimeoutExpired:
        print("✗ pip升级超时")
        print()
        return False
    except Exception as e:
        print(f"✗ pip升级异常: {str(e)}")
        print()
        return False

def install_package(package_name, import_name=None):
    if import_name is None:
        import_name = package_name.lower().replace("-", "_")
    
    print(f"正在安装 {package_name}...")
    
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", package_name,
             "-i", "https://pypi.tuna.tsinghua.edu.cn/simple",
             "--trusted-host", "pypi.tuna.tsinghua.edu.cn"],
            capture_output=True,
            text=True,
            timeout=300
        )
        
        if result.returncode == 0:
            print(f"✓ {package_name} 安装成功")
            
            # 验证安装
            try:
                __import__(import_name)
                print(f"✓ {package_name} 导入测试通过")
            except ImportError as e:
                print(f"✗ {package_name} 导入失败: {str(e)}")
                return False
            
            print()
            return True
        else:
            print(f"✗ {package_name} 安装失败")
            print(f"错误信息: {result.stderr}")
            print()
            return False
            
    except subprocess.TimeoutExpired:
        print(f"✗ {package_name} 安装超时")
        print()
        return False
    except Exception as e:
        print(f"✗ {package_name} 安装异常: {str(e)}")
        print()
        return False

def install_all_packages():
    print("[4/6] 安装所有依赖包...")
    print()
    
    packages = [
        ("PyQt5", "PyQt5"),
        ("python-docx", "docx"),
        ("PyMuPDF", "fitz"),
        ("PyPDF2", "PyPDF2"),
        ("reportlab", "reportlab")
    ]
    
    success_count = 0
    for package_name, import_name in packages:
        if install_package(package_name, import_name):
            success_count += 1
    
    print(f"[4/6] 安装完成: {success_count}/{len(packages)} 个包成功")
    print()
    return success_count == len(packages)

def verify_installation():
    print("[5/6] 验证所有依赖包...")
    print()
    
    packages = [
        ("PyQt5", "PyQt5"),
        ("python-docx", "docx"),
        ("PyMuPDF", "fitz"),
        ("PyPDF2", "PyPDF2"),
        ("reportlab", "reportlab")
    ]
    
    all_ok = True
    for package_name, import_name in packages:
        try:
            __import__(import_name)
            print(f"✓ {package_name}")
        except ImportError:
            print(f"✗ {package_name} (未安装)")
            all_ok = False
    
    print()
    return all_ok

def show_summary():
    print("[6/6] 安装总结...")
    print()
    
    packages = [
        ("PyQt5", "PyQt5"),
        ("python-docx", "docx"),
        ("PyMuPDF", "fitz"),
        ("PyPDF2", "PyPDF2"),
        ("reportlab", "reportlab")
    ]
    
    installed = []
    missing = []
    
    for package_name, import_name in packages:
        try:
            __import__(import_name)
            installed.append(package_name)
        except ImportError:
            missing.append(package_name)
    
    print(f"已安装的包 ({len(installed)}):")
    for pkg in installed:
        print(f"  ✓ {pkg}")
    
    if missing:
        print(f"\n未安装的包 ({len(missing)}):")
        for pkg in missing:
            print(f"  ✗ {pkg}")
    
    print()
    return len(missing) == 0

def main():
    print_header()
    
    if not check_python():
        input("\n按回车键退出...")
        return 1
    
    if not check_pip():
        input("\n按回车键退出...")
        return 1
    
    upgrade_pip()
    
    if not install_all_packages():
        print("\n[警告] 部分依赖包安装失败")
        print("建议:")
        print("1. 检查网络连接")
        print("2. 尝试使用管理员权限运行")
        print("3. 手动安装失败的包")
    
    if not verify_installation():
        print("\n[警告] 部分依赖包验证失败")
    
    if not show_summary():
        print("\n[提示] 安装未完全成功，但可以尝试启动程序")
    
    print("=" * 60)
    print("安装检查完成！")
    print("=" * 60)
    print()
    print("下一步:")
    print("1. 如果所有包都安装成功，可以运行: python main.py")
    print("2. 如果有包安装失败，请查看上方的错误信息")
    print("3. 可以尝试手动安装失败的包")
    print()
    input("按回车键退出...")
    return 0

if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\n[提示] 安装被用户中断")
        input("\n按回车键退出...")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n[错误] 发生未知错误: {str(e)}")
        import traceback
        traceback.print_exc()
        input("\n按回车键退出...")
        sys.exit(1)