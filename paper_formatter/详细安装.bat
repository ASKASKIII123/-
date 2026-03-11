@echo off
chcp 65001 >nul
title 依赖包详细安装工具

echo ========================================
echo   依赖包详细安装工具
echo ========================================
echo.
echo 这个工具会详细显示每个包的安装过程
echo 如果安装失败，会显示具体的错误信息
echo.
echo 正在启动...
echo.

python install_deps_detailed.py

pause