@echo off
chcp 65001 >nul
title 论文格式自动修改软件

echo ========================================
echo    论文格式自动修改软件 v1.0
echo ========================================
echo.

echo [1/4] 检查Python环境...
where python >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python！
    echo.
    echo 请按照以下步骤安装Python：
    echo 1. 访问 https://www.python.org/downloads/
    echo 2. 下载并安装 Python 3.8 或更高版本
    echo 3. 安装时务必勾选 "Add Python to PATH"
    echo 4. 重启此程序
    echo.
    pause
    exit /b 1
)

python --version
if errorlevel 1 (
    echo [错误] Python无法正常运行！
    pause
    exit /b 1
)

echo [成功] Python环境正常
echo.

echo [2/4] 检查依赖包...
python -c "import PyQt5, docx, fitz, PyPDF2, reportlab" >nul 2>&1
if errorlevel 1 (
    echo [提示] 正在安装依赖包（使用清华镜像源加速）...
    echo 这可能需要几分钟，请耐心等待...
    echo.
    python -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn
    python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn
    if errorlevel 1 (
        echo.
        echo [错误] 依赖包安装失败！
        echo.
        echo 请尝试以下方法：
        echo 1. 使用管理员权限运行此脚本
        echo 2. 手动运行: python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
        echo.
        pause
        exit /b 1
    )
)

echo [成功] 依赖包检查通过
echo.

echo [3/4] 检查程序文件...
if not exist "main.py" (
    echo [错误] 未找到 main.py 文件！
    pause
    exit /b 1
)

echo [成功] 程序文件检查通过
echo.

echo [4/4] 启动程序...
echo.
echo ========================================
echo.
python main.py

if errorlevel 1 (
    echo.
    echo [错误] 程序运行出错！
    echo.
    echo 可能的原因：
    echo 1. Python版本过低（需要3.7+）
    echo 2. 依赖包未正确安装
    echo 3. 程序文件损坏
    echo.
    echo 请查看上方的错误信息，或参考 TROUBLESHOOTING.md
    echo.
)

echo.
echo ========================================
echo 程序已关闭
echo ========================================
pause