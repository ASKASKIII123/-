@echo off
chcp 65001 >nul
echo 正在检查Python环境...
python --version
if errorlevel 1 (
    echo 错误：未找到Python，请先安装Python 3.7或更高版本
    pause
    exit /b 1
)
echo.
echo 正在检查依赖包...
python -c "import PyQt5; import docx; import fitz; import PyPDF2; import reportlab" 2>nul
if errorlevel 1 (
    echo 正在安装依赖包...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo 错误：依赖包安装失败
        pause
        exit /b 1
    )
)
echo.
echo 正在启动论文格式自动修改软件...
python main.py
if errorlevel 1 (
    echo.
    echo 程序运行出错，请检查错误信息
)
pause