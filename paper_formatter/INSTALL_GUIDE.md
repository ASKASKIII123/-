# 安装和使用指南

## 常见问题解决

### 问题1：无法识别 pip 命令

**错误信息：**
```
pip : 无法将"pip"项识别为 cmdlet、函数、脚本文件或可运行程序的名称
```

**解决方案：**

1. **使用 python -m pip 代替 pip**
   ```bash
   python -m pip install -r requirements.txt
   ```

2. **检查 Python 是否正确安装**
   ```bash
   python --version
   ```
   如果没有显示版本号，请先安装 Python：https://www.python.org/downloads/

3. **将 Python 添加到系统环境变量**
   - 打开"控制面板" → "系统" → "高级系统设置" → "环境变量"
   - 在"系统变量"中找到"Path"，点击"编辑"
   - 添加 Python 安装路径，例如：`C:\Python39\` 和 `C:\Python39\Scripts\`
   - 重启命令行窗口

### 问题2：依赖包安装失败

**解决方案：**

1. **使用国内镜像源加速安装**
   ```bash
   python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
   ```

2. **升级 pip**
   ```bash
   python -m pip install --upgrade pip
   ```

3. **单独安装每个依赖包**
   ```bash
   python -m pip install PyQt5
   python -m pip install python-docx
   python -m pip install PyPDF2
   python -m pip install PyMuPDF
   python -m pip install reportlab
   ```

### 问题3：程序启动失败

**可能原因和解决方案：**

1. **Python 版本不兼容**
   - 确保使用 Python 3.7 或更高版本
   - 检查命令：`python --version`

2. **依赖包未正确安装**
   - 重新安装依赖包：`python -m pip install -r requirements.txt`

3. **文件路径问题**
   - 确保在正确的目录下运行程序
   - 使用绝对路径运行：`python d:\微信web开发者工具\paper_formatter\main.py`

## 完整安装步骤

### 方法1：使用启动脚本（推荐）

1. 双击 `run.bat` 文件
2. 脚本会自动检查环境并安装依赖
3. 等待程序启动

### 方法2：手动安装

1. **打开命令行窗口**
   - 按 Win+R，输入 `cmd`，按回车

2. **进入项目目录**
   ```bash
   cd d:\微信web开发者工具\paper_formatter
   ```

3. **检查 Python 版本**
   ```bash
   python --version
   ```

4. **安装依赖包**
   ```bash
   python -m pip install -r requirements.txt
   ```

5. **启动程序**
   ```bash
   python main.py
   ```

## 验证安装

运行以下命令验证所有依赖包是否正确安装：

```bash
python -c "import PyQt5; print('PyQt5: OK')"
python -c "import docx; print('python-docx: OK')"
python -c "import fitz; print('PyMuPDF: OK')"
python -c "import PyPDF2; print('PyPDF2: OK')"
python -c "import reportlab; print('reportlab: OK')"
```

如果所有包都显示 "OK"，说明安装成功。

## 使用技巧

### 1. 批量处理文件

如果要处理多个文件，可以创建一个批处理脚本：

```batch
@echo off
for %%f in (*.docx) do (
    echo 正在处理: %%f
    python main.py "%%f"
)
pause
```

### 2. 自定义模板

如果需要创建自定义模板，可以修改以下文件：
- `processors/word_processor.py` 中的模板方法
- `processors/pdf_processor.py` 中的模板方法

### 3. 备份原始文件

建议在格式化前备份原始文件，可以修改程序自动备份：

在 `main.py` 中添加备份功能：

```python
import shutil
import datetime

def backup_file(file_path):
    backup_dir = "backup"
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)
    
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{timestamp}_{os.path.basename(file_path)}"
    backup_path = os.path.join(backup_dir, backup_name)
    
    shutil.copy2(file_path, backup_path)
    return backup_path
```

## 性能优化建议

1. **大文件处理**
   - 对于大型文档，建议分批处理
   - 关闭其他占用内存的程序

2. **PDF 文件优化**
   - 使用压缩功能减小文件大小
   - 避免过度格式化导致文件损坏

3. **定期清理**
   - 删除不需要的临时文件
   - 清理 backup 目录中的旧文件

## 技术支持

如果遇到其他问题，请提供以下信息：
1. Python 版本：`python --version`
2. pip 版本：`python -m pip --version`
3. 操作系统版本
4. 完整的错误信息

## 卸载说明

如果需要卸载程序：

1. 删除项目文件夹
2. 卸载依赖包：
   ```bash
   python -m pip uninstall PyQt5 python-docx PyPDF2 PyMuPDF reportlab
   ```