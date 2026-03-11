# 故障排除指南

## 目录
1. [安装问题](#安装问题)
2. [运行问题](#运行问题)
3. [格式化问题](#格式化问题)
4. [性能问题](#性能问题)

---

## 安装问题

### 问题1: "无法识别pip命令"

**症状：**
```
pip : 无法将"pip"项识别为 cmdlet、函数、脚本文件或可运行程序的名称
```

**解决方案：**

1. **使用 python -m pip**
   ```bash
   python -m pip install -r requirements.txt
   ```

2. **检查Python安装**
   ```bash
   python --version
   ```
   如果没有显示版本号，请从 https://www.python.org/downloads/ 下载并安装Python。

3. **配置环境变量**
   - 打开"控制面板" → "系统" → "高级系统设置" → "环境变量"
   - 在"系统变量"中找到"Path"，点击"编辑"
   - 添加以下路径（根据你的Python安装位置调整）：
     - `C:\Python39\`
     - `C:\Python39\Scripts\`
   - 重启命令行窗口

### 问题2: 依赖包安装失败

**症状：**
```
ERROR: Could not find a version that satisfies the requirement...
```

**解决方案：**

1. **使用国内镜像源**
   ```bash
   python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
   ```

2. **升级pip**
   ```bash
   python -m pip install --upgrade pip
   ```

3. **单独安装每个包**
   ```bash
   python -m pip install PyQt5
   python -m pip install python-docx
   python -m pip install PyPDF2
   python -m pip install PyMuPDF
   python -m pip install reportlab
   ```

4. **使用管理员权限**
   - 右键点击命令提示符，选择"以管理员身份运行"
   - 重新执行安装命令

### 问题3: 网络连接问题

**症状：**
```
ReadTimeoutError: HTTPSConnectionPool...
```

**解决方案：**

1. **增加超时时间**
   ```bash
   python -m pip install -r requirements.txt --timeout 100
   ```

2. **使用国内镜像源**
   ```bash
   python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
   ```

3. **离线安装**
   - 在有网络的机器上下载wheel文件
   - 复制到目标机器
   - 使用 `pip install *.whl` 安装

---

## 运行问题

### 问题1: 程序无法启动

**症状：**
```
ModuleNotFoundError: No module named 'PyQt5'
```

**解决方案：**

1. **检查依赖包**
   ```bash
   python check_deps.py
   ```

2. **重新安装依赖**
   ```bash
   python -m pip install -r requirements.txt
   ```

3. **检查Python版本**
   ```bash
   python --version
   ```
   确保使用Python 3.7或更高版本

### 问题2: 界面显示异常

**症状：**
- 窗口显示不完整
- 按钮无法点击
- 字体显示异常

**解决方案：**

1. **更新PyQt5**
   ```bash
   python -m pip install --upgrade PyQt5
   ```

2. **检查屏幕分辨率**
   - 确保屏幕分辨率至少为1024x768
   - 尝试调整DPI设置

3. **重置程序设置**
   - 删除程序配置文件（如果有）
   - 重新启动程序

### 问题3: 文件选择对话框无法打开

**症状：**
- 点击"选择文件"按钮无反应
- 程序卡死

**解决方案：**

1. **检查文件权限**
   - 确保程序有访问文件系统的权限
   - 尝试以管理员身份运行

2. **检查文件路径**
   - 避免使用过长的文件路径
   - 避免使用特殊字符

3. **重启程序**
   - 关闭程序并重新启动

---

## 格式化问题

### 问题1: Word文档格式化后内容丢失

**症状：**
- 格式化后部分内容消失
- 图片无法显示
- 表格格式错乱

**解决方案：**

1. **检查文件格式**
   - 确保使用.docx格式，不是.doc格式
   - 转换.doc到.docx：在Word中"另存为".docx

2. **备份原始文件**
   - 格式化前始终备份原始文件
   - 使用不同的文件名保存

3. **分步格式化**
   - 先应用预设模板
   - 再逐步调整自定义设置
   - 每步后检查结果

### 问题2: PDF文档格式化失败

**症状：**
```
Error: PDF file is encrypted
```

**解决方案：**

1. **检查PDF是否加密**
   - 尝试在PDF阅读器中打开
   - 如果需要密码，先解密

2. **使用PDF转换工具**
   - 将PDF转换为Word
   - 格式化后再转换回PDF

3. **限制PDF格式化功能**
   - PDF格式化功能有限
   - 主要支持页边距和页码

### 问题3: 引用格式不正确

**症状：**
- 引用格式不符合要求
- 悬挂缩进不正确

**解决方案：**

1. **检查引用格式**
   - 确保引用以[数字]开头
   - 例如：[1]、[2]等

2. **手动调整**
   - 格式化后手动检查引用格式
   - 在Word中调整悬挂缩进

3. **使用正确的引用格式**
   - APA格式：作者(年份)
   - MLA格式：作者 页码

---

## 性能问题

### 问题1: 处理大文件很慢

**症状：**
- 处理大文件需要很长时间
- 程序无响应

**解决方案：**

1. **关闭其他程序**
   - 关闭占用内存的程序
   - 释放系统资源

2. **分批处理**
   - 将大文件分成小文件
   - 分别处理后再合并

3. **优化文档**
   - 删除不必要的格式
   - 压缩图片
   - 减少文档大小

### 问题2: 内存占用过高

**症状：**
- 程序占用大量内存
- 系统变慢

**解决方案：**

1. **处理完成后关闭程序**
   - 格式化完成后及时关闭程序
   - 释放内存

2. **减少同时处理的文件**
   - 一次只处理一个文件
   - 避免批量处理

3. **增加虚拟内存**
   - 增加系统虚拟内存大小
   - 重启计算机

---

## 其他问题

### 问题1: 程序崩溃

**症状：**
- 程序突然关闭
- 显示错误信息

**解决方案：**

1. **查看错误日志**
   - 记录完整的错误信息
   - 检查错误堆栈

2. **重新安装程序**
   - 删除程序文件夹
   - 重新解压或克隆项目
   - 重新安装依赖

3. **检查系统兼容性**
   - 确保操作系统版本兼容
   - 检查是否有冲突的软件

### 问题2: 无法保存文件

**症状：**
- 格式化完成后无法保存
- 提示权限错误

**解决方案：**

1. **检查文件权限**
   - 确保有写入权限
   - 尝试保存到其他位置

2. **检查磁盘空间**
   - 确保有足够的磁盘空间
   - 清理临时文件

3. **使用不同的文件名**
   - 避免覆盖原文件
   - 使用新的文件名

---

## 获取帮助

如果以上解决方案都无法解决你的问题，请：

1. **收集信息**
   - Python版本：`python --version`
   - pip版本：`python -m pip --version`
   - 操作系统版本
   - 完整的错误信息

2. **查看文档**
   - README.md
   - INSTALL_GUIDE.md
   - QUICKSTART.md

3. **测试环境**
   - 使用提供的测试文档
   - 运行 `python create_test_doc.py` 创建测试文档
   - 在测试文档上重现问题

---

## 常用命令

```bash
# 检查Python版本
python --version

# 检查pip版本
python -m pip --version

# 安装依赖
python -m pip install -r requirements.txt

# 检查依赖
python check_deps.py

# 创建测试文档
python create_test_doc.py

# 启动程序
python main.py
```

---

## 预防措施

1. **定期备份**
   - 备份原始文件
   - 备份程序配置

2. **保持更新**
   - 定期更新Python
   - 定期更新依赖包

3. **测试先行**
   - 在测试文档上测试
   - 确认无误后再处理正式文档

4. **记录问题**
   - 记录遇到的问题
   - 记录解决方案
   - 方便下次参考