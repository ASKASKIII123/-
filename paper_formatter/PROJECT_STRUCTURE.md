# 项目结构说明

```
paper_formatter/
│
├── main.py                      # 程序主入口文件
│   └── 启动GUI应用程序
│
├── requirements.txt             # Python依赖包列表
│   ├── PyQt5==5.15.9           # 图形界面框架
│   ├── python-docx==0.8.11     # Word文档处理
│   ├── PyPDF2==3.0.1           # PDF文档处理
│   ├── PyMuPDF==1.23.0         # PDF高级处理
│   └── reportlab==4.0.4        # PDF生成和修改
│
├── run.bat                      # Windows启动脚本
│   └── 自动检查环境和启动程序
│
├── check_deps.py                # 依赖包检查工具
│   └── 检查所有依赖包是否正确安装
│
├── README.md                    # 项目说明文档
│   └── 详细的功能介绍和使用说明
│
├── INSTALL_GUIDE.md             # 安装和故障排除指南
│   └── 常见问题解决方案
│
├── QUICKSTART.md                # 快速启动指南
│   └── 简明的使用流程
│
├── processors/                  # 文档处理模块
│   ├── __init__.py             # 模块初始化文件
│   ├── word_processor.py       # Word文档处理器
│   │   ├── WordProcessor 类
│   │   ├── set_font()          # 设置字体
│   │   ├── set_title_font()    # 设置标题字体
│   │   ├── set_paragraph_format()  # 设置段落格式
│   │   ├── set_page_margins()  # 设置页边距
│   │   ├── set_alignment()     # 设置对齐方式
│   │   ├── format_references_apa()  # APA引用格式
│   │   ├── format_references_mla()  # MLA引用格式
│   │   ├── apply_template()    # 应用预设模板
│   │   └── _apply_*_template() # 各模板实现
│   │
│   └── pdf_processor.py        # PDF文档处理器
│       ├── PDFProcessor 类
│       ├── set_page_margins()  # 设置页边距
│       ├── add_page_numbers()  # 添加页码
│       ├── add_header()        # 添加页眉
│       ├── add_footer()        # 添加页脚
│       ├── set_page_size()     # 设置页面大小
│       ├── compress_pdf()      # 压缩PDF
│       ├── apply_template()    # 应用预设模板
│       └── _apply_*_template() # 各模板实现
│
└── gui/                         # 图形界面模块
    ├── __init__.py             # 模块初始化文件
    └── main_window.py          # 主窗口界面
        ├── MainWindow 类       # 主窗口
        ├── FormatWorker 类     # 格式化工作线程
        ├── create_template_tab()    # 创建模板标签页
        ├── create_font_tab()   # 创建字体设置标签页
        ├── create_paragraph_tab()   # 创建段落格式标签页
        ├── create_page_tab()   # 创建页面布局标签页
        ├── create_reference_tab()  # 创建引用格式标签页
        ├── select_file()       # 选择文件
        ├── start_formatting()  # 开始格式化
        ├── get_settings()     # 获取设置
        └── on_formatting_finished()  # 格式化完成回调
```

## 核心模块说明

### 1. 主程序 (main.py)
- 程序的入口点
- 初始化Qt应用程序
- 创建并显示主窗口

### 2. 文档处理器 (processors/)
- **word_processor.py**: 处理Word文档的格式化
- **pdf_processor.py**: 处理PDF文档的格式化

### 3. 图形界面 (gui/)
- **main_window.py**: 提供用户友好的图形界面
- 包含5个标签页：预设模板、字体设置、段落格式、页面布局、引用格式

## 数据流程

```
用户选择文件
    ↓
用户设置格式参数
    ↓
点击"开始格式化"
    ↓
FormatWorker线程启动
    ↓
根据文件类型选择处理器
    ↓
应用格式设置
    ↓
保存格式化后的文件
    ↓
显示完成消息
```

## 技术架构

### 前端
- **框架**: PyQt5
- **特点**: 跨平台、原生外观、丰富的组件

### 后端
- **Word处理**: python-docx
- **PDF处理**: PyMuPDF (fitz) + PyPDF2 + reportlab
- **特点**: 功能强大、支持多种格式

### 并发处理
- **多线程**: 使用QThread进行后台处理
- **优势**: 不阻塞UI界面，提供更好的用户体验

## 扩展性设计

### 添加新的预设模板
1. 在 `word_processor.py` 中添加 `_apply_*_template()` 方法
2. 在 `pdf_processor.py` 中添加对应的模板方法
3. 在 `main_window.py` 的模板下拉框中添加选项

### 添加新的文件格式支持
1. 在 `processors/` 目录下创建新的处理器文件
2. 实现相应的格式化方法
3. 在 `main_window.py` 中添加文件类型支持

### 添加新的格式功能
1. 在相应的处理器类中添加新方法
2. 在 `main_window.py` 中添加UI控件
3. 在 `get_settings()` 方法中添加参数收集

## 性能优化

1. **多线程处理**: 格式化过程在后台线程执行
2. **进度显示**: 实时显示处理进度
3. **内存管理**: 及时释放文档对象
4. **错误处理**: 完善的异常捕获和错误提示

## 安全性考虑

1. **文件备份**: 不覆盖原文件，创建新文件
2. **输入验证**: 检查文件格式和参数有效性
3. **错误处理**: 防止程序崩溃，提供友好的错误信息
4. **资源释放**: 确保文件和资源正确关闭