#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QComboBox, QSpinBox, 
                             QDoubleSpinBox, QFileDialog, QTabWidget, QGroupBox,
                             QFormLayout, QMessageBox, QProgressBar, QTextEdit,
                             QCheckBox, QListWidget, QListWidgetItem)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
import os
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from processors.word_processor import WordProcessor
from processors.pdf_processor import PDFProcessor

class FormatWorker(QThread):
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(int)
    
    def __init__(self, file_path, settings):
        super().__init__()
        self.file_path = file_path
        self.settings = settings
    
    def run(self):
        try:
            file_ext = os.path.splitext(self.file_path)[1].lower()
            
            if file_ext == '.docx':
                processor = WordProcessor(self.file_path)
                
                if self.settings.get('template'):
                    processor.apply_template(self.settings['template'])
                
                if self.settings.get('font_name'):
                    processor.set_font(
                        font_name=self.settings['font_name'],
                        font_size=self.settings.get('font_size', 12),
                        bold=self.settings.get('bold', False),
                        italic=self.settings.get('italic', False)
                    )
                
                if self.settings.get('title_font_name'):
                    processor.set_title_font(
                        font_name=self.settings['title_font_name'],
                        font_size=self.settings.get('title_font_size', 16),
                        bold=self.settings.get('title_bold', True)
                    )
                
                if self.settings.get('line_spacing'):
                    processor.set_paragraph_format(
                        line_spacing=self.settings['line_spacing'],
                        first_line_indent=self.settings.get('first_line_indent', 2),
                        space_before=self.settings.get('space_before', 0),
                        space_after=self.settings.get('space_after', 0)
                    )
                
                if self.settings.get('margin_top'):
                    processor.set_page_margins(
                        top=self.settings['margin_top'],
                        bottom=self.settings['margin_bottom'],
                        left=self.settings['margin_left'],
                        right=self.settings['margin_right']
                    )
                
                if self.settings.get('reference_format'):
                    if self.settings['reference_format'] == 'apa':
                        processor.format_references_apa()
                    elif self.settings['reference_format'] == 'mla':
                        processor.format_references_mla()
                    elif self.settings['reference_format'] == 'cssci':
                        processor.format_references_cssci()
                
                if self.settings.get('comparison_mode') and self.settings.get('original_file_path'):
                    processor.enable_comparison_mode(self.settings['original_file_path'])
                
                if self.settings.get('chapter_mode') and self.settings.get('chapter_indices'):
                    for chapter_index in self.settings['chapter_indices']:
                        processor.format_chapter(chapter_index, template_name=self.settings.get('template'))
                
                output_path = self.settings.get('output_path')
                processor.save(output_path)
                
            elif file_ext == '.pdf':
                processor = PDFProcessor(self.file_path)
                
                if self.settings.get('template'):
                    processor.apply_template(self.settings['template'])
                
                if self.settings.get('margin_top'):
                    output_path = self.settings.get('output_path', self.file_path)
                    processor.set_page_margins(
                        top=self.settings['margin_top'],
                        bottom=self.settings['margin_bottom'],
                        left=self.settings['margin_left'],
                        right=self.settings['margin_right'],
                        output_path=output_path
                    )
                
                if self.settings.get('add_page_numbers'):
                    output_path = self.settings.get('output_path', self.file_path)
                    processor.add_page_numbers(
                        position=self.settings['page_number_position'],
                        font_size=self.settings['page_number_size'],
                        output_path=output_path
                    )
                
                if self.settings.get('header_text'):
                    output_path = self.settings.get('output_path', self.file_path)
                    processor.add_header(
                        text=self.settings['header_text'],
                        font_size=self.settings['header_size'],
                        output_path=output_path
                    )
                
                if self.settings.get('footer_text'):
                    output_path = self.settings.get('output_path', self.file_path)
                    processor.add_footer(
                        text=self.settings['footer_text'],
                        font_size=self.settings['footer_size'],
                        output_path=output_path
                    )
                
                processor.save(self.settings.get('output_path'))
            
            self.finished.emit(True, "格式修改完成！")
            
        except Exception as e:
            self.finished.emit(False, f"错误: {str(e)}")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.original_file_path = None
        self.worker = None
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle("论文格式自动修改软件")
        self.setGeometry(100, 100, 800, 600)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        
        title_label = QLabel("论文格式自动修改软件")
        title_label.setFont(QFont("Arial", 18, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        file_group = QGroupBox("文件选择")
        file_layout = QHBoxLayout()
        
        self.file_label = QLabel("未选择文件")
        file_layout.addWidget(self.file_label)
        
        select_button = QPushButton("选择文件")
        select_button.clicked.connect(self.select_file)
        file_layout.addWidget(select_button)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        self.tab_widget = QTabWidget()
        self.tab_widget.addTab(self.create_template_tab(), "预设模板")
        self.tab_widget.addTab(self.create_font_tab(), "字体设置")
        self.tab_widget.addTab(self.create_paragraph_tab(), "段落格式")
        self.tab_widget.addTab(self.create_page_tab(), "页面布局")
        self.tab_widget.addTab(self.create_reference_tab(), "引用格式")
        self.tab_widget.addTab(self.create_comparison_tab(), "对比模式")
        self.tab_widget.addTab(self.create_chapter_tab(), "章节处理")
        layout.addWidget(self.tab_widget)
        
        button_layout = QHBoxLayout()
        
        self.format_button = QPushButton("开始格式化")
        self.format_button.clicked.connect(self.start_formatting)
        self.format_button.setEnabled(False)
        button_layout.addWidget(self.format_button)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        button_layout.addWidget(self.progress_bar)
        
        layout.addLayout(button_layout)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(100)
        layout.addWidget(self.log_text)
    
    def create_template_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        self.template_combo = QComboBox()
        self.template_combo.addItems(["", "academic", "journal", "thesis", "conference", "imrad", "cssci"])
        self.template_combo.setItemText(0, "选择模板...")
        form_layout.addRow("预设模板:", self.template_combo)
        
        template_info = QLabel(
            "academic: 学术论文模板\n"
            "journal: 期刊论文模板\n"
            "thesis: 学位论文模板\n"
            "conference: 会议论文模板\n"
            "imrad: IMRAD格式模板\n"
            "cssci: CSSCI格式模板"
        )
        template_info.setStyleSheet("color: gray; font-size: 10px;")
        form_layout.addRow(template_info)
        
        layout.addLayout(form_layout)
        layout.addStretch()
        
        widget.setLayout(layout)
        return widget
    
    def create_font_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        self.font_name_combo = QComboBox()
        self.font_name_combo.addItems(["", "宋体", "黑体", "楷体", "Arial", "Times New Roman", "Calibri"])
        self.font_name_combo.setItemText(0, "选择字体...")
        form_layout.addRow("正文字体:", self.font_name_combo)
        
        self.font_size_spin = QSpinBox()
        self.font_size_spin.setRange(8, 72)
        self.font_size_spin.setValue(12)
        form_layout.addRow("正文字号:", self.font_size_spin)
        
        self.title_font_name_combo = QComboBox()
        self.title_font_name_combo.addItems(["", "黑体", "宋体", "Arial", "Times New Roman"])
        self.title_font_name_combo.setItemText(0, "选择字体...")
        form_layout.addRow("标题字体:", self.title_font_name_combo)
        
        self.title_font_size_spin = QSpinBox()
        self.title_font_size_spin.setRange(8, 72)
        self.title_font_size_spin.setValue(16)
        form_layout.addRow("标题字号:", self.title_font_size_spin)
        
        layout.addLayout(form_layout)
        layout.addStretch()
        
        widget.setLayout(layout)
        return widget
    
    def create_paragraph_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        self.line_spacing_spin = QDoubleSpinBox()
        self.line_spacing_spin.setRange(0.5, 3.0)
        self.line_spacing_spin.setSingleStep(0.1)
        self.line_spacing_spin.setValue(1.5)
        form_layout.addRow("行间距:", self.line_spacing_spin)
        
        self.first_line_indent_spin = QDoubleSpinBox()
        self.first_line_indent_spin.setRange(0, 5.0)
        self.first_line_indent_spin.setSingleStep(0.1)
        self.first_line_indent_spin.setValue(2.0)
        form_layout.addRow("首行缩进(cm):", self.first_line_indent_spin)
        
        self.space_before_spin = QDoubleSpinBox()
        self.space_before_spin.setRange(0, 50)
        self.space_before_spin.setValue(0)
        form_layout.addRow("段前间距(pt):", self.space_before_spin)
        
        self.space_after_spin = QDoubleSpinBox()
        self.space_after_spin.setRange(0, 50)
        self.space_after_spin.setValue(0)
        form_layout.addRow("段后间距(pt):", self.space_after_spin)
        
        layout.addLayout(form_layout)
        layout.addStretch()
        
        widget.setLayout(layout)
        return widget
    
    def create_page_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        self.margin_top_spin = QDoubleSpinBox()
        self.margin_top_spin.setRange(0, 10)
        self.margin_top_spin.setSingleStep(0.1)
        self.margin_top_spin.setValue(2.54)
        form_layout.addRow("上边距(cm):", self.margin_top_spin)
        
        self.margin_bottom_spin = QDoubleSpinBox()
        self.margin_bottom_spin.setRange(0, 10)
        self.margin_bottom_spin.setSingleStep(0.1)
        self.margin_bottom_spin.setValue(2.54)
        form_layout.addRow("下边距(cm):", self.margin_bottom_spin)
        
        self.margin_left_spin = QDoubleSpinBox()
        self.margin_left_spin.setRange(0, 10)
        self.margin_left_spin.setSingleStep(0.1)
        self.margin_left_spin.setValue(3.17)
        form_layout.addRow("左边距(cm):", self.margin_left_spin)
        
        self.margin_right_spin = QDoubleSpinBox()
        self.margin_right_spin.setRange(0, 10)
        self.margin_right_spin.setSingleStep(0.1)
        self.margin_right_spin.setValue(3.17)
        form_layout.addRow("右边距(cm):", self.margin_right_spin)
        
        layout.addLayout(form_layout)
        layout.addStretch()
        
        widget.setLayout(layout)
        return widget
    
    def create_reference_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        self.reference_format_combo = QComboBox()
        self.reference_format_combo.addItems(["", "apa", "mla", "cssci"])
        self.reference_format_combo.setItemText(0, "选择格式...")
        form_layout.addRow("引用格式:", self.reference_format_combo)
        
        ref_info = QLabel(
            "APA: 美国心理学会格式\n"
            "MLA: 现代语言协会格式\n"
            "CSSCI: 中文社会科学引文索引格式"
        )
        ref_info.setStyleSheet("color: gray; font-size: 10px;")
        form_layout.addRow(ref_info)
        
        layout.addLayout(form_layout)
        layout.addStretch()
        
        widget.setLayout(layout)
        return widget
    
    def create_comparison_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        self.comparison_mode_checkbox = QCheckBox("启用对比模式")
        self.comparison_mode_checkbox.setToolTip("启用后，新增内容显示为绿色，删除内容显示为红色删除线")
        form_layout.addRow("对比模式:", self.comparison_mode_checkbox)
        
        self.original_file_label = QLabel("未选择原始文件")
        form_layout.addRow("原始文件:", self.original_file_label)
        
        select_original_button = QPushButton("选择原始文件")
        select_original_button.clicked.connect(self.select_original_file)
        form_layout.addRow("", select_original_button)
        
        comparison_info = QLabel(
            "对比模式说明:\n"
            "• 绿色文字：新增的内容\n"
            "• 红色删除线：删除的内容\n"
            "• 需要选择原始文件进行对比"
        )
        comparison_info.setStyleSheet("color: gray; font-size: 10px;")
        form_layout.addRow(comparison_info)
        
        layout.addLayout(form_layout)
        layout.addStretch()
        
        widget.setLayout(layout)
        return widget
    
    def create_chapter_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        self.chapter_mode_checkbox = QCheckBox("启用分章节处理")
        self.chapter_mode_checkbox.setToolTip("启用后，可以按章节逐个优化，而不是一次性处理全文")
        self.chapter_mode_checkbox.stateChanged.connect(self.on_chapter_mode_changed)
        form_layout.addRow("章节处理:", self.chapter_mode_checkbox)
        
        self.chapter_list_widget = QListWidget()
        self.chapter_list_widget.setEnabled(False)
        form_layout.addRow("章节列表:", self.chapter_list_widget)
        
        self.refresh_chapters_button = QPushButton("刷新章节列表")
        self.refresh_chapters_button.clicked.connect(self.refresh_chapters)
        self.refresh_chapters_button.setEnabled(False)
        form_layout.addRow("", self.refresh_chapters_button)
        
        chapter_info = QLabel(
            "章节处理说明:\n"
            "• 自动识别文档中的章节（摘要、引言、方法等）\n"
            "• 可以选择特定章节进行格式化\n"
            "• 支持中英文章节名称"
        )
        chapter_info.setStyleSheet("color: gray; font-size: 10px;")
        form_layout.addRow(chapter_info)
        
        layout.addLayout(form_layout)
        layout.addStretch()
        
        widget.setLayout(layout)
        return widget
    
    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择文件",
            "",
            "支持的文件 (*.docx *.pdf);;Word文档 (*.docx);;PDF文档 (*.pdf)"
        )
        
        if file_path:
            self.file_path = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.format_button.setEnabled(True)
            self.log_text.append(f"已选择文件: {file_path}")
            self.chapter_list_widget.setEnabled(False)
            self.chapter_list_widget.clear()
            self.chapter_mode_checkbox.setChecked(False)
            self.comparison_mode_checkbox.setChecked(False)
    
    def select_original_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择原始文件",
            "",
            "支持的文件 (*.docx *.pdf);;Word文档 (*.docx);;PDF文档 (*.pdf)"
        )
        
        if file_path:
            self.original_file_path = file_path
            self.original_file_label.setText(os.path.basename(file_path))
            self.log_text.append(f"已选择原始文件: {file_path}")
    
    def refresh_chapters(self):
        if not self.file_path:
            QMessageBox.warning(self, "警告", "请先选择文件！")
            return
        
        try:
            file_ext = os.path.splitext(self.file_path)[1].lower()
            
            if file_ext == '.docx':
                from processors.word_processor import WordProcessor
                processor = WordProcessor(self.file_path)
                chapters = processor.get_chapters()
            elif file_ext == '.pdf':
                from processors.pdf_processor import PDFProcessor
                processor = PDFProcessor(self.file_path)
                chapters = processor.get_chapters()
            else:
                QMessageBox.warning(self, "警告", "不支持的文件格式！")
                return
            
            self.chapter_list_widget.clear()
            
            for i, chapter in enumerate(chapters):
                item = QListWidgetItem(f"{i+1}. {chapter['name']}")
                item.setData(Qt.UserRole, i)
                self.chapter_list_widget.addItem(item)
            
            self.log_text.append(f"找到 {len(chapters)} 个章节")
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"获取章节失败: {str(e)}")
            self.log_text.append(f"错误: {str(e)}")
    
    def on_chapter_mode_changed(self, state):
        if state == Qt.Checked:
            self.chapter_list_widget.setEnabled(True)
            refresh_chapters_button = self.findChild(QPushButton, "refresh_chapters_button")
            if refresh_chapters_button:
                refresh_chapters_button.setEnabled(True)
        else:
            self.chapter_list_widget.setEnabled(False)
            refresh_chapters_button = self.findChild(QPushButton, "refresh_chapters_button")
            if refresh_chapters_button:
                refresh_chapters_button.setEnabled(False)
    
    def start_formatting(self):
        if not self.file_path:
            QMessageBox.warning(self, "警告", "请先选择文件！")
            return
        
        settings = self.get_settings()
        
        if not any(settings.values()):
            QMessageBox.warning(self, "警告", "请至少选择一个格式设置！")
            return
        
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存文件",
            "",
            "支持的文件 (*.docx *.pdf);;Word文档 (*.docx);;PDF文档 (*.pdf)"
        )
        
        if output_path:
            if output_path == self.file_path:
                QMessageBox.warning(self, "警告", "输出文件不能与输入文件相同！")
                return
            
            if output_path == self.original_file_path:
                QMessageBox.warning(self, "警告", "输出文件不能与原始文件相同！")
                return
            
            settings['output_path'] = output_path
            self.format_button.setEnabled(False)
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)
            self.log_text.append("开始格式化...")
            
            self.worker = FormatWorker(self.file_path, settings)
            self.worker.finished.connect(self.on_formatting_finished)
            self.worker.start()
    
    def get_settings(self):
        settings = {}
        
        template = self.template_combo.currentText()
        if template and template != "选择模板...":
            settings['template'] = template
        
        font_name = self.font_name_combo.currentText()
        if font_name and font_name != "选择字体...":
            settings['font_name'] = font_name
            settings['font_size'] = self.font_size_spin.value()
        
        title_font_name = self.title_font_name_combo.currentText()
        if title_font_name and title_font_name != "选择字体...":
            settings['title_font_name'] = title_font_name
            settings['title_font_size'] = self.title_font_size_spin.value()
        
        settings['line_spacing'] = self.line_spacing_spin.value()
        settings['first_line_indent'] = self.first_line_indent_spin.value()
        settings['space_before'] = self.space_before_spin.value()
        settings['space_after'] = self.space_after_spin.value()
        
        settings['margin_top'] = self.margin_top_spin.value()
        settings['margin_bottom'] = self.margin_bottom_spin.value()
        settings['margin_left'] = self.margin_left_spin.value()
        settings['margin_right'] = self.margin_right_spin.value()
        
        ref_format = self.reference_format_combo.currentText()
        if ref_format and ref_format != "选择格式...":
            settings['reference_format'] = ref_format
        
        settings['comparison_mode'] = self.comparison_mode_checkbox.isChecked()
        if self.original_file_path:
            settings['original_file_path'] = self.original_file_path
        
        settings['chapter_mode'] = self.chapter_mode_checkbox.isChecked()
        if self.chapter_mode_checkbox.isChecked():
            selected_items = self.chapter_list_widget.selectedItems()
            if selected_items:
                chapter_indices = [item.data(Qt.UserRole) for item in selected_items]
                settings['chapter_indices'] = chapter_indices
        
        return settings
    
    def on_formatting_finished(self, success, message):
        self.format_button.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        if success:
            self.log_text.append(message)
            QMessageBox.information(self, "成功", message)
        else:
            self.log_text.append(message)
            QMessageBox.critical(self, "错误", message)