#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import PyPDF2
import fitz
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

class PDFProcessor:
    def __init__(self, file_path):
        self.file_path = file_path
        self.temp_file = None
        self.pdf_document = fitz.open(file_path)
    
    def save(self, output_path=None):
        if output_path is None:
            output_path = self.file_path
        
        self.pdf_document.save(output_path)
        if output_path != self.file_path:
            self.pdf_document.close()
            self.pdf_document = fitz.open(output_path)
    
    def get_page_count(self):
        return len(self.pdf_document)
    
    def extract_text(self):
        text = ""
        for page in self.pdf_document:
            text += page.get_text()
        return text
    
    def set_page_margins(self, top=2.54, bottom=2.54, left=3.17, right=3.17, output_path=None):
        if output_path is None:
            output_path = self.file_path.replace('.pdf', '_formatted.pdf')
        
        new_pdf = fitz.open()
        
        for page in self.pdf_document:
            rect = page.rect
            new_rect = fitz.Rect(
                rect.x0 + left * 28.35,
                rect.y0 + top * 28.35,
                rect.x1 - right * 28.35,
                rect.y1 - bottom * 28.35
            )
            
            new_page = new_pdf.new_page(width=rect.width, height=rect.height)
            new_page.show_pdf_page(rect, self.pdf_document, page.number)
        
        new_pdf.save(output_path)
        new_pdf.close()
        self.pdf_document = fitz.open(output_path)
    
    def add_page_numbers(self, position="bottom", font_size=10, output_path=None):
        if output_path is None:
            output_path = self.file_path.replace('.pdf', '_numbered.pdf')
        
        for page_num, page in enumerate(self.pdf_document, 1):
            text = str(page_num)
            rect = page.rect
            
            if position == "bottom":
                x = rect.width / 2
                y = rect.height - 2 * 28.35
            elif position == "top":
                x = rect.width / 2
                y = 2 * 28.35
            else:
                x = rect.width / 2
                y = rect.height - 2 * 28.35
            
            point = fitz.Point(x, y)
            page.insert_text(
                point,
                text,
                fontsize=font_size,
                color=(0, 0, 0)
            )
        
        self.pdf_document.save(output_path)
        self.pdf_document = fitz.open(output_path)
    
    def add_header(self, text, font_size=10, output_path=None):
        if output_path is None:
            output_path = self.file_path.replace('.pdf', '_header.pdf')
        
        for page in self.pdf_document:
            rect = page.rect
            x = rect.width / 2
            y = 1.5 * 28.35
            
            point = fitz.Point(x, y)
            page.insert_text(
                point,
                text,
                fontsize=font_size,
                color=(0, 0, 0)
            )
        
        self.pdf_document.save(output_path)
        self.pdf_document = fitz.open(output_path)
    
    def add_footer(self, text, font_size=10, output_path=None):
        if output_path is None:
            output_path = self.file_path.replace('.pdf', '_footer.pdf')
        
        for page in self.pdf_document:
            rect = page.rect
            x = rect.width / 2
            y = rect.height - 1.5 * 28.35
            
            page.insert_text(
                (x, y),
                text,
                fontsize=font_size,
                color=(0, 0, 0),
                align=1
            )
        
        self.pdf_document.save(output_path)
        self.pdf_document = fitz.open(output_path)
    
    def set_page_size(self, width=21.0, height=29.7, output_path=None):
        if output_path is None:
            output_path = self.file_path.replace('.pdf', '_resized.pdf')
        
        new_pdf = fitz.open()
        
        for page in self.pdf_document:
            new_page = new_pdf.new_page(width=width * 28.35, height=height * 28.35)
            new_page.show_pdf_page(new_page.rect, self.pdf_document, page.number)
        
        new_pdf.save(output_path)
        new_pdf.close()
        self.pdf_document = fitz.open(output_path)
    
    def compress_pdf(self, output_path=None):
        if output_path is None:
            output_path = self.file_path.replace('.pdf', '_compressed.pdf')
        
        self.pdf_document.save(output_path, deflate=True)
        self.pdf_document = fitz.open(output_path)
    
    def apply_template(self, template_name):
        templates = {
            "academic": self._apply_academic_template,
            "journal": self._apply_journal_template,
            "thesis": self._apply_thesis_template,
            "conference": self._apply_conference_template,
            "imrad": self._apply_imrad_template,
            "cssci": self._apply_cssci_template
        }
        
        if template_name in templates:
            templates[template_name]()
    
    def _apply_academic_template(self):
        self.set_page_margins(top=2.54, bottom=2.54, left=3.17, right=3.17)
        self.add_page_numbers(position="bottom", font_size=10)
    
    def _apply_journal_template(self):
        self.set_page_margins(top=2.54, bottom=2.54, left=2.54, right=2.54)
        self.add_page_numbers(position="bottom", font_size=9)
    
    def _apply_thesis_template(self):
        self.set_page_margins(top=3.0, bottom=3.0, left=3.0, right=3.0)
        self.add_page_numbers(position="bottom", font_size=12)
    
    def _apply_conference_template(self):
        self.set_page_margins(top=2.0, bottom=2.0, left=2.5, right=2.5)
        self.add_page_numbers(position="bottom", font_size=9)
    
    def _apply_imrad_template(self):
        self.set_page_margins(top=2.54, bottom=2.54, left=2.54, right=2.54)
        self.add_page_numbers(position="bottom", font_size=10)
    
    def _apply_cssci_template(self):
        self.set_page_margins(top=2.54, bottom=2.54, left=3.17, right=3.17)
        self.add_page_numbers(position="bottom", font_size=10)
    
    def enable_comparison_mode(self, original_file_path):
        original_doc = fitz.open(original_file_path)
        
        for page_num, (orig_page, new_page) in enumerate(zip(original_doc, self.pdf_document)):
            orig_text = orig_page.get_text()
            new_text = new_page.get_text()
            
            if orig_text != new_text:
                self._add_comparison_annotation(new_page, orig_text, new_text)
        
        original_doc.close()
    
    def _add_comparison_annotation(self, page, orig_text, new_text):
        rect = page.rect
        
        if new_text and not orig_text:
            point = fitz.Point(rect.x0 + 50, rect.y0 + 50)
            page.insert_text(
                point,
                "新增内容",
                fontsize=10,
                color=(0, 0.5, 0)
            )
        elif not new_text and orig_text:
            point = fitz.Point(rect.x0 + 50, rect.y0 + 50)
            page.insert_text(
                point,
                "删除内容",
                fontsize=10,
                color=(1, 0, 0)
            )
    
    def get_chapters(self):
        chapters = []
        chapter_keywords = ['摘要', '引言', '方法', '结果', '讨论', '结论', '参考文献',
                        'Abstract', 'Introduction', 'Methods', 'Results', 'Discussion',
                        'Conclusion', 'References', '第一章', '第二章', '第三章',
                        '第四章', '第五章', '第六章', 'Chapter']
        
        for page_num, page in enumerate(self.pdf_document):
            text = page.get_text()
            for keyword in chapter_keywords:
                if keyword in text:
                    chapters.append({
                        'name': keyword,
                        'page_number': page_num + 1,
                        'text': text
                    })
                    break
        
        return chapters
    
    def format_chapter(self, chapter_index, template_name=None):
        chapters = self.get_chapters()
        
        if chapter_index >= len(chapters):
            raise IndexError(f"章节索引 {chapter_index} 超出范围，共有 {len(chapters)} 个章节")
        
        chapter = chapters[chapter_index]
        page_num = chapter['page_number'] - 1
        
        if template_name:
            template_methods = {
                "academic": self._apply_academic_template,
                "journal": self._apply_journal_template,
                "thesis": self._apply_thesis_template,
                "conference": self._apply_conference_template,
                "imrad": self._apply_imrad_template,
                "cssci": self._apply_cssci_template
            }
            
            if template_name in template_methods:
                template_methods[template_name]()
    
    def get_document_info(self):
        return {
            "pages": len(self.pdf_document),
            "file_size": os.path.getsize(self.file_path),
            "text_length": len(self.extract_text()),
            "chapters": len(self.get_chapters())
        }