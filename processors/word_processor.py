#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Word处理器 - 移除编辑保护和VML水印
第一性原理：.docx/.xlsx/.pptx = ZIP压缩包，保护信息 = XML标签
"""

import os
import re
import shutil
import zipfile
import tempfile
import subprocess
import sys
from pathlib import Path
from typing import Set, Tuple, Optional
from .base import DocumentProcessor, ProcessResult


class WordProcessor(DocumentProcessor):
    """Word文档处理器"""
    
    SUPPORTED_EXTENSIONS: Set[str] = {'.doc', '.docx'}
    
    # 处理模式
    MODE_PROTECTION = 'protection'  # 仅去保护
    MODE_WATERMARK = 'watermark'    # 仅去水印
    MODE_ALL = 'all'                # 去保护+去水印
    
    def __init__(self, preview: bool = False, keywords: list = None,
                 mode: str = 'all', preserve_format: bool = True):
        super().__init__(preview, keywords)
        self.mode = mode
        self.preserve_format = preserve_format
    
    def get_description(self) -> str:
        return "Word文档解锁 (保护+水印)"
    
    @staticmethod
    def is_encrypted(file_path: str) -> bool:
        """
        检测文档是否被密码加密
        第一性原理：OOXML格式(.docx)本质是ZIP，如果不是ZIP格式说明被OLE加密
        """
        return not zipfile.is_zipfile(file_path)
    
    def process(self, input_path: str, output_path: str = None,
                progress_callback=None) -> ProcessResult:
        """处理Word文档"""
        input_path = Path(input_path)
        suffix = "_unlocked" if self.mode == self.MODE_PROTECTION else "_处理后"
        
        try:
            # 检测是否为加密文档（非ZIP格式）
            if input_path.suffix.lower() == '.docx' and self.is_encrypted(str(input_path)):
                return ProcessResult(False, message="文档被密码加密，无法处理")
            
            if input_path.suffix.lower() == '.doc':
                # DOC文件必须转换为DOCX才能处理水印
                print(f"  [INFO] 检测到.doc格式，正在转换为.docx...")
                docx_path = self._convert_doc_to_docx(str(input_path))
                
                if docx_path:
                    # 转换成功，处理docx文件
                    output_path = output_path or self.get_default_output(docx_path, suffix)
                    shutil.copy2(docx_path, output_path)
                    success = self._process_docx(output_path)
                    
                    # 清理临时转换文件
                    try:
                        if docx_path != str(input_path):
                            os.remove(docx_path)
                    except:
                        pass
                    
                    if success:
                        msg = f"DOC转换为DOCX并处理成功 (移除{self.stats['removed']}项)"
                        return ProcessResult(True, str(output_path), msg, self.stats['removed'])
                    else:
                        return ProcessResult(False, message="DOCX处理失败")
                else:
                    return ProcessResult(False, message="DOC转换失败，请安装Microsoft Word")
            else:
                # 直接处理.docx文件
                output_path = output_path or self.get_default_output(str(input_path), suffix)
                shutil.copy2(input_path, output_path)
                success = self._process_docx(output_path)
                
                if success:
                    msg = f"Word文档处理成功 (移除{self.stats['removed']}项)"
                    return ProcessResult(True, str(output_path), msg, self.stats['removed'])
                else:
                    return ProcessResult(False, message="Word文档处理失败")
                
        except Exception as e:
            return ProcessResult(False, message=f"处理失败: {str(e)}")

    
    def _process_doc_direct(self, file_path: str) -> bool:
        """直接处理.doc文件"""
        try:
            from docx import Document
            doc = Document(file_path)
            doc.save(file_path)
            return True
        except:
            return True  # 对于.doc文件，复制即可
    
    def _process_docx(self, file_path: str) -> bool:
        """处理.docx文件"""
        if not zipfile.is_zipfile(file_path):
            return False
        
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            with zipfile.ZipFile(file_path, 'r') as zf:
                zf.extractall(temp_path)
            
            modified = False
            
            # 1. 移除保护
            if self.mode in (self.MODE_PROTECTION, self.MODE_ALL):
                settings_xml = temp_path / "word" / "settings.xml"
                if settings_xml.exists():
                    with open(settings_xml, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    patterns = [
                        r'<w:documentProtection[^>]*/?>',
                        r'<w:writeProtection[^>]*/?>',
                    ]
                    original = content
                    for p in patterns:
                        content = re.sub(p, '', content, flags=re.IGNORECASE | re.DOTALL)
                    
                    if content != original:
                        with open(settings_xml, 'w', encoding='utf-8') as f:
                            f.write(content)
                        modified = True
                        self.stats['removed'] += 1
            
            # 2. 移除水印
            if self.mode in (self.MODE_WATERMARK, self.MODE_ALL):
                # 处理页眉中的VML水印
                for hf in temp_path.glob("word/header*.xml"):
                    with open(hf, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    original = content
                    watermark_patterns = [
                        r'<w:pict[^>]*>.*?type="#_x0000_t136".*?</w:pict>',
                        r'<w:pict[^>]*>.*?rotation:.*?</w:pict>',
                        r'<w:pict[^>]*>.*?PowerPlusWaterMarkObject.*?</w:pict>',
                    ]
                    for p in watermark_patterns:
                        content = re.sub(p, '', content, flags=re.IGNORECASE | re.DOTALL)
                    
                    if content != original:
                        with open(hf, 'w', encoding='utf-8') as f:
                            f.write(content)
                        modified = True
                        self.stats['removed'] += 1
                
                # 处理背景水印
                doc_xml = temp_path / "word" / "document.xml"
                if doc_xml.exists():
                    with open(doc_xml, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    original = content
                    content = re.sub(r'<w:background[^>]*>.*?</w:background>', '', content, flags=re.IGNORECASE | re.DOTALL)
                    content = re.sub(r'<w:background[^>]*/>', '', content, flags=re.IGNORECASE)
                    
                    if content != original:
                        with open(doc_xml, 'w', encoding='utf-8') as f:
                            f.write(content)
                        modified = True
                        self.stats['removed'] += 1
            
            # 重新打包
            if modified:
                with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for root, dirs, files in os.walk(temp_path):
                        for file in files:
                            fp = os.path.join(root, file)
                            zf.write(fp, os.path.relpath(fp, temp_path))
            
            return True
    
    def _convert_doc_to_docx(self, doc_path: str) -> Optional[str]:
        """将.doc转换为.docx"""
        doc_path = Path(doc_path)
        docx_path = doc_path.with_suffix('.docx')
        
        # 尝试使用pywin32
        if sys.platform.startswith('win'):
            try:
                import win32com.client
                import pythoncom
                pythoncom.CoInitialize()
                
                word = win32com.client.Dispatch('Word.Application')
                word.Visible = False
                doc = word.Documents.Open(str(doc_path.resolve()))
                doc.SaveAs2(str(docx_path.resolve()), FileFormat=16)
                doc.Close()
                word.Quit()
                pythoncom.CoUninitialize()
                
                if docx_path.exists():
                    return str(docx_path)
            except:
                pass
        
        # 尝试LibreOffice
        try:
            subprocess.run([
                'soffice', '--headless', '--convert-to', 'docx',
                '--outdir', str(doc_path.parent), str(doc_path)
            ], capture_output=True, timeout=30)
            if docx_path.exists():
                return str(docx_path)
        except:
            pass
        
        return None
