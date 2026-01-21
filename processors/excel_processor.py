#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel处理器 - 移除工作表保护和背景水印
第一性原理：.xlsx = ZIP压缩包，保护信息 = XML标签
"""

import os
import re
import shutil
import zipfile
import tempfile
import subprocess
import sys
from pathlib import Path
from typing import Set, Optional
from .base import DocumentProcessor, ProcessResult


class ExcelProcessor(DocumentProcessor):
    """Excel文档处理器"""
    
    SUPPORTED_EXTENSIONS: Set[str] = {'.xls', '.xlsx'}
    
    def __init__(self, preview: bool = False, keywords: list = None,
                 preserve_format: bool = True):
        super().__init__(preview, keywords)
        self.preserve_format = preserve_format
    
    def get_description(self) -> str:
        return "Excel文档解锁 (保护移除)"
    
    @staticmethod
    def is_encrypted(file_path: str) -> bool:
        """
        检测文档是否被密码加密
        第一性原理：OOXML格式(.xlsx)本质是ZIP，如果不是ZIP格式说明被OLE加密
        """
        return not zipfile.is_zipfile(file_path)
    
    def process(self, input_path: str, output_path: str = None,
                progress_callback=None) -> ProcessResult:
        """处理Excel文档"""
        input_path = Path(input_path)
        output_path = output_path or self.get_default_output(str(input_path), "_unlocked")
        
        try:
            # 检测是否为加密文档（非ZIP格式）
            if input_path.suffix.lower() == '.xlsx' and self.is_encrypted(str(input_path)):
                return ProcessResult(False, message="文档被密码加密，无法处理")
            
            shutil.copy2(input_path, output_path)
            
            if input_path.suffix.lower() == '.xls':
                # 转换为xlsx后处理
                xlsx_path = self._convert_xls_to_xlsx(str(input_path))
                if xlsx_path:
                    success = self._process_xlsx(xlsx_path)
                    output_path = xlsx_path
                else:
                    # 转换失败，尝试二进制方法，或复制原文件
                    success = self._process_xls_binary(output_path)
                    if not success:
                        # 直接复制原文件
                        return ProcessResult(True, str(output_path),
                                           "XLS格式无法转换(需Excel/LibreOffice)，已复制原文件", 0)
            else:
                success = self._process_xlsx(output_path)
            
            if success:
                return ProcessResult(True, str(output_path), 
                                   f"Excel文档解锁成功", self.stats['removed'])
            else:
                return ProcessResult(False, message="Excel文档处理失败")
                
        except Exception as e:
            return ProcessResult(False, message=f"处理失败: {str(e)}")
    
    def _process_xlsx(self, file_path: str) -> bool:
        """处理.xlsx文件"""
        if not zipfile.is_zipfile(file_path):
            return False
        
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            with zipfile.ZipFile(file_path, 'r') as zf:
                zf.extractall(temp_path)
            
            modified = False
            
            # 1. 移除工作簿保护
            workbook_xml = temp_path / "xl" / "workbook.xml"
            if workbook_xml.exists():
                with open(workbook_xml, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                original = content
                patterns = [
                    r'<workbookProtection[^>]*/?>',
                    r'<workbookProtection[^>]*>.*?</workbookProtection>',
                ]
                for p in patterns:
                    content = re.sub(p, '', content, flags=re.IGNORECASE | re.DOTALL)
                
                if content != original:
                    with open(workbook_xml, 'w', encoding='utf-8') as f:
                        f.write(content)
                    modified = True
                    self.stats['removed'] += 1
            
            # 2. 移除工作表保护
            for sheet in temp_path.glob("xl/worksheets/sheet*.xml"):
                with open(sheet, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                original = content
                patterns = [
                    r'<sheetProtection[^>]*/?>',
                    r'<sheetProtection[^>]*>.*?</sheetProtection>',
                    r'<picture[^>]*/?>',  # 背景图片
                ]
                for p in patterns:
                    content = re.sub(p, '', content, flags=re.IGNORECASE | re.DOTALL)
                
                if content != original:
                    with open(sheet, 'w', encoding='utf-8') as f:
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
    
    def _process_xls_binary(self, file_path: str) -> bool:
        """二进制方式处理.xls文件"""
        try:
            with open(file_path, 'rb') as f:
                content = f.read()
            
            patterns = [
                b'\x12\x02\x01\x00',
                b'\x13\x02\x01\x00',
            ]
            
            original = content
            for p in patterns:
                content = content.replace(p, b'\x12\x02\x00\x00')
            
            if content != original:
                with open(file_path, 'wb') as f:
                    f.write(content)
                self.stats['removed'] += 1
            
            return True
        except:
            return False
    
    def _convert_xls_to_xlsx(self, xls_path: str) -> Optional[str]:
        """将.xls转换为.xlsx"""
        xls_path = Path(xls_path)
        xlsx_path = xls_path.with_suffix('.xlsx')
        
        # 尝试pywin32
        if sys.platform.startswith('win'):
            try:
                import win32com.client
                import pythoncom
                pythoncom.CoInitialize()
                
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                wb = excel.Workbooks.Open(str(xls_path.resolve()))
                wb.SaveAs(str(xlsx_path.resolve()), FileFormat=51)
                wb.Close()
                excel.Quit()
                pythoncom.CoUninitialize()
                
                if xlsx_path.exists():
                    return str(xlsx_path)
            except:
                pass
        
        # 尝试LibreOffice
        try:
            subprocess.run([
                'soffice', '--headless', '--convert-to', 'xlsx',
                '--outdir', str(xls_path.parent), str(xls_path)
            ], capture_output=True, timeout=60)
            if xlsx_path.exists():
                return str(xlsx_path)
        except:
            pass
        
        return None
