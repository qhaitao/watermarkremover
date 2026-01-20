#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PDF处理器 - 基于变换矩阵检测旋转文本水印 + 密码解密
"""

import re
import math
from pathlib import Path
from typing import Set
from .base import DocumentProcessor, ProcessResult

class PDFProcessor(DocumentProcessor):
    """PDF水印处理器"""
    
    SUPPORTED_EXTENSIONS: Set[str] = {'.pdf'}
    
    def __init__(self, preview: bool = False, keywords: list = None,
                 angle_min: float = 5.0, angle_max: float = 85.0,
                 rotation_threshold: float = 0.1):
        super().__init__(preview, keywords)
        self.angle_min = angle_min
        self.angle_max = angle_max
        self.rotation_threshold = rotation_threshold
        self.detected_patterns: Set[str] = set()
    
    def get_description(self) -> str:
        return "PDF水印去除 (旋转文本检测+密码解密)"
    
    def process(self, input_path: str, output_path: str = None,
                progress_callback=None) -> ProcessResult:
        """处理PDF文件"""
        try:
            import pikepdf
        except ImportError:
            return ProcessResult(False, message="需要安装 pikepdf: pip install pikepdf")
        
        input_path = Path(input_path)
        output_path = output_path or self.get_default_output(str(input_path), "_无水印")
        
        try:
            # 尝试打开PDF，先尝试无密码，再尝试空密码
            pdf = None
            decrypted = False
            
            try:
                pdf = pikepdf.open(input_path)
            except pikepdf.PasswordError:
                # 尝试空密码
                try:
                    pdf = pikepdf.open(input_path, password='')
                    decrypted = True
                    self.stats['removed'] += 1
                    print(f"  [OK] PDF密码解除成功!")
                except:
                    return ProcessResult(False, message="PDF需要密码，无法打开")
            
            self.stats['pages'] = len(pdf.pages)
            
            for page_num, page in enumerate(pdf.pages):
                if progress_callback:
                    progress_callback(page_num + 1, self.stats['pages'], f"处理第{page_num+1}页")
                
                if "/Contents" not in page:
                    continue
                
                contents = page["/Contents"]
                streams = list(contents) if isinstance(contents, pikepdf.Array) else [contents]
                page_removed = 0
                
                for stream in streams:
                    data = stream.read_bytes()
                    text = data.decode('latin1', errors='replace')
                    pattern = re.compile(r'(BT\s+.*?ET)', re.DOTALL)
                    
                    def filter_watermarks(match):
                        nonlocal page_removed
                        block = match.group(1)
                        self.stats['scanned'] += 1
                        
                        # 检测旋转
                        tm_pattern = re.compile(
                            r'([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+Tm'
                        )
                        tm_match = tm_pattern.search(block)
                        
                        if tm_match:
                            a, b, c, d, e, f = map(float, tm_match.groups())
                            if abs(b) > self.rotation_threshold or abs(c) > self.rotation_threshold:
                                angle = abs(math.degrees(math.atan2(b, 1)))
                                if self.angle_min <= angle <= self.angle_max:
                                    page_removed += 1
                                    self.detected_patterns.add(f'旋转({angle:.1f}°)')
                                    if not self.preview:
                                        return ''
                        
                        # 检测关键词
                        for kw in self.keywords:
                            if kw in block:
                                page_removed += 1
                                self.detected_patterns.add(f'关键词({kw})')
                                if not self.preview:
                                    return ''
                        
                        return block
                    
                    new_text = pattern.sub(filter_watermarks, text)
                    if page_removed > 0 and not self.preview:
                        stream.write(new_text.encode('latin1'))
                
                self.stats['removed'] += page_removed
            
            if not self.preview and self.stats['removed'] > 0:
                pdf.save(output_path)
            pdf.close()
            
            patterns_str = ', '.join(self.detected_patterns) if self.detected_patterns else '无'
            action = "扫描" if self.preview else "处理"
            msg = f"{action}完成, 移除{self.stats['removed']}个水印 [{patterns_str}]"
            
            return ProcessResult(
                success=True,
                output_path=str(output_path) if not self.preview else None,
                message=msg,
                removed_count=self.stats['removed'],
                page_count=self.stats['pages']
            )
            
        except Exception as e:
            return ProcessResult(False, message=f"处理失败: {str(e)}")
