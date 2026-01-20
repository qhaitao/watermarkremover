#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PPTX处理器 - 基于OOXML结构检测艺术字水印
"""

import os
import re
import zipfile
import shutil
from pathlib import Path
from typing import Set, List
from xml.etree import ElementTree as ET
from .base import DocumentProcessor, ProcessResult

class PPTXProcessor(DocumentProcessor):
    """PPTX水印处理器"""
    
    SUPPORTED_EXTENSIONS: Set[str] = {'.pptx', '.ppt'}
    
    NAMESPACES = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    }
    
    def __init__(self, preview: bool = False, keywords: list = None,
                 name_patterns: List[str] = None, detect_wordart: bool = True,
                 alpha_threshold: int = 80000):
        super().__init__(preview, keywords)
        self.name_patterns = name_patterns or ['艺术字', 'WordArt', '水印']
        self.detect_wordart = detect_wordart
        self.alpha_threshold = alpha_threshold
        self.detected_patterns: Set[str] = set()
        
        # 注册命名空间
        for prefix, uri in self.NAMESPACES.items():
            ET.register_namespace(prefix, uri)
    
    def get_description(self) -> str:
        return "PPTX水印去除 (艺术字检测)"
    
    def process(self, input_path: str, output_path: str = None,
                progress_callback=None) -> ProcessResult:
        """处理PPTX文件"""
        input_path = Path(input_path)
        output_path = output_path or self.get_default_output(str(input_path), "_无水印")
        
        temp_dir = input_path.parent / f'_pptx_temp_{input_path.stem}'
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
        temp_dir.mkdir()
        
        try:
            # 检测是否为ZIP格式（非加密的Office文档）
            if not zipfile.is_zipfile(input_path):
                return ProcessResult(False, message="文件格式错误或被加密，无法处理")
            
            # 解压
            with zipfile.ZipFile(input_path, 'r') as zf:
                zf.extractall(temp_dir)
            
            # 0. 移除演示文稿保护（presentation.xml）
            pres_xml = temp_dir / 'ppt' / 'presentation.xml'
            if pres_xml.exists():
                with open(pres_xml, 'r', encoding='utf-8') as f:
                    content = f.read()
                original = content
                # 移除修改保护标签
                content = re.sub(r'<p:modifyVerifier[^>]*/>', '', content)
                content = re.sub(r'<p:modifyVerifier[^>]*>.*?</p:modifyVerifier>', '', content, flags=re.DOTALL)
                if content != original:
                    with open(pres_xml, 'w', encoding='utf-8') as f:
                        f.write(content)
                    self.stats['removed'] += 1
                    self.detected_patterns.add('演示文稿保护')
            
            # 处理幻灯片
            slides_dir = temp_dir / 'ppt' / 'slides'
            if slides_dir.exists():
                slide_files = sorted(slides_dir.glob('slide*.xml'),
                                   key=lambda x: int(re.search(r'\d+', x.stem).group()))
                self.stats['pages'] = len(slide_files)
                
                for i, slide_file in enumerate(slide_files):
                    if progress_callback:
                        progress_callback(i + 1, self.stats['pages'], f"处理 {slide_file.name}")
                    
                    root = ET.parse(slide_file).getroot()
                    spTree = root.find('.//p:spTree', self.NAMESPACES)
                    if spTree is None:
                        continue
                    
                    watermarks = []
                    for sp in spTree.findall('p:sp', self.NAMESPACES):
                        self.stats['scanned'] += 1
                        
                        # 检测名称
                        cNvPr = sp.find('.//p:cNvPr', self.NAMESPACES)
                        if cNvPr is not None:
                            name = cNvPr.get('name', '')
                            for pattern in self.name_patterns:
                                if pattern.lower() in name.lower():
                                    watermarks.append(sp)
                                    self.detected_patterns.add(f'名称含"{pattern}"')
                                    break
                            else:
                                # 检测WordArt + 透明
                                if self.detect_wordart:
                                    bodyPr = sp.find('.//a:bodyPr', self.NAMESPACES)
                                    if bodyPr is not None and bodyPr.get('fromWordArt') == '1':
                                        alpha = sp.find('.//a:alpha', self.NAMESPACES)
                                        if alpha is not None:
                                            if int(alpha.get('val', '100000')) < self.alpha_threshold:
                                                watermarks.append(sp)
                                                self.detected_patterns.add('WordArt+透明')
                    
                    if watermarks:
                        self.stats['removed'] += len(watermarks)
                        if not self.preview:
                            for wm in watermarks:
                                spTree.remove(wm)
                            ET.ElementTree(root).write(slide_file, encoding='UTF-8', xml_declaration=True)
            
            # 重新打包
            if not self.preview and self.stats['removed'] > 0:
                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            file_path = Path(root) / file
                            zf.write(file_path, file_path.relative_to(temp_dir))
            
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
        finally:
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
