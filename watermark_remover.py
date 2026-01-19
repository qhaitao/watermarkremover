#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
通用水印去除器 v2.0 - 支持 PDF 和 PPTX
=============================================================================

自动识别文件类型并应用对应的水印去除策略:
    - PDF: 基于变换矩阵旋转检测
    - PPTX: 基于OOXML结构检测

用法:
    python watermark_remover.py input.pdf                # 自动检测类型
    python watermark_remover.py input.pptx -o output.pptx
    python watermark_remover.py input.pdf -k "机密"      # 关键词匹配
    python watermark_remover.py input.pptx --preview     # 预览模式

依赖:
    - PDF: pip install pikepdf
    - PPTX: 无需额外依赖

Author: Linus Torvalds Style
"""

import argparse
import sys
import os
from pathlib import Path
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import List, Set, Optional, Dict, Any

# ============================================================================
#                           通用配置
# ============================================================================

@dataclass
class WatermarkConfig:
    """水印检测配置"""
    text_keywords: List[str] = field(default_factory=list)
    preview: bool = False
    
    # PDF专用
    rotation_threshold: float = 0.1
    angle_min: float = 5.0
    angle_max: float = 85.0
    detect_color: bool = False
    
    # PPTX专用
    name_patterns: List[str] = field(default_factory=lambda: ['艺术字', 'WordArt', '水印'])
    detect_wordart: bool = True
    detect_rotated: bool = True
    alpha_threshold: int = 80000

# ============================================================================
#                           抽象基类
# ============================================================================

class WatermarkRemover(ABC):
    """水印去除器基类"""
    
    def __init__(self, config: WatermarkConfig):
        self.config = config
        self.stats = {'pages': 0, 'removed': 0, 'scanned': 0}
    
    @abstractmethod
    def process(self, input_path: str, output_path: str) -> Dict[str, Any]:
        """处理文件并返回统计信息"""
        pass
    
    def print_header(self, input_name: str, output_name: str, file_type: str):
        mode_str = '[预览模式]' if self.config.preview else '[处理模式]'
        print(f"\n{'='*60}")
        print(f"通用水印去除器 v2.0 - {file_type} {mode_str}")
        print(f"{'='*60}")
        print(f"输入: {input_name}")
        print(f"输出: {output_name}")
        if self.config.text_keywords:
            print(f"关键词: {', '.join(self.config.text_keywords)}")
        print(f"{'='*60}\n")
    
    def print_result(self, patterns: Set[str], output_path: str):
        action = '处理' if not self.config.preview else '扫描'
        removed_word = '移除' if not self.config.preview else '发现'
        
        print(f"\n{'='*60}")
        print(f"{action}完成!")
        print(f"{'='*60}")
        print(f"  页面总数: {self.stats['pages']}")
        print(f"  扫描元素: {self.stats['scanned']}")
        print(f"  {removed_word}水印: {self.stats['removed']}")
        
        if patterns:
            print(f"  检测模式: {', '.join(patterns)}")
        
        if not self.config.preview and self.stats['removed'] > 0:
            print(f"\n  输出文件: {output_path}")
        
        print(f"{'='*60}\n")

# ============================================================================
#                           PDF 水印去除器
# ============================================================================

class PDFWatermarkRemover(WatermarkRemover):
    """PDF水印去除器"""
    
    def __init__(self, config: WatermarkConfig):
        super().__init__(config)
        self.detected_patterns: Set[str] = set()
    
    def process(self, input_path: str, output_path: str) -> Dict[str, Any]:
        import re
        import math
        
        try:
            import pikepdf
        except ImportError:
            print("错误: PDF处理需要 pikepdf 库")
            print("运行: pip install pikepdf")
            sys.exit(1)
        
        input_path = Path(input_path)
        output_path = Path(output_path)
        
        self.print_header(input_path.name, output_path.name, "PDF")
        
        pdf = pikepdf.open(input_path)
        self.stats['pages'] = len(pdf.pages)
        
        print(f"[1/2] 扫描水印... ({self.stats['pages']}页)")
        
        for page_num, page in enumerate(pdf.pages):
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
                        
                        if abs(b) > self.config.rotation_threshold or abs(c) > self.config.rotation_threshold:
                            angle = abs(math.degrees(math.atan2(b, 1)))
                            if self.config.angle_min <= angle <= self.config.angle_max:
                                page_removed += 1
                                self.detected_patterns.add(f'旋转({angle:.1f}°)')
                                if not self.config.preview:
                                    return ''
                    
                    # 检测关键词
                    for kw in self.config.text_keywords:
                        if kw in block:
                            page_removed += 1
                            self.detected_patterns.add(f'关键词({kw})')
                            if not self.config.preview:
                                return ''
                    
                    return block
                
                new_text = pattern.sub(filter_watermarks, text)
                
                if page_removed > 0 and not self.config.preview:
                    stream.write(new_text.encode('latin1'))
            
            if page_removed > 0:
                self.stats['removed'] += page_removed
                print(f"  第{page_num + 1}页: 发现 {page_removed} 个水印")
        
        if not self.config.preview and self.stats['removed'] > 0:
            print(f"\n[2/2] 保存文件...")
            pdf.save(output_path)
        
        pdf.close()
        self.print_result(self.detected_patterns, output_path)
        
        return {'pages': self.stats['pages'], 'removed': self.stats['removed']}

# ============================================================================
#                           PPTX 水印去除器
# ============================================================================

class PPTXWatermarkRemover(WatermarkRemover):
    """PPTX水印去除器"""
    
    def __init__(self, config: WatermarkConfig):
        super().__init__(config)
        self.detected_patterns: Set[str] = set()
    
    def process(self, input_path: str, output_path: str) -> Dict[str, Any]:
        import zipfile
        import shutil
        import re
        from xml.etree import ElementTree as ET
        
        NAMESPACES = {
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        }
        
        for prefix, uri in NAMESPACES.items():
            ET.register_namespace(prefix, uri)
        
        input_path = Path(input_path)
        output_path = Path(output_path)
        
        self.print_header(input_path.name, output_path.name, "PPTX")
        
        temp_dir = input_path.parent / '_pptx_processing'
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
        temp_dir.mkdir()
        
        try:
            print("[1/3] 解压PPTX...")
            with zipfile.ZipFile(input_path, 'r') as zf:
                zf.extractall(temp_dir)
            
            print("\n[2/3] 扫描水印...")
            slides_dir = temp_dir / 'ppt' / 'slides'
            
            if slides_dir.exists():
                slide_files = sorted(slides_dir.glob('slide*.xml'),
                                   key=lambda x: int(re.search(r'\d+', x.stem).group()))
                self.stats['pages'] = len(slide_files)
                
                for slide_file in slide_files:
                    root = ET.parse(slide_file).getroot()
                    spTree = root.find('.//p:spTree', NAMESPACES)
                    
                    if spTree is None:
                        continue
                    
                    watermarks = []
                    
                    for sp in spTree.findall('p:sp', NAMESPACES):
                        self.stats['scanned'] += 1
                        
                        # 检测名称
                        cNvPr = sp.find('.//p:cNvPr', NAMESPACES)
                        if cNvPr is not None:
                            name = cNvPr.get('name', '')
                            for pattern in self.config.name_patterns:
                                if pattern.lower() in name.lower():
                                    watermarks.append(sp)
                                    self.detected_patterns.add(f'名称含"{pattern}"')
                                    break
                            else:
                                # 检测WordArt + 透明
                                if self.config.detect_wordart:
                                    bodyPr = sp.find('.//a:bodyPr', NAMESPACES)
                                    if bodyPr is not None and bodyPr.get('fromWordArt') == '1':
                                        alpha = sp.find('.//a:alpha', NAMESPACES)
                                        if alpha is not None:
                                            if int(alpha.get('val', '100000')) < self.config.alpha_threshold:
                                                watermarks.append(sp)
                                                self.detected_patterns.add('WordArt+透明')
                    
                    if watermarks:
                        self.stats['removed'] += len(watermarks)
                        print(f"  {slide_file.name}: 发现 {len(watermarks)} 个水印")
                        
                        if not self.config.preview:
                            for wm in watermarks:
                                spTree.remove(wm)
                            ET.ElementTree(root).write(slide_file, encoding='UTF-8', xml_declaration=True)
            
            if not self.config.preview and self.stats['removed'] > 0:
                print(f"\n[3/3] 重新打包...")
                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            file_path = Path(root) / file
                            zf.write(file_path, file_path.relative_to(temp_dir))
            
            self.print_result(self.detected_patterns, output_path)
            
        finally:
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
        
        return {'pages': self.stats['pages'], 'removed': self.stats['removed']}

# ============================================================================
#                           工厂函数
# ============================================================================

def get_remover(file_path: str, config: WatermarkConfig) -> WatermarkRemover:
    """根据文件类型返回对应的处理器"""
    ext = Path(file_path).suffix.lower()
    
    if ext == '.pdf':
        return PDFWatermarkRemover(config)
    elif ext in ['.pptx', '.ppt']:
        return PPTXWatermarkRemover(config)
    else:
        print(f"错误: 不支持的文件类型 '{ext}'")
        print("支持的类型: .pdf, .pptx")
        sys.exit(1)

def get_default_output(input_path: str) -> str:
    """生成默认输出文件名"""
    p = Path(input_path)
    return str(p.parent / f"{p.stem}_无水印{p.suffix}")

# ============================================================================
#                           批量处理函数
# ============================================================================

def process_single_file(input_path: str, config: WatermarkConfig) -> dict:
    """处理单个文件"""
    output_path = get_default_output(input_path)
    
    try:
        remover = get_remover(input_path, config)
        result = remover.process(input_path, output_path)
        return {'success': True, 'file': input_path, 'removed': result.get('removed', 0)}
    except Exception as e:
        print(f"\n处理 {input_path} 时发生错误: {str(e)}")
        return {'success': False, 'file': input_path, 'error': str(e)}

def process_batch(file_list: list, config: WatermarkConfig) -> list:
    """批量处理多个文件"""
    results = []
    total = len(file_list)
    
    print(f"\n{'='*60}")
    print(f"批量处理模式 - 共 {total} 个文件")
    print(f"{'='*60}\n")
    
    for i, file_path in enumerate(file_list, 1):
        print(f"\n[{i}/{total}] 处理: {Path(file_path).name}")
        print("-" * 40)
        result = process_single_file(file_path, config)
        results.append(result)
    
    # 汇总结果
    success_count = sum(1 for r in results if r['success'])
    total_removed = sum(r.get('removed', 0) for r in results if r['success'])
    
    print(f"\n\n{'='*60}")
    print(f"批量处理完成!")
    print(f"{'='*60}")
    print(f"  处理文件: {total}")
    print(f"  成功: {success_count}")
    print(f"  失败: {total - success_count}")
    print(f"  总计移除水印: {total_removed}")
    print(f"{'='*60}\n")
    
    return results

# ============================================================================
#                           命令行接口
# ============================================================================

def main():
    # 检测运行模式
    # 1. 无参数: 交互模式(弹窗选择)
    # 2. 参数是文件列表: 拖拽模式(批量处理)
    # 3. 带选项参数: 命令行模式
    
    has_options = any(arg.startswith('-') for arg in sys.argv[1:])
    file_args = [arg for arg in sys.argv[1:] if not arg.startswith('-') and os.path.exists(arg)]
    
    # 判断模式
    interactive_mode = len(sys.argv) == 1
    drag_drop_mode = len(file_args) > 0 and not has_options
    
    if interactive_mode:
        # -------------------- 交互模式 --------------------
        print(f"\n{'='*60}")
        print(f"通用水印去除器 v3.0 - 交互模式")
        print(f"{'='*60}")
        print("提示: 可拖拽文件到本程序图标上进行批量处理")
        print("未检测到文件，请选择...")
        
        try:
            import tkinter as tk
            from tkinter import filedialog
        except ImportError:
            print("错误: 缺少 tkinter 模块")
            input("按回车键退出...")
            sys.exit(1)
            
        root = tk.Tk()
        root.withdraw()
        
        # 支持多选
        file_paths = filedialog.askopenfilenames(
            title="选择要去除水印的文件 (可多选)",
            filetypes=[
                ("PDF & PPTX", "*.pdf;*.pptx;*.ppt"),
                ("PDF Files", "*.pdf"),
                ("PPTX Files", "*.pptx;*.ppt"),
                ("All Files", "*.*")
            ]
        )
        
        if not file_paths:
            print("未选择文件。")
            input("按回车键退出...")
            sys.exit(0)
        
        file_list = list(file_paths)
        print(f"已选择 {len(file_list)} 个文件")
        
        config = WatermarkConfig()  # 默认配置
        
        try:
            if len(file_list) == 1:
                process_single_file(file_list[0], config)
            else:
                process_batch(file_list, config)
        except Exception as e:
            print(f"\n发生错误: {str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            print("\n处理结束.")
            input("按回车键退出...")
    
    elif drag_drop_mode:
        # -------------------- 拖拽模式 --------------------
        print(f"\n{'='*60}")
        print(f"通用水印去除器 v3.0 - 拖拽模式")
        print(f"{'='*60}")
        
        # 过滤支持的文件类型
        supported_ext = {'.pdf', '.pptx', '.ppt'}
        valid_files = [f for f in file_args if Path(f).suffix.lower() in supported_ext]
        
        if not valid_files:
            print("错误: 没有找到支持的文件类型 (PDF/PPTX)")
            input("按回车键退出...")
            sys.exit(1)
        
        skipped = len(file_args) - len(valid_files)
        if skipped > 0:
            print(f"跳过 {skipped} 个不支持的文件")
        
        config = WatermarkConfig()  # 默认配置
        
        try:
            if len(valid_files) == 1:
                process_single_file(valid_files[0], config)
            else:
                process_batch(valid_files, config)
        except Exception as e:
            print(f"\n发生错误: {str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            print("\n处理结束.")
            input("按回车键退出...")
    
    else:
        # -------------------- 命令行模式 --------------------
        parser = argparse.ArgumentParser(
            description='通用水印去除器 v3.0 - 支持 PDF 和 PPTX',
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog="""
示例:
  %(prog)s document.pdf                    # 处理单个文件
  %(prog)s *.pdf                           # 批量处理
  %(prog)s file1.pdf file2.pptx            # 处理多个文件
  %(prog)s input.pdf -o output.pdf         # 指定输出文件
  %(prog)s input.pptx -k "机密"            # 按关键词匹配
  %(prog)s input.pdf --preview             # 只扫描不删除

拖拽模式:
  将文件拖拽到程序图标上即可批量处理
            """
        )
        
        parser.add_argument('input', nargs='+', help='输入文件路径 (支持多个)')
        parser.add_argument('-o', '--output', help='输出文件路径 (仅单文件时有效)')
        parser.add_argument('-k', '--keyword', action='append', default=[],
                           help='水印关键词 (可多次使用)')
        parser.add_argument('--preview', action='store_true',
                           help='预览模式: 只扫描不删除')
        parser.add_argument('--angle-min', type=float, default=5.0,
                           help='[PDF] 最小检测角度 (默认5°)')
        parser.add_argument('--angle-max', type=float, default=85.0,
                           help='[PDF] 最大检测角度 (默认85°)')
        parser.add_argument('--no-wordart', action='store_true',
                           help='[PPTX] 禁用WordArt检测')
        parser.add_argument('--alpha', type=int, default=80000,
                           help='[PPTX] 透明度阈值 (0-100000)')
        
        args = parser.parse_args()
        
        # 验证文件
        valid_files = []
        for f in args.input:
            if os.path.exists(f):
                valid_files.append(f)
            else:
                print(f"警告: 文件不存在 - {f}")
        
        if not valid_files:
            print("错误: 没有有效的输入文件")
            sys.exit(1)
        
        config = WatermarkConfig(
            text_keywords=args.keyword,
            preview=args.preview,
            angle_min=args.angle_min,
            angle_max=args.angle_max,
            detect_wordart=not args.no_wordart,
            alpha_threshold=args.alpha
        )
        
        if len(valid_files) == 1:
            output_path = args.output or get_default_output(valid_files[0])
            remover = get_remover(valid_files[0], config)
            result = remover.process(valid_files[0], output_path)
            sys.exit(0 if result.get('removed', 0) > 0 or args.preview else 1)
        else:
            if args.output:
                print("警告: 批量模式下忽略 -o 参数")
            process_batch(valid_files, config)

if __name__ == '__main__':
    main()
