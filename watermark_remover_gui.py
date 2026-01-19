#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
=============================================================================
通用水印去除器 v3.0 - GUI版
=============================================================================

带图形界面的水印去除工具，支持:
    - 窗口内拖拽文件上传
    - 点击选择文件
    - 批量处理
    - 实时进度显示

依赖: pip install pikepdf tkinterdnd2

Author: Linus Torvalds Style
"""

import sys
import os
import threading
import queue
from pathlib import Path
from dataclasses import dataclass, field
from typing import List, Set, Dict, Any
from abc import ABC, abstractmethod

# ============================================================================
#                           GUI界面
# ============================================================================

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 尝试导入拖拽支持
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

# ============================================================================
#                           水印配置和处理器 (复用核心逻辑)
# ============================================================================

@dataclass
class WatermarkConfig:
    """水印检测配置"""
    text_keywords: List[str] = field(default_factory=list)
    preview: bool = False
    rotation_threshold: float = 0.1
    angle_min: float = 5.0
    angle_max: float = 85.0
    detect_color: bool = False
    name_patterns: List[str] = field(default_factory=lambda: ['艺术字', 'WordArt', '水印'])
    detect_wordart: bool = True
    detect_rotated: bool = True
    alpha_threshold: int = 80000

class WatermarkRemover(ABC):
    def __init__(self, config: WatermarkConfig):
        self.config = config
        self.stats = {'pages': 0, 'removed': 0, 'scanned': 0}
    
    @abstractmethod
    def process(self, input_path: str, output_path: str, progress_callback=None) -> Dict[str, Any]:
        pass

class PDFWatermarkRemover(WatermarkRemover):
    def process(self, input_path: str, output_path: str, progress_callback=None) -> Dict[str, Any]:
        import re
        import math
        import pikepdf
        
        pdf = pikepdf.open(input_path)
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
                                if not self.config.preview:
                                    return ''
                    return block
                
                new_text = pattern.sub(filter_watermarks, text)
                if page_removed > 0 and not self.config.preview:
                    stream.write(new_text.encode('latin1'))
            
            self.stats['removed'] += page_removed
        
        if not self.config.preview and self.stats['removed'] > 0:
            pdf.save(output_path)
        pdf.close()
        
        return {'pages': self.stats['pages'], 'removed': self.stats['removed']}

class PPTXWatermarkRemover(WatermarkRemover):
    def process(self, input_path: str, output_path: str, progress_callback=None) -> Dict[str, Any]:
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
        
        temp_dir = Path(input_path).parent / '_pptx_temp_gui'
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
        temp_dir.mkdir()
        
        try:
            with zipfile.ZipFile(input_path, 'r') as zf:
                zf.extractall(temp_dir)
            
            slides_dir = temp_dir / 'ppt' / 'slides'
            if slides_dir.exists():
                slide_files = sorted(slides_dir.glob('slide*.xml'),
                                   key=lambda x: int(re.search(r'\d+', x.stem).group()))
                self.stats['pages'] = len(slide_files)
                
                for i, slide_file in enumerate(slide_files):
                    if progress_callback:
                        progress_callback(i + 1, self.stats['pages'], f"处理 {slide_file.name}")
                    
                    root = ET.parse(slide_file).getroot()
                    spTree = root.find('.//p:spTree', NAMESPACES)
                    if spTree is None:
                        continue
                    
                    watermarks = []
                    for sp in spTree.findall('p:sp', NAMESPACES):
                        self.stats['scanned'] += 1
                        cNvPr = sp.find('.//p:cNvPr', NAMESPACES)
                        if cNvPr is not None:
                            name = cNvPr.get('name', '')
                            for pattern in self.config.name_patterns:
                                if pattern.lower() in name.lower():
                                    watermarks.append(sp)
                                    break
                    
                    if watermarks:
                        self.stats['removed'] += len(watermarks)
                        if not self.config.preview:
                            for wm in watermarks:
                                spTree.remove(wm)
                            ET.ElementTree(root).write(slide_file, encoding='UTF-8', xml_declaration=True)
            
            if not self.config.preview and self.stats['removed'] > 0:
                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            file_path = Path(root) / file
                            zf.write(file_path, file_path.relative_to(temp_dir))
        finally:
            if temp_dir.exists():
                shutil.rmtree(temp_dir)
        
        return {'pages': self.stats['pages'], 'removed': self.stats['removed']}

def get_remover(file_path: str, config: WatermarkConfig) -> WatermarkRemover:
    ext = Path(file_path).suffix.lower()
    if ext == '.pdf':
        return PDFWatermarkRemover(config)
    elif ext in ['.pptx', '.ppt']:
        return PPTXWatermarkRemover(config)
    raise ValueError(f"不支持的文件类型: {ext}")

def get_default_output(input_path: str) -> str:
    p = Path(input_path)
    return str(p.parent / f"{p.stem}_无水印{p.suffix}")

# ============================================================================
#                           GUI应用
# ============================================================================

class WatermarkRemoverGUI:
    def __init__(self):
        # 创建主窗口
        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        
        self.root.title("水印去除器 v3.0")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # 文件列表
        self.file_list = []
        self.processing = False
        self.message_queue = queue.Queue()
        
        self.setup_ui()
        self.check_queue()
    
    def setup_ui(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="通用水印去除器 v3.0", font=('微软雅黑', 16, 'bold'))
        title_label.pack(pady=(0, 10))
        
        subtitle = ttk.Label(main_frame, text="支持 PDF 和 PPTX 文件", font=('微软雅黑', 10))
        subtitle.pack()
        
        # 拖拽区域
        drop_frame = ttk.LabelFrame(main_frame, text="文件列表", padding="10")
        drop_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 拖拽提示
        if HAS_DND:
            hint_text = "将文件拖拽到此处，或点击下方按钮选择"
        else:
            hint_text = "点击下方按钮选择文件 (安装 tkinterdnd2 可启用拖拽)"
        
        self.drop_label = ttk.Label(drop_frame, text=hint_text, 
                                    font=('微软雅黑', 10), foreground='gray')
        self.drop_label.pack(pady=5)
        
        # 文件列表框
        list_frame = ttk.Frame(drop_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox = tk.Listbox(list_frame, height=10, yscrollcommand=scrollbar.set,
                                  font=('Consolas', 9))
        self.listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)
        
        # 绑定拖拽
        if HAS_DND:
            self.listbox.drop_target_register(DND_FILES)
            self.listbox.dnd_bind('<<Drop>>', self.on_drop)
        
        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        self.btn_add = ttk.Button(btn_frame, text="添加文件", command=self.add_files)
        self.btn_add.pack(side=tk.LEFT, padx=5)
        
        self.btn_clear = ttk.Button(btn_frame, text="清空列表", command=self.clear_files)
        self.btn_clear.pack(side=tk.LEFT, padx=5)
        
        self.btn_remove = ttk.Button(btn_frame, text="移除选中", command=self.remove_selected)
        self.btn_remove.pack(side=tk.LEFT, padx=5)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress.pack(fill=tk.X, pady=5)
        
        # 状态标签
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var, font=('微软雅黑', 9))
        self.status_label.pack()
        
        # 处理按钮
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=10)
        
        self.btn_process = ttk.Button(action_frame, text="开始处理", command=self.start_processing)
        self.btn_process.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        self.btn_preview = ttk.Button(action_frame, text="预览模式", command=self.start_preview)
        self.btn_preview.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
    
    def on_drop(self, event):
        """处理拖拽事件"""
        files = self.root.tk.splitlist(event.data)
        self.add_files_to_list(files)
    
    def add_files(self):
        """打开文件选择对话框"""
        files = filedialog.askopenfilenames(
            title="选择文件",
            filetypes=[
                ("PDF & PPTX", "*.pdf;*.pptx;*.ppt"),
                ("PDF Files", "*.pdf"),
                ("PPTX Files", "*.pptx;*.ppt"),
                ("All Files", "*.*")
            ]
        )
        if files:
            self.add_files_to_list(files)
    
    def add_files_to_list(self, files):
        """添加文件到列表"""
        supported = {'.pdf', '.pptx', '.ppt'}
        added = 0
        for f in files:
            f = f.strip('{}')  # 处理Windows拖拽格式
            if Path(f).suffix.lower() in supported and f not in self.file_list:
                self.file_list.append(f)
                self.listbox.insert(tk.END, Path(f).name)
                added += 1
        
        if added > 0:
            self.status_var.set(f"已添加 {added} 个文件，共 {len(self.file_list)} 个")
    
    def clear_files(self):
        """清空文件列表"""
        self.file_list.clear()
        self.listbox.delete(0, tk.END)
        self.status_var.set("列表已清空")
    
    def remove_selected(self):
        """移除选中的文件"""
        selection = self.listbox.curselection()
        for i in reversed(selection):
            self.listbox.delete(i)
            del self.file_list[i]
        self.status_var.set(f"剩余 {len(self.file_list)} 个文件")
    
    def start_processing(self):
        """开始处理"""
        self._start_work(preview=False)
    
    def start_preview(self):
        """预览模式"""
        self._start_work(preview=True)
    
    def _start_work(self, preview=False):
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加文件")
            return
        
        if self.processing:
            return
        
        self.processing = True
        self.btn_process.config(state='disabled')
        self.btn_preview.config(state='disabled')
        
        thread = threading.Thread(target=self._process_files, args=(preview,))
        thread.daemon = True
        thread.start()
    
    def _process_files(self, preview):
        """后台处理文件"""
        config = WatermarkConfig(preview=preview)
        total_files = len(self.file_list)
        total_removed = 0
        success = 0
        
        for i, file_path in enumerate(self.file_list):
            try:
                self.message_queue.put(('status', f"[{i+1}/{total_files}] 处理: {Path(file_path).name}"))
                self.message_queue.put(('progress', (i / total_files) * 100))
                
                output_path = get_default_output(file_path)
                remover = get_remover(file_path, config)
                
                def progress_cb(current, total, msg):
                    pct = (i + current/total) / total_files * 100
                    self.message_queue.put(('progress', pct))
                    self.message_queue.put(('status', f"[{i+1}/{total_files}] {msg}"))
                
                result = remover.process(file_path, output_path, progress_cb)
                total_removed += result.get('removed', 0)
                success += 1
                
            except Exception as e:
                self.message_queue.put(('status', f"错误: {str(e)}"))
        
        # 完成
        self.message_queue.put(('progress', 100))
        action = "扫描" if preview else "处理"
        self.message_queue.put(('status', f"{action}完成! 成功{success}个，共移除{total_removed}个水印"))
        self.message_queue.put(('done', None))
        
        if not preview and total_removed > 0:
            self.message_queue.put(('msgbox', f"处理完成!\n\n成功: {success}/{total_files}\n移除水印: {total_removed}"))
    
    def check_queue(self):
        """检查消息队列"""
        try:
            while True:
                msg_type, data = self.message_queue.get_nowait()
                if msg_type == 'status':
                    self.status_var.set(data)
                elif msg_type == 'progress':
                    self.progress_var.set(data)
                elif msg_type == 'done':
                    self.processing = False
                    self.btn_process.config(state='normal')
                    self.btn_preview.config(state='normal')
                elif msg_type == 'msgbox':
                    messagebox.showinfo("完成", data)
        except queue.Empty:
            pass
        
        self.root.after(100, self.check_queue)
    
    def run(self):
        self.root.mainloop()

# ============================================================================
#                           主入口
# ============================================================================

def main():
    # 检查是否有命令行参数(拖拽文件到EXE)
    if len(sys.argv) > 1:
        # 命令行/拖拽模式 - 直接处理
        files = [f for f in sys.argv[1:] if os.path.exists(f)]
        if files:
            print("检测到拖拽文件，直接处理...")
            config = WatermarkConfig()
            for f in files:
                try:
                    print(f"\n处理: {f}")
                    output = get_default_output(f)
                    remover = get_remover(f, config)
                    result = remover.process(f, output)
                    print(f"完成! 移除 {result.get('removed', 0)} 个水印")
                except Exception as e:
                    print(f"错误: {e}")
            input("\n按回车键退出...")
            return
    
    # GUI模式
    app = WatermarkRemoverGUI()
    app.run()

if __name__ == '__main__':
    main()
