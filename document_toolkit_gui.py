#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
通用文档解锁工具 v2.0 - 经典浅色风格
设计理念: 传统Windows风格 + 清晰边框 + 专业稳重
"""

import sys
import os
import threading
import queue
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

from processors import PDFProcessor, PPTXProcessor, WordProcessor, ExcelProcessor

# ============================================================================
#                           设计系统 - 经典浅色主题
# ============================================================================

THEME = {
    # 背景
    'bg_main': '#e8e8e8',         # 主背景 - 浅灰
    'bg_white': '#ffffff',        # 卡片/输入区 - 纯白
    'bg_input': '#ffffff',        # 输入框背景
    
    # 文字
    'text_dark': '#000000',       # 主文字 - 黑色
    'text_title': '#1a3a6b',      # 标题 - 深蓝
    'text_link': '#0066cc',       # 链接色 - 蓝色
    'text_muted': '#666666',      # 次要文字
    
    # 边框
    'border': '#888888',          # 主边框
    'border_light': '#aaaaaa',    # 浅边框
    'border_dark': '#555555',     # 深边框
    
    # 状态
    'success': '#008000',         # 成功 - 绿色
    'error': '#cc0000',           # 错误 - 红色
    'warning': '#996600',         # 警告 - 橙色
    
    # 按钮
    'btn_bg': '#f0f0f0',          # 按钮背景
    'btn_hover': '#e0e0e0',       # 按钮悬停
    'btn_active': '#d0d0d0',      # 按钮按下
}

# ============================================================================
#                           多语言支持
# ============================================================================

LANGUAGES = {
    'zh': {
        'app_title': 'FileFree v2.0 - 文件自由',
        'main_title': 'FileFree',
        'file_select': '📁 文件选择',
        'drop_hint': '拖拽文件到这里',
        'drop_sub': '或点击下方按钮选择文件',
        'btn_select': '📂  选择文件',
        'btn_clear': '🗑  清空列表',
        'btn_unlock': '🔓 开始解锁',
        'file_list': '📋 文件列表',
        'col_filename': '文件名',
        'col_format': '格式',
        'col_size': '大小',
        'col_status': '状态',
        'log_section': '📝 处理日志',
        'btn_open_folder': '📁 打开输出文件夹',
        'btn_about': 'ℹ️ 关于',
        'status_ready': '就绪',
        'status_processing': '⏳ 处理中...',
        'status_done': '✅ 完成',
        'status_failed': '❌ 失败',
        'msg_started': '🔧 通用文档解锁工具已启动',
        'msg_formats': '支持格式: Word (.doc, .docx)、Excel (.xls, .xlsx)、PDF (.pdf)、PPT (.ppt, .pptx)',
        'msg_processing': '⚡ 开始处理...',
        'msg_complete': '🎉 完成! 成功 {}/{}',
        'msg_no_files': '请先添加文件',
        'msg_no_output': '请先处理文件',
        'about_title': '关于',
        'about_version': 'v2.0',
        'about_func': '🔧 功能:',
        'about_func_desc': 'Word/Excel保护解除并去水印、PDF/PPTX水印移除',
        'about_format': '📁 格式:',
        'about_format_desc': 'PDF, PPTX, PPT, DOC, DOCX, XLS, XLSX',
        'about_feature': '⚡ 特点:',
        'about_feature_desc': '批量处理、拖拽上传、保持原格式',
        'about_warning': '⚠️ 仅限合法用途，请勿用于未授权文档',
        'about_author': '作者: qin + AI Assistant',
        'about_ok': '确定',
        'lang_switch': '🌐 English',
    },
    'en': {
        'app_title': 'FileFree v2.0 - File Freedom',
        'main_title': 'FileFree',
        'file_select': '📁 File Selection',
        'drop_hint': 'Drop files here',
        'drop_sub': 'or click button below to select',
        'btn_select': '📂  Select Files',
        'btn_clear': '🗑  Clear List',
        'btn_unlock': '🔓 Unlock',
        'file_list': '📋 File List',
        'col_filename': 'Filename',
        'col_format': 'Format',
        'col_size': 'Size',
        'col_status': 'Status',
        'log_section': '📝 Process Log',
        'btn_open_folder': '📁 Open Output Folder',
        'btn_about': 'ℹ️ About',
        'status_ready': 'Ready',
        'status_processing': '⏳ Processing...',
        'status_done': '✅ Done',
        'status_failed': '❌ Failed',
        'msg_started': '🔧 Document Unlocker Started',
        'msg_formats': 'Formats: Word (.doc, .docx), Excel (.xls, .xlsx), PDF (.pdf), PPT (.ppt, .pptx)',
        'msg_processing': '⚡ Processing...',
        'msg_complete': '🎉 Complete! Success {}/{}',
        'msg_no_files': 'Please add files first',
        'msg_no_output': 'Please process files first',
        'about_title': 'About',
        'about_version': 'v2.0',
        'about_func': '🔧 Features:',
        'about_func_desc': 'Remove Word/Excel protection & watermarks, PDF/PPTX watermark removal',
        'about_format': '📁 Formats:',
        'about_format_desc': 'PDF, PPTX, PPT, DOC, DOCX, XLS, XLSX',
        'about_feature': '⚡ Highlights:',
        'about_feature_desc': 'Batch processing, Drag & Drop, Keep original format',
        'about_warning': '⚠️ For legal use only. Do not use on unauthorized documents.',
        'about_author': 'Author: qin + AI Assistant',
        'about_ok': 'OK',
        'lang_switch': '🌐 中文',
    }
}

# 当前语言
CURRENT_LANG = 'zh'

def t(key):
    """获取当前语言的文本"""
    return LANGUAGES[CURRENT_LANG].get(key, key)

# ============================================================================
#                           处理器映射
# ============================================================================

PROCESSOR_MAP = {
    '.pdf': PDFProcessor, '.pptx': PPTXProcessor, '.ppt': PPTXProcessor,
    '.docx': WordProcessor, '.doc': WordProcessor,
    '.xlsx': ExcelProcessor, '.xls': ExcelProcessor,
}

SUPPORTED_EXTENSIONS = set(PROCESSOR_MAP.keys())

# ============================================================================
#                           主应用
# ============================================================================

class DocumentUnlockerGUI:
    def __init__(self):
        global CURRENT_LANG
        
        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        
        self.root.title(t('app_title'))
        self.root.geometry("800x720")
        self.root.minsize(750, 650)
        self.root.configure(bg=THEME['bg_main'])
        self.root.resizable(True, True)
        
        self.file_list = []
        self.processing = False
        self.msg_queue = queue.Queue()
        self.password_queue = queue.Queue()
        self.output_dir = None
        
        self.setup_styles()
        self.setup_ui()
        self.log(t('msg_started'))
        self.log(t('msg_formats'))
        self.check_queue()
    
    def switch_language(self):
        """切换语言"""
        global CURRENT_LANG
        CURRENT_LANG = 'en' if CURRENT_LANG == 'zh' else 'zh'
        # 刷新UI
        self.root.title(t('app_title'))
        self.refresh_ui_text()
    
    def refresh_ui_text(self):
        """刷新界面文本"""
        # 更新标题
        self.title_label.config(text=t('main_title'))
        # 更新拖拽区
        self._draw_drop_zone()
        # 更新按钮
        self.btn_select.config(text=t('btn_select'))
        self.btn_clear.config(text=t('btn_clear'))
        self.btn_unlock.config(text=t('btn_unlock'))
        self.btn_open.config(text=t('btn_open_folder'))
        self.btn_about.config(text=t('btn_about'))
        self.btn_lang.config(text=t('lang_switch'))
        # 更新LabelFrame
        self.select_section.config(text=t('file_select'))
        self.list_section.config(text=t('file_list'))
        self.log_section.config(text=t('log_section'))
        # 更新表头
        self.tree.heading('filename', text=t('col_filename'))
        self.tree.heading('format', text=t('col_format'))
        self.tree.heading('size', text=t('col_size'))
        self.tree.heading('status', text=t('col_status'))
        # 更新状态
        if not self.processing:
            self.status_label.config(text=t('status_ready'))
    
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Treeview 样式
        style.configure('Classic.Treeview',
                       background=THEME['bg_white'],
                       foreground=THEME['text_dark'],
                       fieldbackground=THEME['bg_white'],
                       rowheight=24,
                       borderwidth=1)
        style.configure('Classic.Treeview.Heading',
                       background=THEME['bg_main'],
                       foreground=THEME['text_dark'],
                       font=('微软雅黑', 9),
                       relief='raised')
        style.map('Classic.Treeview', 
                 background=[('selected', '#0078d7')],
                 foreground=[('selected', 'white')])
        
        # 进度条样式
        style.configure('Classic.Horizontal.TProgressbar',
                       background='#0078d7',
                       troughcolor=THEME['bg_white'],
                       borderwidth=1)
    
    def setup_ui(self):
        main = tk.Frame(self.root, bg=THEME['bg_main'], padx=15, pady=10)
        main.pack(fill=tk.BOTH, expand=True)
        
        # ===== 标题 =====
        self.title_label = tk.Label(main, text=t('main_title'), 
                              fg=THEME['text_title'], bg=THEME['bg_main'],
                              font=('微软雅黑', 18, 'bold'))
        self.title_label.pack(pady=(0, 15))
        
        # ===== 文件选择区 =====
        self.select_section = tk.LabelFrame(main, text=t('file_select'), 
                                       fg=THEME['text_dark'], bg=THEME['bg_main'],
                                       font=('微软雅黑', 9))
        self.select_section.pack(fill=tk.X, pady=(0, 10))
        
        # 拖拽区域 - 使用Canvas绘制虚线边框
        drop_container = tk.Frame(self.select_section, bg=THEME['bg_white'],
                                  highlightbackground=THEME['border'],
                                  highlightthickness=1)
        drop_container.pack(fill=tk.X, padx=10, pady=10)
        
        self.drop_canvas = tk.Canvas(drop_container, height=80, bg=THEME['bg_white'], 
                                    highlightthickness=0)
        self.drop_canvas.pack(fill=tk.X, padx=3, pady=3)
        self._draw_drop_zone()
        self.drop_canvas.bind('<Configure>', lambda e: self._draw_drop_zone())
        self.drop_canvas.bind('<Button-1>', lambda e: self.add_files())
        self.drop_canvas.configure(cursor='hand2')
        
        if HAS_DND:
            self.drop_canvas.drop_target_register(DND_FILES)
            self.drop_canvas.dnd_bind('<<Drop>>', self.on_drop)
        
        # ===== 按钮行 =====
        btn_frame = tk.Frame(self.select_section, bg=THEME['bg_main'])
        btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        # 三个等宽按钮
        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)
        btn_frame.columnconfigure(2, weight=1)
        
        self.btn_select = self._create_classic_button(btn_frame, t('btn_select'), self.add_files)
        self.btn_select.grid(row=0, column=0, sticky='ew', padx=(0, 5))
        self.btn_clear = self._create_classic_button(btn_frame, t('btn_clear'), self.clear_files)
        self.btn_clear.grid(row=0, column=1, sticky='ew', padx=5)
        self.btn_unlock = self._create_classic_button(btn_frame, t('btn_unlock'), self.start_process)
        self.btn_unlock.grid(row=0, column=2, sticky='ew', padx=(5, 0))
        
        # ===== 文件列表区 =====
        self.list_section = tk.LabelFrame(main, text=t('file_list'), 
                                     fg=THEME['text_dark'], bg=THEME['bg_main'],
                                     font=('微软雅黑', 9))
        self.list_section.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        list_frame = tk.Frame(self.list_section, bg=THEME['bg_white'],
                             highlightbackground=THEME['border'],
                             highlightthickness=1)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ('filename', 'format', 'size', 'status')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings',
                                style='Classic.Treeview', height=8)
        
        self.tree.heading('filename', text=t('col_filename'))
        self.tree.heading('format', text=t('col_format'))
        self.tree.heading('size', text=t('col_size'))
        self.tree.heading('status', text=t('col_status'))
        
        self.tree.column('filename', width=320)
        self.tree.column('format', width=80, anchor='center')
        self.tree.column('size', width=80, anchor='center')
        self.tree.column('status', width=100, anchor='center')
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        if HAS_DND:
            self.tree.drop_target_register(DND_FILES)
            self.tree.dnd_bind('<<Drop>>', self.on_drop)
        
        # ===== 进度条 =====
        progress_frame = tk.Frame(main, bg=THEME['bg_main'])
        progress_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(progress_frame, variable=self.progress_var,
                                        maximum=100, style='Classic.Horizontal.TProgressbar')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.status_label = tk.Label(progress_frame, text=t('status_ready'), 
                                    fg=THEME['text_muted'], bg=THEME['bg_main'],
                                    font=('微软雅黑', 9), width=8)
        self.status_label.pack(side=tk.RIGHT, padx=(10, 0))
        
        # ===== 处理日志区 =====
        self.log_section = tk.LabelFrame(main, text=t('log_section'), 
                                    fg=THEME['text_dark'], bg=THEME['bg_main'],
                                    font=('微软雅黑', 9))
        self.log_section.pack(fill=tk.X, pady=(0, 10))
        
        log_frame = tk.Frame(self.log_section, bg=THEME['bg_white'],
                            highlightbackground=THEME['border'],
                            highlightthickness=1)
        log_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.log_text = tk.Text(log_frame, height=4, bg=THEME['bg_white'],
                               fg=THEME['text_dark'], font=('Consolas', 9),
                               bd=0, padx=8, pady=5)
        
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ===== 底部按钮 =====
        footer = tk.Frame(main, bg=THEME['bg_main'])
        footer.pack(fill=tk.X)
        
        footer.columnconfigure(0, weight=1)
        footer.columnconfigure(1, weight=1)
        footer.columnconfigure(2, weight=1)
        
        self.btn_open = self._create_classic_button(footer, t('btn_open_folder'), self.open_output_folder)
        self.btn_open.grid(row=0, column=0, sticky='ew', padx=(0, 5))
        self.btn_about = self._create_classic_button(footer, t('btn_about'), self.show_about)
        self.btn_about.grid(row=0, column=1, sticky='ew', padx=5)
        self.btn_lang = self._create_classic_button(footer, t('lang_switch'), self.switch_language)
        self.btn_lang.grid(row=0, column=2, sticky='ew', padx=(5, 0))
    
    def _draw_drop_zone(self):
        """绘制虚线边框拖拽区域"""
        self.drop_canvas.delete('all')
        w = self.drop_canvas.winfo_width() or 700
        h = 80
        
        # 虚线边框
        dash = (6, 4)
        self.drop_canvas.create_rectangle(5, 5, w-5, h-5, 
                                         outline=THEME['border_light'], 
                                         dash=dash, width=1)
        
        # 图标
        self.drop_canvas.create_text(w//2, 22, text="📂", 
                                    font=('Segoe UI', 14))
        # 主文字
        self.drop_canvas.create_text(w//2, 42, text=t('drop_hint'),
                                    fill=THEME['text_link'],
                                    font=('微软雅黑', 11))
        # 副文字
        self.drop_canvas.create_text(w//2, 60, text=t('drop_sub'),
                                    fill=THEME['text_muted'],
                                    font=('微软雅黑', 9))

    
    def _create_classic_button(self, parent, text, command):
        """创建经典Windows风格按钮"""
        btn = tk.Button(parent, text=text, command=command,
                       bg=THEME['btn_bg'], fg=THEME['text_dark'],
                       activebackground=THEME['btn_active'],
                       font=('微软雅黑', 9), bd=1, relief='raised',
                       padx=15, pady=5, cursor='hand2',
                       highlightthickness=0)
        
        btn.bind('<Enter>', lambda e: btn.config(bg=THEME['btn_hover']))
        btn.bind('<Leave>', lambda e: btn.config(bg=THEME['btn_bg']))
        
        return btn
    
    def log(self, msg):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {msg}\n")
        self.log_text.see(tk.END)
    
    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        self.add_files_list(files)
    
    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择文件",
            filetypes=[
                ("所有支持格式", "*.pdf;*.pptx;*.ppt;*.docx;*.doc;*.xlsx;*.xls"),
                ("PDF", "*.pdf"), ("PowerPoint", "*.pptx;*.ppt"),
                ("Word", "*.docx;*.doc"), ("Excel", "*.xlsx;*.xls"),
            ]
        )
        if files:
            self.add_files_list(files)
    
    def add_files_list(self, files):
        added = 0
        for f in files:
            f = f.strip('{}')
            if not os.path.exists(f):
                continue
            ext = Path(f).suffix.lower()
            if ext in SUPPORTED_EXTENSIONS and f not in self.file_list:
                self.file_list.append(f)
                name = Path(f).name
                size = self.format_size(os.path.getsize(f))
                self.tree.insert('', tk.END, values=(name, ext.upper()[1:], size, '等待中'))
                added += 1
        
        if added:
            self.log(f"✅ 已添加 {added} 个文件")
            self.status_label.config(text=f"{len(self.file_list)} 个文件")
    
    def format_size(self, size):
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024:
                return f"{size:.0f}{unit}"
            size /= 1024
        return f"{size:.1f}TB"
    
    def clear_files(self):
        self.file_list.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.progress_var.set(0)
        self.status_label.config(text="就绪")
        self.log("🗑 列表已清空")
    
    def start_process(self):
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加文件")
            return
        if self.processing:
            return
        
        self.processing = True
        self.log("⚡ 开始处理...")
        self.status_label.config(text="处理中...")
        
        thread = threading.Thread(target=self._process_thread, daemon=True)
        thread.start()
    
    def _process_thread(self):
        total = len(self.file_list)
        success_count = 0
        items = self.tree.get_children()
        
        for i, (item_id, fp) in enumerate(zip(items, self.file_list)):
            progress = ((i + 1) / total) * 100
            self.msg_queue.put(('progress', progress))
            self.msg_queue.put(('tree_update', (item_id, '⏳ 处理中...')))
            
            try:
                ext = Path(fp).suffix.lower()
                processor = PROCESSOR_MAP[ext](preview=False)
                result = processor.process(fp)
                
                if result.success:
                    success_count += 1
                    self.msg_queue.put(('tree_update', (item_id, '✅ 完成')))
                    self.msg_queue.put(('log', f"✅ {Path(fp).name}"))
                    if result.output_path:
                        self.output_dir = str(Path(result.output_path).parent)
                else:
                    self.msg_queue.put(('tree_update', (item_id, '❌ 失败')))
                    self.msg_queue.put(('log', f"❌ {Path(fp).name}: {result.message}"))
            except Exception as e:
                self.msg_queue.put(('tree_update', (item_id, '❌ 错误')))
                self.msg_queue.put(('log', f"❌ {Path(fp).name}: {str(e)}"))
        
        self.msg_queue.put(('log', f"🎉 完成! 成功 {success_count}/{total}"))
        self.msg_queue.put(('done', success_count))
    
    def check_queue(self):
        try:
            while True:
                msg_type, data = self.msg_queue.get_nowait()
                if msg_type == 'log':
                    self.log(data)
                elif msg_type == 'progress':
                    self.progress_var.set(data)
                elif msg_type == 'tree_update':
                    item_id, status = data
                    values = list(self.tree.item(item_id, 'values'))
                    values[3] = status
                    self.tree.item(item_id, values=values)
                elif msg_type == 'done':
                    self.processing = False
                    self.status_label.config(text=f"完成 ({data}/{len(self.file_list)})")
                    if data > 0:
                        messagebox.showinfo("完成", f"🎉 处理完成!\n成功: {data}/{len(self.file_list)}")
        except queue.Empty:
            pass
        self.root.after(100, self.check_queue)
    
    def open_output_folder(self):
        if self.output_dir and os.path.exists(self.output_dir):
            os.startfile(self.output_dir)
        else:
            messagebox.showinfo("提示", "请先处理文件")
    
    def show_about(self):
        about_win = tk.Toplevel(self.root)
        about_win.title("关于")
        about_win.geometry("400x420")
        about_win.configure(bg=THEME['bg_main'])
        about_win.resizable(False, False)
        about_win.transient(self.root)
        about_win.grab_set()
        
        main = tk.Frame(about_win, bg=THEME['bg_main'], padx=20, pady=15)
        main.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        tk.Label(main, text="通用文档解锁工具", fg=THEME['text_title'], 
                bg=THEME['bg_main'], font=('微软雅黑', 14, 'bold')).pack()
        tk.Label(main, text="v2.0", fg=THEME['text_muted'], 
                bg=THEME['bg_main'], font=('微软雅黑', 10)).pack()
        
        # 分隔线
        tk.Frame(main, bg=THEME['border'], height=1).pack(fill=tk.X, pady=15)
        
        # 功能说明
        info_frame = tk.Frame(main, bg=THEME['bg_white'],
                             highlightbackground=THEME['border'],
                             highlightthickness=1)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        info_items = [
            ("🔧 功能:", "Word/Excel保护解除并去水印、PDF/PPTX水印移除"),
            ("📁 格式:", "PDF, PPTX, PPT, DOC, DOCX, XLS, XLSX"),
            ("⚡ 特点:", "批量处理、拖拽上传、保持原格式"),
        ]
        for label, content in info_items:
            row = tk.Frame(info_frame, bg=THEME['bg_white'])
            row.pack(fill=tk.X, padx=10, pady=3)
            tk.Label(row, text=label, fg=THEME['text_title'], bg=THEME['bg_white'],
                    font=('微软雅黑', 9, 'bold'), width=8, anchor='w').pack(side=tk.LEFT)
            tk.Label(row, text=content, fg=THEME['text_dark'], bg=THEME['bg_white'],
                    font=('微软雅黑', 9)).pack(side=tk.LEFT)
        
        # 警告
        warning_frame = tk.Frame(main, bg='#fff3cd',
                                highlightbackground='#ffc107',
                                highlightthickness=1)
        warning_frame.pack(fill=tk.X, pady=10)
        tk.Label(warning_frame, text="⚠️ 仅限合法用途，请勿用于未授权文档",
                fg='#856404', bg='#fff3cd',
                font=('微软雅黑', 9), pady=8).pack()
        
        # 作者信息
        tk.Label(main, text="作者: qin + AI Assistant", fg=THEME['text_muted'],
                bg=THEME['bg_main'], font=('微软雅黑', 9)).pack(pady=(10, 0))
        tk.Label(main, text="© 2026", fg=THEME['text_muted'],
                bg=THEME['bg_main'], font=('微软雅黑', 9)).pack()
        
        # 确定按钮
        self._create_classic_button(main, "确定", about_win.destroy).pack(pady=(15, 0))
    
    def run(self):
        self.root.mainloop()


def main():
    if len(sys.argv) > 1:
        files = [f for f in sys.argv[1:] if os.path.exists(f)]
        if files:
            for fp in files:
                ext = Path(fp).suffix.lower()
                if ext in PROCESSOR_MAP:
                    result = PROCESSOR_MAP[ext](preview=False).process(fp)
                    print(f"{Path(fp).name}: {result.message}")
            input("\n按回车键退出...")
            return
    
    app = DocumentUnlockerGUI()
    app.run()


if __name__ == '__main__':
    main()
