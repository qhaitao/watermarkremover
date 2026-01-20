# 通用文档解锁工具 / Document Unlocker v2.0

一款基于 Python 的文档解锁工具，支持移除 Word、Excel、PDF、PPTX 文档的编辑保护和水印。

A Python-based document unlocker that removes edit protection and watermarks from Word, Excel, PDF, and PPTX files.

## ✨ 功能特性 / Features

| 格式 Format | 编辑保护移除 Protection | 水印移除 Watermark | 说明 Description |
|------|-------------|---------|------|
| Word (.docx) | ✅ | ✅ | 文档保护、VML水印 |
| Excel (.xlsx) | ✅ | ✅ | 工作簿/工作表保护 |
| PDF (.pdf) | - | ✅ | 旋转文本水印 |
| PPTX (.pptx) | ✅ | ✅ | 演示文稿保护、艺术字水印 |

### 🌐 多语言支持 / Multilingual

- 中文 / English 界面切换
- 点击底部「🌐」按钮切换语言

> ⚠️ **注意**：本工具无法处理**密码加密**的文档（需要密码才能打开的文档）。

## 🔬 技术原理

基于**第一性原理**实现：

```
.docx / .xlsx / .pptx = ZIP 压缩包
编辑保护 = XML 中的标签
↓
解压 → 删除保护标签 → 重新打包
```

## 🚀 快速开始

### 方式一：直接运行（推荐）

下载 [Releases](https://github.com/qhaitao/watermarkremover/releases) 中的 `DocumentUnlocker.exe`，双击运行。

### 方式二：源码运行

```bash
git clone https://github.com/qhaitao/watermarkremover.git
cd watermarkremover
pip install -r requirements.txt
python document_toolkit_gui.py
```

## 📦 依赖

- Python 3.8+
- pikepdf (PDF处理)
- tkinterdnd2 (拖拽功能，可选)
- pywin32 (Windows下.doc/.xls转换，可选)

## 🏗️ 项目结构

```
├── document_toolkit_gui.py   # GUI主程序
├── processors/               # 文档处理器
│   ├── __init__.py
│   ├── base.py              # 抽象基类
│   ├── word_processor.py    # Word处理器
│   ├── excel_processor.py   # Excel处理器
│   ├── pdf_processor.py     # PDF处理器
│   └── pptx_processor.py    # PPTX处理器
├── requirements.txt
└── README.md
```

## 📋 使用说明

1. **选择文件**：拖拽文件到窗口或点击"选择文件"按钮
2. **开始解锁**：点击"开始解锁"按钮
3. **查看结果**：处理后的文件保存在原文件同目录，以 `_unlocked` 或 `_无水印` 后缀命名

## ⚠️ 免责声明

本工具仅供学习和合法用途。请勿用于未经授权的文档解锁。使用者需自行承担相关法律责任。

## 📄 License

MIT License

## 🙏 致谢

- [pikepdf](https://github.com/pikepdf/pikepdf) - PDF处理库
- [tkinterdnd2](https://github.com/pmgagne/tkinterdnd2) - Tkinter拖拽扩展
