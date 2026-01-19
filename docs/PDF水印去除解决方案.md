# PDF水印去除解决方案 v2.0

## 问题描述

需要去除PDF文件中的斜向文字水印：
**`秦海涛-财务部/中国邮政集团有限公司广东省分公司 2026-01-06`**

## 核心挑战

1. 水印与正文内容在空间上重叠
2. 传统的redact(涂抹)方法会同时删除水印覆盖区域的正文
3. 需要精确区分水印和正文

## 解决方案：基于旋转矩阵的识别

### 核心原理

PDF中的文字通过"变换矩阵"(Transformation Matrix)控制位置和方向。

**文字矩阵格式：** `a b c d e f Tm`

- `a, d` - 缩放因子
- `b, c` - 旋转因子
- `e, f` - 位移

**关键区别：**
- **正常水平文字**：`b ≈ 0, c ≈ 0`
- **斜向水印**：`b ≠ 0` 或 `c ≠ 0`（表示有旋转）

### 检测特征

| 特征 | 检测方式 |
|------|----------|
| 旋转角度 | 变换矩阵b/c分量非零 (5°~85°) |
| 关键词 | 文本内容匹配 |
| 颜色 | 灰色系 (可选) |

## 依赖安装

```bash
pip install pikepdf
```

## 使用方法

### 基础用法
```bash
python pdf_watermark_remover.py input.pdf
```

### 高级用法
```bash
# 指定输出文件
python pdf_watermark_remover.py input.pdf -o output.pdf

# 按关键词匹配水印
python pdf_watermark_remover.py input.pdf -k "机密" -k "内部"

# 预览模式（只扫描不删除）
python pdf_watermark_remover.py input.pdf --preview

# 设置角度范围
python pdf_watermark_remover.py input.pdf --angle-min 20 --angle-max 60

# 启用颜色检测
python pdf_watermark_remover.py input.pdf --detect-color
```

### 命令行参数

| 参数 | 说明 |
|------|------|
| `input` | 输入PDF文件路径 |
| `-o, --output` | 输出文件路径 |
| `-k, --keyword` | 水印关键词 (可多次使用) |
| `--preview` | 预览模式: 只扫描不删除 |
| `--angle-min` | 最小检测角度 (默认5°) |
| `--angle-max` | 最大检测角度 (默认85°) |
| `--threshold` | 旋转检测阈值 (默认0.1) |
| `--detect-color` | 启用颜色检测 |

## 为什么这个方法有效

| 方法 | 问题 |
|------|------|
| 按关键词redact | 会删除水印区域内的所有内容，包括正文 |
| 按颜色redact | 水印颜色与某些正文元素可能相同 |
| **按旋转角度过滤** | ✅ 正文始终是水平的，水印是斜向的，可精确区分 |

## 适用场景

此方法适用于：
- 斜向/倾斜的文字水印
- 水印与正文在同一图层
- 水印覆盖在正文上方

不适用于：
- 水平方向的水印
- 图像水印
- 背景图层水印

## 测试结果

```
输入: 寄递业务外包采购案例汇编（2025年版）.pdf
页面总数: 37
扫描文本块: 2704
发现水印: 49 (角度29.8°)
```

## 输出文件

- `寄递业务外包采购案例汇编（2025年版）_无水印.pdf`
- ✅ 水印完全去除
- ✅ 正文内容完整
- ✅ 排版格式保留
