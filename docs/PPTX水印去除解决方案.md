# PPTX水印去除解决方案 v2.0

## 问题描述

需要去除PPTX文件中的斜向文字水印，例如：
**`张三-部门/某公司 2026-01-01`**

## 核心挑战

1. 水印以"艺术字"(WordArt)形式嵌入每页
2. 每页4个水印元素，覆盖整个页面
3. 60页共240个水印，手动删除极其繁琐
4. 水印元素被锁定(`noTextEdit="1"`)

## 解决方案：基于OOXML结构的识别

### 核心原理

PPTX本质是ZIP压缩包，内含XML结构。水印以`<p:sp>`(shape)元素存在于每个`slide*.xml`中。

**PPTX文件结构：**
```
presentation.pptx (ZIP)
    ├── [Content_Types].xml
    ├── ppt/
    │   ├── slides/
    │   │   ├── slide1.xml    ← 水印在这里
    │   │   ├── slide2.xml
    │   │   └── ...
    │   ├── slideMasters/
    │   └── slideLayouts/
    └── ...
```

### 检测特征

| 特征 | 检测方式 |
|------|----------|
| 名称模式 | `name`包含"艺术字"、"WordArt"、"水印"等 |
| WordArt属性 | `fromWordArt="1"` |
| 透明度 | `alpha < 80%` |
| 旋转角度 | 25°~50°倾斜 |
| 颜色 | 灰色系(C0C0C0等) |

## 依赖安装

无需安装额外依赖，使用Python标准库即可。

## 使用方法

### 基础用法
```bash
python pptx_watermark_remover.py input.pptx
```

### 高级用法
```bash
# 指定输出文件
python pptx_watermark_remover.py input.pptx -o output.pptx

# 按关键词匹配水印
python pptx_watermark_remover.py input.pptx -k "机密" -k "内部"

# 预览模式（只扫描不删除）
python pptx_watermark_remover.py input.pptx --preview

# 禁用WordArt检测
python pptx_watermark_remover.py input.pptx --no-wordart

# 禁用旋转元素检测
python pptx_watermark_remover.py input.pptx --no-rotation
```

### 命令行参数

| 参数 | 说明 |
|------|------|
| `input` | 输入PPTX文件路径 |
| `-o, --output` | 输出文件路径 |
| `-k, --keyword` | 水印关键词 (可多次使用) |
| `--preview` | 预览模式: 只扫描不删除 |
| `--no-wordart` | 禁用WordArt检测 |
| `--no-rotation` | 禁用旋转元素检测 |
| `--alpha` | 透明度阈值 (0-100000, 默认80000) |

## 为什么这个方法有效

| 方法 | 问题 |
|------|------|
| 手动逐页删除 | 60页×4个=240次操作，效率极低 |
| 母版删除 | 此水印不在母版中，无效 |
| 第三方工具 | 可能损坏格式或收费 |
| **基于XML结构过滤** | ✅ 精确识别水印元素，一键全部移除 |

## 适用场景

此方法适用于：
- WordArt艺术字水印
- 每页重复的平铺水印
- OA/DMS系统自动添加的下载水印
- 半透明文字水印

不适用于：
- 图片水印（需图像处理）
- 嵌入母版的水印（需修改slideMaster）
- PDF转换后的水印（应使用PDF水印工具）

## 测试结果

```
输入: example.pptx
页面总数: 60
扫描元素: 668
发现水印: 240 (名称含"艺术字")
```

## 输出文件

- `example_无水印.pptx`
- ✅ 240个水印完全去除
- ✅ 其他形状元素完整保留
- ✅ 图片图表完整保留
