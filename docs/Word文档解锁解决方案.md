# Word 文档解锁解决方案

## 技术原理

### 第一性原理

```
.docx = ZIP 压缩包
编辑保护 = XML 中的标签
水印 = 页眉中的 VML 图形
```

### 文件结构

```
document.docx (ZIP)
├── [Content_Types].xml
├── _rels/
├── docProps/
└── word/
    ├── document.xml      # 主文档内容
    ├── settings.xml      # 文档设置（保护信息）
    ├── header1.xml       # 页眉（水印位置）
    ├── header2.xml
    └── ...
```

## 保护移除

### 文档保护标签

位置：`word/settings.xml`

```xml
<!-- 文档保护 -->
<w:documentProtection w:edit="readOnly" w:enforcement="1" 
    w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" 
    w:cryptAlgorithmSid="14" w:cryptSpinCount="100000" 
    w:hash="..." w:salt="..."/>

<!-- 写保护 -->
<w:writeProtection w:recommended="1"/>
```

### 解决方案

直接删除这些标签：

```python
patterns = [
    r'<w:documentProtection[^>]*/?>',
    r'<w:writeProtection[^>]*/?>',
]
for p in patterns:
    content = re.sub(p, '', content, flags=re.IGNORECASE | re.DOTALL)
```

## 水印移除

### VML 水印结构

位置：`word/header*.xml`

```xml
<w:pict>
    <v:shape id="PowerPlusWaterMarkObject" 
        style="rotation:-45" 
        type="#_x0000_t136">
        <v:textpath string="机密文件"/>
    </v:shape>
</w:pict>
```

### 水印特征

| 特征 | 说明 |
|------|------|
| `type="#_x0000_t136"` | WordArt 艺术字类型 |
| `rotation:` | 旋转角度（通常 -45° 或 45°） |
| `PowerPlusWaterMarkObject` | Word 水印标识 |

### 解决方案

```python
watermark_patterns = [
    r'<w:pict[^>]*>.*?type="#_x0000_t136".*?</w:pict>',
    r'<w:pict[^>]*>.*?rotation:.*?</w:pict>',
    r'<w:pict[^>]*>.*?PowerPlusWaterMarkObject.*?</w:pict>',
]
```

### 背景水印

位置：`word/document.xml`

```xml
<w:background w:color="FFFFFF">
    <v:background>...</v:background>
</w:background>
```

## 加密检测

### 原理

- 正常 .docx 是 ZIP 格式
- 加密后变成 OLE/CFB 格式

```python
def is_encrypted(file_path: str) -> bool:
    return not zipfile.is_zipfile(file_path)
```

## 处理流程

```
1. 检测文件格式
   ├── 非 ZIP 格式 → 加密文档，无法处理
   └── ZIP 格式 → 继续
   
2. 解压 ZIP

3. 处理 XML
   ├── settings.xml → 删除保护标签
   ├── header*.xml → 删除 VML 水印
   └── document.xml → 删除背景水印

4. 重新打包 ZIP
```

## 注意事项

- `.doc` 格式需要先转换为 `.docx`（需要 Microsoft Word 或 LibreOffice）
- 密码加密的文档无法处理（需要知道密码）
- 编辑保护 ≠ 密码加密
