# Excel 文档解锁解决方案

## 技术原理

### 第一性原理

```
.xlsx = ZIP 压缩包
工作簿保护 = workbook.xml 中的标签
工作表保护 = sheet*.xml 中的标签
```

### 文件结构

```
workbook.xlsx (ZIP)
├── [Content_Types].xml
├── _rels/
├── docProps/
└── xl/
    ├── workbook.xml           # 工作簿设置
    ├── styles.xml             # 样式定义
    ├── sharedStrings.xml      # 共享字符串
    └── worksheets/
        ├── sheet1.xml         # 工作表1
        ├── sheet2.xml         # 工作表2
        └── ...
```

## 保护类型

### 1. 工作簿保护

防止添加/删除/重命名工作表。

位置：`xl/workbook.xml`

```xml
<workbookProtection 
    workbookPassword="CC7D"
    lockStructure="1" 
    lockWindows="1"/>
```

### 2. 工作表保护

防止编辑单元格内容。

位置：`xl/worksheets/sheet*.xml`

```xml
<sheetProtection 
    password="CC7D" 
    sheet="1" 
    objects="1" 
    scenarios="1"
    selectLockedCells="1"/>
```

### 3. 单元格锁定

配合工作表保护使用。

```xml
<c r="A1" s="1">
    <v>锁定的值</v>
</c>
<!-- s="1" 引用的样式包含 locked 属性 -->
```

## 解决方案

### 移除工作簿保护

```python
patterns = [
    r'<workbookProtection[^>]*/>',
    r'<workbookProtection[^>]*>.*?</workbookProtection>',
]
for p in patterns:
    content = re.sub(p, '', content, flags=re.IGNORECASE | re.DOTALL)
```

### 移除工作表保护

```python
patterns = [
    r'<sheetProtection[^>]*/>',
    r'<sheetProtection[^>]*>.*?</sheetProtection>',
]
```

### 移除背景图片

```python
# 背景图片
r'<picture[^>]*/?>',
```

## 密码哈希分析

### Excel 2007-2010 (旧算法)

```
密码 → 简单哈希 → 16位值
算法：XOR + 位移
弱点：可暴力破解
```

### Excel 2013+ (SHA-512)

```
密码 → PBKDF2-SHA512 → 哈希值
迭代：100,000次
强度：无法暴力破解
```

### 关键洞察

> **删除保护标签 ≠ 破解密码**
> 
> 我们不需要知道密码是什么，只需要删除保护标签即可。
> Excel 只是检查标签是否存在，而不是验证密码。

## 加密检测

### 原理

```python
def is_encrypted(file_path: str) -> bool:
    # 正常 xlsx 是 ZIP 格式
    # 加密后变成 OLE/CFB 格式
    return not zipfile.is_zipfile(file_path)
```

### 文件头对比

| 类型 | 文件头 (Hex) |
|------|-------------|
| 正常 ZIP | `50 4B 03 04` |
| OLE 加密 | `D0 CF 11 E0` |

## 处理流程

```
1. 检测文件格式
   ├── 非 ZIP 格式 → 加密文档，无法处理
   └── ZIP 格式 → 继续
   
2. 解压 ZIP

3. 处理 XML
   ├── xl/workbook.xml → 删除 workbookProtection
   └── xl/worksheets/sheet*.xml → 删除 sheetProtection

4. 重新打包 ZIP
```

## 特殊情况

### .xls 格式 (Excel 97-2003)

- 不是 ZIP 格式，是 BIFF 二进制格式
- 需要先转换为 .xlsx（需要 Microsoft Excel 或 LibreOffice）
- 或使用二进制方式处理：

```python
# 旧版 Excel 保护标志位
patterns = [
    b'\x12\x02\x01\x00',  # 保护开启
    b'\x13\x02\x01\x00',
]
# 替换为保护关闭
content = content.replace(p, b'\x12\x02\x00\x00')
```

## 注意事项

- 密码加密（打开密码）无法处理
- 编辑保护（工作表/工作簿保护）可以移除
- VBA 宏保护需要单独处理
- 隐藏工作表可通过修改 `state` 属性显示
