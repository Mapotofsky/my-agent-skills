---
name: pdf-reader
description: 读取与解析PDF文件文本与表格。用户要求读取、解析、提取或摘要PDF内容时调用。
dependency:
  python:
    - PyPDF2>=3.0.0
---

# PDF文档读取专家（PDF）

## 任务目标
- 用于读取.pdf文件并提取文本内容与基本结构
- 能力包含：
  - 逐页提取纯文本
  - 提供合并后的全文
  - 基本页统计信息
  - 简单元数据读取（标题、作者等，如可用）
- 触发条件：
  - 用户提供.pdf文件路径
  - 用户要求读取、解析或提取PDF内容

## 前置准备
- 依赖说明：
  ```
  pdfminer.six>=20221105
  PyPDF2>=3.0.0
  ```
- 无需额外文件准备

## 操作步骤

### 步骤1：读取PDF文件的统计信息
调用 `scripts/read_pdf.py` ，并使用参数 `--include-content false` 读取 statistics，再判断是否需要分批读取指定范围：
- 输入：PDF文件路径
- 参数：`--include-content false`：仅读取统计信息，不包含文本内容
- 输出：结构化JSON数据

### 步骤2：解析PDF文件
调用 `scripts/read_pdf.py` 读取PDF：
- 输入：PDF文件路径
- 可选：页范围、是否输出content
- 输出：结构化JSON数据

### 步骤3：输出结构化内容
输出包含：
- content：合并后的纯文本
- metadata：可用时的PDF元数据
- statistics：页数与字符数统计

## 输出格式
```
{
  "success": true/false,
  "file_path": "输入路径",
  "content": "合并后的纯文本",
  "metadata": {
    "title": "可选",
    "author": "可选",
    "creator": "可选",
    "producer": "可选",
    "subject": "可选"
  },
  "statistics": {
    "page_count": 0,
    "char_count": 0
  },
  "error": "错误信息或null"
}
```

## 调用命令示例
先获取统计信息：
```
python .trae\skills\pdf-reader\scripts\read_pdf.py D:\docs\whitepaper.pdf --include-content false
```

文本量较小时，直接读取完整内容：
```
python .trae\skills\pdf-reader\scripts\read_pdf.py D:\docs\whitepaper.pdf
```

文本量较大时，按页范围读取：
```
python .trae\skills\pdf-reader\scripts\read_pdf.py D:\docs\whitepaper.pdf --page-start 1 --page-end 5
```

## 资源索引
- 读取脚本：见 [scripts/read_pdf.py](scripts/read_pdf.py)
  - 用途：读取PDF并提取每页文本与元数据
  - 参数：file_path - PDF文件路径
  - 适用场景：任意PDF文档解析

## 使用示例
```
功能：读取项目白皮书.pdf
输入：D:\docs\whitepaper.pdf
执行方式：
1. 调用read_pdf.py读取文件
2. 输出每页与全文结构化结果
3. 需要时再进行摘要或分析
```

## 适用场景
- 文档内容提取与检索
- 报告内容摘要前的文本读取
- 法规、论文PDF文本解析
