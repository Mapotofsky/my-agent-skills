---
name: "docx-reader"
description: "Reads .docx Word files and extracts text content. Invoke when user asks to read, parse, or summarize .docx documents."
dependency:
  python:
    - python-docx>=1.1.0
---

# Word文档读取专家（DOCX）

## 任务目标
- 本Skill用于：读取.docx格式的Word文件并提取文本内容
- 能力包含：
  - 解析段落文本
  - 解析表格内容
  - 输出结构化JSON数据，便于后续摘要或分析
- 触发条件：
  - 用户提供.docx文件路径
  - 用户要求读取、解析或提取Word内容

## 前置准备
- 依赖说明：
  ```
  python-docx>=1.1.0
  ```
- 无需额外文件准备

## 操作步骤

### 步骤1：读取docx文件的统计信息
调用 `scripts/read_docx.py` ，并使用参数 `--include-content false` 读取 statistics，再判断是否需要分批读取指定范围：
- 输入：docx文件路径
- 参数：`--include-content false`：仅读取统计信息，不包含文本内容
- 输出：结构化JSON数据

### 步骤2：解析docx文件
调用 `scripts/read_docx.py` 读取docx：
- 输入：docx文件路径
- 可选：段落范围、表格范围、是否输出content
- 输出：结构化JSON数据

### 步骤3：输出结构化内容
输出包含：
- content：合并后的纯文本内容
- statistics：基本统计信息

## 输出格式
```
{
  "success": true/false,
  "file_path": "输入路径",
  "content": "合并后的纯文本",
  "statistics": {
    "paragraph_count": 0,
    "table_count": 0,
    "table_row_counts": []
  },
  "error": "错误信息或null"
}
```

## 调用命令示例
先获取统计信息：
```
python .trae\skills\docx-reader\scripts\read_docx.py D:\docs\requirements.docx --include-content false
```

文本量较小时，直接读取完整内容：
```
python .trae\skills\docx-reader\scripts\read_docx.py D:\docs\requirements.docx
```

文本量较大时，按段落或表格范围读取：
```
python .trae\skills\docx-reader\scripts\read_docx.py D:\docs\requirements.docx --paragraph-start 1 --paragraph-end 20
```
```
python .trae\skills\docx-reader\scripts\read_docx.py D:\docs\requirements.docx --table-start 2 --table-end 3
```

## 资源索引
- 读取脚本：见 [scripts/read_docx.py](scripts/read_docx.py)
  - 用途：读取docx并提取文本与表格内容
  - 参数：file_path - docx文件路径
  - 适用场景：任意Word文档解析

## 使用示例
```
功能：读取项目需求说明.docx
输入：D:\docs\requirements.docx
执行方式：
1. 智能体调用read_docx.py读取文件
2. 输出段落与表格的结构化结果
3. 需要时再进行摘要或分析
```

## 适用场景
- 文档内容提取
- 需求文档解析
- 报告内容摘要前的文本读取
- 表格数据转结构化处理
