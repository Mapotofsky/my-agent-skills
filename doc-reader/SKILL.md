---
name: "doc-reader"
description: "Reads .doc Word files and extracts text content. Invoke when user asks to read, parse, or summarize .doc documents."
dependency:
  python:
    - pywin32>=306
---

# Word文档读取专家（DOC）

## 任务目标
- 本Skill用于：读取.doc格式的Word文件并提取文本内容
- 能力包含：
  - 解析段落文本
  - 解析表格内容
  - 输出结构化JSON数据，便于后续摘要或分析
- 触发条件：
  - 用户提供.doc文件路径
  - 用户要求读取、解析或提取Word内容

## 前置准备
- 依赖说明：
  ```
  pywin32>=306
  ```
- 无需额外文件准备

## 操作步骤

### 步骤1：解析doc文件
调用 `scripts/read_doc.py` 读取doc：
- 输入：doc文件路径
- 输出：结构化JSON数据

### 步骤2：输出结构化内容
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

## 资源索引
- 读取脚本：见 [scripts/read_doc.py](scripts/read_doc.py)
  - 用途：读取doc并提取文本与表格内容
  - 参数：file_path - doc文件路径
  - 适用场景：任意Word文档解析

## 使用示例
```
功能：读取项目需求说明.doc
输入：D:\docs\requirements.doc
执行方式：
1. 智能体调用read_doc.py读取文件
2. 输出段落与表格的结构化结果
3. 需要时再进行摘要或分析
```

## 适用场景
- 文档内容提取
- 需求文档解析
- 报告内容摘要前的文本读取
- 表格数据转结构化处理
