---
name: "pptx-reader"
description: "Reads .pptx PowerPoint files and extracts full content or slide content lists. Invoke when user asks to read, parse, or summarize .pptx documents."
dependency:
  python:
    - python-pptx>=0.6.21
---

# PowerPoint文档读取专家（PPTX）

## 任务目标
- 本Skill用于：读取.pptx格式的PowerPoint文件并提取文本内容
- 能力包含：
  - 解析幻灯片文本
  - 解析表格内容
  - 输出结构化JSON数据，便于后续摘要或分析
  - 支持输出完整内容或幻灯片内容列表
- 触发条件：
  - 用户提供.pptx文件路径
  - 用户要求读取、解析或提取PPT内容

## 前置准备
- 依赖说明：
  ```
  python-pptx>=0.6.21
  ```
- 无需额外文件准备

## 操作步骤

### 步骤1：读取pptx文件的统计信息
调用 `scripts/read_pptx.py` ，并使用参数 `--include-content false` 读取 statistics，再判断是否需要分批读取指定范围：
- 输入：pptx文件路径
- 参数：`--include-content false`：仅读取统计信息，不包含文本内容
- 输出：结构化JSON数据

### 步骤2：解析pptx文件
调用 `scripts/read_pptx.py` 读取pptx：
- 输入：pptx文件路径
- 可选：幻灯片范围、是否输出content、输出模式
- 输出：结构化JSON数据

### 步骤3：输出结构化内容
输出包含：
- content：合并后的纯文本内容（output-mode 为 full 时）
- slides：幻灯片内容列表（output-mode 为 list 时）
- statistics：基本统计信息

## 输出格式
```
{
  "success": true/false,
  "file_path": "输入路径",
  "content": "合并后的纯文本",
  "slides": [
    {
      "slide_index": 1,
      "texts": "每页合并后的纯文本",
      "tables": ["表格A", "表格B"]
    }
  ],
  "statistics": {
    "slide_count": 0,
    "table_count": 0,
    "table_row_counts": [],
    "char_count": 0
  },
  "error": "错误信息或null"
}
```

## 调用命令示例
先获取统计信息：
```
python .trae\skills\pptx-reader\scripts\read_pptx.py D:\docs\slides.pptx --include-content false
```

文本量较小时，直接读取完整内容：
```
python .trae\skills\pptx-reader\scripts\read_pptx.py D:\docs\slides.pptx
```

文本量较大时，按幻灯片范围读取：
```
python .trae\skills\pptx-reader\scripts\read_pptx.py D:\docs\slides.pptx --slide-start 1 --slide-end 5
```

需要幻灯片内容列表：
```
python .trae\skills\pptx-reader\scripts\read_pptx.py D:\docs\slides.pptx --output-mode list
```

## 资源索引
- 读取脚本：见 [scripts/read_pptx.py](scripts/read_pptx.py)
  - 用途：读取pptx并提取幻灯片文本与表格内容
  - 参数：file_path - pptx文件路径
  - 适用场景：任意PowerPoint文档解析

## 使用示例
```
功能：读取项目路演PPT
输入：D:\docs\slides.pptx
执行方式：
1. 智能体调用read_pptx.py读取文件
2. 输出幻灯片与表格的结构化结果
3. 需要时再进行摘要或分析
```

## 适用场景
- 文档内容提取
- 汇报材料解析
- 报告内容摘要前的文本读取
- 表格数据转结构化处理
