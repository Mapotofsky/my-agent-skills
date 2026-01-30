import json
import os
import sys
from docx import Document
from docx.opc.exceptions import PackageNotFoundError


def extract_docx(file_path: str) -> dict:
    if not os.path.isfile(file_path):
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "statistics": {
                "paragraph_count": 0,
                "table_count": 0,
                "table_row_counts": [],
                "char_count": 0
            },
            "error": "文件不存在"
        }

    if not file_path.lower().endswith(".docx"):
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "statistics": {
                "paragraph_count": 0,
                "table_count": 0,
                "table_row_counts": [],
                "char_count": 0
            },
            "error": "仅支持.docx格式"
        }

    try:
        document = Document(file_path)
    except PackageNotFoundError:
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "statistics": {
                "paragraph_count": 0,
                "table_count": 0,
                "table_row_counts": [],
                "char_count": 0
            },
            "error": "无法读取docx文件"
        }
    except Exception as exc:
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "statistics": {
                "paragraph_count": 0,
                "table_count": 0,
                "table_row_counts": [],
                "char_count": 0
            },
            "error": str(exc)
        }

    paragraphs = [p.text.strip() for p in document.paragraphs if p.text and p.text.strip()]
    tables = []
    table_row_counts = []

    for table in document.tables:
        table_rows = []
        for row in table.rows:
            row_cells = [cell.text.strip() for cell in row.cells]
            table_rows.append(row_cells)
        tables.append(table_rows)
        table_row_counts.append(len(table_rows))

    table_text_blocks = []
    for table in tables:
        rows = ["\t".join(cells) for cells in table]
        table_text_blocks.append("\n".join(rows))

    content_parts = []
    if paragraphs:
        content_parts.append("\n".join(paragraphs))
    if table_text_blocks:
        content_parts.append("\n\n".join(table_text_blocks))

    content = "\n\n".join(content_parts).strip()
    char_count = len(content)

    return {
        "success": True,
        "file_path": file_path,
        "content": content,
        "statistics": {
            "paragraph_count": len(paragraphs),
            "table_count": len(tables),
            "table_row_counts": table_row_counts,
            "char_count": char_count
        },
        "error": None
    }


def main():
    if len(sys.argv) != 2:
        print("用法: python read_docx.py <DOCX_PATH>")
        print("示例: python read_docx.py D:\\docs\\requirements.docx")
        sys.exit(1)

    file_path = sys.argv[1]
    result = extract_docx(file_path)
    print(json.dumps(result, ensure_ascii=False, indent=2))

    if not result["success"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
