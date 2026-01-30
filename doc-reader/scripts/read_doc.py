import json
import os
import sys
import win32com.client


def extract_doc(file_path: str) -> dict:
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

    if not file_path.lower().endswith(".doc"):
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
            "error": "仅支持.doc格式"
        }

    try:
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        doc = word_app.Documents.Open(os.path.abspath(file_path), ReadOnly=True)
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

    paragraphs = []
    tables = []
    table_row_counts = []

    try:
        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            if text:
                paragraphs.append(text)

        for table in doc.Tables:
            table_rows = []
            for r in range(1, table.Rows.Count + 1):
                row_obj = table.Rows(r)
                row_cells = []
                for cell in row_obj.Cells:
                    cell_text = cell.Range.Text
                    cell_text = cell_text.replace("\r", "").replace("\x07", "").strip()
                    row_cells.append(cell_text)
                table_rows.append(row_cells)
            tables.append(table_rows)
            table_row_counts.append(len(table_rows))
    finally:
        doc.Close(False)
        word_app.Quit()

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
        print("用法: python read_doc.py <DOC_PATH>")
        print("示例: python read_doc.py D:\\docs\\requirements.doc")
        sys.exit(1)

    file_path = sys.argv[1]
    result = extract_doc(file_path)
    print(json.dumps(result, ensure_ascii=False, indent=2))

    if not result["success"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
