import json
import os
import sys
from PyPDF2 import PdfReader


def extract_pdf(file_path: str) -> dict:
    if not os.path.isfile(file_path):
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "metadata": {},
            "statistics": {
                "page_count": 0,
                "char_count": 0
            },
            "error": "文件不存在"
        }

    if not file_path.lower().endswith(".pdf"):
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "metadata": {},
            "statistics": {
                "page_count": 0,
                "char_count": 0
            },
            "error": "仅支持.pdf格式"
        }

    try:
        reader = PdfReader(file_path)
        if reader.is_encrypted:
            try:
                if reader.decrypt("") == 0:
                    return {
                        "success": False,
                        "file_path": file_path,
                        "content": "",
                        "metadata": {},
                        "statistics": {
                            "page_count": 0,
                            "char_count": 0
                        },
                        "error": "PDF已加密，无法读取"
                    }
            except Exception:
                return {
                    "success": False,
                    "file_path": file_path,
                    "content": "",
                    "metadata": {},
                    "statistics": {
                        "page_count": 0,
                        "char_count": 0
                    },
                    "error": "PDF已加密，无法读取"
                }
    except Exception as exc:
        return {
            "success": False,
            "file_path": file_path,
            "content": "",
            "metadata": {},
            "statistics": {
                "page_count": 0,
                "char_count": 0
            },
            "error": str(exc)
        }

    pages = []
    char_count = 0
    for page in reader.pages:
        text = page.extract_text() or ""
        page_text = text.strip()
        pages.append(page_text)
        char_count += len(page_text)

    content = "\n\n".join([p for p in pages if p]).strip()
    metadata = {}
    if reader.metadata:
        metadata = {
            "title": reader.metadata.title,
            "author": reader.metadata.author,
            "creator": reader.metadata.creator,
            "producer": reader.metadata.producer,
            "subject": reader.metadata.subject
        }

    return {
        "success": True,
        "file_path": file_path,
        "content": content,
        "metadata": metadata,
        "statistics": {
            "page_count": len(pages),
            "char_count": char_count
        },
        "error": None
    }


def main():
    if len(sys.argv) != 2:
        print("用法: python read_pdf.py <PDF_PATH>")
        print("示例: python read_pdf.py D:\\docs\\whitepaper.pdf")
        sys.exit(1)

    file_path = sys.argv[1]
    result = extract_pdf(file_path)
    print(json.dumps(result, ensure_ascii=False, indent=2))

    if not result["success"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
