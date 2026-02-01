import argparse
import json
import os
import sys
from typing import Optional, Tuple
from PyPDF2 import PdfReader


def parse_bool(value: str) -> bool:
    value = value.lower()
    if value in ("true", "1", "yes", "y"):
        return True
    if value in ("false", "0", "no", "n"):
        return False
    raise ValueError("include-content 仅支持 true/false")


def normalize_range(total: int, start: Optional[int], end: Optional[int]) -> Tuple[int, int]:
    if total <= 0:
        return 0, 0
    if start is None:
        start = 1
    if end is None:
        end = total
    start = max(1, int(start))
    end = min(total, int(end))
    if start > end:
        return 0, 0
    return start - 1, end


def extract_pdf(
    file_path: str,
    page_start: Optional[int] = None,
    page_end: Optional[int] = None,
    include_content: bool = True
) -> dict:
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

    page_start_index, page_end_index = normalize_range(len(pages), page_start, page_end)
    selected_pages = pages[page_start_index:page_end_index]

    content = "\n\n".join([p for p in selected_pages if p]).strip()
    char_count = len(content)
    if not include_content:
        content = ""
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
            "page_count": len(selected_pages),
            "char_count": char_count
        },
        "error": None
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("file_path")
    parser.add_argument("--page-start", type=int, default=None)
    parser.add_argument("--page-end", type=int, default=None)
    parser.add_argument("--include-content", type=parse_bool, default=True)

    args = parser.parse_args()
    result = extract_pdf(
        args.file_path,
        page_start=args.page_start,
        page_end=args.page_end,
        include_content=args.include_content
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))

    if not result["success"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
