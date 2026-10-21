from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from backend.quality import inspect_conversion, write_report  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Đối chiếu nội dung/cấu trúc PDF và DOCX sau chuyển đổi.")
    parser.add_argument("source_pdf", type=Path)
    parser.add_argument("output_docx", type=Path)
    parser.add_argument("--page-policy", choices=("source", "administrative"), default="source")
    parser.add_argument("--mode", choices=("editable", "layout"), default="editable")
    parser.add_argument("--render", action="store_true", help="Render DOCX bằng LibreOffice và so sánh page count/pixel sai khác.")
    parser.add_argument("--output", type=Path, help="Đường dẫn JSON; mặc định dùng tên DOCX với hậu tố .quality.json.")
    args = parser.parse_args()

    report = inspect_conversion(
        args.source_pdf,
        args.output_docx,
        page_policy=args.page_policy,
        mode=args.mode,
        render=args.render,
    )
    output_path = args.output or args.output_docx.with_suffix(".quality.json")
    write_report(report, output_path)
    comparison = report["content_comparison"]
    print(json.dumps({
        "report": str(output_path),
        "source_pages": report["source_pages"],
        "docx_sections": len(report["word_structure"]["sections"]),
        "native_tables": report["word_structure"]["table_shapes"],
        "image_count": report["word_structure"]["image_count"],
        "recall_percent": comparison["recall_percent"],
        "precision_percent": comparison["precision_percent"],
        "findings": report["findings"],
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
