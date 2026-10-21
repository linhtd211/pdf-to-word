from __future__ import annotations

import json
import re
import shutil
import subprocess
import tempfile
import unicodedata
from collections import Counter
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import Path
from typing import Any, Iterator, cast

import fitz
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from PIL import Image, ImageChops


def _tokens(text: str) -> list[str]:
    normalized = unicodedata.normalize("NFKC", text).casefold()
    return re.findall(r"(?u)[^\W_]+", normalized)


def _source_text(pdf_path: Path) -> tuple[str, int, int, list[tuple[float, float]]]:
    chunks: list[str] = []
    scanned_pages = 0
    source_tables = 0
    page_sizes: list[tuple[float, float]] = []
    with fitz.open(pdf_path) as pdf:
        for page_index in range(len(pdf)):
            page = cast(Any, pdf[page_index])
            page_sizes.append((float(page.rect.width), float(page.rect.height)))
            page_text = page.get_text("text", sort=True).strip()
            if not page_text and page.get_images(full=True):
                scanned_pages += 1
            else:
                for block in page.get_text("dict", sort=True).get("blocks", []):
                    if block.get("type") != 0:
                        continue
                    for line in block.get("lines", []):
                        line_text = "".join(str(span.get("text", "")) for span in line.get("spans", [])).strip()
                        bbox = line.get("bbox", (0, 0, 0, 0))
                        center = (float(bbox[0]) + float(bbox[2])) / 2
                        if (
                            line_text == str(page_index + 1)
                            and abs(center - float(page.rect.width) / 2) <= float(page.rect.width) * 0.07
                            and float(bbox[1]) <= min(45.0, float(page.rect.height) * 0.06)
                        ):
                            continue
                        if line_text:
                            chunks.append(line_text)
            try:
                source_tables += len(page.find_tables().tables)
            except Exception:
                pass
    return "\n".join(chunks), scanned_pages, source_tables, page_sizes


def _iter_blocks(parent: Any) -> Iterator[tuple[str, Any]]:
    element = getattr(parent, "_element", None)
    if element is None:
        return
    for child in element.iterchildren():
        if child.tag == qn("w:p"):
            paragraph = Paragraph(child, parent)
            if paragraph._p.xpath(".//w:instrText | .//w:fldSimple"):
                continue
            yield "paragraph", paragraph.text
        elif child.tag == qn("w:tbl"):
            yield "table", Table(child, parent)


def _table_content(table: Any) -> Iterator[tuple[str, Any]]:
    seen_cells: set[Any] = set()
    for row in table.rows:
        for cell in row.cells:
            identity = cell._tc
            if identity in seen_cells:
                continue
            seen_cells.add(identity)
            yield from _iter_blocks(cell)


def _walk_docx_content(document: Any) -> tuple[list[str], int, list[dict[str, Any]]]:
    text_chunks: list[str] = []
    paragraph_count = 0
    table_shapes: list[dict[str, Any]] = []

    def walk(items: Iterator[tuple[str, Any]]) -> Iterator[tuple[str, Any]]:
        for kind, value in items:
            if kind == "paragraph":
                yield kind, value
            else:
                table = value
                cells_seen: set[Any] = set()
                for row in table.rows:
                    for cell in row.cells:
                        identity = cell._tc
                        if identity in cells_seen:
                            continue
                        cells_seen.add(identity)
                        yield from walk(_iter_blocks(cell))

    body_items = list(_iter_blocks(document._body))
    for kind, value in body_items:
        if kind == "paragraph":
            if value.strip():
                paragraph_count += 1
                text_chunks.append(value)
        else:
            table = value
            table_shapes.append({"rows": len(table.rows), "columns": len(table.columns)})
            for nested_kind, nested_value in walk(iter(_table_content(table))):
                if nested_kind == "paragraph" and str(nested_value).strip():
                    paragraph_count += 1
                    text_chunks.append(str(nested_value))

    seen_story_parts: set[Any] = set()
    for section in document.sections:
        for story in (section.header, section.first_page_header, section.footer, section.first_page_footer):
            identity = story._element
            if identity in seen_story_parts:
                continue
            seen_story_parts.add(identity)
            for kind, value in walk(iter(_iter_blocks(story))):
                if kind == "paragraph" and str(value).strip():
                    paragraph_count += 1
                    text_chunks.append(str(value))
    return text_chunks, paragraph_count, table_shapes


def _inspect_structure(docx_path: Path) -> dict[str, Any]:
    document = Document(str(docx_path))
    text_chunks, paragraph_count, table_shapes = _walk_docx_content(document)
    sections: list[dict[str, Any]] = []
    for section in document.sections:
        section = cast(Any, section)
        width = section.page_width.pt if section.page_width else 0.0
        height = section.page_height.pt if section.page_height else 0.0
        sections.append({
            "width_cm": round(width / 72 * 2.54, 2),
            "height_cm": round(height / 72 * 2.54, 2),
            "orientation": "landscape" if width > height else "portrait",
            "margins_cm": {
                "top": round(section.top_margin.pt / 72 * 2.54, 2),
                "bottom": round(section.bottom_margin.pt / 72 * 2.54, 2),
                "left": round(section.left_margin.pt / 72 * 2.54, 2),
                "right": round(section.right_margin.pt / 72 * 2.54, 2),
            },
        })
    xml = document._element
    anchor_count = len(xml.xpath(".//wp:anchor"))
    inline_count = len(document.inline_shapes)
    return {
        "paragraph_count": paragraph_count,
        "native_tables": table_shapes,
        "table_count": len(table_shapes),
        "table_shapes": [f"{table['rows']}x{table['columns']}" for table in table_shapes],
        "horizontal_merge_markers": len(xml.xpath(".//w:gridSpan")),
        "vertical_merge_markers": len(xml.xpath(".//w:vMerge")),
        "image_count": anchor_count + inline_count,
        "image_anchors": anchor_count,
        "inline_images": inline_count,
        "textboxes": len(xml.xpath(".//w:txbxContent")),
        "hard_page_breaks": len(xml.xpath(".//w:br[@w:type='page']")),
        "line_breaks": len(xml.xpath(".//w:br[not(@w:type) or @w:type='textWrapping']")),
        "sections": sections,
        "text": "\n".join(text_chunks),
    }


def _render_docx(docx_path: Path, render_dir: Path, timeout: int = 120) -> Path:
    soffice = shutil.which("soffice")
    if not soffice:
        raise RuntimeError("LibreOffice (soffice) is not installed")
    render_dir.mkdir(parents=True, exist_ok=True)
    profile = render_dir / "lo-profile"
    command = [
        soffice,
        f"-env:UserInstallation={profile.as_uri()}",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(render_dir),
        str(docx_path),
    ]
    result = subprocess.run(command, capture_output=True, text=True, timeout=timeout, check=False)
    output = render_dir / f"{docx_path.stem}.pdf"
    if result.returncode != 0 or not output.is_file():
        raise RuntimeError((result.stderr or result.stdout or "LibreOffice did not create a PDF").strip())
    return output


def _page_geometry_metrics(page: Any) -> dict[str, Any]:
    width, height = float(page.rect.width), float(page.rect.height)
    blocks: list[tuple[str, tuple[float, float, float, float]]] = []
    for block in page.get_text("dict", sort=True).get("blocks", []):
        if block.get("type") not in {0, 1} or not block.get("bbox"):
            continue
        raw_bbox = block["bbox"]
        bbox = (float(raw_bbox[0]), float(raw_bbox[1]), float(raw_bbox[2]), float(raw_bbox[3]))
        blocks.append(("text" if block.get("type") == 0 else "image", bbox))
    if not blocks:
        return {
            "content_bbox": None,
            "content_bbox_area_percent": 0.0,
            "max_internal_blank_gap_pt": 0.0,
            "max_internal_blank_gap_percent": 0.0,
            "text_image_overlap_pairs": 0,
            "offpage_blocks": 0,
        }
    x0 = min(box[0] for _, box in blocks)
    y0 = min(box[1] for _, box in blocks)
    x1 = max(box[2] for _, box in blocks)
    y1 = max(box[3] for _, box in blocks)
    intervals = sorted((max(0.0, box[1]), min(height, box[3])) for _, box in blocks if box[3] > box[1])
    merged: list[list[float]] = []
    for top, bottom in intervals:
        if not merged or top > merged[-1][1]:
            merged.append([top, bottom])
        else:
            merged[-1][1] = max(merged[-1][1], bottom)
    gaps = [max(0.0, right[0] - left[1]) for left, right in zip(merged, merged[1:])]
    overlap_pairs = 0
    for index, (kind_a, box_a) in enumerate(blocks):
        for kind_b, box_b in blocks[index + 1 :]:
            if kind_a == kind_b:
                continue
            intersection = max(0.0, min(box_a[2], box_b[2]) - max(box_a[0], box_b[0])) * max(
                0.0, min(box_a[3], box_b[3]) - max(box_a[1], box_b[1])
            )
            area_a = max(1.0, (box_a[2] - box_a[0]) * (box_a[3] - box_a[1]))
            area_b = max(1.0, (box_b[2] - box_b[0]) * (box_b[3] - box_b[1]))
            if intersection / min(area_a, area_b) > 0.15:
                overlap_pairs += 1
    offpage = sum(
        box[0] < -1 or box[1] < -1 or box[2] > width + 1 or box[3] > height + 1
        for _, box in blocks
    )
    content_area = max(0.0, x1 - x0) * max(0.0, y1 - y0)
    max_gap = max(gaps, default=0.0)
    return {
        "content_bbox": [round(x0, 1), round(y0, 1), round(x1, 1), round(y1, 1)],
        "content_bbox_area_percent": round(100 * content_area / max(1.0, width * height), 2),
        "max_internal_blank_gap_pt": round(max_gap, 1),
        "max_internal_blank_gap_percent": round(100 * max_gap / max(1.0, height), 2),
        "text_image_overlap_pairs": overlap_pairs,
        "offpage_blocks": int(offpage),
    }


def _visual_compare(source_pdf: Path, rendered_pdf: Path) -> dict[str, Any]:
    page_results: list[dict[str, Any]] = []
    with fitz.open(source_pdf) as source, fitz.open(rendered_pdf) as rendered:
        for index in range(max(len(source), len(rendered))):
            if index >= len(source) or index >= len(rendered):
                extra_page = cast(Any, rendered[index]) if index < len(rendered) else None
                source_page = cast(Any, source[index]) if index < len(source) else None
                page_results.append({
                    "page": index + 1,
                    "comparable": False,
                    "reason": "page_count_mismatch",
                    "source_geometry": _page_geometry_metrics(source_page) if source_page is not None else None,
                    "output_geometry": _page_geometry_metrics(extra_page) if extra_page is not None else None,
                })
                continue
            source_page, output_page = cast(Any, source[index]), cast(Any, rendered[index])
            source_geometry = _page_geometry_metrics(source_page)
            output_geometry = _page_geometry_metrics(output_page)
            size_error = max(
                abs(float(source_page.rect.width) - float(output_page.rect.width)),
                abs(float(source_page.rect.height) - float(output_page.rect.height)),
            )
            if size_error > 2.0:
                page_results.append({
                    "page": index + 1,
                    "comparable": False,
                    "reason": "page_size_differs",
                    "source_pt": [round(source_page.rect.width, 1), round(source_page.rect.height, 1)],
                    "output_pt": [round(output_page.rect.width, 1), round(output_page.rect.height, 1)],
                    "source_geometry": source_geometry,
                    "output_geometry": output_geometry,
                })
                continue
            source_image = Image.open(BytesIO(source_page.get_pixmap(matrix=fitz.Matrix(0.65, 0.65), alpha=False).tobytes("png"))).convert("RGB")
            output_image = Image.open(BytesIO(output_page.get_pixmap(matrix=fitz.Matrix(0.65, 0.65), alpha=False).tobytes("png"))).convert("RGB")
            if source_image.size != output_image.size:
                output_image = output_image.resize(source_image.size, Image.Resampling.LANCZOS)
            difference = ImageChops.difference(source_image, output_image).convert("L")
            histogram = difference.histogram()
            total = max(1, source_image.width * source_image.height)
            changed = sum(histogram[25:])
            mean_error = sum(index * count for index, count in enumerate(histogram)) / total
            page_results.append({
                "page": index + 1,
                "comparable": True,
                "changed_pixel_percent": round(100 * changed / total, 2),
                "mean_absolute_pixel_error": round(mean_error, 2),
                "source_geometry": source_geometry,
                "output_geometry": output_geometry,
            })
        return {
            "source_pages": len(source),
            "rendered_pages": len(rendered),
            "page_count_delta": len(rendered) - len(source),
            "pages": page_results,
        }


def inspect_conversion(
    source_pdf: Path,
    output_docx: Path,
    *,
    page_policy: str = "source",
    mode: str = "editable",
    ocr_pages: list[dict[str, Any]] | None = None,
    render: bool = False,
) -> dict[str, Any]:
    source_text, scanned_pages, source_table_count, source_page_sizes = _source_text(source_pdf)
    structure = _inspect_structure(output_docx)
    source_words = _tokens(source_text)
    output_words = _tokens(structure.pop("text"))
    findings: list[dict[str, Any]] = []

    if source_words:
        source_counter, output_counter = Counter(source_words), Counter(output_words)
        matched = sum((source_counter & output_counter).values())
        missing = source_counter - output_counter
        extra = output_counter - source_counter
        recall = matched / len(source_words)
        precision = matched / max(1, len(output_words)) if not scanned_pages else None
        sampled_source, sampled_output = source_words[:3000], output_words[:3000]
        order_similarity = (
            SequenceMatcher(None, sampled_source, sampled_output, autojunk=False).ratio()
            if not scanned_pages
            else None
        )
        comparison: dict[str, Any] = {
            "applicable": True,
            "basis": "native PDF text vs DOCX text, Unicode NFKC and case-insensitive word tokens",
            "source_word_count": len(source_words),
            "docx_word_count": len(output_words),
            "matched_word_count": matched,
            "missing_word_count": sum(missing.values()),
            "extra_word_count": sum(extra.values()),
            "recall_percent": round(recall * 100, 2),
            "precision_percent": round(precision * 100, 2) if precision is not None else None,
            "order_similarity_percent_first_3000_words": round(order_similarity * 100, 2) if order_similarity is not None else None,
            "missing_examples": [word for word, count in missing.most_common(15) for _ in range(min(count, 3))],
            "extra_examples": [word for word, count in extra.most_common(15) for _ in range(min(count, 3))],
        }
        if recall < 0.98 or (precision is not None and precision < 0.98):
            findings.append({
                "severity": "critical" if recall < 0.95 else "major",
                "code": "text_token_mismatch",
                "message": "Nội dung PDF và DOCX có chênh lệch token; cần rà soát phần bị mất/lặp.",
                "details": {"recall_percent": comparison["recall_percent"], "precision_percent": comparison["precision_percent"], "missing_examples": comparison["missing_examples"], "extra_examples": comparison["extra_examples"]},
            })
        if order_similarity is not None and order_similarity < 0.90:
            findings.append({
                "severity": "major",
                "code": "text_order_differs",
                "message": "Thứ tự token giữa PDF và DOCX khác đáng kể; có thể do thứ tự đọc bảng/cột hoặc thứ tự vùng.",
                "details": {"order_similarity_percent_first_3000_words": comparison["order_similarity_percent_first_3000_words"]},
            })
    else:
        comparison = {
            "applicable": False,
            "basis": "source PDF has no extractable native text; OCR output cannot be independently compared with source text",
            "source_word_count": 0,
            "docx_word_count": len(output_words),
            "matched_word_count": 0,
            "missing_word_count": None,
            "extra_word_count": None,
            "recall_percent": None,
            "precision_percent": None,
            "order_similarity_percent_first_3000_words": None,
            "missing_examples": [],
            "extra_examples": [],
        }
        findings.append({
            "severity": "caution",
            "code": "ocr_text_not_comparable",
            "message": "PDF scan không có text gốc để so sánh tự động; cần đọc soát OCR, đặc biệt dấu tiếng Việt và vùng bảng.",
            "details": {"scanned_pages": scanned_pages},
        })

    if structure["textboxes"]:
        findings.append({
            "severity": "major",
            "code": "unexpected_textboxes",
            "message": "DOCX có TextBox; nội dung văn bản thường nên dùng paragraph/run/table chuẩn Word.",
            "details": {"count": structure["textboxes"]},
        })
    if mode == "editable" and structure["hard_page_breaks"]:
        findings.append({
            "severity": "caution",
            "code": "hard_page_breaks_present",
            "message": "DOCX có ngắt trang cứng; kiểm tra chúng có cần thiết theo page policy hay không.",
            "details": {"count": structure["hard_page_breaks"], "page_policy": page_policy},
        })
    if source_table_count and not structure["table_count"]:
        findings.append({
            "severity": "major",
            "code": "source_tables_not_editable",
            "message": "PDF có bảng được nhận diện nhưng DOCX không có bảng Word native.",
            "details": {"source_tables": source_table_count, "docx_tables": structure["table_count"]},
        })

    ocr_page_metrics = ocr_pages or []
    low_confidence_pages = [
        int(page["page"])
        for page in ocr_page_metrics
        if page.get("mean_confidence") is not None and float(page["mean_confidence"]) < 55
    ]
    if low_confidence_pages:
        findings.append({
            "severity": "caution",
            "code": "ocr_confidence_low",
            "message": "Một số trang scan có confidence OCR thấp; cần đọc soát dấu tiếng Việt và nội dung trong bảng.",
            "details": {"pages": low_confidence_pages},
        })

    report: dict[str, Any] = {
        "source_name": source_pdf.name,
        "output_name": output_docx.name,
        "page_policy": page_policy,
        "mode": mode,
        "source_pages": len(source_page_sizes),
        "source_page_sizes_pt": [[round(width, 1), round(height, 1)] for width, height in source_page_sizes],
        "source_scanned_pages": scanned_pages,
        "ocr": {
            "pages": ocr_page_metrics,
            "mean_confidence": round(
                sum(float(page["mean_confidence"]) for page in ocr_page_metrics if page.get("mean_confidence") is not None)
                / max(1, sum(page.get("mean_confidence") is not None for page in ocr_page_metrics)),
                2,
            ) if ocr_page_metrics else None,
        },
        "source_detected_tables": source_table_count,
        "content_comparison": comparison,
        "word_structure": structure,
        "findings": findings,
        "render_comparison": None,
    }
    if render:
        try:
            with tempfile.TemporaryDirectory(prefix="pdfword-quality-") as temporary_directory:
                rendered = _render_docx(output_docx, Path(temporary_directory))
                report["render_comparison"] = _visual_compare(source_pdf, rendered)
            layout = report["render_comparison"]
            if layout["page_count_delta"] != 0:
                severity = "major" if page_policy == "source" and mode == "layout" else "caution" if page_policy == "source" else "information"
                findings.append({
                    "severity": severity,
                    "code": "rendered_page_count_differs",
                    "message": "Số trang render khác PDF nguồn; đây có thể là reflow có chủ đích hoặc sai lệch cần rà soát.",
                    "details": {"source_pages": layout["source_pages"], "rendered_pages": layout["rendered_pages"], "page_policy": page_policy},
                })
            large_gaps: list[dict[str, Any]] = []
            overlaps: list[dict[str, Any]] = []
            offpage: list[dict[str, Any]] = []
            for page_result in layout["pages"]:
                output_geometry = page_result.get("output_geometry")
                source_geometry = page_result.get("source_geometry")
                if not output_geometry:
                    continue
                output_gap = float(output_geometry["max_internal_blank_gap_percent"])
                source_gap = float(source_geometry["max_internal_blank_gap_percent"]) if source_geometry else 0.0
                if output_gap >= 22 and output_gap - source_gap >= 12:
                    large_gaps.append({"page": page_result["page"], "source_gap_percent": source_gap, "output_gap_percent": output_gap})
                if int(output_geometry["text_image_overlap_pairs"]) > 0:
                    overlaps.append({"page": page_result["page"], "pairs": output_geometry["text_image_overlap_pairs"]})
                if int(output_geometry["offpage_blocks"]) > 0:
                    offpage.append({"page": page_result["page"], "blocks": output_geometry["offpage_blocks"]})
            if large_gaps:
                findings.append({
                    "severity": "major" if page_policy == "source" else "caution",
                    "code": "large_internal_blank_gap",
                    "message": "Một số trang Word có khoảng trống dọc lớn hơn rõ rệt so với trang PDF tương ứng.",
                    "details": {"pages": large_gaps},
                })
            if overlaps:
                findings.append({
                    "severity": "caution",
                    "code": "possible_text_image_overlap",
                    "message": "Có vùng chữ và hình ảnh chồng lấn trong PDF render; cần kiểm tra trực quan để phân biệt chữ ký/dấu có chủ đích.",
                    "details": {"pages": overlaps},
                })
            if offpage:
                findings.append({
                    "severity": "major",
                    "code": "content_outside_page",
                    "message": "Một số khối nội dung nằm ngoài mép trang trong bản render Word.",
                    "details": {"pages": offpage},
                })
        except Exception as exc:
            report["render_comparison"] = {"available": False, "error": str(exc)}
            findings.append({
                "severity": "caution",
                "code": "render_unavailable",
                "message": "Không thể tạo hoặc so sánh PDF render của DOCX trong môi trường này.",
                "details": {"error": str(exc)},
            })
    report["findings"] = findings
    return report


def write_report(report: dict[str, Any], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
