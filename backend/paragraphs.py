from __future__ import annotations

import re
from copy import deepcopy
from typing import Any


_PARAGRAPH_START = re.compile(
    r"^(?:điều\s+\d+\b|căn cứ\b|theo đề nghị\b|quyết định\s*:?$|phụ lục\b|"
    r"nơi nhận\s*:|[-–•]\s|\d+[.)]\s)",
    re.IGNORECASE,
)
_TERMINAL = re.compile(r"[.!?;:…][\"'”’)]*$")
_CONTINUATION_WORDS = {
    "và", "hoặc", "của", "được", "theo", "trong", "với", "khi", "tại", "để", "do",
    "nhằm", "các", "người", "bảo", "chi", "thực", "đồng", "mà", "nếu", "như", "cho",
}


def line_text(line: dict[str, Any]) -> str:
    return "".join(str(span.get("text", "")) for span in line.get("spans", [])).strip()


def block_text(block: dict[str, Any]) -> str:
    return " ".join(line_text(line) for line in block.get("lines", []) if line_text(line)).strip()


def _line_font_signature(line: dict[str, Any]) -> tuple[str, float, int] | None:
    spans = [span for span in line.get("spans", []) if str(span.get("text", "")).strip()]
    if not spans:
        return None
    span = spans[-1]
    return (
        str(span.get("font", "")).split("+")[-1].lower(),
        round(float(span.get("size", 0)), 1),
        int(span.get("flags", 0)) & (2 | 16),
    )


def starts_independent_paragraph(text: str) -> bool:
    return bool(_PARAGRAPH_START.match(text.strip()))


def split_lines_into_paragraphs(lines: list[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    """Separate true headings/list items/visible paragraph gaps; keep ordinary wraps together."""
    groups: list[list[dict[str, Any]]] = []
    previous: dict[str, Any] | None = None
    for line in lines:
        text = line_text(line)
        if not text:
            continue
        current_box = line.get("bbox", (0, 0, 0, 0))
        starts_new = starts_independent_paragraph(text)
        if previous is not None:
            previous_box = previous.get("bbox", (0, 0, 0, 0))
            line_height = max(1.0, float(previous_box[3]) - float(previous_box[1]))
            vertical_gap = float(current_box[1]) - float(previous_box[3])
            # A PDF line wrap has almost no blank gap; a paragraph break normally has a
            # visible baseline gap. Keep this conservative to avoid splitting legal text.
            if vertical_gap > max(9.0, line_height * 0.80):
                starts_new = True
        if not groups or starts_new:
            groups.append([line])
        else:
            groups[-1].append(line)
        previous = line
    return groups


def can_join_text_blocks(
    previous: dict[str, Any],
    current: dict[str, Any],
    *,
    page_width: float,
    page_height: float | None = None,
    across_pages: bool = False,
) -> bool:
    """True only for a clear wrapped continuation, not a new sentence/heading/list item."""
    previous_lines = [line for line in previous.get("lines", []) if line_text(line)]
    current_lines = [line for line in current.get("lines", []) if line_text(line)]
    if not previous_lines or not current_lines:
        return False
    previous_text = block_text(previous)
    current_text = block_text(current)
    first_current = line_text(current_lines[0]).lstrip()
    if not previous_text or not current_text or starts_independent_paragraph(first_current):
        return False
    if _TERMINAL.search(previous_text):
        return False
    first_char = next((char for char in first_current if char.isalpha()), "")
    first_word = first_current.split(maxsplit=1)[0].strip(".,;:()[]{}") if first_current else ""
    if not first_char or not (first_char.islower() or first_word.lower() in _CONTINUATION_WORDS):
        return False

    previous_box = previous.get("bbox", (0, 0, 0, 0))
    current_box = current.get("bbox", (0, 0, 0, 0))
    previous_first = previous_lines[0].get("bbox", previous_box)
    current_first = current_lines[0].get("bbox", current_box)
    if abs(float(previous_first[0]) - float(current_first[0])) > max(12.0, page_width * 0.025):
        return False

    previous_signature = _line_font_signature(previous_lines[-1])
    current_signature = _line_font_signature(current_lines[0])
    if previous_signature and current_signature:
        if previous_signature[2] != current_signature[2] or abs(previous_signature[1] - current_signature[1]) > 1.5:
            return False

    if across_pages:
        if page_height is None:
            return False
        return float(previous_box[3]) >= page_height * 0.84 and float(current_box[1]) <= min(48.0, page_height * 0.07)

    line_height = max(1.0, float(previous_box[3]) - float(previous_box[1]))
    gap = float(current_box[1]) - float(previous_box[3])
    return -2.0 <= gap <= max(9.0, line_height * 0.65)


def merge_text_blocks(previous: dict[str, Any], current: dict[str, Any]) -> dict[str, Any]:
    merged = deepcopy(previous)
    merged["lines"] = [*previous.get("lines", []), *current.get("lines", [])]
    boxes = [previous.get("bbox"), current.get("bbox")]
    boxes = [box for box in boxes if box and len(box) >= 4]
    if boxes:
        merged["bbox"] = (
            min(float(box[0]) for box in boxes),
            min(float(box[1]) for box in boxes),
            max(float(box[2]) for box in boxes),
            max(float(box[3]) for box in boxes),
        )
    return merged


def merge_adjacent_text_events(events: list[tuple[float, int, dict[str, Any]]], page_width: float) -> list[tuple[float, int, dict[str, Any]]]:
    """Join only consecutive text events; tables, images and footers are hard boundaries."""
    result: list[tuple[float, int, dict[str, Any]]] = []
    for event in events:
        top, kind, item = event
        is_text = kind == 1 and item.get("type") == 0
        if result:
            previous_top, previous_kind, previous_item = result[-1]
            previous_is_text = previous_kind == 1 and previous_item.get("type") == 0
            if previous_is_text and is_text and can_join_text_blocks(
                previous_item, item, page_width=page_width
            ):
                merged = merge_text_blocks(previous_item, item)
                result[-1] = (min(previous_top, top), previous_kind, merged)
                continue
        result.append(event)
    return result
