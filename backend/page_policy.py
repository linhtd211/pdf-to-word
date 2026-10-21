from __future__ import annotations

from dataclasses import dataclass
from statistics import median
from typing import Literal

from docx.enum.section import WD_ORIENT
from docx.shared import Inches

PagePolicy = Literal["source", "administrative"]
A4_WIDTH_PT = 595.2756
A4_HEIGHT_PT = 841.8898
MIN_SOURCE_MARGIN_PT = 14.17  # 0.5 cm; PDFs do not normally store Word page margins.
MAX_SOURCE_MARGIN_PT = 99.21  # 3.5 cm; avoid deriving giant blank areas from sparse text.


@dataclass(frozen=True)
class PageSetup:
    width_pt: float
    height_pt: float
    top_pt: float
    bottom_pt: float
    left_pt: float
    right_pt: float

    @property
    def orientation(self) -> str:
        return "landscape" if self.width_pt > self.height_pt else "portrait"

    @property
    def key(self) -> tuple[float, ...]:
        return tuple(round(value, 1) for value in (
            self.width_pt, self.height_pt, self.top_pt, self.bottom_pt, self.left_pt, self.right_pt
        ))


def _content_boxes(page) -> list[tuple[float, float, float, float]]:
    boxes: list[tuple[float, float, float, float]] = []
    for block in page.get_text("dict", sort=True).get("blocks", []):
        bbox = block.get("bbox")
        if bbox and len(bbox) >= 4 and (block.get("type") == 0 or block.get("type") == 1):
            boxes.append((float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3])))
    try:
        boxes.extend(
            (float(table.bbox[0]), float(table.bbox[1]), float(table.bbox[2]), float(table.bbox[3]))
            for table in page.find_tables().tables
            if table.bbox is not None and len(table.bbox) >= 4
        )
    except Exception:
        pass
    return boxes


def _clamp_source_margin(value: float) -> float:
    return min(MAX_SOURCE_MARGIN_PT, max(MIN_SOURCE_MARGIN_PT, value))


def resolve_page_setups(pdf, policy: PagePolicy, administrative_document: bool = False) -> list[PageSetup]:
    """Resolve stable settings per source page without treating sparse content as huge margins.

    PDF files generally contain page size and drawing coordinates, not Word page-margin
    metadata. Source margins are therefore inferred conservatively, clamped to 0.5–3.5 cm,
    and median-smoothed across pages of equal size. Administrative mode normalizes to A4.
    """
    pages = list(pdf)
    if policy == "administrative":
        raw: list[PageSetup] = []
        for page in pages:
            width = float(page.rect.width)
            height = float(page.rect.height)
            wide_tables: list[float] = []
            try:
                wide_tables = [float(table.bbox[2]) - float(table.bbox[0]) for table in page.find_tables().tables]
            except Exception:
                pass
            portrait_content_width = A4_WIDTH_PT - (3.0 + 1.5) * 72 / 2.54
            landscape = width > height or any(table_width > portrait_content_width for table_width in wide_tables)
            page_width, page_height = (A4_HEIGHT_PT, A4_WIDTH_PT) if landscape else (A4_WIDTH_PT, A4_HEIGHT_PT)
            raw.append(PageSetup(
                width_pt=page_width,
                height_pt=page_height,
                top_pt=2.0 * 72 / 2.54,
                bottom_pt=2.5 * 72 / 2.54,
                left_pt=3.0 * 72 / 2.54,
                right_pt=1.5 * 72 / 2.54,
            ))
        return raw

    raw: list[PageSetup] = []
    groups: dict[tuple[float, float], list[tuple[int, tuple[float, float, float, float]]]] = {}
    for page_index, page in enumerate(pages):
        width, height = float(page.rect.width), float(page.rect.height)
        boxes = _content_boxes(page)
        if boxes:
            margins = (
                min(box[1] for box in boxes),
                height - max(box[3] for box in boxes),
                min(box[0] for box in boxes),
                width - max(box[2] for box in boxes),
            )
        else:
            margins = (45.0, 45.0, 45.0, 45.0)
        clamped = (
            _clamp_source_margin(margins[0]),
            _clamp_source_margin(margins[1]),
            _clamp_source_margin(margins[2]),
            _clamp_source_margin(margins[3]),
        )
        is_standard_a4_portrait = (
            width <= height
            and abs(width - A4_WIDTH_PT) <= 2
            and abs(height - A4_HEIGHT_PT) <= 2
        )
        if administrative_document and is_standard_a4_portrait:
            clamped = (2.0 * 72 / 2.54, 2.5 * 72 / 2.54, 3.0 * 72 / 2.54, 1.5 * 72 / 2.54)
        raw.append(PageSetup(width, height, *clamped))
        groups.setdefault((round(width, 1), round(height, 1)), []).append((page_index, clamped))

    stable = list(raw)
    for entries in groups.values():
        medians = tuple(median(values[index] for _, values in entries) for index in range(4))
        for page_index, _ in entries:
            original = raw[page_index]
            stable[page_index] = PageSetup(
                original.width_pt,
                original.height_pt,
                *(_clamp_source_margin(value) for value in medians),
            )
    return stable


def apply_page_setup(section, setup: PageSetup) -> None:
    section.page_width = Inches(setup.width_pt / 72)
    section.page_height = Inches(setup.height_pt / 72)
    section.orientation = WD_ORIENT.LANDSCAPE if setup.width_pt > setup.height_pt else WD_ORIENT.PORTRAIT
    section.top_margin = Inches(setup.top_pt / 72)
    section.bottom_margin = Inches(setup.bottom_pt / 72)
    section.left_margin = Inches(setup.left_pt / 72)
    section.right_margin = Inches(setup.right_pt / 72)
