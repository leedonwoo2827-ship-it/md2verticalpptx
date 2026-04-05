"""
Extended Markdown parser.
Parses ---slide blocks with @field: value syntax into SlideData objects.

Adapted from pptx-vertical-writer/src/md_parser.py.
"""

from __future__ import annotations

import re
from typing import Any

from .models import SlideData


def parse_md(md_text: str) -> tuple[dict[str, str], list[SlideData]]:
    """
    Parse extended MD format into (config, slides).

    Format:
    ---config
    reference_pptx: path/to/ref.pptx
    ---

    ---slide
    template: T1
    ref_slide: 5
    ---
    @governing_message: text...
    @note: slide notes text
    """
    config: dict[str, str] = {}

    # Extract config block
    config_match = re.search(r'---config\s*\n(.*?)\n---', md_text, re.DOTALL)
    if config_match:
        for line in config_match.group(1).strip().split('\n'):
            line = line.strip()
            if ':' in line:
                key, val = line.split(':', 1)
                config[key.strip()] = val.strip()

    # Split on ---slide blocks (with optional # [SNNN] comment line)
    slide_blocks = re.split(r'---slide\s*\n(?:#\s*\[S\d+\].*\n)?', md_text)

    slides: list[SlideData] = []
    for idx, block in enumerate(slide_blocks[1:]):
        slide = _parse_slide_block(block, idx)
        if slide:
            slides.append(slide)

    return config, slides


def _parse_slide_block(block: str, index: int) -> SlideData | None:
    """Parse a single slide block into SlideData."""
    header_match = re.match(r'(.*?)\n---\s*\n(.*)', block, re.DOTALL)

    if header_match:
        header_text = header_match.group(1)
        body_text = header_match.group(2)
    else:
        header_text = block
        body_text = ''

    slide = SlideData(index=index)

    # Parse header
    for line in header_text.strip().split('\n'):
        line = line.strip()
        if line.startswith('template:'):
            slide.template = line.split(':', 1)[1].strip()
        elif line.startswith('ref_slide:'):
            try:
                slide.ref_slide = int(line.split(':', 1)[1].strip())
            except ValueError:
                pass
        elif line.startswith('reference_pptx:'):
            slide.reference_pptx = line.split(':', 1)[1].strip()

    if not body_text:
        body_text = header_text if not header_match else ''

    # Parse body
    if body_text:
        _parse_body(body_text.strip(), slide)

    # Extract @note from fields (should not go to shape replacement)
    if 'note' in slide.fields:
        slide.note = slide.fields.pop('note')

    return slide if slide.template or slide.ref_slide else None


def _parse_body(body: str, slide: SlideData) -> None:
    """Parse body: @fields, markdown tables, bullets."""
    lines = body.split('\n')
    current_field: str | None = None
    current_value_lines: list[str] = []
    current_table: list[str] = []
    in_table = False

    def flush_field():
        nonlocal current_field, current_value_lines
        if current_field:
            value = '\n'.join(current_value_lines).strip()
            slide.fields[current_field] = value
            current_field = None
            current_value_lines = []

    def flush_table():
        nonlocal current_table, in_table
        if current_table:
            parsed = _parse_md_table(current_table)
            if parsed:
                slide.tables.append(parsed)
            current_table = []
            in_table = False

    for line in lines:
        stripped = line.strip()

        # Markdown table detection
        if '|' in stripped and stripped.startswith('|'):
            if not in_table:
                flush_field()
            in_table = True
            current_table.append(stripped)
            continue
        elif in_table:
            flush_table()

        # @field: value
        field_match = re.match(r'^@(\S+?):\s*(.*)', stripped)
        if field_match:
            flush_field()
            current_field = field_match.group(1)
            value = field_match.group(2)
            current_value_lines = [value] if value else []
            continue

        # Continuation of current field (bullet or plain text)
        if current_field and stripped:
            current_value_lines.append(stripped)
            continue

        # Empty line ends current field
        if not stripped and current_field:
            flush_field()

    # Flush remaining
    flush_field()
    flush_table()


def _parse_md_table(table_lines: list[str]) -> dict[str, Any] | None:
    """Parse markdown table into 2D array."""
    if len(table_lines) < 2:
        return None

    rows = []
    for line in table_lines:
        cells = [c.strip() for c in line.strip('|').split('|')]
        # Skip separator lines (---|---)
        if all(re.match(r'^[-:]+$', c) for c in cells if c):
            continue
        rows.append(cells)

    if not rows:
        return None

    return {
        'headers': rows[0] if rows else [],
        'rows': rows[1:] if len(rows) > 1 else [],
        'raw_rows': rows,
    }
