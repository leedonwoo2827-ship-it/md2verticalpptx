"""
PowerPoint COM slide builder.
Clones template slides, replaces text, adds slide notes, merges files.

Adapted from pptx-vertical-writer/src/slide_builder.py.
"""

from __future__ import annotations

import gc
import json
import os
import shutil
import time
from typing import Callable

import comtypes.client

from .models import SlideData


# ---------------------------------------------------------------------------
# Slide index
# ---------------------------------------------------------------------------

def load_slide_index(json_path: str) -> dict:
    """Load slide_index.json and return {slide_number: slide_info} mapping."""
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return {s['slide_number']: s for s in data.get('slides', [])}


# ---------------------------------------------------------------------------
# COM lifecycle
# ---------------------------------------------------------------------------

def _get_powerpoint():
    """Create PowerPoint COM object."""
    pp = comtypes.client.CreateObject('Powerpoint.Application')
    pp.Visible = 1
    return pp


def _quit_powerpoint(pp):
    """Quit PowerPoint COM and force process cleanup."""
    try:
        pp.Quit()
    except Exception:
        pass
    del pp
    gc.collect()
    time.sleep(0.5)


# ---------------------------------------------------------------------------
# PPTX verification
# ---------------------------------------------------------------------------

def _verify_pptx(path: str, min_size: int = 4096) -> bool:
    """Verify PPTX file is valid (size + ZIP magic number)."""
    if not os.path.exists(path):
        return False
    if os.path.getsize(path) < min_size:
        return False
    with open(path, 'rb') as f:
        magic = f.read(4)
    return magic == b'PK\x03\x04'


# ---------------------------------------------------------------------------
# Shape text replacement
# ---------------------------------------------------------------------------

def replace_shape_text_com(shape_com, new_text: str) -> bool:
    """Replace shape text via COM, preserving first-run formatting."""
    try:
        if not shape_com.HasTextFrame:
            return False
        shape_com.TextFrame.TextRange.Text = new_text
        return True
    except Exception:
        return False


def replace_table_cell_com(table_com, row: int, col: int, new_text: str) -> bool:
    """Replace table cell text (1-based index)."""
    try:
        cell = table_com.Cell(row, col)
        cell.Shape.TextFrame.TextRange.Text = new_text
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Unreplaced shape marking
# ---------------------------------------------------------------------------

MARKER = '★미교체★ '

SKIP_MARK_ROLES = {
    'group_decoration', 'decoration', 'empty_shape', 'unknown',
    'governing_label', 'chapter_label', 'page_number', 'number_circle',
    'image',
}


def _build_name_map(shapes_com) -> dict:
    """Build name → COM shape mapping."""
    name_map = {}
    for i in range(1, shapes_com.Count + 1):
        try:
            shape = shapes_com(i)
            name_map[shape.Name] = shape
        except Exception:
            pass
    return name_map


def _get_shape_name(slide_info: dict, meta_index: int) -> str:
    """Get shape name from slide_index metadata by index."""
    shapes_meta = slide_info.get('shapes', [])
    if 0 <= meta_index < len(shapes_meta):
        return shapes_meta[meta_index].get('name', '')
    return ''


def mark_unreplaced_shapes(slide_com, slide_info: dict, fields: dict) -> None:
    """Mark shapes that were not replaced with ★미교체★ prefix."""
    role_map = slide_info.get('role_map', {})
    shapes_meta = slide_info.get('shapes', [])
    shapes_com = slide_com.Shapes
    name_map = _build_name_map(shapes_com)

    replaced_names: set[str] = set()

    # Card tables
    for n, si in enumerate(role_map.get('card_table', []), 1):
        if f'카드{n}_제목' in fields or f'카드{n}_내용' in fields:
            replaced_names.add(_get_shape_name(slide_info, si))

    # General roles
    role_groups = {
        'governing_message': role_map.get('governing_message', []),
        'breadcrumb': role_map.get('breadcrumb', []),
        'section_title': role_map.get('section_title', []),
        'content': sorted(role_map.get('content_box', []) + role_map.get('content_shape', [])),
        'heading': role_map.get('heading_box', []),
        'label': sorted(role_map.get('label_box', []) + role_map.get('label_shape', [])),
        'text': role_map.get('text_content', []),
    }

    for role_name, indices in role_groups.items():
        for n, si in enumerate(indices, 1):
            individual_key = f'{role_name}_{n}'
            if individual_key in fields or role_name in fields:
                replaced_names.add(_get_shape_name(slide_info, si))

    # Mark unreplaced shapes
    for si, meta in enumerate(shapes_meta):
        role = meta.get('role', 'unknown')
        if role in SKIP_MARK_ROLES:
            continue

        shape_name = meta.get('name', '')
        if shape_name in replaced_names:
            continue

        text = meta.get('text', '')
        if not text or len(text) <= 5:
            continue

        shape = name_map.get(shape_name)
        if shape:
            try:
                if shape.HasTextFrame:
                    current = shape.TextFrame.TextRange.Text
                    if current and not current.startswith(MARKER):
                        shape.TextFrame.TextRange.Text = MARKER + current
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Field application (text replacement)
# ---------------------------------------------------------------------------

def _collect_all_text_shapes(shapes_com) -> list:
    """Collect all text-capable shapes recursively (including group items)."""
    result = []
    for i in range(1, shapes_com.Count + 1):
        shape = shapes_com(i)
        try:
            if shape.Type == 6:  # msoGroup
                for gi in range(1, shape.GroupItems.Count + 1):
                    gshape = shape.GroupItems(gi)
                    try:
                        if gshape.HasTextFrame:
                            result.append(gshape)
                    except Exception:
                        pass
            elif shape.HasTextFrame:
                result.append(shape)
        except Exception:
            pass
    return result


def _find_shape_by_role_hint(text_shapes: list, role_hint: str):
    """Find shape by role hint (fallback matching)."""
    for shape in text_shapes:
        try:
            name = shape.Name.lower()
            if role_hint == 'governing_message':
                if '부제목' in name or 'subtitle' in name:
                    return shape
            elif role_hint == 'breadcrumb':
                if '제목' in name and '부제목' not in name:
                    return shape
                if 'title' in name and 'subtitle' not in name:
                    return shape
            elif role_hint in ('content_1', 'content'):
                if '둥근' in name or '양쪽' in name or '모서리' in name:
                    return shape
        except Exception:
            pass
    return None


def apply_fields_com(slide_com, slide_info: dict, fields: dict, tables: list | None = None) -> None:
    """Apply field data to a PowerPoint COM slide via shape name matching."""
    role_map = slide_info.get('role_map', {})
    shapes_com = slide_com.Shapes

    name_map = _build_name_map(shapes_com)
    all_text_shapes = _collect_all_text_shapes(shapes_com)
    applied_fields: set[str] = set()

    def get_com_shape(meta_index: int):
        shape_name = _get_shape_name(slide_info, meta_index)
        if shape_name and shape_name in name_map:
            return name_map[shape_name]
        com_idx = meta_index + 1
        if 1 <= com_idx <= shapes_com.Count:
            return shapes_com(com_idx)
        return None

    role_groups = {
        'governing_message': role_map.get('governing_message', []),
        'breadcrumb': role_map.get('breadcrumb', []),
        'section_title': role_map.get('section_title', []),
        'content': sorted(role_map.get('content_box', []) + role_map.get('content_shape', [])),
        'heading': role_map.get('heading_box', []),
        'label': sorted(role_map.get('label_box', []) + role_map.get('label_shape', [])),
        'text': role_map.get('text_content', []),
    }

    # Text shape replacement (all roles with _N indexing)
    for role_name, indices in role_groups.items():
        for n, si in enumerate(indices, 1):
            individual_key = f'{role_name}_{n}'
            if individual_key in fields:
                shape = get_com_shape(si)
                if shape and replace_shape_text_com(shape, fields[individual_key]):
                    applied_fields.add(individual_key)
            elif role_name in fields:
                shape = get_com_shape(si)
                if shape and replace_shape_text_com(shape, fields[role_name]):
                    applied_fields.add(role_name)

    # Card table replacement (카드N_제목, 카드N_내용)
    card_indices = role_map.get('card_table', [])
    for card_num, si in enumerate(card_indices, 1):
        title_key = f'카드{card_num}_제목'
        content_key = f'카드{card_num}_내용'
        shape = get_com_shape(si)
        if shape and shape.HasTable:
            table = shape.Table
            if title_key in fields and table.Rows.Count >= 1:
                replace_table_cell_com(table, 1, 1, fields[title_key])
            if content_key in fields and table.Rows.Count >= 2:
                replace_table_cell_com(table, 2, 1, fields[content_key])

    # text_content replacement (backward compat)
    for t_num, si in enumerate(role_map.get('text_content', []), 1):
        key = f'text_{t_num}'
        if key in fields:
            shape = get_com_shape(si)
            if shape:
                replace_shape_text_com(shape, fields[key])

    # Data table replacement
    if tables:
        dt_indices = role_map.get('data_table', [])
        for ti, table_data in enumerate(tables):
            if ti < len(dt_indices):
                si = dt_indices[ti]
                shape = get_com_shape(si)
                if shape and shape.HasTable:
                    all_rows = table_data.get('raw_rows', [])
                    if not all_rows:
                        all_rows = [table_data.get('headers', [])] + table_data.get('rows', [])
                    table_com = shape.Table
                    for ri, row_data in enumerate(all_rows):
                        if ri >= table_com.Rows.Count:
                            break
                        for ci, cell_text in enumerate(row_data):
                            if ci >= table_com.Columns.Count:
                                break
                            replace_table_cell_com(table_com, ri + 1, ci + 1, str(cell_text))

    # Fallback: match remaining fields via group-internal search
    remaining = {k: v for k, v in fields.items() if k not in applied_fields}
    if remaining:
        if 'governing_message' in remaining:
            shape = _find_shape_by_role_hint(all_text_shapes, 'governing_message')
            if shape:
                replace_shape_text_com(shape, remaining.pop('governing_message'))

        if 'breadcrumb' in remaining:
            shape = _find_shape_by_role_hint(all_text_shapes, 'breadcrumb')
            if shape:
                replace_shape_text_com(shape, remaining.pop('breadcrumb'))

        content_fields = sorted([k for k in remaining if k.startswith('content_')])
        if content_fields:
            available = []
            for shape in all_text_shapes:
                try:
                    t = shape.TextFrame.TextRange.Text
                    if t and '████' in t:
                        available.append(shape)
                except Exception:
                    pass
            for i, key in enumerate(content_fields):
                if i < len(available):
                    replace_shape_text_com(available[i], remaining[key])


# ---------------------------------------------------------------------------
# Slide notes
# ---------------------------------------------------------------------------

def set_slide_notes(slide_com, note_text: str) -> bool:
    """Write text to the slide's notes page via COM."""
    try:
        notes_page = slide_com.NotesPage
        # NotesPage.Shapes(2) is the notes text placeholder by convention
        notes_shape = notes_page.Shapes(2)
        notes_shape.TextFrame.TextRange.Text = note_text
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Single slide builder
# ---------------------------------------------------------------------------

def build_single_slide(
    slide: SlideData,
    slides_info: dict,
    output_path: str,
    slides_dir: str,
    write_notes: bool = True,
) -> str:
    """
    Build a 1-slide PPTX: copy template → COM text replacement → save.

    Args:
        slide: Parsed slide data
        slides_info: {slide_number: slide_info} from slide_index.json
        output_path: Where to save the result
        slides_dir: Directory containing S*.pptx template files
        write_notes: Whether to write @note to slide notes

    Returns:
        Path to the generated PPTX file.
    """
    ref_slide_num = slide.ref_slide
    if ref_slide_num is None:
        raise ValueError('ref_slide not specified')

    slide_file = os.path.join(slides_dir, f'S{ref_slide_num:04d}.pptx')
    if not os.path.exists(slide_file):
        raise FileNotFoundError(f'Template not found: {slide_file}')

    # Copy template
    shutil.copy2(slide_file, output_path)

    fields = slide.fields
    tables = slide.tables
    note = slide.note

    if not fields and not tables and not note:
        return output_path

    # Open in COM and replace text
    pp = _get_powerpoint()
    try:
        prs = pp.Presentations.Open(os.path.abspath(output_path), WithWindow=False)
        slide_info = slides_info.get(ref_slide_num, {})

        if fields or tables:
            apply_fields_com(prs.Slides(1), slide_info, fields, tables)

        if write_notes and note:
            set_slide_notes(prs.Slides(1), note)

        prs.Save()
        prs.Close()
    finally:
        _quit_powerpoint(pp)

    if not _verify_pptx(output_path):
        raise RuntimeError(f'PPTX file corrupted: {output_path}')

    return output_path


# ---------------------------------------------------------------------------
# Merge
# ---------------------------------------------------------------------------

def _merge_pptx_files(part_files: list[str], output_path: str,
                      on_progress: Callable[[int, int], None] | None = None) -> None:
    """Merge multiple PPTX files into one using InsertFromFile (no clipboard)."""
    if len(part_files) == 1:
        shutil.copy2(part_files[0], output_path)
        return

    pp = _get_powerpoint()
    try:
        target_prs = pp.Presentations.Open(os.path.abspath(part_files[0]), WithWindow=False)

        for i, part_file in enumerate(part_files[1:], 1):
            abs_path = os.path.abspath(part_file)
            insert_at = target_prs.Slides.Count
            try:
                target_prs.Slides.InsertFromFile(abs_path, insert_at)
            except Exception as e:
                print(f'  Warning: InsertFromFile failed for {part_file}: {e}')

            if on_progress:
                on_progress(i, len(part_files) - 1)

        target_prs.SaveAs(os.path.abspath(output_path), 24)
        target_prs.Close()
    finally:
        _quit_powerpoint(pp)


def merge_pptx_files(
    part_files: list[str],
    output_path: str,
    batch_size: int = 25,
    on_progress: Callable[[int, int], None] | None = None,
) -> None:
    """
    Merge PPTX files with 2-stage batching for large sets.
    Does NOT delete source files (only intermediate merge artifacts).
    """
    if not part_files:
        raise ValueError('No files to merge')

    # Filter valid files
    valid = [f for f in part_files if _verify_pptx(f)]
    if not valid:
        raise ValueError('All PPTX files are corrupted')

    if len(valid) == 1:
        shutil.copy2(valid[0], output_path)
        return

    if len(valid) <= batch_size:
        _merge_pptx_files(valid, output_path, on_progress=on_progress)
        return

    # 2-stage merge
    intermediate_files: list[str] = []
    for i in range(0, len(valid), batch_size):
        batch = valid[i:i + batch_size]
        if len(batch) == 1:
            intermediate_files.append(batch[0])
            continue
        intermediate = f'{output_path}.merge_{i}.pptx'
        _merge_pptx_files(batch, intermediate)
        intermediate_files.append(intermediate)

    _merge_pptx_files(intermediate_files, output_path, on_progress=on_progress)

    # Clean up intermediate files only (not source slides)
    for f in intermediate_files:
        if f != output_path and f not in valid and os.path.exists(f):
            try:
                os.remove(f)
            except Exception:
                pass
