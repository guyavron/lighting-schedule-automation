#!/usr/bin/env python3
"""
Insert one fixture row into the studio's lighting legend (מקרא תאורה), using
the column layout found in the ACTUAL workbook's header row (row 8) — never a
hardcoded column letter. Real project files customize the base template
(add/reorder/rename columns, e.g. קרקע/מרתף splits), so column positions are
not stable across files.

Any column the skill doesn't explicitly recognize (quantities, prices,
totals, project-specific splits, etc.) is left completely untouched — the
skill only ever writes to: סימון, תיאור, יצרן, מתח_עמעום (מתח + פרוטוקול
עמעום together), גוון, and תמונה (image embed). סעיף is auto-numbered.

Usage:
    python3 insert_fixture.py <xlsx_path> --category ceiling|wall|floor \
        --data '{"סימון":"SH-2","תיאור":"...","יצרן":"VELTRANO","מתח_עמעום":"48V\nDIM DALI","גוון":"הכל לבן"}' \
        [--image /path/to/photo.jpg] \
        [--uncertain גוון,תמונה] \
        [--out /path/to/output.xlsx]

--uncertain takes canonical field names (not column letters) — column
letters are not stable across project files.
"""
import argparse
import json
import re
import copy
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter, coordinate_to_tuple
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.styles import Alignment, PatternFill
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor
from openpyxl.cell.cell import MergedCell
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

from xlsx_io import WorkbookError, open_workbook, save_workbook

# Hebrew output on Windows too, when stdout is a pipe (the default there is cp1252)
for _s in (sys.stdout, sys.stderr):
    _s.reconfigure(encoding="utf-8", errors="replace")

SHEET_NAME = "אומדן"
HEADER_ROW = 8
CATEGORY_TEXT = {
    "ceiling": "מנורות הזנה מהתקרה",
    "wall": "מנורות הזנה מהקיר",
    "floor": "מנורות הזנה מהרצפה",
}
SUMMARY_MARKER = 'סה"כ כולל'   # substring identifying the summary row
TOTAL_HEADER = 'סה"כ'          # exact header text for the per-row total column

UNCERTAIN_FILL = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
LINE_HEIGHT_PT = 24.75   # calibrated against a real single (unmerged) legend row: 4 lines -> 99pt
MIN_ROW_HEIGHT = 30.75   # matches the skeleton's default single-row height

# Canonical fields the skill is allowed to write.
WRITABLE_FIELDS = {"סימון", "תיאור", "יצרן", "מתח_עמעום", "גוון"}


def normalize(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v not in (None, "") else ""


def classify_header(norm):
    """Map a normalized header-row cell to a canonical field name, or None if
    this is a column the skill doesn't know how to fill (and must leave alone)."""
    if norm.startswith("סעיף"):
        # prefix match: some real files have stray leftover text appended to
        # this header cell (seen in practice) — "סעיף" itself is unambiguous.
        return "seif"
    if norm == "סימון":
        return "סימון"
    if norm == "תיאור":
        return "תיאור"
    if norm == "יצרן":
        return "יצרן"
    if norm == "גוון":
        return "גוון"
    if norm == "תמונה":
        return "תמונה"
    if "מתח" in norm or "עמעום" in norm:
        return "מתח_עמעום"
    return None


def build_header_map(ws, header_row=HEADER_ROW):
    header_map = {}
    for col in range(1, ws.max_column + 1):
        raw = ws.cell(row=header_row, column=col).value
        if raw is None:
            continue
        canon = classify_header(normalize(raw))
        if canon and canon not in header_map:
            header_map[canon] = col
    return header_map


def find_category_blocks(ws):
    """Search the whole sheet for each category's header text (column position
    varies), plus the summary row. Returns {label: (header_row, first, last)}."""
    header_rows = {}
    summary_row = None
    max_col = ws.max_column
    for row in range(HEADER_ROW + 1, ws.max_row + 1):
        for col in range(1, max_col + 1):
            val = ws.cell(row=row, column=col).value
            if not val:
                continue
            val_clean = normalize(val)
            for label, text in CATEGORY_TEXT.items():
                if label not in header_rows and val_clean == text.strip():
                    header_rows[label] = row
            if summary_row is None and SUMMARY_MARKER in val_clean:
                summary_row = row

    if not header_rows:
        raise ValueError("No category headers found — check CATEGORY_TEXT against this workbook's actual text.")
    if summary_row is None:
        raise ValueError(f"Summary row (containing '{SUMMARY_MARKER}') not found.")

    ordered = sorted(header_rows.items(), key=lambda kv: kv[1])
    blocks = {}
    for i, (label, header_row) in enumerate(ordered):
        next_row = ordered[i + 1][1] if i + 1 < len(ordered) else summary_row
        blocks[label] = (header_row, header_row + 1, next_row - 1)
    return blocks, summary_row


def find_empty_row(ws, marker_col, first, last):
    """A row only counts as free if it's a genuinely independent cell. Some
    real project files still have legacy multi-row-merged fixture entries
    (from before this rule existed); a non-anchor cell of such a merge reads
    back as None but is NOT free space — it belongs to the fixture above it."""
    for row in range(first, last + 1):
        cell = ws.cell(row=row, column=marker_col)
        if isinstance(cell, MergedCell):
            continue
        if cell.value in (None, ""):
            return row
    return None


def copy_row_style(ws, src_row, dst_row, max_col):
    """Copy borders/fonts/alignment from the row above -- but NOT an orange
    "check this" fill. The row above is often a fixture that was just flagged;
    copying its fill painted confirmed values in the new row orange too, with no
    comment explaining why (found in testing, 10.2026)."""
    for col in range(1, max_col + 1):
        src = ws.cell(row=src_row, column=col)
        dst = ws.cell(row=dst_row, column=col)
        dst._style = copy.copy(src._style)
        if _is_uncertain_fill(dst.fill):
            dst.fill = PatternFill(fill_type=None)


def _is_uncertain_fill(fill):
    rgb = str(getattr(fill.fgColor, "rgb", "") or "").upper()
    return fill.fill_type == "solid" and rgb.endswith("FFC000")


def unmerge_for_insert(ws, insert_at):
    """openpyxl's insert_rows() leaves stale MergedCell placeholder objects
    behind (they don't move/convert cleanly), which then silently swallow any
    value written into them. The safe fix is to unmerge every affected range
    BEFORE insert_rows() runs (while the registry and cells are still
    consistent), and re-merge at corrected coordinates afterwards. Returns the
    list of (min_col, min_row, max_col, max_row, straddles) to re-merge."""
    to_restore = []
    for merged_range in list(ws.merged_cells.ranges):
        if merged_range.min_row >= insert_at:
            to_restore.append((merged_range.min_col, merged_range.min_row,
                                merged_range.max_col, merged_range.max_row, False))
            ws.unmerge_cells(str(merged_range))
        elif merged_range.max_row >= insert_at:
            # range straddles the insertion point -> extend it instead of shifting it
            to_restore.append((merged_range.min_col, merged_range.min_row,
                                merged_range.max_col, merged_range.max_row, True))
            ws.unmerge_cells(str(merged_range))
    return to_restore


def remerge_after_insert(ws, to_restore, n=1):
    for min_col, min_row, max_col, max_row, straddles in to_restore:
        new_min_row = min_row if straddles else min_row + n
        new_max_row = max_row + n
        ws.merge_cells(start_row=new_min_row, start_column=min_col,
                        end_row=new_max_row, end_column=max_col)


def realize_image_anchor(img):
    """openpyxl's ws.add_image() just stores the string anchor ('G11') as-is
    on img.anchor -- it's only converted into a real OneCellAnchor object at
    save time (see _check_anchor in openpyxl.drawing.spreadsheet_drawing),
    which normally doesn't matter. But this skill's own row-shift helpers
    (shift_images_for_insert / _for_delete) need a real object with a
    mutable ._from.row to move an image when a row above it is inserted or
    removed later in the SAME run -- e.g. this fixture's own just-embedded
    photo, when the skeleton-row cleanup below deletes a row above it.
    Against a plain string that mutation is a silent no-op. Do the same
    conversion _check_anchor would do, immediately, so every image (freshly
    added or reloaded from a previous save) behaves the same way from here
    on."""
    anchor = img.anchor
    if isinstance(anchor, str):
        row, col = coordinate_to_tuple(anchor.upper())
        new_anchor = OneCellAnchor()
        new_anchor._from.row = row - 1
        new_anchor._from.col = col - 1
        if getattr(img, "width", None) and getattr(img, "height", None):
            new_anchor.ext.width = pixels_to_EMU(img.width)
            new_anchor.ext.height = pixels_to_EMU(img.height)
        img.anchor = new_anchor


def shift_images_for_insert(ws, insert_at, n=1):
    """Embedded fixture photos have the same problem as merged ranges:
    insert_rows() moves cell CONTENT down but leaves every image's anchor
    pinned to its original row. Left unfixed, a later fixture forced into the
    same category silently detaches every earlier photo from its row — the
    text drifts down, the photo doesn't. insert_at is a 1-indexed Excel row."""
    for img in getattr(ws, "_images", []):
        anchor = img.anchor
        frm = getattr(anchor, "_from", None)
        if frm is not None and frm.row >= insert_at - 1:  # anchor rows are 0-indexed
            frm.row += n
        to = getattr(anchor, "to", None)
        if to is not None and to.row >= insert_at - 1:
            to.row += n


def shift_row_heights_for_insert(ws, insert_at, n=1):
    """insert_rows() moves cell CONTENT down but does NOT move row_dimensions
    (height, hidden, outline level) at all -- every such entry stays pinned to
    its old row number while the data now silently sits one row further down.
    Left unfixed, any row below an insertion point ends up with a height that
    belonged to whatever used to occupy that row number, not the content now
    there -- e.g. a later fixture's description row can inherit a stray
    leftover 12.75pt (Excel's unset default) from a spacer row that used to be
    there, clipping its product photo down to invisible even though the
    anchor and image data are both correct. Shift explicitly, highest row
    first so we don't clobber not-yet-moved entries, then clear the height of
    the freshly inserted (still blank) row itself so it doesn't inherit a
    stale value either."""
    existing_rows = sorted([r for r in ws.row_dimensions.keys() if r >= insert_at], reverse=True)
    for r in existing_rows:
        src = ws.row_dimensions[r]
        dst = ws.row_dimensions[r + n]
        dst.height = src.height
        dst.hidden = src.hidden
        dst.outlineLevel = src.outlineLevel
    for r in range(insert_at, insert_at + n):
        ws.row_dimensions[r].height = None


def unmerge_for_delete(ws, del_row, n=1):
    """Mirror of unmerge_for_insert, for ws.delete_rows(). Returns a list of
    (min_col, min_row, max_col, max_row) to re-merge after the delete, with
    row coordinates already corrected. A range entirely inside the deleted
    rows is simply dropped (not restored)."""
    to_restore = []
    del_last = del_row + n - 1
    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))
        min_row, max_row = merged_range.min_row, merged_range.max_row
        if max_row < del_row:
            to_restore.append((merged_range.min_col, min_row, merged_range.max_col, max_row))
        elif min_row > del_last:
            to_restore.append((merged_range.min_col, min_row - n, merged_range.max_col, max_row - n))
        else:
            new_min = min_row if min_row < del_row else del_row
            new_max = max_row - n if max_row > del_last else del_row - 1
            if new_max >= new_min:
                to_restore.append((merged_range.min_col, new_min, merged_range.max_col, new_max))
    return to_restore


def remerge_after_delete(ws, to_restore):
    for min_col, min_row, max_col, max_row in to_restore:
        ws.merge_cells(start_row=min_row, start_column=min_col,
                        end_row=max_row, end_column=max_col)


def shift_images_for_delete(ws, del_row, n=1):
    """Mirror of shift_images_for_insert. Any image anchored inside the
    deleted rows is dropped along with them; anything below moves up."""
    del_last0 = del_row + n - 2  # 0-indexed last deleted row
    del_first0 = del_row - 1
    keep = []
    for img in getattr(ws, "_images", []):
        anchor = img.anchor
        frm = getattr(anchor, "_from", None)
        if frm is not None and del_first0 <= frm.row <= del_last0:
            continue  # anchored inside the deleted range -> drop the image
        if frm is not None and frm.row > del_last0:
            frm.row -= n
        to = getattr(anchor, "to", None)
        if to is not None and to.row > del_last0:
            to.row -= n
        keep.append(img)
    ws._images = keep


def shift_row_heights_for_delete(ws, del_row, n=1):
    """Mirror of shift_row_heights_for_insert."""
    del_last = del_row + n - 1
    for r in range(del_row, del_last + 1):
        if r in ws.row_dimensions:
            del ws.row_dimensions[r]
    existing_rows = sorted(r for r in ws.row_dimensions.keys() if r > del_last)
    for r in existing_rows:
        src = ws.row_dimensions[r]
        dst = ws.row_dimensions[r - n]
        dst.height = src.height
        dst.hidden = src.hidden
        dst.outlineLevel = src.outlineLevel
        del ws.row_dimensions[r]


# The example row baked into the base template carries this exact placeholder text.
_SKELETON_MARKERS = ("[סוג הגוף]", "[שם הדגם]")


def _is_skeleton_text(text):
    """The template's example row is recognised by its fixed placeholder text, NOT by a
    leading "[": a real description can legitimately start with a bracket ("[Ø55] ..."),
    and the old "starts with [" test deleted such a fixture."""
    t = text.strip()
    return t.startswith(_SKELETON_MARKERS[0]) or _SKELETON_MARKERS[1] in t


def strip_skeleton_rows(ws):
    """Remove any leftover template/example row (identified by the exact placeholder
    text of the base skeleton, "[סוג הגוף]" / "[שם הדגם]") from every category block,
    so a delivered legend only ever shows real fixtures that were actually added.
    Safe to call unconditionally on every save: a real description that starts with a
    bracket, e.g. "[Ø55] ...", is kept. Returns the list of removed row numbers (as
    they were before the removal), so a caller can correct a row number it is holding."""
    header_map = build_header_map(ws)
    desc_col = header_map.get("תיאור")
    if not desc_col:
        return []

    blocks, _ = find_category_blocks(ws)
    to_delete = []
    for _, (_, first, last) in blocks.items():
        for r in range(first, last + 1):
            cell = ws.cell(row=r, column=desc_col)
            if isinstance(cell, MergedCell):
                continue
            v = cell.value
            if isinstance(v, CellRichText):
                text = "".join(p.text if hasattr(p, "text") else str(p) for p in v)
            else:
                text = str(v) if v is not None else ""
            if _is_skeleton_text(text):
                to_delete.append(r)

    for r in sorted(to_delete, reverse=True):
        pending = unmerge_for_delete(ws, r)
        shift_images_for_delete(ws, r)
        shift_row_heights_for_delete(ws, r)
        ws.delete_rows(r, 1)
        remerge_after_delete(ws, pending)

    if not to_delete:
        return []

    # renumber סעיף within each block now that skeleton rows are gone, and
    # keep the summary SUM formula's range valid
    blocks, summary_row = find_category_blocks(ws)
    seif_col = header_map.get("seif")
    if seif_col:
        for _, (_, first, last) in blocks.items():
            n = 1
            for r in range(first, last + 1):
                c = ws.cell(row=r, column=seif_col)
                if isinstance(c, MergedCell) or c.value in (None, ""):
                    continue
                c.value = float(n)
                n += 1

    max_col = ws.max_column
    total_col = None
    for col in range(1, max_col + 1):
        if normalize(ws.cell(row=HEADER_ROW, column=col).value) == TOTAL_HEADER:
            total_col = col
            break
    if total_col:
        first_block_first_row = min(b[1] for b in blocks.values())
        new_last_row = summary_row - 1
        col_letter = get_column_letter(total_col)
        cell = ws.cell(row=summary_row, column=total_col)
        if isinstance(cell.value, str) and cell.value.upper().startswith("=SUM("):
            cell.value = f"=SUM({col_letter}{first_block_first_row}:{col_letter}{new_last_row})"

    return to_delete


REVIEW_NOTE = (
    "לעובד/ת: נא לעבור על כל השורות המסומנות בכתום בקובץ, לוודא את הנתונים "
    "ולתקן במידת הצורך, ולהסיר את הצביעה הכתומה מתא שנבדק. לאחר שהאקסל מתוקן "
    "ומאושר מבחינתכם, נא לשלוח אותו חזרה. "
    "(גוף חדש עם תא שעדיין כתום לא ייכנס לבנק — כדי שערך שלא נבדק לא יתפזר לפרויקטים הבאים.)"
)
MISSING_NOTE = "חסר: לא התקבל ערך לשדה הזה. נא להשלים ידנית — לא מנחשים."
NO_IMAGE_NOTE = "חסר: לא צורפה תמונת מוצר. נא להוסיף ידנית."


def add_review_note(ws):
    """Attach the standing review request as a comment in the header area, so
    it travels with the file itself (Excel, email, WhatsApp — not just this
    chat) and doesn't disturb the studio's fixed visual layout. A2 is used
    deliberately rather than a header field cell like C2 (project name) or C5
    (light colour) -- those can already carry their own "missing value"
    warning comment, which a blanket overwrite here would silently erase."""
    cell = ws["A2"]
    cell.comment = Comment(REVIEW_NOTE, "mikra skill")


def estimate_row_height(text, col_width):
    if not text:
        return MIN_ROW_HEIGHT
    chars_per_line = max(int(col_width or 33), 10)
    visual_lines = 0
    for line in str(text).split("\n"):
        ln = len(line)
        visual_lines += max(1, -(-ln // chars_per_line))  # ceil division
    return max(MIN_ROW_HEIGHT, visual_lines * LINE_HEIGHT_PT)


UNCERTAIN_FONT = InlineFont(color="FFC000", b=True)


def build_rich_text(value, snippet):
    """Highlight just the uncertain snippet within a longer field (typically תיאור),
    instead of flagging the whole cell — so a shaky beam-angle line doesn't turn an
    otherwise-confirmed multi-line description entirely orange. Returns a
    CellRichText with only `snippet` bolded/coloured, or None if `snippet` isn't
    found verbatim in `value` (caller should fall back to whole-cell fill)."""
    value = str(value)
    idx = value.find(snippet)
    if idx == -1:
        return None
    before, after = value[:idx], value[idx + len(snippet):]
    parts = []
    if before:
        parts.append(before)
    parts.append(TextBlock(UNCERTAIN_FONT, snippet))
    if after:
        parts.append(after)
    return CellRichText(*parts)


def next_seif_number(ws, seif_col, first, last):
    nums = [ws.cell(row=r, column=seif_col).value for r in range(first, last + 1)]
    nums = [n for n in nums if isinstance(n, (int, float))]
    return (max(nums) + 1) if nums else 1


def insert_fixture(xlsx_path, category, data, image_path=None, uncertain_fields=None,
                    uncertain_snippets=None, out_path=None):
    uncertain_fields = set(uncertain_fields or [])
    uncertain_snippets = dict(uncertain_snippets or {})
    # rich_text=True is required or openpyxl silently flattens any rich-text
    # highlighting written by an earlier run of this script the moment this
    # workbook gets reloaded and resaved for a later, unrelated fixture
    if category not in CATEGORY_TEXT:
        raise ValueError(f"Unknown category '{category}'. Use one of: {list(CATEGORY_TEXT)}")
    wb = open_workbook(xlsx_path, sheet=SHEET_NAME, rich_text=True)
    ws = wb[SHEET_NAME]
    max_col = ws.max_column

    header_map = build_header_map(ws)

    unknown_keys = set(data) - WRITABLE_FIELDS
    if unknown_keys:
        raise ValueError(f"Unsupported data keys (this skill never writes these): {unknown_keys}")
    missing_cols = [k for k in data if k not in header_map]
    if missing_cols:
        raise ValueError(f"Workbook header row (row {HEADER_ROW}) has no column recognized for: {missing_cols}")

    # A field the sheet has a column for, but that came in empty (or not at all),
    # is flagged -- an empty cell that LOOKS finished is how a gap goes unnoticed.
    # סימון is the exception: it's a project-specific code, blank by design.
    data = dict(data)
    missing_fields = set()
    for field in ("תיאור", "יצרן", "מתח_עמעום", "גוון"):
        if field in header_map and not str(data.get(field) or "").strip():
            data[field] = ""
            missing_fields.add(field)

    blocks, summary_row = find_category_blocks(ws)
    if category not in blocks:
        raise ValueError(f"Category '{category}' not found. Available: {list(blocks)}")
    header_row, first, last = blocks[category]

    desc_col = header_map.get("תיאור")
    target_row = find_empty_row(ws, desc_col, first, last) if desc_col else None

    if target_row is None:
        # block is full: insert a new row right before the next section (or the summary row)
        insert_at = last + 1
        pending_merges = unmerge_for_insert(ws, insert_at)
        ws.insert_rows(insert_at)
        remerge_after_insert(ws, pending_merges)
        shift_images_for_insert(ws, insert_at)
        shift_row_heights_for_insert(ws, insert_at)
        copy_row_style(ws, last, insert_at, max_col)
        target_row = insert_at
        last += 1
        summary_row += 1  # everything from insert_at down (incl. summary row) shifted by 1

    seif_col = header_map.get("seif")
    if seif_col:
        seif = next_seif_number(ws, seif_col, first, last)
        ws.cell(row=target_row, column=seif_col, value=float(seif))

    for field, value in data.items():
        col = header_map[field]
        snippet = uncertain_snippets.get(field)
        rich = build_rich_text(value, snippet) if snippet else None

        if rich is not None:
            # only the flagged phrase is marked (bold/orange text) — the rest of the
            # cell (and its background) stays exactly as any confirmed value would
            cell = ws.cell(row=target_row, column=col, value=rich)
            cell.comment = Comment(f'לבדיקה: "{snippet}" — לא נמצא בבירור בPDF, נא לאמת ידנית.', "mikra skill")
        else:
            cell = ws.cell(row=target_row, column=col, value=value)
            # snippet was given but not found verbatim in value -> fall back to
            # flagging the whole cell, regardless of whether it's also in
            # uncertain_fields (a caller shouldn't have to pass both)
            if field in missing_fields:
                cell.fill = UNCERTAIN_FILL
                cell.comment = Comment(MISSING_NOTE, "mikra skill")
            elif field in uncertain_fields or snippet is not None:
                cell.fill = UNCERTAIN_FILL
                if snippet is not None:
                    cell.comment = Comment(
                        f'לבדיקה: הביטוי "{snippet}" לא נמצא בטקסט בדיוק — נא לאמת את כל השדה.',
                        "mikra skill")
                else:
                    cell.comment = Comment("לבדיקה: לא נמצא בבירור בPDF, נא לאמת ידנית.", "mikra skill")
        cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True, readingOrder=2)

    # row height sized to fit the description (no multi-row merging — single row per fixture)
    desc_val = data.get("תיאור", "")
    col_width = 33
    if desc_col:
        col_width = ws.column_dimensions[get_column_letter(desc_col)].width or 33
    ws.row_dimensions[target_row].height = estimate_row_height(desc_val, col_width)

    image_col = header_map.get("תמונה")
    if image_path and image_col:
        img = XLImage(image_path)
        row_h_px = (ws.row_dimensions[target_row].height or MIN_ROW_HEIGHT) * 96 / 72
        col_w_chars = ws.column_dimensions[get_column_letter(image_col)].width or 20
        col_w_px = col_w_chars * 7 + 5          # Excel's character-width -> pixel rule of thumb
        # fit inside the cell both ways: scaling by row height alone let a wide photo
        # (a linear profile) spill over the neighbouring columns
        scale = min(row_h_px / img.height, col_w_px / img.width)
        img.width, img.height = img.width * scale * 0.9, img.height * scale * 0.9
        anchor = f"{get_column_letter(image_col)}{target_row}"
        ws.add_image(img, anchor)
        realize_image_anchor(img)
        if "תמונה" in uncertain_fields:
            img_cell = ws.cell(row=target_row, column=image_col)
            img_cell.fill = UNCERTAIN_FILL
            img_cell.comment = Comment(
                "לבדיקה: לוודא שזו התמונה הנכונה של הגוף.", "mikra skill")
    elif image_path and not image_col:
        print("WARNING: workbook has no 'תמונה' column recognized — image not embedded.")
    elif image_col and not image_path:
        img_cell = ws.cell(row=target_row, column=image_col)
        img_cell.fill = UNCERTAIN_FILL
        img_cell.comment = Comment(NO_IMAGE_NOTE, "mikra skill")

    # keep the summary SUM formula's range valid after any insertion
    total_col = None
    for col in range(1, max_col + 1):
        if normalize(ws.cell(row=HEADER_ROW, column=col).value) == TOTAL_HEADER:
            total_col = col
            break
    if total_col:
        first_block_first_row = min(b[1] for b in blocks.values())
        new_last_row = summary_row - 1
        col_letter = get_column_letter(total_col)
        cell = ws.cell(row=summary_row, column=total_col)
        if isinstance(cell.value, str) and cell.value.upper().startswith("=SUM("):
            cell.value = f"=SUM({col_letter}{first_block_first_row}:{col_letter}{new_last_row})"

    # every save is a potential export to the person doing the work: strip any
    # leftover skeleton/example rows and make sure the standing review note
    # is attached, automatically, every time
    removed_rows = strip_skeleton_rows(ws)
    removed = len(removed_rows)
    # the returned row number must point at the fixture AFTER the cleanup: deleting
    # the template row above it moves it up (callers pass this row on, e.g. to
    # mark_bank_status.py -- an off-by-one would comment the wrong fixture)
    target_row -= sum(1 for r in removed_rows if r < target_row)
    add_review_note(ws)

    out_path = out_path or xlsx_path
    save_workbook(wb, out_path)
    msg = f"Inserted row {target_row} in block '{category}'."  # final position
    if removed:
        msg += f" Removed {removed} leftover skeleton row(s)."
    if missing_fields:
        msg += f" Flagged empty field(s): {', '.join(sorted(missing_fields))}."
    msg += f" Saved to {out_path}"
    print(msg)
    return target_row


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx_path")
    ap.add_argument("--category", required=True, choices=list(CATEGORY_TEXT))
    ap.add_argument("--data", required=True, help="JSON dict of canonical field name -> value")
    ap.add_argument("--image", default=None)
    ap.add_argument("--uncertain", default="", help="comma-separated canonical field names, e.g. גוון,תמונה")
    ap.add_argument("--uncertain-snippets", default="",
                     help='JSON dict of field name -> exact substring to flag within that field '
                          '(e.g. \'{"תיאור":"זוית פיזור 38 מעלות"}\') — highlights only that '
                          'phrase instead of the whole cell; takes priority over --uncertain '
                          'for the same field')
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    try:
        data = json.loads(args.data)
        snippets = json.loads(args.uncertain_snippets) if args.uncertain_snippets else {}
    except json.JSONDecodeError as e:
        raise SystemExit(f"--data / --uncertain-snippets is not valid JSON: {e}")
    uncertain = [f.strip() for f in args.uncertain.split(",") if f.strip()]
    if args.image and not Path(args.image).exists():
        raise SystemExit(f"Image not found: {args.image}")
    try:
        insert_fixture(args.xlsx_path, args.category, data, args.image, uncertain, snippets, args.out)
    except (WorkbookError, ValueError) as e:
        raise SystemExit(f"שגיאה: {e}")


if __name__ == "__main__":
    main()
