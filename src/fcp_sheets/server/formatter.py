"""Response formatting utilities for sheets queries."""

from __future__ import annotations

from datetime import datetime

from openpyxl.cell import Cell
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

from fcp_sheets.model.refs import index_to_col


def format_cell_addr(col: int, row: int) -> str:
    """Format a (col, row) pair as an A1-style address."""
    return f"{index_to_col(col)}{row}"


def format_range(
    min_row: int, min_col: int, max_row: int, max_col: int
) -> str:
    """Format a bounding rectangle as A1:Z99."""
    return f"{index_to_col(min_col)}{min_row}:{index_to_col(max_col)}{max_row}"


def truncate_list(items: list[str], max_items: int = 8) -> str:
    """Join items, showing 'and N more' if truncated."""
    if len(items) <= max_items:
        return ", ".join(items)
    shown = ", ".join(items[:max_items])
    return f"{shown} ... +{len(items) - max_items} more"


# ---------------------------------------------------------------------------
# Extended formatters for Wave 3 queries
# ---------------------------------------------------------------------------


def format_cell_value(value) -> str:
    """Format a cell value for display, keeping it token-efficient."""
    if value is None:
        return "(empty)"
    if isinstance(value, str):
        if value.startswith("="):
            return value
        return repr(value) if len(value) > 30 else f'"{value}"'
    if isinstance(value, bool):
        return str(value).upper()
    if isinstance(value, float):
        # Show as int if whole number
        if value == int(value):
            return str(int(value))
        return f"{value:g}"
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return str(value)


def format_value_type(value) -> str:
    """Return a short type label for a cell value."""
    if value is None:
        return "empty"
    if isinstance(value, str):
        if value.startswith("="):
            return "formula"
        return "text"
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, datetime):
        return "date"
    return "other"


def format_font(font: Font) -> str:
    """Format font information token-efficiently."""
    parts: list[str] = []
    name = font.name or "Calibri"
    parts.append(name)
    if font.size:
        parts.append(f"{int(font.size)}pt")
    flags: list[str] = []
    if font.bold:
        flags.append("bold")
    if font.italic:
        flags.append("italic")
    if font.underline:
        flags.append("underline")
    if font.strike:
        flags.append("strike")
    if flags:
        parts.append(" ".join(flags))
    if font.color and font.color.type == "rgb":
        rgb = font.color.rgb
        if isinstance(rgb, str) and rgb != "00000000":
            # openpyxl stores as AARRGGBB
            if len(rgb) == 8:
                parts.append(f"#{rgb[2:]}")
            else:
                parts.append(f"#{rgb}")
    return " ".join(parts)


def format_fill(fill: PatternFill) -> str:
    """Format fill information."""
    if fill.fill_type is None or fill.fill_type == "none":
        return "(none)"
    color = fill.start_color
    if color and color.type == "rgb":
        rgb = color.rgb
        if isinstance(rgb, str) and rgb != "00000000":
            if len(rgb) == 8:
                hex_str = f"#{rgb[2:]}"
            else:
                hex_str = f"#{rgb}"
            return f"{hex_str} ({fill.fill_type})"
    return f"({fill.fill_type})"


def format_alignment(alignment: Alignment) -> str:
    """Format alignment information."""
    parts: list[str] = []
    if alignment.horizontal:
        parts.append(alignment.horizontal)
    if alignment.vertical and alignment.vertical != "bottom":
        parts.append(f"v:{alignment.vertical}")
    if alignment.wrap_text:
        parts.append("wrap")
    if alignment.indent:
        parts.append(f"indent:{alignment.indent}")
    if alignment.text_rotation:
        parts.append(f"rotate:{alignment.text_rotation}")
    return " ".join(parts) if parts else "(default)"


def format_border_side(side: Side) -> str:
    """Format a single border side."""
    if side is None or side.style is None:
        return ""
    parts = [side.style]
    if side.color and side.color.type == "rgb":
        rgb = side.color.rgb
        if isinstance(rgb, str) and rgb != "00000000":
            if len(rgb) == 8:
                parts.append(f"#{rgb[2:]}")
            else:
                parts.append(f"#{rgb}")
    return " ".join(parts)


def format_border(border: Border) -> str:
    """Format border information."""
    sides: list[str] = []
    for name, side in [
        ("top", border.top),
        ("bottom", border.bottom),
        ("left", border.left),
        ("right", border.right),
    ]:
        desc = format_border_side(side)
        if desc:
            sides.append(f"{name} {desc}")
    return ", ".join(sides) if sides else "(none)"


def format_table_row(values: list[str], col_widths: list[int] | None = None) -> str:
    """Format a list of values as a pipe-delimited table row.

    If col_widths provided, pad each cell.  Otherwise just join with |.
    """
    if col_widths:
        parts = []
        for i, val in enumerate(values):
            w = col_widths[i] if i < len(col_widths) else len(val)
            parts.append(val.ljust(w))
        return "|".join(parts)
    return "|".join(values)
