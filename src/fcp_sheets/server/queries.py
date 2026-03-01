"""Query handlers — read-only inspection of workbook state.

MVP queries: plan/map, stats, status, history.
Extended queries (Wave 3): describe, peek, list, find.
"""

from __future__ import annotations

import re
from datetime import datetime

from fcp_sheets.model.index import SheetIndex
from fcp_sheets.model.refs import (
    index_to_col,
    col_to_index,
    parse_cell_ref,
    parse_range_ref,
    CellRef,
    RangeRef,
)
from fcp_sheets.model.snapshot import SheetsModel
from fcp_sheets.server.formatter import (
    format_range,
    format_cell_addr,
    truncate_list,
    format_cell_value,
    format_value_type,
    format_font,
    format_fill,
    format_alignment,
    format_border,
    format_table_row,
)


def dispatch_query(query: str, model: SheetsModel, index: SheetIndex) -> str:
    """Route a query string to the appropriate handler."""
    query = query.strip()
    parts = query.split(None, 1)
    command = parts[0].lower() if parts else ""
    args = parts[1] if len(parts) > 1 else ""

    handlers = {
        "plan": _query_plan,
        "map": _query_plan,
        "stats": _query_stats,
        "status": _query_status,
        "history": _query_history,
        "describe": _query_describe,
        "peek": _query_peek,
        "list": _query_list,
        "find": _query_find,
    }

    handler = handlers.get(command)
    if handler is None:
        return (
            f"! Unknown query: {command!r}\n"
            "  try: plan, stats, status, history, describe, peek, list, find"
        )

    return handler(args, model, index)


# ---------------------------------------------------------------------------
# plan / map — primary overview
# ---------------------------------------------------------------------------

def _query_plan(args: str, model: SheetsModel, index: SheetIndex) -> str:
    """Show workbook topology — the most important query."""
    wb = model.wb
    lines: list[str] = []

    # Header
    saved = f", saved: {model.file_path}" if model.file_path else ", unsaved"
    lines.append(f'Workbook: "{model.title}" ({len(wb.sheetnames)} sheets{saved})')

    for ws in wb.worksheets:
        sheet_name = ws.title
        active = " [active]" if ws == wb.active else ""
        hidden = " [hidden]" if ws.sheet_state == "hidden" else ""
        lines.append(f"\n  Sheet: {sheet_name}{active}{hidden}")

        bounds = index.get_bounds(sheet_name)
        if bounds:
            min_row, min_col, max_row, max_col = bounds
            range_str = format_range(min_row, min_col, max_row, max_col)

            # Data bounds + freeze + filter
            meta_parts = [f"data: {range_str}"]
            if ws.freeze_panes:
                meta_parts.append(f"frozen: {ws.freeze_panes}")
            if ws.auto_filter and ws.auto_filter.ref:
                meta_parts.append(f"filter: {ws.auto_filter.ref}")
            lines.append(f"    {' | '.join(meta_parts)}")

            # Column headers (first row)
            cols = []
            for col in range(min_col, min(max_col + 1, min_col + 8)):
                cell = ws.cell(row=min_row, column=col)
                val = cell.value
                if val is not None:
                    cols.append(f"{index_to_col(col)}:{val}")
            if cols:
                lines.append(f"    cols: {truncate_list(cols)}")

            # Formula patterns
            formula_patterns = _detect_formula_patterns(ws, bounds)
            if formula_patterns:
                lines.append(f"    formulas: {' | '.join(formula_patterns)}")

            # Tables
            if ws.tables:
                for table_name in ws.tables:
                    table = ws.tables[table_name]
                    lines.append(f"    table: {table_name} {table.ref}")

            # Charts
            if ws._charts:
                for chart in ws._charts:
                    chart_title = chart.title or "Untitled"
                    chart_type = type(chart).__name__
                    lines.append(f"    chart: {chart_type} \"{chart_title}\"")

            # Conditional formatting
            if ws.conditional_formatting:
                cf_count = len(list(ws.conditional_formatting))
                if cf_count > 0:
                    lines.append(f"    cond-fmt: {cf_count} rule(s)")

            # Next empty row
            lines.append(f"    next-empty: row:{max_row + 1}")
        else:
            lines.append("    (empty)")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# stats
# ---------------------------------------------------------------------------

def _query_stats(args: str, model: SheetsModel, index: SheetIndex) -> str:
    """Quick summary of workbook contents."""
    wb = model.wb
    total_data = 0
    total_formula = 0
    total_charts = 0
    total_tables = 0
    total_merged = 0
    total_cond_fmt = 0

    for ws in wb.worksheets:
        bounds = index.get_bounds(ws.title)
        if bounds:
            min_row, min_col, max_row, max_col = bounds
            for row in ws.iter_rows(
                min_row=min_row, max_row=max_row,
                min_col=min_col, max_col=max_col,
            ):
                for cell in row:
                    if cell.value is not None:
                        if isinstance(cell.value, str) and cell.value.startswith("="):
                            total_formula += 1
                        else:
                            total_data += 1

        total_charts += len(ws._charts)
        total_tables += len(ws.tables)
        total_merged += len(ws.merged_cells.ranges)
        total_cond_fmt += len(list(ws.conditional_formatting))

    named_ranges = len(list(wb.defined_names)) if wb.defined_names else 0

    lines = [
        f'Workbook: "{model.title}"',
        f"  Sheets: {len(wb.sheetnames)} ({', '.join(wb.sheetnames)})",
        f"  Data cells: {total_data:,} | Formula cells: {total_formula:,}",
        f"  Tables: {total_tables} | Charts: {total_charts} | Named ranges: {named_ranges}",
        f"  Merged regions: {total_merged} | Conditional formats: {total_cond_fmt}",
    ]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# status
# ---------------------------------------------------------------------------

def _query_status(args: str, model: SheetsModel, index: SheetIndex) -> str:
    """Session status."""
    saved = model.file_path or "unsaved"
    active = model.wb.active.title if model.wb.active else "none"
    return (
        f'Session: "{model.title}"\n'
        f"  File: {saved}\n"
        f"  Sheets: {len(model.wb.sheetnames)}\n"
        f"  Active: {active}"
    )


# ---------------------------------------------------------------------------
# history
# ---------------------------------------------------------------------------

def _query_history(args: str, model: SheetsModel, index: SheetIndex) -> str:
    """Show recent operations — managed by session layer."""
    return "History managed by session layer. Use sheets_session for undo/redo."


# ---------------------------------------------------------------------------
# describe — detailed view of sheet, cell, or range
# ---------------------------------------------------------------------------

def _query_describe(args: str, model: SheetsModel, index: SheetIndex) -> str:
    """Describe a sheet, cell, or range."""
    args = args.strip()
    if not args:
        return "! Usage: describe SHEET | describe CELL | describe RANGE"

    wb = model.wb

    # Check if it's a sheet name (exact match first, then case-insensitive)
    if args in wb.sheetnames:
        return _describe_sheet(args, model, index)
    for name in wb.sheetnames:
        if name.lower() == args.lower():
            return _describe_sheet(name, model, index)

    # Check if it's a cell reference (possibly with sheet prefix)
    cell_ref = parse_cell_ref(args)
    if cell_ref:
        return _describe_cell(cell_ref, model, index)

    # Check if it's a range reference
    range_ref = parse_range_ref(args)
    if isinstance(range_ref, RangeRef):
        return _describe_range(range_ref, model, index)

    return f"! Cannot resolve: {args!r}. Provide a sheet name, cell (A1), or range (A1:D10)."


def _describe_sheet(sheet_name: str, model: SheetsModel, index: SheetIndex) -> str:
    """Detailed sheet description."""
    wb = model.wb
    ws = wb[sheet_name]
    lines: list[str] = [f"Sheet: {sheet_name}"]

    bounds = index.get_bounds(sheet_name)
    if not bounds:
        lines.append("  (empty)")
        return "\n".join(lines)

    min_row, min_col, max_row, max_col = bounds
    range_str = format_range(min_row, min_col, max_row, max_col)

    # Meta line
    meta_parts = [f"data: {range_str}"]
    if ws.freeze_panes:
        meta_parts.append(f"frozen: {ws.freeze_panes}")
    if ws.auto_filter and ws.auto_filter.ref:
        meta_parts.append(f"filter: {ws.auto_filter.ref}")
    lines.append(f"  {' | '.join(meta_parts)}")

    # Column analysis
    lines.append("  columns:")
    for col in range(min_col, max_col + 1):
        col_letter = index_to_col(col)
        header_val = ws.cell(row=min_row, column=col).value
        header = str(header_val) if header_val is not None else ""

        # Scan column for type info (skip header row)
        values: list = []
        formulas: list[str] = []
        unique_texts: set[str] = set()
        nums: list[float] = []
        data_count = 0

        for row in range(min_row + 1, max_row + 1):
            val = ws.cell(row=row, column=col).value
            if val is None:
                continue
            data_count += 1
            if isinstance(val, str) and val.startswith("="):
                formulas.append(val)
            elif isinstance(val, (int, float)) and not isinstance(val, bool):
                nums.append(float(val))
            elif isinstance(val, str):
                unique_texts.add(val)
            values.append(val)

        # Determine dominant type
        if formulas and len(formulas) >= len(values) / 2:
            col_type = "formula"
        elif nums and len(nums) >= len(values) / 2:
            col_type = "number"
        else:
            col_type = "text"

        # Build column line
        header_display = header[:12] if header else "(no header)"
        col_line = f"    {col_letter:<3}{header_display:<14}{col_type:<9}{data_count} values"

        if col_type == "text" and unique_texts:
            col_line += f"   unique:{len(unique_texts)}"
        elif col_type == "number" and nums:
            mn = int(min(nums)) if min(nums) == int(min(nums)) else min(nums)
            mx = int(max(nums)) if max(nums) == int(max(nums)) else max(nums)
            col_line += f"   range:{mn}..{mx}"
        elif col_type == "formula" and formulas:
            # Show pattern from first formula
            pattern = re.sub(r"\d+", "N", formulas[0])
            col_line += f"   pattern:{pattern}"
            if nums:
                mn = int(min(nums)) if min(nums) == int(min(nums)) else min(nums)
                mx = int(max(nums)) if max(nums) == int(max(nums)) else max(nums)
                col_line += f"  range:{mn}..{mx}"

        lines.append(col_line)

    # Sample data (rows 2-4, i.e. first 3 data rows after header)
    sample_start = min_row + 1
    sample_end = min(min_row + 3, max_row)
    if sample_start <= max_row:
        lines.append(f"  sample (rows {sample_start}-{sample_end}):")
        # Build header
        headers = []
        for col in range(min_col, max_col + 1):
            val = ws.cell(row=min_row, column=col).value
            headers.append(str(val) if val is not None else index_to_col(col))

        col_widths = [max(8, len(h)) for h in headers]
        header_line = "    | " + " | ".join(h.ljust(w) for h, w in zip(headers, col_widths)) + " |"
        lines.append(header_line)

        for row in range(sample_start, sample_end + 1):
            row_vals = []
            for col in range(min_col, max_col + 1):
                val = ws.cell(row=row, column=col).value
                row_vals.append(_compact_value(val))
            row_line = "    | " + " | ".join(v.ljust(w) for v, w in zip(row_vals, col_widths)) + " |"
            lines.append(row_line)

    # Formula groupings
    formula_groups = _get_formula_groups(ws, bounds)
    if formula_groups:
        lines.append("  formulas:")
        for group in formula_groups[:5]:
            lines.append(f"    {group}")

    # Merged regions
    merged = list(ws.merged_cells.ranges)
    if merged:
        merged_strs = [str(m) for m in merged]
        lines.append(f"  merged: {truncate_list(merged_strs, 5)}")
    else:
        lines.append("  merged: (none)")

    # Conditional formatting
    cf_rules = list(ws.conditional_formatting)
    if cf_rules:
        for cf in cf_rules[:5]:
            # cf is a ConditionalFormattingList entry
            range_str_cf = str(cf)
            for rule in cf.rules:
                rule_type = _describe_cf_rule(rule)
                lines.append(f"  cond-fmt: {range_str_cf} {rule_type}")
        if len(cf_rules) > 5:
            lines.append(f"  ... +{len(cf_rules) - 5} more rules")
    else:
        lines.append("  cond-fmt: (none)")

    # Tables
    if ws.tables:
        for table_name in ws.tables:
            table = ws.tables[table_name]
            style_name = ""
            if table.tableStyleInfo and table.tableStyleInfo.name:
                style_name = f" style:{table.tableStyleInfo.name}"
            banded = ""
            if table.tableStyleInfo:
                flags = []
                if table.tableStyleInfo.showRowStripes:
                    flags.append("banded-rows")
                if table.tableStyleInfo.showColumnStripes:
                    flags.append("banded-cols")
                if flags:
                    banded = " " + " ".join(flags)
            lines.append(f"  table: {table_name} {table.ref}{style_name}{banded}")
    else:
        lines.append("  tables: (none)")

    return "\n".join(lines)


def _describe_cell(ref: CellRef, model: SheetsModel, index: SheetIndex) -> str:
    """Detailed cell description."""
    wb = model.wb

    # Determine sheet
    if ref.sheet and ref.sheet in wb.sheetnames:
        ws = wb[ref.sheet]
        sheet_name = ref.sheet
    else:
        ws = wb.active
        sheet_name = ws.title

    cell = ws.cell(row=ref.row, column=ref.col)
    addr = format_cell_addr(ref.col, ref.row)
    lines: list[str] = [f"Cell {addr} on {sheet_name}"]

    # Value
    val = cell.value
    val_type = format_value_type(val)
    val_display = format_cell_value(val)
    lines.append(f"  value: {val_display} ({val_type})")

    # Font
    lines.append(f"  font: {format_font(cell.font)}")

    # Fill
    lines.append(f"  fill: {format_fill(cell.fill)}")

    # Alignment
    lines.append(f"  alignment: {format_alignment(cell.alignment)}")

    # Border
    lines.append(f"  border: {format_border(cell.border)}")

    # Number format
    if cell.number_format and cell.number_format != "General":
        lines.append(f"  number-format: {cell.number_format}")

    # Part-of table?
    in_table = None
    for table_name in ws.tables:
        table = ws.tables[table_name]
        rr = parse_range_ref(table.ref)
        if isinstance(rr, RangeRef):
            if (rr.start.row <= ref.row <= rr.end.row and
                    rr.start.col <= ref.col <= rr.end.col):
                in_table = table_name
                break

    # Merged?
    in_merged = False
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            in_merged = True
            break

    parts = []
    if in_table:
        parts.append(f"table {in_table}")
    if in_merged:
        parts.append("merged: yes")
    else:
        parts.append("merged: no")
    lines.append(f"  part-of: {', '.join(parts)}")

    return "\n".join(lines)


def _describe_range(ref: RangeRef, model: SheetsModel, index: SheetIndex) -> str:
    """Range summary description."""
    wb = model.wb

    if ref.sheet and ref.sheet in wb.sheetnames:
        ws = wb[ref.sheet]
        sheet_name = ref.sheet
    else:
        ws = wb.active
        sheet_name = ws.title

    start_addr = format_cell_addr(ref.start.col, ref.start.row)
    end_addr = format_cell_addr(ref.end.col, ref.end.row)
    lines: list[str] = [f"Range {start_addr}:{end_addr} on {sheet_name}"]

    # Scan the range
    total = 0
    non_empty = 0
    type_counts: dict[str, int] = {}
    nums: list[float] = []

    for row in range(ref.start.row, ref.end.row + 1):
        for col in range(ref.start.col, ref.end.col + 1):
            total += 1
            val = ws.cell(row=row, column=col).value
            vtype = format_value_type(val)
            type_counts[vtype] = type_counts.get(vtype, 0) + 1
            if val is not None:
                non_empty += 1
                if isinstance(val, (int, float)) and not isinstance(val, bool):
                    nums.append(float(val))

    rows = ref.end.row - ref.start.row + 1
    cols = ref.end.col - ref.start.col + 1
    lines.append(f"  size: {rows} rows x {cols} cols ({total} cells, {non_empty} non-empty)")

    # Type breakdown
    type_parts = []
    for t in ["text", "number", "formula", "date", "empty"]:
        if t in type_counts and type_counts[t] > 0:
            type_parts.append(f"{t}:{type_counts[t]}")
    if type_parts:
        lines.append(f"  types: {', '.join(type_parts)}")

    # Number range
    if nums:
        mn = int(min(nums)) if min(nums) == int(min(nums)) else min(nums)
        mx = int(max(nums)) if max(nums) == int(max(nums)) else max(nums)
        lines.append(f"  number-range: {mn}..{mx}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# peek — token-efficient data view
# ---------------------------------------------------------------------------

_PEEK_ROW_CAP = 50

def _query_peek(args: str, model: SheetsModel, index: SheetIndex) -> str:
    """Show actual cell data for a range."""
    args = args.strip()
    if not args:
        return "! Usage: peek RANGE (e.g. peek A1:E10)"

    wb = model.wb

    # Parse range ref (may include sheet prefix)
    range_ref = parse_range_ref(args)
    if not isinstance(range_ref, RangeRef):
        # Maybe single cell — treat as 1x1 range
        cell_ref = parse_cell_ref(args)
        if cell_ref:
            range_ref = RangeRef(
                start=CellRef(col=cell_ref.col, row=cell_ref.row),
                end=CellRef(col=cell_ref.col, row=cell_ref.row),
                sheet=cell_ref.sheet,
            )
        else:
            # Maybe it's a sheet name — peek entire data bounds
            if args in wb.sheetnames:
                bounds = index.get_bounds(args)
                if not bounds:
                    return f"Sheet {args} is empty"
                min_row, min_col, max_row, max_col = bounds
                range_ref = RangeRef(
                    start=CellRef(col=min_col, row=min_row),
                    end=CellRef(col=max_col, row=max_row),
                    sheet=args,
                )
            else:
                return f"! Invalid range: {args!r}"

    # Determine sheet
    if range_ref.sheet and range_ref.sheet in wb.sheetnames:
        ws = wb[range_ref.sheet]
        sheet_name = range_ref.sheet
    else:
        ws = wb.active
        sheet_name = ws.title

    start_row = range_ref.start.row
    end_row = range_ref.end.row
    start_col = range_ref.start.col
    end_col = range_ref.end.col
    num_cols = end_col - start_col + 1
    num_rows = end_row - start_row + 1

    # Enforce row cap
    capped = False
    actual_end_row = end_row
    if num_rows > _PEEK_ROW_CAP:
        actual_end_row = start_row + _PEEK_ROW_CAP - 1
        capped = True

    start_addr = format_cell_addr(start_col, start_row)
    end_addr = format_cell_addr(end_col, end_row)

    lines: list[str] = [f"{start_addr}:{end_addr} on {sheet_name}"]

    if num_cols < 12:
        # Narrow mode — compact table
        # Build all row data first to compute column widths
        all_rows: list[list[str]] = []
        for row in range(start_row, actual_end_row + 1):
            row_vals: list[str] = []
            for col in range(start_col, end_col + 1):
                val = ws.cell(row=row, column=col).value
                row_vals.append(_compact_value(val))
            all_rows.append(row_vals)

        if not all_rows:
            lines.append("  (empty range)")
            return "\n".join(lines)

        # Compute column widths
        col_widths = [0] * num_cols
        for row_vals in all_rows:
            for i, v in enumerate(row_vals):
                col_widths[i] = max(col_widths[i], len(v))

        # Cap widths
        col_widths = [min(w, 20) for w in col_widths]

        # Output rows
        for row_vals in all_rows:
            cells = [v[:20].ljust(col_widths[i]) for i, v in enumerate(row_vals)]
            lines.append("  " + "|".join(cells))
    else:
        # Wide mode — vertical per row
        for row in range(start_row, actual_end_row + 1):
            lines.append(f"  Row {row} on {sheet_name}")
            shown = 0
            hidden = 0
            for col in range(start_col, end_col + 1):
                val = ws.cell(row=row, column=col).value
                if shown < 10:
                    col_letter = index_to_col(col)
                    lines.append(f"    {col_letter} = {_compact_value(val)}")
                    shown += 1
                else:
                    hidden += 1
            if hidden > 0:
                lines.append(f"    ...+{hidden} more cols")

    if capped:
        remaining = num_rows - _PEEK_ROW_CAP
        lines.append(f"  ... +{remaining} more rows (capped at {_PEEK_ROW_CAP})")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# list — sub-command dispatcher
# ---------------------------------------------------------------------------

def _query_list(args: str, model: SheetsModel, index: SheetIndex) -> str:
    """List sub-command dispatcher."""
    args = args.strip().lower()

    subcmds = {
        "sheets": _list_sheets,
        "charts": _list_charts,
        "formulas": _list_formulas,
        "styles": _list_styles,
        "names": _list_names,
        "tables": _list_tables,
    }

    handler = subcmds.get(args)
    if handler is None:
        available = ", ".join(sorted(subcmds.keys()))
        return f"! Usage: list {available}"

    return handler(model, index)


def _list_sheets(model: SheetsModel, index: SheetIndex) -> str:
    """List sheets with basic stats."""
    wb = model.wb
    lines: list[str] = [f"Sheets ({len(wb.sheetnames)}):"]

    for ws in wb.worksheets:
        sheet_name = ws.title
        bounds = index.get_bounds(sheet_name)
        if bounds:
            min_row, min_col, max_row, max_col = bounds
            rows = max_row - min_row + 1
            cols = max_col - min_col + 1

            # Count non-empty cells
            data_cells = 0
            for row in ws.iter_rows(
                min_row=min_row, max_row=max_row,
                min_col=min_col, max_col=max_col,
            ):
                for cell in row:
                    if cell.value is not None:
                        data_cells += 1

            active = " [active]" if ws == wb.active else ""
            hidden = " [hidden]" if ws.sheet_state == "hidden" else ""
            lines.append(
                f"  {sheet_name}{active}{hidden}: "
                f"{rows} rows, {cols} cols, {data_cells} data cells"
            )
        else:
            active = " [active]" if ws == wb.active else ""
            hidden = " [hidden]" if ws.sheet_state == "hidden" else ""
            lines.append(f"  {sheet_name}{active}{hidden}: (empty)")

    return "\n".join(lines)


def _list_charts(model: SheetsModel, index: SheetIndex) -> str:
    """List all charts across all sheets."""
    wb = model.wb
    charts_found: list[str] = []

    for ws in wb.worksheets:
        for chart in ws._charts:
            chart_title = chart.title or "Untitled"
            chart_type = type(chart).__name__
            # Try to get anchor position
            anchor = ""
            if hasattr(chart, "anchor") and chart.anchor:
                anchor = f" at:{chart.anchor}"
            charts_found.append(
                f"  {ws.title}: {chart_type} \"{chart_title}\"{anchor}"
            )

    if not charts_found:
        return "Charts: (none)"

    lines = [f"Charts ({len(charts_found)}):"]
    lines.extend(charts_found)
    return "\n".join(lines)


def _list_formulas(model: SheetsModel, index: SheetIndex) -> str:
    """List formula cells grouped by pattern."""
    wb = model.wb
    # pattern -> list of (sheet, addr, formula)
    patterns: dict[str, list[tuple[str, str, str]]] = {}

    for ws in wb.worksheets:
        sheet_name = ws.title
        bounds = index.get_bounds(sheet_name)
        if not bounds:
            continue
        min_row, min_col, max_row, max_col = bounds
        for row in ws.iter_rows(
            min_row=min_row, max_row=max_row,
            min_col=min_col, max_col=max_col,
        ):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    addr = format_cell_addr(cell.column, cell.row)
                    pattern = re.sub(r"\d+", "N", cell.value)
                    key = f"{sheet_name}|{pattern}"
                    if key not in patterns:
                        patterns[key] = []
                    patterns[key].append((sheet_name, addr, cell.value))

    if not patterns:
        return "Formulas: (none)"

    total = sum(len(v) for v in patterns.values())
    lines: list[str] = [f"Formulas ({total} cells, {len(patterns)} patterns):"]

    for key, entries in list(patterns.items())[:15]:
        sheet_name = entries[0][0]
        if len(entries) == 1:
            _, addr, formula = entries[0]
            lines.append(f"  {sheet_name}!{addr}  {formula}")
        else:
            first_addr = entries[0][1]
            last_addr = entries[-1][1]
            sample = entries[0][2]
            lines.append(
                f"  {sheet_name}!{first_addr}:{last_addr}  {sample}  ({len(entries)} cells)"
            )

    if len(patterns) > 15:
        lines.append(f"  ... +{len(patterns) - 15} more patterns")

    return "\n".join(lines)


def _list_styles(model: SheetsModel, index: SheetIndex) -> str:
    """List defined named styles."""
    wb = model.wb
    styles = list(wb.named_styles)

    if not styles:
        return "Named styles: (none)"

    lines: list[str] = [f"Named styles ({len(styles)}):"]
    for style in styles[:20]:
        name = style if isinstance(style, str) else style.name
        lines.append(f"  {name}")

    if len(styles) > 20:
        lines.append(f"  ... +{len(styles) - 20} more")

    return "\n".join(lines)


def _list_names(model: SheetsModel, index: SheetIndex) -> str:
    """List named ranges with scope."""
    wb = model.wb
    # wb.defined_names iterates over name strings; .values() gives DefinedName objects
    names = list(wb.defined_names.values()) if wb.defined_names else []

    if not names:
        return "Named ranges: (none)"

    lines: list[str] = [f"Named ranges ({len(names)}):"]
    for defn in names[:20]:
        scope = "workbook"
        if defn.localSheetId is not None:
            try:
                scope = wb.sheetnames[defn.localSheetId]
            except IndexError:
                scope = f"sheet#{defn.localSheetId}"

        # Get destinations
        dests = []
        try:
            for title, coord in defn.destinations:
                dests.append(f"{title}!{coord}")
        except Exception:
            dests.append(str(defn.value))

        dest_str = ", ".join(dests) if dests else str(defn.value)
        lines.append(f"  {defn.name}: {dest_str} (scope:{scope})")

    if len(names) > 20:
        lines.append(f"  ... +{len(names) - 20} more")

    return "\n".join(lines)


def _list_tables(model: SheetsModel, index: SheetIndex) -> str:
    """List Excel tables with range and style."""
    wb = model.wb
    tables_found: list[str] = []

    for ws in wb.worksheets:
        for table_name in ws.tables:
            table = ws.tables[table_name]
            style_name = ""
            if table.tableStyleInfo and table.tableStyleInfo.name:
                style_name = f" style:{table.tableStyleInfo.name}"
            tables_found.append(
                f"  {ws.title}: {table_name} {table.ref}{style_name}"
            )

    if not tables_found:
        return "Tables: (none)"

    lines = [f"Tables ({len(tables_found)}):"]
    lines.extend(tables_found)
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# find — search cell values
# ---------------------------------------------------------------------------

_FIND_MAX_RESULTS = 50

def _query_find(args: str, model: SheetsModel, index: SheetIndex) -> str:
    """Search cell values or formulas across all sheets."""
    args = args.strip()
    if not args:
        return "! Usage: find TEXT | find formula:PATTERN"

    wb = model.wb

    # Determine search mode
    if args.lower().startswith("formula:"):
        pattern = args[8:]
        return _find_formulas(pattern, model, index)

    # Text search
    search_text = args.lower()
    results: list[str] = []

    for ws in wb.worksheets:
        sheet_name = ws.title
        bounds = index.get_bounds(sheet_name)
        if not bounds:
            continue
        min_row, min_col, max_row, max_col = bounds
        for row in ws.iter_rows(
            min_row=min_row, max_row=max_row,
            min_col=min_col, max_col=max_col,
        ):
            for cell in row:
                if cell.value is None:
                    continue
                val_str = str(cell.value)
                if search_text in val_str.lower():
                    addr = format_cell_addr(cell.column, cell.row)
                    display = _compact_value(cell.value)
                    results.append(f"  {sheet_name}!{addr} = {display}")
                    if len(results) >= _FIND_MAX_RESULTS:
                        break
            if len(results) >= _FIND_MAX_RESULTS:
                break
        if len(results) >= _FIND_MAX_RESULTS:
            break

    if not results:
        return f"find {args!r}: no matches"

    lines = [f"find {args!r}: {len(results)} match(es)"]
    lines.extend(results)
    if len(results) >= _FIND_MAX_RESULTS:
        lines.append(f"  ... capped at {_FIND_MAX_RESULTS} results")
    return "\n".join(lines)


def _find_formulas(pattern: str, model: SheetsModel, index: SheetIndex) -> str:
    """Search formula text across all sheets."""
    wb = model.wb
    search_pattern = pattern.lower()
    results: list[str] = []

    for ws in wb.worksheets:
        sheet_name = ws.title
        bounds = index.get_bounds(sheet_name)
        if not bounds:
            continue
        min_row, min_col, max_row, max_col = bounds
        for row in ws.iter_rows(
            min_row=min_row, max_row=max_row,
            min_col=min_col, max_col=max_col,
        ):
            for cell in row:
                if not isinstance(cell.value, str) or not cell.value.startswith("="):
                    continue
                if search_pattern in cell.value.lower():
                    addr = format_cell_addr(cell.column, cell.row)
                    results.append(f"  {sheet_name}!{addr} = {cell.value}")
                    if len(results) >= _FIND_MAX_RESULTS:
                        break
            if len(results) >= _FIND_MAX_RESULTS:
                break
        if len(results) >= _FIND_MAX_RESULTS:
            break

    if not results:
        return f"find formula:{pattern!r}: no matches"

    lines = [f"find formula:{pattern!r}: {len(results)} match(es)"]
    lines.extend(results)
    if len(results) >= _FIND_MAX_RESULTS:
        lines.append(f"  ... capped at {_FIND_MAX_RESULTS} results")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _detect_formula_patterns(ws, bounds: tuple[int, int, int, int]) -> list[str]:
    """Detect formula patterns in a sheet for the plan query."""
    min_row, min_col, max_row, max_col = bounds
    patterns: dict[str, list[str]] = {}  # pattern -> [cell_addrs]

    for row in ws.iter_rows(
        min_row=min_row, max_row=max_row,
        min_col=min_col, max_col=max_col,
    ):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                # Normalize formula pattern (strip row numbers for grouping)
                pattern = re.sub(r"\d+", "N", cell.value)
                addr = f"{index_to_col(cell.column)}{cell.row}"
                if pattern not in patterns:
                    patterns[pattern] = []
                patterns[pattern].append(addr)

    result = []
    for pattern, addrs in patterns.items():
        if len(addrs) == 1:
            result.append(f"{addrs[0]} {pattern.replace('N', 'N')}")
        else:
            # Show range
            first, last = addrs[0], addrs[-1]
            result.append(f"{first}:{last} pattern:{pattern}")

    return result[:5]  # Cap to avoid token explosion


def _compact_value(val) -> str:
    """Compact display of a cell value for peek/describe."""
    if val is None:
        return ""
    if isinstance(val, str):
        if val.startswith("="):
            return val
        return val[:30]
    if isinstance(val, bool):
        return str(val).upper()
    if isinstance(val, float):
        if val == int(val):
            return str(int(val))
        return f"{val:g}"
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    return str(val)


def _get_formula_groups(ws, bounds: tuple[int, int, int, int]) -> list[str]:
    """Get formula groupings for describe output."""
    min_row, min_col, max_row, max_col = bounds
    # pattern -> (first_addr, last_addr, count, sample)
    groups: dict[str, tuple[str, str, int, str]] = {}

    for row in ws.iter_rows(
        min_row=min_row, max_row=max_row,
        min_col=min_col, max_col=max_col,
    ):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                pattern = re.sub(r"\d+", "N", cell.value)
                addr = format_cell_addr(cell.column, cell.row)
                if pattern not in groups:
                    groups[pattern] = (addr, addr, 1, cell.value)
                else:
                    first, _, count, sample = groups[pattern]
                    groups[pattern] = (first, addr, count + 1, sample)

    result: list[str] = []
    for pattern, (first, last, count, sample) in groups.items():
        if count == 1:
            result.append(f"{first}   {sample}")
        else:
            fill_note = "fill pattern" if count > 1 else ""
            result.append(f"{first}:{last}   {sample}  ({count} cells, {fill_note})")

    return result


def _describe_cf_rule(rule) -> str:
    """Describe a conditional formatting rule briefly."""
    rule_type = type(rule).__name__
    # Try to extract useful info
    if hasattr(rule, "type") and rule.type:
        return rule.type
    return rule_type
