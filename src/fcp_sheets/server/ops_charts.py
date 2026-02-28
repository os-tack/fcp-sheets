"""Chart operation handlers — chart add/series/axis/remove."""

from __future__ import annotations

from openpyxl.chart import Reference

from fcp_core import OpResult, ParsedOp

from fcp_sheets.lib.chart_types import get_chart_class
from fcp_sheets.model.refs import parse_cell_ref, parse_range_ref, RangeRef
from fcp_sheets.server.resolvers import SheetsOpContext


def _get_title_text(title_obj) -> str | None:
    """Extract the plain text string from an openpyxl Title object.

    openpyxl wraps title strings in Title objects with nested rich text.
    This extracts the actual string from charts, axes, etc.
    """
    if title_obj is None:
        return None
    if isinstance(title_obj, str):
        return title_obj
    # Title object — extract text from rich text runs
    try:
        for p in title_obj.tx.rich.p:
            for r in p.r:
                return r.t
    except (AttributeError, TypeError):
        pass
    return None


def _get_chart_title_text(chart) -> str | None:
    """Extract the plain text title string from an openpyxl chart."""
    return _get_title_text(chart.title)


def _find_chart(ws, title: str):
    """Find a chart by title in the worksheet."""
    for chart in ws._charts:
        if _get_chart_title_text(chart) == title:
            return chart
    return None


def _parse_range_to_reference(ws, range_str: str) -> Reference | None:
    """Parse a range string like 'A1:D5' into an openpyxl Reference."""
    ref = parse_range_ref(range_str)
    if isinstance(ref, RangeRef):
        return Reference(
            worksheet=ws,
            min_col=ref.start.col,
            min_row=ref.start.row,
            max_col=ref.end.col,
            max_row=ref.end.row,
        )
    # Try single cell as a 1x1 range
    cell = parse_cell_ref(range_str)
    if cell:
        return Reference(
            worksheet=ws,
            min_col=cell.col,
            min_row=cell.row,
            max_col=cell.col,
            max_row=cell.row,
        )
    return None


def _chart_add(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    """chart add TYPE [title:"TEXT"] data:RANGE [categories:RANGE] [at:CELL] [size:WxH] [legend:POS] [style:N]"""
    if len(op.positionals) < 2:
        return OpResult(
            success=False,
            message='Usage: chart add TYPE [title:"TEXT"] data:RANGE [categories:RANGE] [at:CELL] [size:WxH]',
        )

    chart_type = op.positionals[1]

    # Look up chart class
    try:
        cls, grouping, type_override = get_chart_class(chart_type)
    except ValueError as e:
        return OpResult(success=False, message=str(e))

    # Data range is required
    data_range = op.params.get("data")
    if not data_range:
        return OpResult(success=False, message="Missing required param: data:RANGE")

    ws = ctx.active_sheet

    # Create chart instance
    chart = cls()

    # Set title
    title = op.params.get("title")
    if title:
        chart.title = title

    # Set grouping/type override
    if grouping is not None:
        chart.grouping = grouping
    if type_override is not None:
        chart.type = type_override

    # Set style
    style_num = op.params.get("style")
    if style_num:
        try:
            chart.style = int(style_num)
        except (ValueError, TypeError):
            pass

    # Parse data range
    data_ref = _parse_range_to_reference(ws, data_range)
    if data_ref is None:
        return OpResult(success=False, message=f"Invalid data range: {data_range!r}")

    chart.add_data(data_ref, titles_from_data=True)

    # Parse categories range
    cat_range = op.params.get("categories")
    if cat_range:
        cat_ref = _parse_range_to_reference(ws, cat_range)
        if cat_ref is None:
            return OpResult(success=False, message=f"Invalid categories range: {cat_range!r}")
        chart.set_categories(cat_ref)

    # Set legend position
    legend_pos = op.params.get("legend")
    if legend_pos:
        if chart.legend is not None:
            chart.legend.position = legend_pos
        else:
            from openpyxl.chart.legend import Legend
            chart.legend = Legend()
            chart.legend.position = legend_pos

    # Parse size WxH (in approximate screen units, convert to cm)
    size_str = op.params.get("size")
    if size_str and "x" in size_str.lower():
        parts = size_str.lower().split("x")
        try:
            w, h = float(parts[0]), float(parts[1])
            chart.width = w / 50  # approximate conversion to cm
            chart.height = h / 50
        except (ValueError, IndexError):
            pass

    # Set anchor position
    anchor = op.params.get("at")
    if anchor:
        chart.anchor = anchor

    # Check for duplicate title
    if title and _find_chart(ws, title) is not None:
        return OpResult(success=False, message=f"Chart with title {title!r} already exists")

    ws.add_chart(chart, anchor or "E5")

    label = title or chart_type
    return OpResult(success=True, message=f"Chart '{label}' added ({chart_type})", prefix="+")


def _chart_series(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    """chart series CHART_TITLE data:RANGE [title:"TEXT"]"""
    if len(op.positionals) < 2:
        return OpResult(success=False, message='Usage: chart series CHART_TITLE data:RANGE [title:"TEXT"]')

    chart_title = op.positionals[1]
    ws = ctx.active_sheet
    chart = _find_chart(ws, chart_title)
    if chart is None:
        return OpResult(success=False, message=f"Chart not found: {chart_title!r}")

    data_range = op.params.get("data")
    if not data_range:
        return OpResult(success=False, message="Missing required param: data:RANGE")

    data_ref = _parse_range_to_reference(ws, data_range)
    if data_ref is None:
        return OpResult(success=False, message=f"Invalid data range: {data_range!r}")

    chart.add_data(data_ref, titles_from_data=True)

    series_title = op.params.get("title", data_range)
    return OpResult(success=True, message=f"Series '{series_title}' added to chart '{chart_title}'", prefix="~")


def _chart_axis(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    """chart axis CHART_TITLE x|y [title:"TEXT"] [min:N] [max:N] [fmt:FORMAT]"""
    if len(op.positionals) < 3:
        return OpResult(
            success=False,
            message='Usage: chart axis CHART_TITLE x|y [title:"TEXT"] [min:N] [max:N] [fmt:FORMAT]',
        )

    chart_title = op.positionals[1]
    axis_name = op.positionals[2].lower()

    ws = ctx.active_sheet
    chart = _find_chart(ws, chart_title)
    if chart is None:
        return OpResult(success=False, message=f"Chart not found: {chart_title!r}")

    if axis_name not in ("x", "y"):
        return OpResult(success=False, message=f"Invalid axis: {axis_name!r}. Use 'x' or 'y'.")

    axis = chart.x_axis if axis_name == "x" else chart.y_axis

    # Set axis title
    axis_title = op.params.get("title")
    if axis_title:
        axis.title = axis_title

    # Set min
    min_val = op.params.get("min")
    if min_val is not None:
        try:
            axis.scaling.min = float(min_val)
        except (ValueError, TypeError):
            pass

    # Set max
    max_val = op.params.get("max")
    if max_val is not None:
        try:
            axis.scaling.max = float(max_val)
        except (ValueError, TypeError):
            pass

    # Set number format
    fmt = op.params.get("fmt")
    if fmt:
        axis.numFmt = fmt

    parts = []
    if axis_title:
        parts.append(f"title={axis_title!r}")
    if min_val:
        parts.append(f"min={min_val}")
    if max_val:
        parts.append(f"max={max_val}")
    if fmt:
        parts.append(f"fmt={fmt}")

    detail = ", ".join(parts) if parts else "updated"
    return OpResult(success=True, message=f"Chart '{chart_title}' {axis_name}-axis: {detail}", prefix="*")


def _chart_remove(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    """chart remove CHART_TITLE"""
    if len(op.positionals) < 2:
        return OpResult(success=False, message="Usage: chart remove CHART_TITLE")

    chart_title = op.positionals[1]
    ws = ctx.active_sheet

    for i, chart in enumerate(ws._charts):
        if _get_chart_title_text(chart) == chart_title:
            ws._charts.pop(i)
            return OpResult(success=True, message=f"Chart '{chart_title}' removed", prefix="-")

    return OpResult(success=False, message=f"Chart not found: {chart_title!r}")


# Sub-command dispatch
_CHART_SUBCMDS: dict[str, callable] = {
    "add": _chart_add,
    "series": _chart_series,
    "axis": _chart_axis,
    "remove": _chart_remove,
}


def op_chart(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    """Main chart verb dispatcher."""
    if not op.positionals:
        return OpResult(
            success=False,
            message="Usage: chart add|series|axis|remove ...",
        )

    subcmd = op.positionals[0].lower()
    handler = _CHART_SUBCMDS.get(subcmd)
    if handler is None:
        available = ", ".join(sorted(_CHART_SUBCMDS.keys()))
        return OpResult(success=False, message=f"Unknown chart sub-command: {subcmd!r}. Available: {available}")

    return handler(op, ctx)


HANDLERS: dict[str, callable] = {
    "chart": op_chart,
}
