"""Tests for chart operations — chart add/series/axis/remove."""

from __future__ import annotations

import pytest
from fcp_core import ParsedOp

from fcp_sheets.server.ops_charts import op_chart, _get_chart_title_text, _get_title_text
from fcp_sheets.server.resolvers import SheetsOpContext


def _setup_data(ctx: SheetsOpContext):
    """Populate cells with sample data for chart creation."""
    ws = ctx.active_sheet
    # Headers in row 1
    ws.cell(row=1, column=1, value="Category")
    ws.cell(row=1, column=2, value="Q1")
    ws.cell(row=1, column=3, value="Q2")
    ws.cell(row=1, column=4, value="Q3")
    # Data rows
    for row in range(2, 6):
        ws.cell(row=row, column=1, value=f"Item {row - 1}")
        ws.cell(row=row, column=2, value=row * 10)
        ws.cell(row=row, column=3, value=row * 15)
        ws.cell(row=row, column=4, value=row * 20)


class TestChartAdd:
    def test_add_bar_chart(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"data": "B1:D5", "title": "Sales"},
            raw='chart add bar title:"Sales" data:B1:D5',
        )
        result = op_chart(op, ctx)
        assert result.success
        assert "Sales" in result.message
        assert len(ctx.active_sheet._charts) == 1
        assert _get_chart_title_text(ctx.active_sheet._charts[0]) == "Sales"

    def test_add_column_chart(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="chart",
            positionals=["add", "column"],
            params={"data": "B1:D5", "title": "Revenue"},
            raw='chart add column title:"Revenue" data:B1:D5',
        )
        result = op_chart(op, ctx)
        assert result.success
        chart = ctx.active_sheet._charts[0]
        assert chart.type == "col"
        assert chart.grouping == "clustered"

    def test_add_line_chart(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="chart",
            positionals=["add", "line"],
            params={"data": "B1:D5", "title": "Trend"},
            raw='chart add line title:"Trend" data:B1:D5',
        )
        result = op_chart(op, ctx)
        assert result.success
        assert len(ctx.active_sheet._charts) == 1

    def test_add_pie_chart(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="chart",
            positionals=["add", "pie"],
            params={"data": "B1:B5", "title": "Distribution"},
            raw='chart add pie title:"Distribution" data:B1:B5',
        )
        result = op_chart(op, ctx)
        assert result.success

    def test_add_scatter_chart(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="chart",
            positionals=["add", "scatter"],
            params={"data": "B1:C5", "title": "XY"},
            raw='chart add scatter title:"XY" data:B1:C5',
        )
        result = op_chart(op, ctx)
        assert result.success

    def test_add_chart_with_categories(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"data": "B1:D5", "categories": "A2:A5", "title": "WithCat"},
            raw='chart add bar title:"WithCat" data:B1:D5 categories:A2:A5',
        )
        result = op_chart(op, ctx)
        assert result.success

    def test_add_chart_with_size(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"data": "B1:D5", "title": "Sized", "size": "750x400"},
            raw='chart add bar title:"Sized" data:B1:D5 size:750x400',
        )
        result = op_chart(op, ctx)
        assert result.success
        chart = ctx.active_sheet._charts[0]
        assert chart.width == pytest.approx(15.0)  # 750/50
        assert chart.height == pytest.approx(8.0)  # 400/50

    def test_add_chart_with_at_position(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"data": "B1:D5", "title": "Placed", "at": "F8"},
            raw='chart add bar title:"Placed" data:B1:D5 at:F8',
        )
        result = op_chart(op, ctx)
        assert result.success

    def test_add_chart_with_legend(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"data": "B1:D5", "title": "Legended", "legend": "b"},
            raw='chart add bar title:"Legended" data:B1:D5 legend:b',
        )
        result = op_chart(op, ctx)
        assert result.success
        chart = ctx.active_sheet._charts[0]
        assert chart.legend.position == "b"

    def test_add_chart_with_style(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"data": "B1:D5", "title": "Styled", "style": "10"},
            raw='chart add bar title:"Styled" data:B1:D5 style:10',
        )
        result = op_chart(op, ctx)
        assert result.success
        chart = ctx.active_sheet._charts[0]
        assert chart.style == 10

    def test_add_stacked_bar(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="chart",
            positionals=["add", "stacked-bar"],
            params={"data": "B1:D5", "title": "Stacked"},
            raw='chart add stacked-bar title:"Stacked" data:B1:D5',
        )
        result = op_chart(op, ctx)
        assert result.success
        chart = ctx.active_sheet._charts[0]
        assert chart.grouping == "stacked"
        assert chart.type == "bar"

    def test_add_chart_no_title(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="chart",
            positionals=["add", "line"],
            params={"data": "B1:D5"},
            raw="chart add line data:B1:D5",
        )
        result = op_chart(op, ctx)
        assert result.success
        assert "line" in result.message

    def test_add_duplicate_title_error(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op1 = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"data": "B1:D5", "title": "Dup"},
            raw='chart add bar title:"Dup" data:B1:D5',
        )
        op_chart(op1, ctx)

        op2 = ParsedOp(
            verb="chart",
            positionals=["add", "line"],
            params={"data": "B1:D5", "title": "Dup"},
            raw='chart add line title:"Dup" data:B1:D5',
        )
        result = op_chart(op2, ctx)
        assert not result.success
        assert "already exists" in result.message

    def test_add_invalid_chart_type(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="chart",
            positionals=["add", "sparkle"],
            params={"data": "B1:D5"},
            raw="chart add sparkle data:B1:D5",
        )
        result = op_chart(op, ctx)
        assert not result.success
        assert "Unknown chart type" in result.message

    def test_add_missing_data_range(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"title": "NoData"},
            raw='chart add bar title:"NoData"',
        )
        result = op_chart(op, ctx)
        assert not result.success
        assert "data:RANGE" in result.message

    def test_add_missing_args(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="chart",
            positionals=["add"],
            params={},
            raw="chart add",
        )
        result = op_chart(op, ctx)
        assert not result.success


class TestChartSeries:
    def test_series_add(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        # First create a chart
        op_add = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"data": "B1:B5", "title": "MySeries"},
            raw='chart add bar title:"MySeries" data:B1:B5',
        )
        op_chart(op_add, ctx)

        # Add another series
        op_series = ParsedOp(
            verb="chart",
            positionals=["series", "MySeries"],
            params={"data": "C1:C5"},
            raw="chart series MySeries data:C1:C5",
        )
        result = op_chart(op_series, ctx)
        assert result.success
        assert "Series" in result.message

    def test_series_chart_not_found(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="chart",
            positionals=["series", "NoChart"],
            params={"data": "B1:B5"},
            raw="chart series NoChart data:B1:B5",
        )
        result = op_chart(op, ctx)
        assert not result.success
        assert "not found" in result.message

    def test_series_missing_data(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op_add = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"data": "B1:B5", "title": "SerData"},
            raw='chart add bar title:"SerData" data:B1:B5',
        )
        op_chart(op_add, ctx)

        op_series = ParsedOp(
            verb="chart",
            positionals=["series", "SerData"],
            params={},
            raw="chart series SerData",
        )
        result = op_chart(op_series, ctx)
        assert not result.success


class TestChartAxis:
    def test_axis_x_title(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op_add = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"data": "B1:D5", "title": "AxisChart"},
            raw='chart add bar title:"AxisChart" data:B1:D5',
        )
        op_chart(op_add, ctx)

        op_axis = ParsedOp(
            verb="chart",
            positionals=["axis", "AxisChart", "x"],
            params={"title": "Time"},
            raw='chart axis AxisChart x title:"Time"',
        )
        result = op_chart(op_axis, ctx)
        assert result.success
        chart = ctx.active_sheet._charts[0]
        assert _get_title_text(chart.x_axis.title) == "Time"

    def test_axis_y_title_and_bounds(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op_add = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"data": "B1:D5", "title": "YAxis"},
            raw='chart add bar title:"YAxis" data:B1:D5',
        )
        op_chart(op_add, ctx)

        op_axis = ParsedOp(
            verb="chart",
            positionals=["axis", "YAxis", "y"],
            params={"title": "Revenue", "min": "0", "max": "100"},
            raw='chart axis YAxis y title:"Revenue" min:0 max:100',
        )
        result = op_chart(op_axis, ctx)
        assert result.success
        chart = ctx.active_sheet._charts[0]
        assert _get_title_text(chart.y_axis.title) == "Revenue"
        assert chart.y_axis.scaling.min == 0.0
        assert chart.y_axis.scaling.max == 100.0

    def test_axis_format(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op_add = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"data": "B1:D5", "title": "FmtAxis"},
            raw='chart add bar title:"FmtAxis" data:B1:D5',
        )
        op_chart(op_add, ctx)

        op_axis = ParsedOp(
            verb="chart",
            positionals=["axis", "FmtAxis", "y"],
            params={"fmt": "$#,##0"},
            raw="chart axis FmtAxis y fmt:$#,##0",
        )
        result = op_chart(op_axis, ctx)
        assert result.success
        chart = ctx.active_sheet._charts[0]
        assert chart.y_axis.numFmt.formatCode == "$#,##0"

    def test_axis_chart_not_found(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="chart",
            positionals=["axis", "NoChart", "x"],
            params={"title": "X"},
            raw='chart axis NoChart x title:"X"',
        )
        result = op_chart(op, ctx)
        assert not result.success
        assert "not found" in result.message

    def test_axis_invalid_axis(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op_add = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"data": "B1:D5", "title": "BadAxis"},
            raw='chart add bar title:"BadAxis" data:B1:D5',
        )
        op_chart(op_add, ctx)

        op_axis = ParsedOp(
            verb="chart",
            positionals=["axis", "BadAxis", "z"],
            params={},
            raw="chart axis BadAxis z",
        )
        result = op_chart(op_axis, ctx)
        assert not result.success
        assert "Invalid axis" in result.message


class TestChartRemove:
    def test_remove_chart(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op_add = ParsedOp(
            verb="chart",
            positionals=["add", "bar"],
            params={"data": "B1:D5", "title": "ToRemove"},
            raw='chart add bar title:"ToRemove" data:B1:D5',
        )
        op_chart(op_add, ctx)
        assert len(ctx.active_sheet._charts) == 1

        op_rm = ParsedOp(
            verb="chart",
            positionals=["remove", "ToRemove"],
            params={},
            raw="chart remove ToRemove",
        )
        result = op_chart(op_rm, ctx)
        assert result.success
        assert len(ctx.active_sheet._charts) == 0

    def test_remove_chart_not_found(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="chart",
            positionals=["remove", "Ghost"],
            params={},
            raw="chart remove Ghost",
        )
        result = op_chart(op, ctx)
        assert not result.success
        assert "not found" in result.message


class TestChartDispatch:
    def test_unknown_subcmd(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="chart",
            positionals=["resize", "Foo"],
            params={},
            raw="chart resize Foo",
        )
        result = op_chart(op, ctx)
        assert not result.success
        assert "Unknown chart sub-command" in result.message

    def test_no_positionals(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="chart",
            positionals=[],
            params={},
            raw="chart",
        )
        result = op_chart(op, ctx)
        assert not result.success
