"""Tests for conditional formatting operations — cond-fmt verb."""

from __future__ import annotations

import pytest
from fcp_core import ParsedOp

from fcp_sheets.server.ops_cond_fmt import op_cond_fmt
from fcp_sheets.server.resolvers import SheetsOpContext


def _setup_data(ctx: SheetsOpContext):
    """Populate cells with numeric data for conditional formatting."""
    ws = ctx.active_sheet
    for row in range(1, 11):
        ws.cell(row=row, column=1, value=row * 10)
        ws.cell(row=row, column=2, value=row * 5)


class TestColorScale:
    def test_two_color_scale(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "color-scale"],
            params={},
            raw="cond-fmt A1:A10 color-scale",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success
        assert "2-color" in result.message
        ws = ctx.active_sheet
        rules = list(ws.conditional_formatting)
        assert len(rules) > 0

    def test_two_color_scale_custom_colors(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "color-scale"],
            params={"min-color": "#FF0000", "max-color": "#00FF00"},
            raw="cond-fmt A1:A10 color-scale min-color:#FF0000 max-color:#00FF00",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success
        assert "2-color" in result.message

    def test_three_color_scale(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "color-scale"],
            params={"min-color": "#FF0000", "mid-color": "#FFFF00", "max-color": "#00FF00"},
            raw="cond-fmt A1:A10 color-scale min-color:#FF0000 mid-color:#FFFF00 max-color:#00FF00",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success
        assert "3-color" in result.message

    def test_color_scale_with_named_colors(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "color-scale"],
            params={"min-color": "red", "max-color": "green"},
            raw="cond-fmt A1:A10 color-scale min-color:red max-color:green",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success


class TestDataBar:
    def test_data_bar_default(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "data-bar"],
            params={},
            raw="cond-fmt A1:A10 data-bar",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success
        assert "Data bar" in result.message

    def test_data_bar_custom_color(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "data-bar"],
            params={"color": "#FF6600"},
            raw="cond-fmt A1:A10 data-bar color:#FF6600",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success


class TestIconSet:
    def test_icon_set_arrows(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "icon-set"],
            params={"icons": "arrows"},
            raw="cond-fmt A1:A10 icon-set icons:arrows",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success
        assert "arrows" in result.message

    def test_icon_set_flags(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "icon-set"],
            params={"icons": "flags"},
            raw="cond-fmt A1:A10 icon-set icons:flags",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success

    def test_icon_set_traffic(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "icon-set"],
            params={"icons": "traffic"},
            raw="cond-fmt A1:A10 icon-set icons:traffic",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success

    def test_icon_set_default(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "icon-set"],
            params={},
            raw="cond-fmt A1:A10 icon-set",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success

    def test_icon_set_invalid(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "icon-set"],
            params={"icons": "sparkles"},
            raw="cond-fmt A1:A10 icon-set icons:sparkles",
        )
        result = op_cond_fmt(op, ctx)
        assert not result.success
        assert "Unknown icon set" in result.message


class TestCellIs:
    def test_cell_is_greater_than(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "cell-is", "gt", "50"],
            params={"fill": "#C6EFCE"},
            raw="cond-fmt A1:A10 cell-is gt 50 fill:#C6EFCE",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success
        assert "Cell-is" in result.message

    def test_cell_is_less_than(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "cell-is", "lt", "30"],
            params={"fill": "#FFC7CE"},
            raw="cond-fmt A1:A10 cell-is lt 30 fill:#FFC7CE",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success

    def test_cell_is_between(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "cell-is", "between", "20", "80"],
            params={"fill": "#FFEB9C"},
            raw="cond-fmt A1:A10 cell-is between 20 80 fill:#FFEB9C",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success

    def test_cell_is_between_missing_value2(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "cell-is", "between", "20"],
            params={},
            raw="cond-fmt A1:A10 cell-is between 20",
        )
        result = op_cond_fmt(op, ctx)
        assert not result.success
        assert "two values" in result.message

    def test_cell_is_equal(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "cell-is", "eq", "50"],
            params={"fill": "#C6EFCE", "color": "#006100", "bold": "true"},
            raw="cond-fmt A1:A10 cell-is eq 50 fill:#C6EFCE color:#006100",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success

    def test_cell_is_with_bold(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "cell-is", "gt", "50", "bold"],
            params={"fill": "#C6EFCE"},
            raw="cond-fmt A1:A10 cell-is gt 50 bold fill:#C6EFCE",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success

    def test_cell_is_unknown_operator(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "cell-is", "like", "50"],
            params={},
            raw="cond-fmt A1:A10 cell-is like 50",
        )
        result = op_cond_fmt(op, ctx)
        assert not result.success
        assert "Unknown operator" in result.message

    def test_cell_is_missing_args(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "cell-is", "gt"],
            params={},
            raw="cond-fmt A1:A10 cell-is gt",
        )
        result = op_cond_fmt(op, ctx)
        assert not result.success


class TestFormula:
    def test_formula_rule(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "formula", "=A1>AVERAGE($A$1:$A$10)"],
            params={"fill": "#C6EFCE"},
            raw="cond-fmt A1:A10 formula =A1>AVERAGE($A$1:$A$10) fill:#C6EFCE",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success
        assert "Formula rule" in result.message

    def test_formula_missing_formula(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "formula"],
            params={},
            raw="cond-fmt A1:A10 formula",
        )
        result = op_cond_fmt(op, ctx)
        assert not result.success


class TestDuplicateUnique:
    def test_duplicate_rule(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "duplicate"],
            params={},
            raw="cond-fmt A1:A10 duplicate",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success
        assert "Duplicate" in result.message

    def test_duplicate_custom_fill(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "duplicate"],
            params={"fill": "#FF0000"},
            raw="cond-fmt A1:A10 duplicate fill:#FF0000",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success

    def test_unique_rule(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "unique"],
            params={},
            raw="cond-fmt A1:A10 unique",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success
        assert "Unique" in result.message


class TestTopBottom:
    def test_top_n(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "top", "3"],
            params={},
            raw="cond-fmt A1:A10 top 3",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success
        assert "Top 3" in result.message

    def test_bottom_n(self, ctx: SheetsOpContext):
        _setup_data(ctx)
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "bottom", "5"],
            params={},
            raw="cond-fmt A1:A10 bottom 5",
        )
        result = op_cond_fmt(op, ctx)
        assert result.success
        assert "Bottom 5" in result.message

    def test_top_missing_n(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "top"],
            params={},
            raw="cond-fmt A1:A10 top",
        )
        result = op_cond_fmt(op, ctx)
        assert not result.success

    def test_bottom_invalid_n(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "bottom", "abc"],
            params={},
            raw="cond-fmt A1:A10 bottom abc",
        )
        result = op_cond_fmt(op, ctx)
        assert not result.success


class TestCondFmtDispatch:
    def test_unknown_type(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10", "sparkline"],
            params={},
            raw="cond-fmt A1:A10 sparkline",
        )
        result = op_cond_fmt(op, ctx)
        assert not result.success
        assert "Unknown cond-fmt type" in result.message

    def test_missing_positionals(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="cond-fmt",
            positionals=["A1:A10"],
            params={},
            raw="cond-fmt A1:A10",
        )
        result = op_cond_fmt(op, ctx)
        assert not result.success

    def test_no_positionals(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="cond-fmt",
            positionals=[],
            params={},
            raw="cond-fmt",
        )
        result = op_cond_fmt(op, ctx)
        assert not result.success
