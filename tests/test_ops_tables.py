"""Tests for table operations — table add/remove."""

from __future__ import annotations

import pytest
from fcp_core import ParsedOp

from fcp_sheets.server.ops_tables import op_table
from fcp_sheets.server.resolvers import SheetsOpContext


def _setup_table_data(ctx: SheetsOpContext):
    """Set up data suitable for a table."""
    ws = ctx.active_sheet
    ws.cell(row=1, column=1, value="Name")
    ws.cell(row=1, column=2, value="Age")
    ws.cell(row=1, column=3, value="Score")
    for row in range(2, 6):
        ws.cell(row=row, column=1, value=f"Person {row - 1}")
        ws.cell(row=row, column=2, value=20 + row)
        ws.cell(row=row, column=3, value=80 + row)


class TestTableAdd:
    def test_add_basic_table(self, ctx: SheetsOpContext):
        _setup_table_data(ctx)
        op = ParsedOp(
            verb="table",
            positionals=["add", "People"],
            params={"range": "A1:C5"},
            raw="table add People range:A1:C5",
        )
        result = op_table(op, ctx)
        assert result.success
        assert "People" in result.message
        ws = ctx.active_sheet
        assert len(list(ws.tables.values())) == 1

    def test_add_table_with_style(self, ctx: SheetsOpContext):
        _setup_table_data(ctx)
        op = ParsedOp(
            verb="table",
            positionals=["add", "Styled"],
            params={"range": "A1:C5", "style": "TableStyleMedium2"},
            raw="table add Styled range:A1:C5 style:TableStyleMedium2",
        )
        result = op_table(op, ctx)
        assert result.success
        tbl = list(ctx.active_sheet.tables.values())[0]
        assert tbl.tableStyleInfo.name == "TableStyleMedium2"

    def test_add_table_with_shorthand_style(self, ctx: SheetsOpContext):
        _setup_table_data(ctx)
        op = ParsedOp(
            verb="table",
            positionals=["add", "Short"],
            params={"range": "A1:C5", "style": "medium5"},
            raw="table add Short range:A1:C5 style:medium5",
        )
        result = op_table(op, ctx)
        assert result.success
        tbl = list(ctx.active_sheet.tables.values())[0]
        assert tbl.tableStyleInfo.name == "TableStyleMedium5"

    def test_add_table_with_banded_rows(self, ctx: SheetsOpContext):
        _setup_table_data(ctx)
        op = ParsedOp(
            verb="table",
            positionals=["add", "Banded", "banded-rows"],
            params={"range": "A1:C5"},
            raw="table add Banded range:A1:C5 banded-rows",
        )
        result = op_table(op, ctx)
        assert result.success
        tbl = list(ctx.active_sheet.tables.values())[0]
        assert tbl.tableStyleInfo.showRowStripes is True

    def test_add_table_with_banded_cols(self, ctx: SheetsOpContext):
        _setup_table_data(ctx)
        op = ParsedOp(
            verb="table",
            positionals=["add", "BCols", "banded-cols"],
            params={"range": "A1:C5"},
            raw="table add BCols range:A1:C5 banded-cols",
        )
        result = op_table(op, ctx)
        assert result.success
        tbl = list(ctx.active_sheet.tables.values())[0]
        assert tbl.tableStyleInfo.showColumnStripes is True

    def test_add_table_with_first_last_col(self, ctx: SheetsOpContext):
        _setup_table_data(ctx)
        op = ParsedOp(
            verb="table",
            positionals=["add", "FL", "first-col", "last-col"],
            params={"range": "A1:C5"},
            raw="table add FL range:A1:C5 first-col last-col",
        )
        result = op_table(op, ctx)
        assert result.success
        tbl = list(ctx.active_sheet.tables.values())[0]
        assert tbl.tableStyleInfo.showFirstColumn is True
        assert tbl.tableStyleInfo.showLastColumn is True

    def test_add_duplicate_name_error(self, ctx: SheetsOpContext):
        _setup_table_data(ctx)
        op1 = ParsedOp(
            verb="table",
            positionals=["add", "DupTable"],
            params={"range": "A1:C5"},
            raw="table add DupTable range:A1:C5",
        )
        op_table(op1, ctx)

        op2 = ParsedOp(
            verb="table",
            positionals=["add", "DupTable"],
            params={"range": "A1:C5"},
            raw="table add DupTable range:A1:C5",
        )
        result = op_table(op2, ctx)
        assert not result.success
        assert "already exists" in result.message

    def test_add_missing_range(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="table",
            positionals=["add", "NoRange"],
            params={},
            raw="table add NoRange",
        )
        result = op_table(op, ctx)
        assert not result.success
        assert "range:RANGE" in result.message

    def test_add_invalid_style(self, ctx: SheetsOpContext):
        _setup_table_data(ctx)
        op = ParsedOp(
            verb="table",
            positionals=["add", "BadStyle"],
            params={"range": "A1:C5", "style": "NonExistentStyle"},
            raw="table add BadStyle range:A1:C5 style:NonExistentStyle",
        )
        result = op_table(op, ctx)
        assert not result.success
        assert "Unknown table style" in result.message

    def test_add_missing_name(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="table",
            positionals=["add"],
            params={"range": "A1:C5"},
            raw="table add range:A1:C5",
        )
        result = op_table(op, ctx)
        assert not result.success


class TestTableRemove:
    def test_remove_table(self, ctx: SheetsOpContext):
        _setup_table_data(ctx)
        op_add = ParsedOp(
            verb="table",
            positionals=["add", "ToRemove"],
            params={"range": "A1:C5"},
            raw="table add ToRemove range:A1:C5",
        )
        op_table(op_add, ctx)
        assert len(list(ctx.active_sheet.tables.values())) == 1

        op_rm = ParsedOp(
            verb="table",
            positionals=["remove", "ToRemove"],
            params={},
            raw="table remove ToRemove",
        )
        result = op_table(op_rm, ctx)
        assert result.success
        assert len(list(ctx.active_sheet.tables.values())) == 0

    def test_remove_table_not_found(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="table",
            positionals=["remove", "Ghost"],
            params={},
            raw="table remove Ghost",
        )
        result = op_table(op, ctx)
        assert not result.success
        assert "not found" in result.message


class TestTableDispatch:
    def test_unknown_subcmd(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="table",
            positionals=["update", "Foo"],
            params={},
            raw="table update Foo",
        )
        result = op_table(op, ctx)
        assert not result.success
        assert "Unknown table sub-command" in result.message

    def test_no_positionals(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="table",
            positionals=[],
            params={},
            raw="table",
        )
        result = op_table(op, ctx)
        assert not result.success
