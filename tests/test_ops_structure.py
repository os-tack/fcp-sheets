"""Tests for structure operation handlers — merge, freeze, filter, width, height, etc."""

from __future__ import annotations

import pytest
from fcp_core import ParsedOp

from fcp_sheets.adapter import SheetsAdapter
from fcp_sheets.model.snapshot import SheetsModel
from fcp_sheets.server.ops_structure import (
    op_merge,
    op_unmerge,
    op_freeze,
    op_unfreeze,
    op_filter,
    op_width,
    op_height,
    op_hide_col,
    op_hide_row,
    op_unhide_col,
    op_unhide_row,
    op_group_rows,
    op_group_cols,
    op_ungroup_rows,
    op_ungroup_cols,
)
from fcp_sheets.server.resolvers import SheetsOpContext


# ── Merge / Unmerge ──────────────────────────────────────────────────────────


class TestMerge:
    def test_merge_basic(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="merge", positionals=["A1:D1"], raw="merge A1:D1")
        result = op_merge(op, ctx)
        assert result.success
        ws = ctx.wb.active
        merged = [str(m) for m in ws.merged_cells.ranges]
        assert "A1:D1" in merged

    def test_merge_with_alignment(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="merge", positionals=["A1:D1"],
            params={"align": "center"}, raw="merge A1:D1 align:center",
        )
        result = op_merge(op, ctx)
        assert result.success
        cell = ctx.wb.active.cell(row=1, column=1)
        assert cell.alignment.horizontal == "center"

    def test_merge_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="merge", positionals=[], raw="merge")
        result = op_merge(op, ctx)
        assert not result.success

    def test_merge_invalid_range(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="merge", positionals=["INVALID"], raw="merge INVALID")
        result = op_merge(op, ctx)
        assert not result.success


class TestUnmerge:
    def test_unmerge(self, ctx: SheetsOpContext):
        ws = ctx.wb.active
        ws.merge_cells("A1:D1")
        merged_before = [str(m) for m in ws.merged_cells.ranges]
        assert len(merged_before) > 0

        op = ParsedOp(verb="unmerge", positionals=["A1:D1"], raw="unmerge A1:D1")
        result = op_unmerge(op, ctx)
        assert result.success
        merged_after = [str(m) for m in ws.merged_cells.ranges]
        assert "A1:D1" not in merged_after

    def test_unmerge_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="unmerge", positionals=[], raw="unmerge")
        result = op_unmerge(op, ctx)
        assert not result.success


# ── Freeze / Unfreeze ────────────────────────────────────────────────────────


class TestFreeze:
    def test_freeze_top_row(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="freeze", positionals=["A2"], raw="freeze A2")
        result = op_freeze(op, ctx)
        assert result.success
        assert ctx.wb.active.freeze_panes == "A2"

    def test_freeze_first_column(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="freeze", positionals=["B1"], raw="freeze B1")
        result = op_freeze(op, ctx)
        assert result.success
        assert ctx.wb.active.freeze_panes == "B1"

    def test_freeze_both(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="freeze", positionals=["B2"], raw="freeze B2")
        result = op_freeze(op, ctx)
        assert result.success
        assert ctx.wb.active.freeze_panes == "B2"

    def test_freeze_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="freeze", positionals=[], raw="freeze")
        result = op_freeze(op, ctx)
        assert not result.success

    def test_freeze_invalid_ref(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="freeze", positionals=["NOPE"], raw="freeze NOPE")
        result = op_freeze(op, ctx)
        assert not result.success


class TestUnfreeze:
    def test_unfreeze(self, ctx: SheetsOpContext):
        ws = ctx.wb.active
        ws.freeze_panes = "A2"
        assert ws.freeze_panes == "A2"

        op = ParsedOp(verb="unfreeze", positionals=[], raw="unfreeze")
        result = op_unfreeze(op, ctx)
        assert result.success
        assert ws.freeze_panes is None


# ── Filter ───────────────────────────────────────────────────────────────────


class TestFilter:
    def test_filter_set(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="filter", positionals=["A1:D10"], raw="filter A1:D10")
        result = op_filter(op, ctx)
        assert result.success
        assert ctx.wb.active.auto_filter.ref == "A1:D10"

    def test_filter_off(self, ctx: SheetsOpContext):
        ws = ctx.wb.active
        ws.auto_filter.ref = "A1:D10"

        op = ParsedOp(verb="filter", positionals=["off"], raw="filter off")
        result = op_filter(op, ctx)
        assert result.success
        assert ws.auto_filter.ref is None

    def test_filter_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="filter", positionals=[], raw="filter")
        result = op_filter(op, ctx)
        assert not result.success

    def test_filter_invalid_range(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="filter", positionals=["INVALID"], raw="filter INVALID")
        result = op_filter(op, ctx)
        assert not result.success


# ── Width ────────────────────────────────────────────────────────────────────


class TestWidth:
    def test_set_single_column_width(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="width", positionals=["A", "12"], raw="width A 12")
        result = op_width(op, ctx)
        assert result.success
        assert ctx.wb.active.column_dimensions["A"].width == 12.0

    def test_set_column_range_width(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="width", positionals=["A:C", "15"], raw="width A:C 15")
        result = op_width(op, ctx)
        assert result.success
        for col in ("A", "B", "C"):
            assert ctx.wb.active.column_dimensions[col].width == 15.0

    def test_auto_width(self, ctx: SheetsOpContext):
        ws = ctx.wb.active
        ws.cell(row=1, column=1, value="Short")
        ws.cell(row=2, column=1, value="A much longer value here")
        ctx.index.expand_bounds("Sheet1", 1, 1)
        ctx.index.expand_bounds("Sheet1", 2, 1)

        op = ParsedOp(verb="width", positionals=["A", "auto"], raw="width A auto")
        result = op_width(op, ctx)
        assert result.success
        width = ctx.wb.active.column_dimensions["A"].width
        # Auto width should be based on the longest string
        assert width > 10  # "A much longer value here" is ~24 chars

    def test_width_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="width", positionals=["A"], raw="width A")
        result = op_width(op, ctx)
        assert not result.success

    def test_width_invalid_size(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="width", positionals=["A", "abc"], raw="width A abc")
        result = op_width(op, ctx)
        assert not result.success


# ── Height ───────────────────────────────────────────────────────────────────


class TestHeight:
    def test_set_single_row_height(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="height", positionals=["1", "30"], raw="height 1 30")
        result = op_height(op, ctx)
        assert result.success
        assert ctx.wb.active.row_dimensions[1].height == 30.0

    def test_set_row_range_height(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="height", positionals=["1:3", "25"], raw="height 1:3 25")
        result = op_height(op, ctx)
        assert result.success
        for row in (1, 2, 3):
            assert ctx.wb.active.row_dimensions[row].height == 25.0

    def test_height_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="height", positionals=["1"], raw="height 1")
        result = op_height(op, ctx)
        assert not result.success

    def test_height_invalid_size(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="height", positionals=["1", "abc"], raw="height 1 abc")
        result = op_height(op, ctx)
        assert not result.success


# ── Hide / Unhide ────────────────────────────────────────────────────────────


class TestHideCol:
    def test_hide_single_column(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="hide-col", positionals=["B"], raw="hide-col B")
        result = op_hide_col(op, ctx)
        assert result.success
        assert ctx.wb.active.column_dimensions["B"].hidden is True

    def test_hide_column_range(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="hide-col", positionals=["B:D"], raw="hide-col B:D")
        result = op_hide_col(op, ctx)
        assert result.success
        for col in ("B", "C", "D"):
            assert ctx.wb.active.column_dimensions[col].hidden is True

    def test_hide_col_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="hide-col", positionals=[], raw="hide-col")
        result = op_hide_col(op, ctx)
        assert not result.success


class TestHideRow:
    def test_hide_single_row(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="hide-row", positionals=["3"], raw="hide-row 3")
        result = op_hide_row(op, ctx)
        assert result.success
        assert ctx.wb.active.row_dimensions[3].hidden is True

    def test_hide_row_range(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="hide-row", positionals=["2:4"], raw="hide-row 2:4")
        result = op_hide_row(op, ctx)
        assert result.success
        for row in (2, 3, 4):
            assert ctx.wb.active.row_dimensions[row].hidden is True

    def test_hide_row_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="hide-row", positionals=[], raw="hide-row")
        result = op_hide_row(op, ctx)
        assert not result.success


class TestUnhideCol:
    def test_unhide_column(self, ctx: SheetsOpContext):
        ws = ctx.wb.active
        ws.column_dimensions["C"].hidden = True
        assert ws.column_dimensions["C"].hidden is True

        op = ParsedOp(verb="unhide-col", positionals=["C"], raw="unhide-col C")
        result = op_unhide_col(op, ctx)
        assert result.success
        assert ws.column_dimensions["C"].hidden is False

    def test_unhide_col_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="unhide-col", positionals=[], raw="unhide-col")
        result = op_unhide_col(op, ctx)
        assert not result.success


class TestUnhideRow:
    def test_unhide_row(self, ctx: SheetsOpContext):
        ws = ctx.wb.active
        ws.row_dimensions[5].hidden = True
        assert ws.row_dimensions[5].hidden is True

        op = ParsedOp(verb="unhide-row", positionals=["5"], raw="unhide-row 5")
        result = op_unhide_row(op, ctx)
        assert result.success
        assert ws.row_dimensions[5].hidden is False

    def test_unhide_row_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="unhide-row", positionals=[], raw="unhide-row")
        result = op_unhide_row(op, ctx)
        assert not result.success


# ── Group / Ungroup ──────────────────────────────────────────────────────────


class TestGroupRows:
    def test_group_rows_basic(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="group-rows", positionals=["2:5"], raw="group-rows 2:5")
        result = op_group_rows(op, ctx)
        assert result.success
        assert "Grouped rows" in result.message

    def test_group_rows_collapsed(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="group-rows", positionals=["2:5", "collapse"],
            raw="group-rows 2:5 collapse",
        )
        result = op_group_rows(op, ctx)
        assert result.success
        assert "collapsed" in result.message

    def test_group_rows_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="group-rows", positionals=[], raw="group-rows")
        result = op_group_rows(op, ctx)
        assert not result.success


class TestGroupCols:
    def test_group_cols_basic(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="group-cols", positionals=["B:D"], raw="group-cols B:D")
        result = op_group_cols(op, ctx)
        assert result.success
        assert "Grouped columns" in result.message

    def test_group_cols_collapsed(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="group-cols", positionals=["B:D", "collapse"],
            raw="group-cols B:D collapse",
        )
        result = op_group_cols(op, ctx)
        assert result.success
        assert "collapsed" in result.message

    def test_group_cols_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="group-cols", positionals=[], raw="group-cols")
        result = op_group_cols(op, ctx)
        assert not result.success


class TestUngroupRows:
    def test_ungroup_rows(self, ctx: SheetsOpContext):
        ws = ctx.wb.active
        ws.row_dimensions.group(2, 5, outline_level=1)

        op = ParsedOp(verb="ungroup-rows", positionals=["2:5"], raw="ungroup-rows 2:5")
        result = op_ungroup_rows(op, ctx)
        assert result.success
        for row in range(2, 6):
            assert ws.row_dimensions[row].outlineLevel == 0

    def test_ungroup_rows_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="ungroup-rows", positionals=[], raw="ungroup-rows")
        result = op_ungroup_rows(op, ctx)
        assert not result.success


class TestUngroupCols:
    def test_ungroup_cols(self, ctx: SheetsOpContext):
        ws = ctx.wb.active
        ws.column_dimensions.group("B", "D", outline_level=1)

        op = ParsedOp(verb="ungroup-cols", positionals=["B:D"], raw="ungroup-cols B:D")
        result = op_ungroup_cols(op, ctx)
        assert result.success
        for col in ("B", "C", "D"):
            assert ws.column_dimensions[col].outlineLevel == 0

    def test_ungroup_cols_no_args(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="ungroup-cols", positionals=[], raw="ungroup-cols")
        result = op_ungroup_cols(op, ctx)
        assert not result.success
