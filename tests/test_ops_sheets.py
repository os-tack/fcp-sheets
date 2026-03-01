"""Tests for sheet management operations."""

from __future__ import annotations

import pytest
from fcp_core import ParsedOp, EventLog

from fcp_sheets.adapter import SheetsAdapter
from fcp_sheets.model.snapshot import SheetsModel
from fcp_sheets.server.ops_sheets import op_sheet
from fcp_sheets.server.resolvers import SheetsOpContext


class TestSheetAdd:
    def test_add_sheet(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="sheet", positionals=["add", "Revenue"], raw="sheet add Revenue")
        result = op_sheet(op, ctx)
        assert result.success
        assert "Revenue" in ctx.wb.sheetnames

    def test_add_sheet_at_position(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="sheet", positionals=["add", "First"],
            params={"at": "0"}, raw="sheet add First at:0",
        )
        result = op_sheet(op, ctx)
        assert result.success
        assert ctx.wb.sheetnames[0] == "First"

    def test_add_duplicate(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="sheet", positionals=["add", "Sheet1"], raw="sheet add Sheet1")
        result = op_sheet(op, ctx)
        assert not result.success
        assert "already exists" in result.message


class TestSheetRemove:
    def test_remove_sheet(self, ctx: SheetsOpContext):
        # Add a second sheet first
        ctx.wb.create_sheet("ToRemove")
        op = ParsedOp(verb="sheet", positionals=["remove", "ToRemove"], raw="sheet remove ToRemove")
        result = op_sheet(op, ctx)
        assert result.success
        assert "ToRemove" not in ctx.wb.sheetnames

    def test_remove_last_sheet(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="sheet", positionals=["remove", "Sheet1"], raw="sheet remove Sheet1")
        result = op_sheet(op, ctx)
        assert not result.success
        assert "last sheet" in result.message

    def test_remove_nonexistent(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="sheet", positionals=["remove", "Nope"], raw="sheet remove Nope")
        result = op_sheet(op, ctx)
        assert not result.success


class TestSheetRename:
    def test_rename(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="sheet", positionals=["rename", "Sheet1", "Revenue"],
            raw='sheet rename Sheet1 "Revenue"',
        )
        result = op_sheet(op, ctx)
        assert result.success
        assert "Revenue" in ctx.wb.sheetnames
        assert "Sheet1" not in ctx.wb.sheetnames

    def test_rename_nonexistent(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="sheet", positionals=["rename", "Nope", "Something"],
            raw='sheet rename Nope "Something"',
        )
        result = op_sheet(op, ctx)
        assert not result.success

    def test_rename_to_existing(self, ctx: SheetsOpContext):
        ctx.wb.create_sheet("Other")
        op = ParsedOp(
            verb="sheet", positionals=["rename", "Sheet1", "Other"],
            raw='sheet rename Sheet1 "Other"',
        )
        result = op_sheet(op, ctx)
        assert not result.success


class TestSheetCopy:
    def test_copy(self, ctx: SheetsOpContext):
        # Put data in original
        ctx.wb.active.cell(row=1, column=1, value="test")
        op = ParsedOp(
            verb="sheet", positionals=["copy", "Sheet1", "Sheet1 Copy"],
            raw='sheet copy Sheet1 "Sheet1 Copy"',
        )
        result = op_sheet(op, ctx)
        assert result.success
        assert "Sheet1 Copy" in ctx.wb.sheetnames
        assert ctx.wb["Sheet1 Copy"].cell(row=1, column=1).value == "test"


class TestSheetHideUnhide:
    def test_hide(self, ctx: SheetsOpContext):
        ctx.wb.create_sheet("Hidden")
        op = ParsedOp(verb="sheet", positionals=["hide", "Hidden"], raw="sheet hide Hidden")
        result = op_sheet(op, ctx)
        assert result.success
        assert ctx.wb["Hidden"].sheet_state == "hidden"

    def test_unhide(self, ctx: SheetsOpContext):
        ctx.wb.create_sheet("Hidden")
        ctx.wb["Hidden"].sheet_state = "hidden"
        op = ParsedOp(verb="sheet", positionals=["unhide", "Hidden"], raw="sheet unhide Hidden")
        result = op_sheet(op, ctx)
        assert result.success
        assert ctx.wb["Hidden"].sheet_state == "visible"


class TestSheetActivate:
    def test_activate(self, ctx: SheetsOpContext):
        ctx.wb.create_sheet("Revenue")
        op = ParsedOp(verb="sheet", positionals=["activate", "Revenue"], raw="sheet activate Revenue")
        result = op_sheet(op, ctx)
        assert result.success
        assert ctx.wb.active.title == "Revenue"
        assert ctx.index.active_sheet == "Revenue"

    def test_activate_nonexistent(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="sheet", positionals=["activate", "Nope"], raw="sheet activate Nope")
        result = op_sheet(op, ctx)
        assert not result.success

    def test_activate_returns_star_prefix(self, ctx: SheetsOpContext):
        ctx.wb.create_sheet("Revenue")
        op = ParsedOp(verb="sheet", positionals=["activate", "Revenue"], raw="sheet activate Revenue")
        result = op_sheet(op, ctx)
        assert result.prefix == "*"

    def test_activate_survives_remove_of_other_sheet(self, ctx: SheetsOpContext):
        """Add Revenue, activate it, remove Sheet1 — Revenue should still be active."""
        ctx.wb.create_sheet("Revenue")
        ctx.index.active_sheet = "Revenue"
        ctx.wb.active = ctx.wb.sheetnames.index("Revenue")

        # Remove Sheet1 (not the active one)
        op_remove = ParsedOp(verb="sheet", positionals=["remove", "Sheet1"], raw="sheet remove Sheet1")
        result = op_sheet(op_remove, ctx)
        assert result.success
        assert ctx.index.active_sheet == "Revenue"

    def test_remove_active_sheet_falls_back(self, ctx: SheetsOpContext):
        """Remove the active sheet — should fall back to remaining sheet."""
        ctx.wb.create_sheet("Revenue")
        ctx.index.active_sheet = "Sheet1"

        op_remove = ParsedOp(verb="sheet", positionals=["remove", "Sheet1"], raw="sheet remove Sheet1")
        result = op_sheet(op_remove, ctx)
        assert result.success
        assert ctx.index.active_sheet == "Revenue"
        assert "Sheet1" not in ctx.wb.sheetnames


class TestSheetMissingArgs:
    def test_no_action(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="sheet", positionals=[], raw="sheet")
        result = op_sheet(op, ctx)
        assert not result.success

    def test_unknown_action(self, ctx: SheetsOpContext):
        op = ParsedOp(verb="sheet", positionals=["destroy", "Sheet1"], raw="sheet destroy Sheet1")
        result = op_sheet(op, ctx)
        assert not result.success
