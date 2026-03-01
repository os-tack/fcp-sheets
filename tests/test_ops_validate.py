"""Tests for data validation operations — validate verb."""

from __future__ import annotations

import pytest
from fcp_core import ParsedOp

from fcp_sheets.server.ops_validate import op_validate
from fcp_sheets.server.resolvers import SheetsOpContext


class TestValidateList:
    def test_list_with_inline_values(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["A1:A10", "list", "Yes,No,Maybe"],
            params={},
            raw="validate A1:A10 list Yes,No,Maybe",
        )
        result = op_validate(op, ctx)
        assert result.success
        assert "List validation" in result.message
        ws = ctx.active_sheet
        dvs = ws.data_validations.dataValidation
        assert len(dvs) == 1
        assert dvs[0].type == "list"

    def test_list_with_multiple_positional_values(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["B1:B5", "list", "Red", "Green", "Blue"],
            params={},
            raw="validate B1:B5 list Red Green Blue",
        )
        result = op_validate(op, ctx)
        assert result.success
        ws = ctx.active_sheet
        dvs = ws.data_validations.dataValidation
        assert len(dvs) == 1
        assert '"Red,Green,Blue"' in dvs[0].formula1

    def test_list_with_multiword_comma_items(self, ctx: SheetsOpContext):
        """Multi-word items like 'On Track' should stay intact when comma-separated."""
        # Simulates tokenizer output for: validate H4:H12 list Exceeded,On Track,At Risk
        op = ParsedOp(
            verb="validate",
            positionals=["H4:H12", "list", "Exceeded,On", "Track,At", "Risk"],
            params={},
            raw="validate H4:H12 list Exceeded,On Track,At Risk",
        )
        result = op_validate(op, ctx)
        assert result.success
        ws = ctx.active_sheet
        dvs = ws.data_validations.dataValidation
        assert len(dvs) == 1
        assert dvs[0].formula1 == '"Exceeded,On Track,At Risk"'

    def test_list_with_range(self, ctx: SheetsOpContext):
        # Set up options in a column
        ws = ctx.active_sheet
        for i, val in enumerate(["Opt1", "Opt2", "Opt3"], start=1):
            ws.cell(row=i, column=5, value=val)

        op = ParsedOp(
            verb="validate",
            positionals=["A1:A10", "list"],
            params={"range": "E1:E3"},
            raw="validate A1:A10 list range:E1:E3",
        )
        result = op_validate(op, ctx)
        assert result.success
        dvs = ws.data_validations.dataValidation
        assert len(dvs) == 1
        assert dvs[0].formula1 == "E1:E3"

    def test_list_missing_values(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["A1:A10", "list"],
            params={},
            raw="validate A1:A10 list",
        )
        result = op_validate(op, ctx)
        assert not result.success


class TestValidateNumber:
    def test_number_greater_than(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["A1:A10", "number", "gt", "0"],
            params={},
            raw="validate A1:A10 number gt 0",
        )
        result = op_validate(op, ctx)
        assert result.success
        assert "Number" in result.message
        ws = ctx.active_sheet
        dvs = ws.data_validations.dataValidation
        assert len(dvs) == 1
        assert dvs[0].type == "decimal"
        assert dvs[0].operator == "greaterThan"

    def test_number_between(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["A1:A10", "number", "between", "1", "100"],
            params={},
            raw="validate A1:A10 number between 1 100",
        )
        result = op_validate(op, ctx)
        assert result.success
        ws = ctx.active_sheet
        dvs = ws.data_validations.dataValidation
        assert dvs[0].formula1 == "1"
        assert dvs[0].formula2 == "100"

    def test_number_between_missing_value2(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["A1:A10", "number", "between", "1"],
            params={},
            raw="validate A1:A10 number between 1",
        )
        result = op_validate(op, ctx)
        assert not result.success
        assert "two values" in result.message

    def test_number_less_than_or_equal(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["B1:B5", "number", "lte", "999"],
            params={},
            raw="validate B1:B5 number lte 999",
        )
        result = op_validate(op, ctx)
        assert result.success

    def test_number_unknown_operator(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["A1:A10", "number", "like", "5"],
            params={},
            raw="validate A1:A10 number like 5",
        )
        result = op_validate(op, ctx)
        assert not result.success
        assert "Unknown operator" in result.message


class TestValidateDate:
    def test_date_greater_than(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["A1:A10", "date", "gt", "2024-01-01"],
            params={},
            raw="validate A1:A10 date gt 2024-01-01",
        )
        result = op_validate(op, ctx)
        assert result.success
        ws = ctx.active_sheet
        dvs = ws.data_validations.dataValidation
        assert dvs[0].type == "date"


class TestValidateLength:
    def test_length_less_than(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["A1:A10", "length", "lt", "255"],
            params={},
            raw="validate A1:A10 length lt 255",
        )
        result = op_validate(op, ctx)
        assert result.success
        ws = ctx.active_sheet
        dvs = ws.data_validations.dataValidation
        assert dvs[0].type == "textLength"


class TestValidateCustom:
    def test_custom_formula(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["A1:A10", "custom", "=AND(A1>0,A1<100)"],
            params={},
            raw="validate A1:A10 custom =AND(A1>0,A1<100)",
        )
        result = op_validate(op, ctx)
        assert result.success
        assert "Custom validation" in result.message
        ws = ctx.active_sheet
        dvs = ws.data_validations.dataValidation
        assert dvs[0].type == "custom"
        assert dvs[0].formula1 == "=AND(A1>0,A1<100)"

    def test_custom_missing_formula(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["A1:A10", "custom"],
            params={},
            raw="validate A1:A10 custom",
        )
        result = op_validate(op, ctx)
        assert not result.success


class TestValidateOff:
    def test_validate_off(self, ctx: SheetsOpContext):
        # First add a validation
        op_add = ParsedOp(
            verb="validate",
            positionals=["A1:A10", "list", "Yes,No"],
            params={},
            raw="validate A1:A10 list Yes,No",
        )
        op_validate(op_add, ctx)
        ws = ctx.active_sheet
        assert len(ws.data_validations.dataValidation) == 1

        # Now remove it
        op_off = ParsedOp(
            verb="validate",
            positionals=["off", "A1:A10"],
            params={},
            raw="validate off A1:A10",
        )
        result = op_validate(op_off, ctx)
        assert result.success
        assert "Removed" in result.message
        assert len(ws.data_validations.dataValidation) == 0

    def test_validate_off_not_found(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["off", "Z1:Z10"],
            params={},
            raw="validate off Z1:Z10",
        )
        result = op_validate(op, ctx)
        assert not result.success
        assert "No validation found" in result.message

    def test_validate_off_missing_range(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["off"],
            params={},
            raw="validate off",
        )
        result = op_validate(op, ctx)
        assert not result.success


class TestValidateDispatch:
    def test_unknown_type(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["A1:A10", "regex"],
            params={},
            raw="validate A1:A10 regex",
        )
        result = op_validate(op, ctx)
        assert not result.success
        assert "Unknown validation type" in result.message

    def test_no_positionals(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=[],
            params={},
            raw="validate",
        )
        result = op_validate(op, ctx)
        assert not result.success

    def test_missing_type(self, ctx: SheetsOpContext):
        op = ParsedOp(
            verb="validate",
            positionals=["A1:A10"],
            params={},
            raw="validate A1:A10",
        )
        result = op_validate(op, ctx)
        assert not result.success
