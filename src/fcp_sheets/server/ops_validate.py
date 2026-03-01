"""Data validation handlers — validate variants.

Supports: list, number, date, length, custom, off.
"""

from __future__ import annotations

from openpyxl.worksheet.datavalidation import DataValidation

from fcp_core import OpResult, ParsedOp

from fcp_sheets.server.resolvers import SheetsOpContext

# Operator mapping
_OP_MAP: dict[str, str] = {
    "gt": "greaterThan",
    "lt": "lessThan",
    "gte": "greaterThanOrEqual",
    "lte": "lessThanOrEqual",
    "eq": "equal",
    "neq": "notEqual",
    "ne": "notEqual",
    "between": "between",
    "not-between": "notBetween",
}

# Type mapping from DSL names to openpyxl validation types
_TYPE_MAP: dict[str, str] = {
    "number": "decimal",
    "whole": "whole",
    "decimal": "decimal",
    "date": "date",
    "length": "textLength",
    "text-length": "textLength",
    "time": "time",
}


def _validate_list(op: ParsedOp, ctx: SheetsOpContext, range_str: str) -> OpResult:
    """validate RANGE list VALUES|range:RANGE

    Values can be comma-separated inline or a range reference.
    """
    ws = ctx.active_sheet

    # Check for range: param (list from range)
    list_range = op.params.get("range")
    if list_range:
        formula1 = list_range
    else:
        # Inline values from remaining positionals
        if len(op.positionals) < 3:
            return OpResult(
                success=False,
                message="Usage: validate RANGE list Val1,Val2,Val3 | validate RANGE list range:A1:A10",
            )
        # Reconstruct the value string — the tokenizer may have split
        # multi-word items on spaces (e.g. "Exceeded,On Track,At Risk"
        # becomes positionals ["Exceeded,On", "Track,At", "Risk"]).
        raw = " ".join(op.positionals[2:])
        if "," in raw:
            # Comma-delimited: split on commas, strip whitespace per item
            items = [item.strip() for item in raw.split(",") if item.strip()]
            values = ",".join(items)
        else:
            # Space-delimited: each positional is a separate item
            values = ",".join(op.positionals[2:])
        formula1 = f'"{values}"'

    dv = DataValidation(type="list", formula1=formula1, allow_blank=True)
    dv.add(range_str)
    ws.add_data_validation(dv)

    return OpResult(success=True, message=f"List validation on {range_str}", prefix="+")


def _validate_typed(op: ParsedOp, ctx: SheetsOpContext, range_str: str) -> OpResult:
    """validate RANGE number|date|length OP VALUE [VALUE2]"""
    ws = ctx.active_sheet

    if len(op.positionals) < 4:
        return OpResult(
            success=False,
            message="Usage: validate RANGE number|date|length OP VALUE [VALUE2]",
        )

    type_name = op.positionals[1].lower()
    val_type = _TYPE_MAP.get(type_name)
    if val_type is None:
        available = ", ".join(sorted(_TYPE_MAP.keys()))
        return OpResult(success=False, message=f"Unknown validation type: {type_name!r}. Available: {available}")

    op_name = op.positionals[2].lower()
    mapped_op = _OP_MAP.get(op_name)
    if mapped_op is None:
        available = ", ".join(sorted(_OP_MAP.keys()))
        return OpResult(success=False, message=f"Unknown operator: {op_name!r}. Available: {available}")

    value1 = op.positionals[3]

    kwargs: dict = {
        "type": val_type,
        "operator": mapped_op,
        "formula1": str(value1),
        "allow_blank": True,
    }

    # For between/not-between, need a second value
    if mapped_op in ("between", "notBetween"):
        if len(op.positionals) < 5:
            return OpResult(success=False, message=f"Operator '{op_name}' requires two values")
        value2 = op.positionals[4]
        kwargs["formula2"] = str(value2)

    dv = DataValidation(**kwargs)
    dv.add(range_str)
    ws.add_data_validation(dv)

    return OpResult(success=True, message=f"{type_name.capitalize()} validation ({op_name}) on {range_str}", prefix="+")


def _validate_custom(op: ParsedOp, ctx: SheetsOpContext, range_str: str) -> OpResult:
    """validate RANGE custom FORMULA"""
    ws = ctx.active_sheet

    if len(op.positionals) < 3:
        return OpResult(success=False, message="Usage: validate RANGE custom FORMULA")

    formula = op.positionals[2]

    dv = DataValidation(type="custom", formula1=formula, allow_blank=True)
    dv.add(range_str)
    ws.add_data_validation(dv)

    return OpResult(success=True, message=f"Custom validation on {range_str}", prefix="+")


def _validate_off(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    """validate off RANGE"""
    if len(op.positionals) < 2:
        return OpResult(success=False, message="Usage: validate off RANGE")

    range_str = op.positionals[1]
    ws = ctx.active_sheet

    # Remove data validations that cover this range
    to_remove = []
    for dv in ws.data_validations.dataValidation:
        # Check if any of the dv's cell ranges overlap with the target range
        for cr in dv.cells.ranges:
            if str(cr) == range_str:
                to_remove.append(dv)
                break

    if not to_remove:
        return OpResult(success=False, message=f"No validation found on {range_str}")

    for dv in to_remove:
        ws.data_validations.dataValidation.remove(dv)

    count = len(to_remove)
    return OpResult(
        success=True,
        message=f"Removed {count} validation(s) from {range_str}",
        prefix="-",
    )


# Sub-type dispatch
_VALIDATE_TYPES: dict[str, callable] = {
    "list": _validate_list,
    "custom": _validate_custom,
    # Typed validations handled by _validate_typed
}


def op_validate(op: ParsedOp, ctx: SheetsOpContext) -> OpResult:
    """Main validate verb dispatcher.

    Syntax: validate RANGE TYPE [params...] | validate off RANGE
    """
    if not op.positionals:
        return OpResult(success=False, message="Usage: validate RANGE TYPE [params...] | validate off RANGE")

    # Handle "validate off RANGE"
    if op.positionals[0].lower() == "off":
        return _validate_off(op, ctx)

    if len(op.positionals) < 2:
        return OpResult(success=False, message="Usage: validate RANGE TYPE [params...]")

    range_str = op.positionals[0]
    val_type = op.positionals[1].lower()

    # Check specific handlers first
    handler = _VALIDATE_TYPES.get(val_type)
    if handler:
        return handler(op, ctx, range_str)

    # Check if it's a typed validation (number, date, length, etc.)
    if val_type in _TYPE_MAP:
        return _validate_typed(op, ctx, range_str)

    available = ", ".join(sorted(list(_VALIDATE_TYPES.keys()) + list(_TYPE_MAP.keys())))
    return OpResult(success=False, message=f"Unknown validation type: {val_type!r}. Available: {available}")


HANDLERS: dict[str, callable] = {
    "validate": op_validate,
}
