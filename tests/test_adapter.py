"""Tests for SheetsAdapter — session lifecycle, undo/redo, batch atomicity."""

from __future__ import annotations

import os
import tempfile

import pytest
from fcp_core import EventLog, ParsedOp, parse_op

from fcp_sheets.adapter import SheetsAdapter
from fcp_sheets.model.snapshot import SheetsModel


class TestSessionLifecycle:
    def test_create_empty(self, adapter: SheetsAdapter):
        model = adapter.create_empty("Test", {})
        assert model.title == "Test"
        assert len(model.wb.sheetnames) == 1
        assert model.wb.sheetnames[0] == "Sheet1"

    def test_create_with_multiple_sheets(self, adapter: SheetsAdapter):
        model = adapter.create_empty("Multi", {"sheets": "3"})
        assert len(model.wb.sheetnames) == 3
        assert model.wb.sheetnames == ["Sheet1", "Sheet2", "Sheet3"]

    def test_serialize_deserialize(self, adapter: SheetsAdapter):
        model = adapter.create_empty("Save Test", {})
        log = EventLog()

        # Set a cell
        op = parse_op("set A1 42")
        adapter.dispatch_op(op, model, log)

        # Save
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name

        try:
            adapter.serialize(model, path)
            assert os.path.exists(path)

            # Load
            loaded = adapter.deserialize(path)
            assert loaded.wb.active.cell(row=1, column=1).value == 42
        finally:
            os.unlink(path)

    def test_round_trip(self, adapter: SheetsAdapter):
        """Create → set cells → save → reopen → verify."""
        model = adapter.create_empty("Round Trip", {})
        log = EventLog()

        ops = [
            parse_op("set A1 Name"),
            parse_op("set B1 Score"),
            parse_op("set A2 Alice"),
            parse_op("set B2 95"),
        ]
        for op in ops:
            adapter.dispatch_op(op, model, log)

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name

        try:
            adapter.serialize(model, path)
            loaded = adapter.deserialize(path)

            assert loaded.wb.active.cell(row=1, column=1).value == "Name"
            assert loaded.wb.active.cell(row=1, column=2).value == "Score"
            assert loaded.wb.active.cell(row=2, column=1).value == "Alice"
            assert loaded.wb.active.cell(row=2, column=2).value == 95
        finally:
            os.unlink(path)


class TestUndoRedo:
    def test_undo_single_op(self, adapter: SheetsAdapter):
        model = adapter.create_empty("Undo Test", {})
        log = EventLog()

        op = parse_op("set A1 42")
        adapter.dispatch_op(op, model, log)
        assert model.wb.active.cell(row=1, column=1).value == 42

        # Undo
        events = log.undo()
        assert len(events) == 1
        adapter.reverse_event(events[0], model)
        # After undo, cell should be empty (None) since we restored pre-set snapshot
        assert model.wb.active.cell(row=1, column=1).value is None

    def test_redo_after_undo(self, adapter: SheetsAdapter):
        model = adapter.create_empty("Redo Test", {})
        log = EventLog()

        op = parse_op("set A1 Hello")
        adapter.dispatch_op(op, model, log)

        events = log.undo()
        adapter.reverse_event(events[0], model)
        assert model.wb.active.cell(row=1, column=1).value is None

        replayed = log.redo()
        adapter.replay_event(replayed[0], model)
        assert model.wb.active.cell(row=1, column=1).value == "Hello"

    def test_undo_multiple_ops(self, adapter: SheetsAdapter):
        model = adapter.create_empty("Multi Undo", {})
        log = EventLog()

        adapter.dispatch_op(parse_op("set A1 1"), model, log)
        adapter.dispatch_op(parse_op("set A2 2"), model, log)
        adapter.dispatch_op(parse_op("set A3 3"), model, log)

        assert model.wb.active.cell(row=3, column=1).value == 3

        # Undo last op
        events = log.undo()
        adapter.reverse_event(events[0], model)
        assert model.wb.active.cell(row=3, column=1).value is None
        assert model.wb.active.cell(row=2, column=1).value == 2

    def test_snapshot_round_trip_fidelity(self, adapter: SheetsAdapter):
        """Byte snapshot preserves all workbook state."""
        model = adapter.create_empty("Fidelity", {})
        log = EventLog()

        # Set various data types
        adapter.dispatch_op(parse_op("set A1 42"), model, log)
        adapter.dispatch_op(parse_op('set B1 "Hello"'), model, log)
        adapter.dispatch_op(parse_op("set C1 =A1*2"), model, log)

        # Snapshot
        snapshot = model.snapshot()
        assert len(snapshot) > 0

        # Modify
        model.wb.active.cell(row=1, column=1, value=999)

        # Restore
        model.restore(snapshot)
        assert model.wb.active.cell(row=1, column=1).value == 42
        assert model.wb.active.cell(row=2, column=2).value is None  # wasn't set


class TestBatchAtomicity:
    """C7: Batch failure should rollback all ops."""

    def test_batch_rollback_on_invalid_verb(self, adapter: SheetsAdapter):
        """If any op in a batch fails, all prior ops are rolled back."""
        model = adapter.create_empty("Batch Test", {})
        log = EventLog()

        # First set a baseline value
        adapter.dispatch_op(parse_op("set A1 baseline"), model, log)

        # Take snapshot of current state
        before_batch = model.wb.active.cell(row=1, column=1).value
        assert before_batch == "baseline"

        # The batch atomicity is handled by main.py's execute_ops,
        # so we simulate it here
        pre_batch = model.snapshot()

        # Op 1: succeeds
        op1 = parse_op("set B1 100")
        r1 = adapter.dispatch_op(op1, model, log)
        assert r1.success

        # Op 2: fails (stub verb)
        op2 = parse_op("chart bar A1:D10")
        r2 = adapter.dispatch_op(op2, model, log)
        assert not r2.success  # NotImplementedError caught

        # Rollback
        model.restore(pre_batch)
        adapter.rebuild_indices(model)

        # Baseline should be intact, B1 should be rolled back
        assert model.wb.active.cell(row=1, column=1).value == "baseline"
        assert model.wb.active.cell(row=1, column=2).value is None


class TestDataBlockMode:
    def test_basic_data_block(self, adapter: SheetsAdapter):
        model = adapter.create_empty("Data Block", {})
        log = EventLog()

        # Start data block
        r1 = adapter.dispatch_op(parse_op("data A1"), model, log)
        assert r1.success

        # Feed CSV lines (these come through as parsed ops with verb being the CSV)
        adapter.dispatch_op(
            ParsedOp(verb="name,score", positionals=[], raw="Name,Score"),
            model, log,
        )
        adapter.dispatch_op(
            ParsedOp(verb="alice,95", positionals=[], raw="Alice,95"),
            model, log,
        )

        # End data block
        r_end = adapter.dispatch_op(parse_op("data end"), model, log)
        assert r_end.success
        assert "2 rows" in r_end.message

        # Verify cells
        ws = model.wb.active
        assert ws.cell(row=1, column=1).value == "Name"
        assert ws.cell(row=1, column=2).value == "Score"
        assert ws.cell(row=2, column=1).value == "Alice"
        assert ws.cell(row=2, column=2).value == 95

    def test_data_end_without_start(self, adapter: SheetsAdapter):
        model = adapter.create_empty("No Start", {})
        log = EventLog()
        r = adapter.dispatch_op(parse_op("data end"), model, log)
        assert not r.success

    def test_data_block_with_formulas(self, adapter: SheetsAdapter):
        model = adapter.create_empty("Formulas", {})
        log = EventLog()

        adapter.dispatch_op(parse_op("data A1"), model, log)
        adapter.dispatch_op(
            ParsedOp(verb="a,b,total", raw="A,B,Total"), model, log,
        )
        adapter.dispatch_op(
            ParsedOp(verb="10,20,=a2+b2", raw="10,20,=A2+B2"), model, log,
        )
        adapter.dispatch_op(parse_op("data end"), model, log)

        ws = model.wb.active
        assert ws.cell(row=2, column=3).value == "=A2+B2"

    def test_data_block_leading_zeros(self, adapter: SheetsAdapter):
        """C1: Leading zeros preserved as text."""
        model = adapter.create_empty("C1", {})
        log = EventLog()

        adapter.dispatch_op(parse_op("data A1"), model, log)
        adapter.dispatch_op(
            ParsedOp(verb="zip", raw="ZipCode"), model, log,
        )
        adapter.dispatch_op(
            ParsedOp(verb="01234", raw="01234"), model, log,
        )
        adapter.dispatch_op(parse_op("data end"), model, log)

        ws = model.wb.active
        assert ws.cell(row=2, column=1).value == "01234"  # Text, not 1234


class TestActiveSheetTracking:
    """Bug 3: Name-based active sheet tracking."""

    def test_activate_persists_through_ops(self, adapter: SheetsAdapter):
        """After activate, subsequent ops target the activated sheet."""
        model = adapter.create_empty("Track Test", {})
        log = EventLog()

        # Add and activate Revenue
        adapter.dispatch_op(parse_op("sheet add Revenue"), model, log)
        adapter.dispatch_op(parse_op("sheet activate Revenue"), model, log)

        # Set a cell — should go to Revenue
        adapter.dispatch_op(parse_op("set A1 test"), model, log)

        assert model.wb["Revenue"].cell(row=1, column=1).value == "test"
        assert adapter.index.active_sheet == "Revenue"

    def test_activate_not_overwritten_by_adapter(self, adapter: SheetsAdapter):
        """adapter.dispatch_op must not overwrite index.active_sheet."""
        model = adapter.create_empty("No Overwrite", {})
        log = EventLog()

        adapter.dispatch_op(parse_op("sheet add Revenue"), model, log)
        adapter.dispatch_op(parse_op("sheet activate Revenue"), model, log)

        # Do a normal op — previously line 186 overwrote active_sheet
        adapter.dispatch_op(parse_op("set A1 42"), model, log)

        assert adapter.index.active_sheet == "Revenue"

    def test_multi_sheet_workflow(self, adapter: SheetsAdapter):
        """Add → activate → write → activate back → write → verify."""
        model = adapter.create_empty("Multi", {})
        log = EventLog()

        adapter.dispatch_op(parse_op("sheet add Revenue"), model, log)
        adapter.dispatch_op(parse_op("sheet add Metrics"), model, log)

        adapter.dispatch_op(parse_op("sheet activate Revenue"), model, log)
        adapter.dispatch_op(parse_op('set A1 "revenue data"'), model, log)

        adapter.dispatch_op(parse_op("sheet activate Metrics"), model, log)
        adapter.dispatch_op(parse_op('set A1 "metrics data"'), model, log)

        assert model.wb["Revenue"].cell(row=1, column=1).value == "revenue data"
        assert model.wb["Metrics"].cell(row=1, column=1).value == "metrics data"


class TestDataBlockStructuralVerbs:
    """Bug 3: Structural verbs must not be swallowed during data block mode."""

    def test_sheet_during_data_block_auto_flushes(self, adapter: SheetsAdapter):
        """If 'sheet activate' arrives during data block, auto-flush first."""
        model = adapter.create_empty("Auto Flush", {})
        log = EventLog()

        # Start data block
        adapter.dispatch_op(parse_op("data A1"), model, log)
        adapter.dispatch_op(
            ParsedOp(verb="name,score", raw="Name,Score"), model, log,
        )

        # Sheet command during data block — should auto-flush
        adapter.dispatch_op(parse_op("sheet add Revenue"), model, log)

        # Data should have been flushed
        ws = model.wb["Sheet1"]
        assert ws.cell(row=1, column=1).value == "Name"
        assert ws.cell(row=1, column=2).value == "Score"

        # Sheet should have been created
        assert "Revenue" in model.wb.sheetnames


class TestGetDigest:
    def test_digest(self, adapter: SheetsAdapter):
        model = adapter.create_empty("Digest", {})
        digest = adapter.get_digest(model)
        assert "1 sheets" in digest
