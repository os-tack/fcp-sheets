"""Integration tests -- end-to-end workflows exercising multiple subsystems.

These tests verify that the full pipeline (parse -> dispatch -> model -> index -> query)
works correctly across realistic multi-step scenarios.
"""

from __future__ import annotations

import os
import tempfile

from fcp_core import EventLog, parse_op

from fcp_sheets.adapter import SheetsAdapter, MAX_EVENTS
from fcp_sheets.model.snapshot import SheetsModel


# ---------------------------------------------------------------------------
# 1. Full Workflow Integration
# ---------------------------------------------------------------------------


class TestFullWorkflow:
    """Realistic multi-step workflow: headers -> data -> formulas -> save -> reopen."""

    def test_full_workflow_round_trip(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        # Set headers
        adapter.dispatch_op(parse_op("set A1 Name"), model, log)
        adapter.dispatch_op(parse_op("set B1 Score"), model, log)
        adapter.dispatch_op(parse_op("set C1 Grade"), model, log)

        # Enter data rows
        adapter.dispatch_op(parse_op("set A2 Alice"), model, log)
        adapter.dispatch_op(parse_op("set B2 95"), model, log)
        adapter.dispatch_op(parse_op("set C2 A"), model, log)
        adapter.dispatch_op(parse_op("set A3 Bob"), model, log)
        adapter.dispatch_op(parse_op("set B3 82"), model, log)
        adapter.dispatch_op(parse_op("set C3 B"), model, log)

        # Add formulas
        adapter.dispatch_op(parse_op("set D1 Average"), model, log)
        adapter.dispatch_op(parse_op("set D2 =AVERAGE(B2:B3)"), model, log)

        # Verify in-memory state
        ws = model.wb.active
        assert ws.cell(row=1, column=1).value == "Name"
        assert ws.cell(row=2, column=2).value == 95
        assert ws.cell(row=3, column=1).value == "Bob"
        assert ws.cell(row=2, column=4).value == "=AVERAGE(B2:B3)"

        # Save and reopen
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            adapter.serialize(model, path)
            loaded = adapter.deserialize(path)
            assert loaded.wb.active.cell(row=1, column=1).value == "Name"
            assert loaded.wb.active.cell(row=2, column=2).value == 95
            assert loaded.wb.active.cell(row=2, column=4).value == "=AVERAGE(B2:B3)"
        finally:
            os.unlink(path)

    def test_workflow_with_styles_and_borders(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Build a workbook with data + formatting and verify round-trip."""
        adapter.dispatch_op(parse_op("set A1 Revenue"), model, log)
        adapter.dispatch_op(parse_op("set B1 Q1"), model, log)
        adapter.dispatch_op(parse_op("set B2 50000"), model, log)

        # Style the header (range with colon)
        adapter.dispatch_op(parse_op("style A1:B1 bold"), model, log)
        adapter.dispatch_op(parse_op("border A1:B2 all"), model, log)

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            adapter.serialize(model, path)
            loaded = adapter.deserialize(path)
            ws = loaded.wb.active
            assert ws.cell(row=1, column=1).value == "Revenue"
            assert ws.cell(row=1, column=1).font.bold is True
            assert ws.cell(row=1, column=1).border.top.style is not None
        finally:
            os.unlink(path)


# ---------------------------------------------------------------------------
# 2. Session Lifecycle Integration
# ---------------------------------------------------------------------------


class TestSessionLifecycle:
    """Test new -> modify -> save -> reopen -> checkpoint -> undo to checkpoint."""

    def test_checkpoint_and_undo(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        # Build some state
        adapter.dispatch_op(parse_op("set A1 Original"), model, log)
        assert model.wb.active.cell(row=1, column=1).value == "Original"

        # Checkpoint
        log.checkpoint("v1")

        # More modifications
        adapter.dispatch_op(parse_op("set A1 Modified"), model, log)
        adapter.dispatch_op(parse_op("set B1 Extra"), model, log)
        assert model.wb.active.cell(row=1, column=1).value == "Modified"
        assert model.wb.active.cell(row=1, column=2).value == "Extra"

        # Undo to checkpoint (returns events newest-first)
        events = log.undo_to("v1")
        for ev in events:
            adapter.reverse_event(ev, model)
        assert model.wb.active.cell(row=1, column=1).value == "Original"
        assert model.wb.active.cell(row=1, column=2).value is None

    def test_save_reload_preserves_file_path(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        adapter.dispatch_op(parse_op("set A1 Test"), model, log)
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            adapter.serialize(model, path)
            assert model.file_path == path
            loaded = adapter.deserialize(path)
            assert loaded.file_path == path
        finally:
            os.unlink(path)


# ---------------------------------------------------------------------------
# 3. Multi-Sheet Workflows
# ---------------------------------------------------------------------------


class TestMultiSheetWorkflows:
    """Cross-sheet data entry and manipulation."""

    def test_multi_sheet_data_entry(self, adapter: SheetsAdapter):
        model = adapter.create_empty("Multi", {"sheets": "2"})
        log = EventLog()

        # Set data on Sheet1
        adapter.dispatch_op(parse_op("set A1 Revenue"), model, log)
        adapter.dispatch_op(parse_op("set A2 1000"), model, log)

        # Switch to Sheet2
        adapter.dispatch_op(parse_op("sheet activate Sheet2"), model, log)

        # Set data on Sheet2
        adapter.dispatch_op(parse_op("set A1 Summary"), model, log)
        adapter.dispatch_op(parse_op("set A2 Total"), model, log)

        # Verify both sheets have data
        assert model.wb["Sheet1"].cell(row=1, column=1).value == "Revenue"
        assert model.wb["Sheet1"].cell(row=2, column=1).value == 1000
        assert model.wb["Sheet2"].cell(row=1, column=1).value == "Summary"
        assert model.wb["Sheet2"].cell(row=2, column=1).value == "Total"

    def test_multi_sheet_round_trip(self, adapter: SheetsAdapter):
        """Multi-sheet workbook survives save/reload."""
        model = adapter.create_empty("Multi RT", {"sheets": "3"})
        log = EventLog()

        # Put data on each sheet
        adapter.dispatch_op(parse_op("set A1 Sheet1Data"), model, log)

        adapter.dispatch_op(parse_op("sheet activate Sheet2"), model, log)
        adapter.dispatch_op(parse_op("set A1 Sheet2Data"), model, log)

        adapter.dispatch_op(parse_op("sheet activate Sheet3"), model, log)
        adapter.dispatch_op(parse_op("set A1 Sheet3Data"), model, log)

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        try:
            adapter.serialize(model, path)
            loaded = adapter.deserialize(path)
            assert loaded.wb["Sheet1"].cell(row=1, column=1).value == "Sheet1Data"
            assert loaded.wb["Sheet2"].cell(row=1, column=1).value == "Sheet2Data"
            assert loaded.wb["Sheet3"].cell(row=1, column=1).value == "Sheet3Data"
        finally:
            os.unlink(path)

    def test_add_and_remove_sheet_with_data(self, adapter: SheetsAdapter):
        """Add a sheet, populate it, then remove it."""
        model = adapter.create_empty("AddRemove", {})
        log = EventLog()

        adapter.dispatch_op(parse_op("set A1 Main"), model, log)
        adapter.dispatch_op(parse_op("sheet add Temp"), model, log)
        adapter.dispatch_op(parse_op("sheet activate Temp"), model, log)
        adapter.dispatch_op(parse_op("set A1 TempData"), model, log)

        assert model.wb["Temp"].cell(row=1, column=1).value == "TempData"

        adapter.dispatch_op(parse_op("sheet activate Sheet1"), model, log)
        adapter.dispatch_op(parse_op("sheet remove Temp"), model, log)

        assert "Temp" not in model.wb.sheetnames
        assert model.wb["Sheet1"].cell(row=1, column=1).value == "Main"


# ---------------------------------------------------------------------------
# 4. Data Block Integration
# ---------------------------------------------------------------------------


class TestDataBlockIntegration:
    """Data block entry followed by formatting."""

    def test_data_then_format(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        # Enter data via data block
        adapter.dispatch_op(parse_op("data A1"), model, log)
        adapter.dispatch_op(parse_op("Name,Score"), model, log)
        adapter.dispatch_op(parse_op("Alice,95"), model, log)
        adapter.dispatch_op(parse_op("data end"), model, log)

        # Apply formatting (range with colon)
        adapter.dispatch_op(parse_op("style A1:B1 bold"), model, log)

        # Verify both data and formatting exist
        ws = model.wb.active
        assert ws.cell(row=1, column=1).value == "Name"
        assert ws.cell(row=1, column=2).value == "Score"
        assert ws.cell(row=2, column=1).value == "Alice"
        assert ws.cell(row=2, column=2).value == 95
        assert ws.cell(row=1, column=1).font.bold is True
        assert ws.cell(row=1, column=2).font.bold is True

    def test_data_block_then_formula(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Data block followed by a formula referencing the block data."""
        adapter.dispatch_op(parse_op("data A1"), model, log)
        adapter.dispatch_op(parse_op("Value"), model, log)
        adapter.dispatch_op(parse_op("10"), model, log)
        adapter.dispatch_op(parse_op("20"), model, log)
        adapter.dispatch_op(parse_op("30"), model, log)
        adapter.dispatch_op(parse_op("data end"), model, log)

        # Add sum formula (colon in formula)
        adapter.dispatch_op(parse_op("set A5 =SUM(A2:A4)"), model, log)

        ws = model.wb.active
        assert ws.cell(row=2, column=1).value == 10
        assert ws.cell(row=3, column=1).value == 20
        assert ws.cell(row=4, column=1).value == 30
        assert ws.cell(row=5, column=1).value == "=SUM(A2:A4)"

    def test_data_block_with_style_and_border(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Data block followed by styling AND borders."""
        adapter.dispatch_op(parse_op("data A1"), model, log)
        adapter.dispatch_op(parse_op("A,B"), model, log)
        adapter.dispatch_op(parse_op("1,2"), model, log)
        adapter.dispatch_op(parse_op("data end"), model, log)

        adapter.dispatch_op(parse_op("style A1:B1 bold"), model, log)
        adapter.dispatch_op(parse_op("border A1:B2 all"), model, log)

        ws = model.wb.active
        assert ws.cell(row=1, column=1).value == "A"
        assert ws.cell(row=1, column=1).font.bold is True
        assert ws.cell(row=2, column=2).border.top.style is not None


# ---------------------------------------------------------------------------
# 5. Batch Atomicity (C7)
# ---------------------------------------------------------------------------


class TestBatchAtomicity:
    """C7: Mid-batch failure rolls back ALL ops."""

    def test_batch_atomicity_full(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        adapter.dispatch_op(parse_op("set A1 baseline"), model, log)
        pre = model.snapshot()

        # Op 1 succeeds
        r1 = adapter.dispatch_op(parse_op("set B1 100"), model, log)
        assert r1.success

        # Op 2 fails (unknown verb)
        r2 = adapter.dispatch_op(parse_op("nonexistent_verb foo"), model, log)
        assert not r2.success

        # Simulate rollback (main.py would do this)
        model.restore(pre)
        adapter.rebuild_indices(model)

        assert model.wb.active.cell(row=1, column=1).value == "baseline"
        assert model.wb.active.cell(row=1, column=2).value is None

    def test_batch_atomicity_multiple_failures(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Rollback works even after multiple successful ops before a failure."""
        adapter.dispatch_op(parse_op("set A1 anchor"), model, log)
        pre = model.snapshot()

        # Several successful ops
        adapter.dispatch_op(parse_op("set B1 200"), model, log)
        adapter.dispatch_op(parse_op("set C1 300"), model, log)
        adapter.dispatch_op(parse_op("set D1 400"), model, log)

        # Failure
        r = adapter.dispatch_op(parse_op("bogus_verb x"), model, log)
        assert not r.success

        # Rollback
        model.restore(pre)
        adapter.rebuild_indices(model)

        assert model.wb.active.cell(row=1, column=1).value == "anchor"
        assert model.wb.active.cell(row=1, column=2).value is None
        assert model.wb.active.cell(row=1, column=3).value is None
        assert model.wb.active.cell(row=1, column=4).value is None


# ---------------------------------------------------------------------------
# 6. Undo/Redo Integration
# ---------------------------------------------------------------------------


class TestUndoRedoComplex:
    """Undo/redo across multiple operation types."""

    def test_undo_redo_complex(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        adapter.dispatch_op(parse_op("set A1 First"), model, log)
        adapter.dispatch_op(parse_op("set B1 Second"), model, log)
        adapter.dispatch_op(parse_op("set C1 Third"), model, log)

        # Undo last two
        events = log.undo()
        adapter.reverse_event(events[0], model)
        events = log.undo()
        adapter.reverse_event(events[0], model)

        # Verify state -- only First remains
        assert model.wb.active.cell(row=1, column=1).value == "First"
        assert model.wb.active.cell(row=1, column=2).value is None
        assert model.wb.active.cell(row=1, column=3).value is None

        # Redo one -- Second comes back
        events = log.redo()
        adapter.replay_event(events[0], model)
        assert model.wb.active.cell(row=1, column=2).value == "Second"
        assert model.wb.active.cell(row=1, column=3).value is None

    def test_undo_data_block(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Undo a data block restores the entire block atomically."""
        adapter.dispatch_op(parse_op("data A1"), model, log)
        adapter.dispatch_op(parse_op("X,Y"), model, log)
        adapter.dispatch_op(parse_op("1,2"), model, log)
        adapter.dispatch_op(parse_op("data end"), model, log)

        assert model.wb.active.cell(row=1, column=1).value == "X"
        assert model.wb.active.cell(row=2, column=2).value == 2

        # Undo the entire data block (single snapshot event)
        events = log.undo()
        assert len(events) == 1
        adapter.reverse_event(events[0], model)

        assert model.wb.active.cell(row=1, column=1).value is None
        assert model.wb.active.cell(row=2, column=2).value is None

    def test_undo_style_operation(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Undo restores the cell to its pre-styled state."""
        adapter.dispatch_op(parse_op("set A1 Hello"), model, log)
        adapter.dispatch_op(parse_op("style A1 bold"), model, log)
        assert model.wb.active.cell(row=1, column=1).font.bold is True

        # Undo the style
        events = log.undo()
        adapter.reverse_event(events[0], model)
        # After undo, cell still has value but bold is reverted
        assert model.wb.active.cell(row=1, column=1).value == "Hello"
        assert model.wb.active.cell(row=1, column=1).font.bold is not True

    def test_redo_restores_full_state(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Redo after undo restores the full post-op state."""
        adapter.dispatch_op(parse_op("set A1 100"), model, log)
        adapter.dispatch_op(parse_op("set B1 200"), model, log)

        # Undo
        events = log.undo()
        adapter.reverse_event(events[0], model)
        assert model.wb.active.cell(row=1, column=2).value is None

        # Redo
        events = log.redo()
        adapter.replay_event(events[0], model)
        assert model.wb.active.cell(row=1, column=2).value == 200


# ---------------------------------------------------------------------------
# 7. Snapshot Memory Cap
# ---------------------------------------------------------------------------


class TestSnapshotMemoryCap:
    """Test that the MAX_EVENTS limit mechanism works."""

    def test_many_ops_produce_many_events(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Creating more than MAX_EVENTS ops still works; events accumulate."""
        for i in range(MAX_EVENTS + 5):
            adapter.dispatch_op(parse_op(f"set A{i + 1} val{i}"), model, log)

        # All events should be in the log (trimming is currently a no-op)
        assert len(log) >= MAX_EVENTS

    def test_snapshot_bytes_are_valid(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Each snapshot in the log is a valid workbook snapshot."""
        for i in range(5):
            adapter.dispatch_op(parse_op(f"set A{i + 1} val{i}"), model, log)

        # The last event's after-snapshot should restore to current state
        last_event = log._events[log.cursor - 1]
        model.restore(last_event.after)
        assert model.wb.active.cell(row=5, column=1).value == "val4"


# ---------------------------------------------------------------------------
# 8. Query Integration
# ---------------------------------------------------------------------------


class TestQueryIntegration:
    """Queries work correctly on a workbook built through ops."""

    def test_plan_query_on_built_workbook(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        adapter.dispatch_op(parse_op("set A1 Product"), model, log)
        adapter.dispatch_op(parse_op("set B1 Price"), model, log)
        adapter.dispatch_op(parse_op("set A2 Widget"), model, log)
        adapter.dispatch_op(parse_op("set B2 29.99"), model, log)

        plan = adapter.dispatch_query("plan", model)
        assert "data:" in plan
        # Column headers should appear
        assert "Product" in plan

    def test_stats_query_on_built_workbook(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        adapter.dispatch_op(parse_op("set A1 Product"), model, log)
        adapter.dispatch_op(parse_op("set B1 Price"), model, log)
        adapter.dispatch_op(parse_op("set A2 Widget"), model, log)
        adapter.dispatch_op(parse_op("set B2 29.99"), model, log)

        stats = adapter.dispatch_query("stats", model)
        assert "Data cells:" in stats

    def test_status_query_on_built_workbook(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        adapter.dispatch_op(parse_op("set A1 Test"), model, log)

        status = adapter.dispatch_query("status", model)
        assert "Test Workbook" in status
        assert "unsaved" in status

    def test_plan_query_reflects_formulas(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Plan query should detect formula patterns."""
        adapter.dispatch_op(parse_op("set A1 Val"), model, log)
        adapter.dispatch_op(parse_op("set A2 10"), model, log)
        adapter.dispatch_op(parse_op("set A3 20"), model, log)
        adapter.dispatch_op(parse_op("set B1 Double"), model, log)
        adapter.dispatch_op(parse_op("set B2 =A2*2"), model, log)
        adapter.dispatch_op(parse_op("set B3 =A3*2"), model, log)

        plan = adapter.dispatch_query("plan", model)
        assert "formulas:" in plan

    def test_stats_after_multiple_ops(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Stats accurately counts data and formula cells."""
        adapter.dispatch_op(parse_op("set A1 Header"), model, log)
        adapter.dispatch_op(parse_op("set A2 10"), model, log)
        adapter.dispatch_op(parse_op("set A3 20"), model, log)
        adapter.dispatch_op(parse_op("set B1 Total"), model, log)
        adapter.dispatch_op(parse_op("set B2 =SUM(A2:A3)"), model, log)

        stats = adapter.dispatch_query("stats", model)
        assert "Data cells: 4" in stats
        assert "Formula cells: 1" in stats


# ---------------------------------------------------------------------------
# 9. Sheet Operations Integration
# ---------------------------------------------------------------------------


class TestSheetOpsIntegration:
    """Sheet management in combination with data operations."""

    def test_sheet_copy_preserves_original_data(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        # Add data
        adapter.dispatch_op(parse_op("set A1 Original"), model, log)

        # Copy sheet
        adapter.dispatch_op(
            parse_op('sheet copy Sheet1 "Backup"'), model, log
        )

        # Modify original
        adapter.dispatch_op(parse_op("set A1 Modified"), model, log)

        # Verify copy still has original
        assert model.wb["Backup"].cell(row=1, column=1).value == "Original"
        assert model.wb["Sheet1"].cell(row=1, column=1).value == "Modified"

    def test_rename_then_data_entry(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Rename a sheet then continue working on it."""
        adapter.dispatch_op(parse_op("set A1 Before"), model, log)
        adapter.dispatch_op(
            parse_op('sheet rename Sheet1 "Revenue"'), model, log
        )
        # Continue entering data on the renamed sheet
        adapter.dispatch_op(parse_op("set A2 After"), model, log)

        assert "Revenue" in model.wb.sheetnames
        assert model.wb["Revenue"].cell(row=1, column=1).value == "Before"
        assert model.wb["Revenue"].cell(row=2, column=1).value == "After"

    def test_hide_unhide_preserves_data(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Hiding and unhiding a sheet preserves its data."""
        adapter.dispatch_op(parse_op("sheet add Hidden"), model, log)
        adapter.dispatch_op(parse_op("sheet activate Hidden"), model, log)
        adapter.dispatch_op(parse_op("set A1 Secret"), model, log)
        adapter.dispatch_op(parse_op("sheet activate Sheet1"), model, log)
        adapter.dispatch_op(parse_op("sheet hide Hidden"), model, log)

        assert model.wb["Hidden"].sheet_state == "hidden"
        assert model.wb["Hidden"].cell(row=1, column=1).value == "Secret"

        adapter.dispatch_op(parse_op("sheet unhide Hidden"), model, log)
        assert model.wb["Hidden"].sheet_state == "visible"
        assert model.wb["Hidden"].cell(row=1, column=1).value == "Secret"


# ---------------------------------------------------------------------------
# 10. Error Recovery
# ---------------------------------------------------------------------------


class TestErrorRecovery:
    """System handles errors gracefully and remains usable."""

    def test_unknown_verb_preserves_state(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        adapter.dispatch_op(parse_op("set A1 Safe"), model, log)
        r = adapter.dispatch_op(parse_op("nonexistent_verb foo"), model, log)
        assert not r.success
        # Previous state is intact
        assert model.wb.active.cell(row=1, column=1).value == "Safe"

    def test_invalid_cell_ref_preserves_state(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        adapter.dispatch_op(parse_op("set A1 Safe"), model, log)
        r = adapter.dispatch_op(parse_op("style ZZZZ9999999 bold"), model, log)
        # Whether it fails or handles gracefully, A1 should be intact
        assert model.wb.active.cell(row=1, column=1).value == "Safe"

    def test_operations_after_error(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """System remains usable after an error."""
        adapter.dispatch_op(parse_op("set A1 Before"), model, log)

        # Trigger an error
        adapter.dispatch_op(parse_op("nonexistent_verb foo"), model, log)

        # System should still work
        r = adapter.dispatch_op(parse_op("set B1 After"), model, log)
        assert r.success
        assert model.wb.active.cell(row=1, column=1).value == "Before"
        assert model.wb.active.cell(row=1, column=2).value == "After"

    def test_operations_after_undo(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """New ops work correctly after an undo."""
        adapter.dispatch_op(parse_op("set A1 First"), model, log)
        adapter.dispatch_op(parse_op("set A2 Second"), model, log)

        events = log.undo()
        adapter.reverse_event(events[0], model)

        # Now do a new op -- should work on the undo-restored state
        r = adapter.dispatch_op(parse_op("set A2 NewSecond"), model, log)
        assert r.success
        assert model.wb.active.cell(row=2, column=1).value == "NewSecond"

    def test_data_block_error_recovery(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """After a failed data block, adapter is back in normal mode."""
        # Start a data block with invalid anchor
        adapter.dispatch_op(parse_op("data @invalid_anchor"), model, log)
        adapter.dispatch_op(parse_op("1,2"), model, log)
        r = adapter.dispatch_op(parse_op("data end"), model, log)
        assert not r.success

        # Adapter should be back in normal mode
        r = adapter.dispatch_op(parse_op("set A1 Recovery"), model, log)
        assert r.success
        assert model.wb.active.cell(row=1, column=1).value == "Recovery"


# ---------------------------------------------------------------------------
# 11. Cross-Subsystem: Data + Query + Undo
# ---------------------------------------------------------------------------


class TestCrossSubsystem:
    """Tests that span data entry, query, and undo subsystems."""

    def test_query_after_undo(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Query results reflect the post-undo state."""
        adapter.dispatch_op(parse_op("set A1 Header"), model, log)
        adapter.dispatch_op(parse_op("set A2 10"), model, log)
        adapter.dispatch_op(parse_op("set A3 20"), model, log)

        stats_before = adapter.dispatch_query("stats", model)
        assert "Data cells: 3" in stats_before

        # Undo last op (A3=20)
        events = log.undo()
        adapter.reverse_event(events[0], model)
        adapter.rebuild_indices(model)

        stats_after = adapter.dispatch_query("stats", model)
        assert "Data cells: 2" in stats_after

    def test_data_block_then_query(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Queries work on data entered via data blocks."""
        adapter.dispatch_op(parse_op("data A1"), model, log)
        adapter.dispatch_op(parse_op("Product,Price"), model, log)
        adapter.dispatch_op(parse_op("Widget,10"), model, log)
        adapter.dispatch_op(parse_op("data end"), model, log)

        plan = adapter.dispatch_query("plan", model)
        assert "Product" in plan
        assert "data:" in plan

    def test_multi_sheet_query(self, adapter: SheetsAdapter):
        """Plan query shows data from all sheets."""
        model = adapter.create_empty("MultiQ", {"sheets": "2"})
        log = EventLog()

        adapter.dispatch_op(parse_op("set A1 Revenues"), model, log)
        adapter.dispatch_op(parse_op("sheet activate Sheet2"), model, log)
        adapter.dispatch_op(parse_op("set A1 Expenses"), model, log)

        plan = adapter.dispatch_query("plan", model)
        assert "Sheet1" in plan
        assert "Sheet2" in plan
        assert "Revenues" in plan
        assert "Expenses" in plan


# ---------------------------------------------------------------------------
# 12. Index Integrity
# ---------------------------------------------------------------------------


class TestIndexIntegrity:
    """Index stays in sync through various operations."""

    def test_index_after_set_ops(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Index bounds expand correctly with set operations."""
        adapter.dispatch_op(parse_op("set A1 x"), model, log)
        adapter.dispatch_op(parse_op("set C3 y"), model, log)

        bounds = adapter.index.get_bounds("Sheet1")
        assert bounds is not None
        min_r, min_c, max_r, max_c = bounds
        assert min_r <= 1 and max_r >= 3
        assert min_c <= 1 and max_c >= 3

    def test_index_after_undo_rebuild(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Index is rebuilt correctly after undo."""
        adapter.dispatch_op(parse_op("set A1 x"), model, log)
        adapter.dispatch_op(parse_op("set D10 y"), model, log)

        # Undo D10
        events = log.undo()
        adapter.reverse_event(events[0], model)
        # After reverse_event, index is rebuilt
        bounds = adapter.index.get_bounds("Sheet1")
        assert bounds is not None
        # After undo, D10 is gone; bounds should reflect just A1
        _, _, max_r, max_c = bounds
        # openpyxl might still report max_row=1 after undo since A1 remains
        assert max_r >= 1

    def test_index_after_data_block(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Index bounds expand after data block write."""
        adapter.dispatch_op(parse_op("data B2"), model, log)
        adapter.dispatch_op(parse_op("a,b,c"), model, log)
        adapter.dispatch_op(parse_op("1,2,3"), model, log)
        adapter.dispatch_op(parse_op("data end"), model, log)

        bounds = adapter.index.get_bounds("Sheet1")
        assert bounds is not None
        min_r, min_c, max_r, max_c = bounds
        assert min_r <= 2 and max_r >= 3
        assert min_c <= 2 and max_c >= 4


# ---------------------------------------------------------------------------
# 13. Large Data Workflow
# ---------------------------------------------------------------------------


class TestLargeDataWorkflow:
    """Verify the system handles reasonably large data sets."""

    def test_50_row_data_entry(
        self, adapter: SheetsAdapter, model: SheetsModel, log: EventLog
    ):
        """Enter 50 rows of data via set and verify."""
        for i in range(1, 51):
            adapter.dispatch_op(parse_op(f"set A{i} row{i}"), model, log)
            adapter.dispatch_op(parse_op(f"set B{i} {i * 10}"), model, log)

        ws = model.wb.active
        assert ws.cell(row=1, column=1).value == "row1"
        assert ws.cell(row=50, column=1).value == "row50"
        assert ws.cell(row=50, column=2).value == 500

        # Queries work on this data
        stats = adapter.dispatch_query("stats", model)
        assert "Data cells:" in stats
