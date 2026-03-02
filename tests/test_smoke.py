"""Smoke test: verify the entry point module can be imported.

This catches dead imports, missing symbols, and broken top-level code
that unit tests miss because they import submodules directly.
"""


def test_main_imports():
    """Importing main should not raise ImportError."""
    from fcp_sheets import main  # noqa: F401


def test_entry_point_callable():
    """The entry point function should exist and be callable."""
    from fcp_sheets.main import main

    assert callable(main)
