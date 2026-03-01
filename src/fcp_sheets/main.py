"""fcp-sheets — Spreadsheet File Context Protocol MCP server."""

from __future__ import annotations

import re

from fcp_core.server import create_fcp_server

from fcp_sheets.adapter import SheetsAdapter
from fcp_sheets.server.reference_card import EXTRA_SECTIONS
from fcp_sheets.server.verb_registry import VERBS

# Column range pattern: 1-3 letters colon 1-3 letters (e.g. B:G, AA:ZZ)
_COL_RANGE_RE = re.compile(r"^[A-Za-z]{1,3}:[A-Za-z]{1,3}$")


def _is_sheets_positional(token: str) -> bool:
    """Domain-level positional override for spreadsheet column ranges.

    Column ranges like B:G are ambiguous with key:value pairs at the
    grammar level, so they must be handled here at the domain level.
    """
    return bool(_COL_RANGE_RE.match(token))


adapter = SheetsAdapter()

mcp = create_fcp_server(
    domain="sheets",
    adapter=adapter,
    verbs=VERBS,
    extra_sections=EXTRA_SECTIONS,
    is_positional=_is_sheets_positional,
    name="sheets-fcp",
    instructions="Spreadsheet File Context Protocol. Call sheets_help for the reference card.",
)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
