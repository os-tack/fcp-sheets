"""fcp-sheets — Spreadsheet File Context Protocol MCP server."""

from __future__ import annotations

from fcp_core.server import create_fcp_server

from fcp_sheets.adapter import SheetsAdapter
from fcp_sheets.server.reference_card import EXTRA_SECTIONS
from fcp_sheets.server.verb_registry import VERBS

adapter = SheetsAdapter()

mcp = create_fcp_server(
    domain="sheets",
    adapter=adapter,
    verbs=VERBS,
    extra_sections=EXTRA_SECTIONS,
    name="sheets-fcp",
    instructions="Spreadsheet File Context Protocol. Call sheets_help for the reference card.",
)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
