"""Structured verb registry — all ~43 verb specifications for fcp-sheets."""

from __future__ import annotations

from fcp_core import VerbSpec

VERBS: list[VerbSpec] = [
    # --- CELLS ---
    VerbSpec(
        verb="set",
        syntax='set CELL VALUE [fmt:FORMAT]',
        category="cells",
        params=["fmt"],
        description="Set a cell value, formula, or formatted number.",
    ),
    VerbSpec(
        verb="data",
        syntax="data ANCHOR ... data end",
        category="cells",
        description="Bulk data entry — CSV or markdown table rows between 'data ANCHOR' and 'data end'. Preferred over set for any multi-cell input.",
    ),
    VerbSpec(
        verb="fill",
        syntax="fill SRC dir:down|right [to:CELL|count:N] [until:COL]",
        category="cells",
        params=["dir", "to", "count", "until"],
        description="Drag-fill formula/pattern down or right.",
    ),
    VerbSpec(
        verb="clear",
        syntax="clear RANGE [all]",
        category="cells",
        description="Clear cell contents (optionally including formatting).",
    ),

    # --- SHEETS ---
    VerbSpec(
        verb="sheet",
        syntax="sheet add|remove|rename|copy|hide|unhide|activate NAME [at:N]",
        category="sheets",
        params=["at"],
        description="Worksheet management.",
    ),

    # --- STYLE ---
    VerbSpec(
        verb="style",
        syntax="style RANGE|@SEL [bold] [italic] [underline] [strike] [font:NAME] [size:N] [color:#HEX] [fill:#HEX] [align:ALIGN] [valign:VALIGN] [wrap] [indent:N] [rotate:N] [fmt:FORMAT]",
        category="style",
        params=["font", "size", "color", "fill", "fill-pattern", "align", "valign", "indent", "rotate", "fmt"],
        description="Apply formatting to cells.",
    ),
    VerbSpec(
        verb="border",
        syntax="border RANGE|@SEL SIDES [line:STYLE] [color:#HEX]",
        category="style",
        params=["line", "color"],
        description="Apply borders to cells.",
    ),
    VerbSpec(
        verb="define-style",
        syntax="define-style NAME [font:F] [size:N] [bold] [fill:#HEX] [color:#HEX] [fmt:FORMAT] [align:A] [border:SIDES-STYLE]",
        category="style",
        params=["font", "size", "fill", "color", "fmt", "align", "border"],
        description="Define a reusable named style.",
    ),
    VerbSpec(
        verb="apply-style",
        syntax="apply-style NAME RANGE|@SEL",
        category="style",
        description="Apply a named style to a range.",
    ),

    # --- STRUCTURE ---
    VerbSpec(
        verb="merge",
        syntax="merge RANGE [align:center]",
        category="structure",
        params=["align"],
        description="Merge cells.",
    ),
    VerbSpec(
        verb="unmerge",
        syntax="unmerge RANGE",
        category="structure",
        description="Unmerge cells.",
    ),
    VerbSpec(
        verb="freeze",
        syntax="freeze CELL",
        category="structure",
        description="Freeze panes at cell position.",
    ),
    VerbSpec(
        verb="unfreeze",
        syntax="unfreeze",
        category="structure",
        description="Remove freeze panes.",
    ),
    VerbSpec(
        verb="filter",
        syntax="filter RANGE | filter off",
        category="structure",
        description="Add/remove auto-filter.",
    ),
    VerbSpec(
        verb="width",
        syntax="width COL|RANGE SIZE|auto",
        category="structure",
        description="Set column width.",
    ),
    VerbSpec(
        verb="height",
        syntax="height ROW|RANGE SIZE",
        category="structure",
        description="Set row height.",
    ),
    VerbSpec(
        verb="hide-col",
        syntax="hide-col COL|RANGE",
        category="structure",
        description="Hide columns.",
    ),
    VerbSpec(
        verb="hide-row",
        syntax="hide-row ROW|RANGE",
        category="structure",
        description="Hide rows.",
    ),
    VerbSpec(
        verb="unhide-col",
        syntax="unhide-col COL|RANGE",
        category="structure",
        description="Unhide columns.",
    ),
    VerbSpec(
        verb="unhide-row",
        syntax="unhide-row ROW|RANGE",
        category="structure",
        description="Unhide rows.",
    ),
    VerbSpec(
        verb="group-rows",
        syntax="group-rows RANGE [collapse]",
        category="structure",
        description="Outline group rows.",
    ),
    VerbSpec(
        verb="group-cols",
        syntax="group-cols RANGE [collapse]",
        category="structure",
        description="Outline group columns.",
    ),
    VerbSpec(
        verb="ungroup-rows",
        syntax="ungroup-rows RANGE",
        category="structure",
        description="Remove row grouping.",
    ),
    VerbSpec(
        verb="ungroup-cols",
        syntax="ungroup-cols RANGE",
        category="structure",
        description="Remove column grouping.",
    ),

    # --- CHARTS ---
    VerbSpec(
        verb="chart",
        syntax='chart add TYPE [title:"TEXT"] data:RANGE [categories:RANGE] [at:CELL] [size:WxH] [legend:POS] [style:N]',
        category="charts",
        params=["title", "data", "categories", "at", "size", "legend", "style"],
        description="Create, modify, or remove charts.",
    ),

    # --- TABLES ---
    VerbSpec(
        verb="table",
        syntax="table add NAME range:RANGE [style:STYLE] [banded-rows] [banded-cols] [first-col] [last-col]",
        category="tables",
        params=["range", "style"],
        description="Create or remove Excel tables.",
    ),

    # --- CONDITIONAL FORMATTING ---
    VerbSpec(
        verb="cond-fmt",
        syntax="cond-fmt RANGE TYPE [params...]\n  Types: cell-is OP VALUE | formula =EXPR | color-scale | data-bar | icon-set | duplicate | unique | top N | bottom N",
        category="conditional-formatting",
        params=["min-color", "max-color", "mid-color", "color", "icons", "fill", "bold"],
        description="Apply conditional formatting rules.",
    ),

    # --- DATA VALIDATION ---
    VerbSpec(
        verb="validate",
        syntax="validate RANGE TYPE [params...] | validate off RANGE",
        category="data-validation",
        description="Apply data validation rules.",
    ),

    # --- EDITING ---
    VerbSpec(
        verb="remove",
        syntax="remove @SELECTOR",
        category="editing",
        description="Delete matched content.",
    ),
    VerbSpec(
        verb="copy",
        syntax="copy RANGE to:CELL [sheet:NAME]",
        category="editing",
        params=["to", "sheet"],
        description="Copy range.",
    ),
    VerbSpec(
        verb="move",
        syntax="move RANGE to:CELL [sheet:NAME]",
        category="editing",
        params=["to", "sheet"],
        description="Move range.",
    ),
    VerbSpec(
        verb="sort",
        syntax="sort RANGE by:COL [dir:asc|desc] [by2:COL dir2:asc|desc]",
        category="editing",
        params=["by", "dir", "by2", "dir2"],
        description="Sort range by column(s).",
    ),
    VerbSpec(
        verb="insert-row",
        syntax="insert-row ROW [count:N]",
        category="editing",
        params=["count"],
        description="Insert rows.",
    ),
    VerbSpec(
        verb="insert-col",
        syntax="insert-col COL [count:N]",
        category="editing",
        params=["count"],
        description="Insert columns.",
    ),
    VerbSpec(
        verb="delete-row",
        syntax="delete-row ROW [count:N]",
        category="editing",
        params=["count"],
        description="Delete rows.",
    ),
    VerbSpec(
        verb="delete-col",
        syntax="delete-col COL [count:N]",
        category="editing",
        params=["count"],
        description="Delete columns.",
    ),

    # --- MISC ---
    VerbSpec(
        verb="name",
        syntax="name define|remove NAME [range:RANGE] [scope:SHEET]",
        category="misc",
        params=["range", "scope"],
        description="Define or remove named ranges.",
    ),
    VerbSpec(
        verb="image",
        syntax="image CELL path:PATH [size:WxH]",
        category="misc",
        params=["path", "size"],
        description="Insert image at cell.",
    ),
    VerbSpec(
        verb="link",
        syntax='link CELL url:URL [text:"TEXT"] | link CELL sheet:NAME!CELL | link off CELL',
        category="misc",
        params=["url", "text", "sheet"],
        description="Add/remove hyperlinks.",
    ),
    VerbSpec(
        verb="comment",
        syntax='comment CELL "TEXT" | comment off CELL',
        category="misc",
        description="Add/remove cell comments.",
    ),
    VerbSpec(
        verb="protect",
        syntax="protect [password:PWD]",
        category="misc",
        params=["password"],
        description="Protect sheet.",
    ),
    VerbSpec(
        verb="unprotect",
        syntax="unprotect [password:PWD]",
        category="misc",
        params=["password"],
        description="Unprotect sheet.",
    ),
    VerbSpec(
        verb="lock",
        syntax="lock RANGE",
        category="misc",
        description="Lock cell range.",
    ),
    VerbSpec(
        verb="unlock",
        syntax="unlock RANGE",
        category="misc",
        description="Unlock cell range.",
    ),
    VerbSpec(
        verb="page-setup",
        syntax="page-setup [orient:landscape|portrait] [paper:letter|a4|legal] [margins:T,R,B,L] [header:TEXT] [footer:TEXT] [print-area:RANGE] [fit-width:N] [fit-height:N] [gridlines] [center-h] [center-v]",
        category="misc",
        params=["orient", "paper", "margins", "header", "footer", "print-area", "fit-width", "fit-height"],
        description="Configure page setup for printing.",
    ),
]

VERB_MAP: dict[str, VerbSpec] = {v.verb: v for v in VERBS}
