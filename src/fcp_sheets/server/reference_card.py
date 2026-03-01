"""Extra reference card sections for the sheets tool description."""

from __future__ import annotations

EXTRA_SECTIONS: dict[str, str] = {
    "Cell References": "A1 single | A1:D10 range | B:B column | 3:3 row | Sheet2!A1 cross-sheet\n  @bottom_left @bottom_right @right_top (spatial anchors, +N offset)",
    "Number Formats": (
        "General | 0 | 0.00 | #,##0 | $#,##0 | $#,##0.00\n"
        "  0% | 0.00% | yyyy-mm-dd | mm/dd/yyyy | hh:mm:ss | @"
    ),
    "Colors": (
        "#4472C4 blue  #ED7D31 orange  #A5A5A5 gray  #FFC000 gold\n"
        "  #5B9BD5 lt-blue  #70AD47 green  #FF0000 red  #00B050 dk-green\n"
        "  #C6EFCE good-fill  #FFC7CE bad-fill  #FFEB9C neutral-fill"
    ),
    "Chart Types": (
        "bar, column, line, pie, scatter, area, doughnut, radar, bubble\n"
        "  stacked-bar, stacked-column, stacked-area\n"
        "  100-bar, 100-column, 100-area\n"
        "  bar-3d, column-3d, line-3d, pie-3d, area-3d"
    ),
    "Selectors": (
        "@sheet:NAME  @range:A1:Z99  @row:N  @col:A  @type:formula|number|text|date|empty\n"
        "  @table:NAME  @name:NAME  @all  @recent  @recent:N  @not:TYPE:VALUE\n"
        "  Combine to intersect: @sheet:Revenue @col:E @type:formula"
    ),
    "Border Styles": "thin | medium | thick | dashed | dotted | double | hair\n  Sides: all | outline | top | bottom | left | right | inner | h | v",
    "Cond-Fmt Operators": "gt | lt | gte | lte | eq | neq | between | not-between",
    "Table Styles": "TableStyleLight1-21 | TableStyleMedium1-28 | TableStyleDark1-11",
    "Data Block Formats": (
        "CSV:  data A1\\n"
        "        Name,Age,City\\n"
        "        Alice,30,NYC\\n"
        "        Bob,25,LA\\n"
        "      data end\\n"
        "  Markdown:  data A1\\n"
        "               | Name | Age | City |\\n"
        "               |------|-----|------|\\n"
        "               | Alice | 30 | NYC |\\n"
        "             data end\\n"
        "  Formulas: =SUM(A1:A10), =B2*0.22, =C3/C$8 all work inside data blocks\\n"
        "  Types: bare numbers → numeric, =expr → formula, \"quoted\" → text, 007 → text"
    ),
    "Response Prefixes": (
        "+  cell/data created    ~  chart/table created\n"
        "  *  style/format modified  -  cell/range removed\n"
        "  !  error or meta         @  bulk/selector operation"
    ),
    "Example Workflow": (
        "1. sheets_session('new \"Q4 Report\"')\n"
        "  2. sheets(['sheet add Revenue'])\n"
        "  3. sheets(['merge A1:F1', 'set A1 \"Revenue Summary\"',\n"
        "             'style A1:F1 bold size:16 fill:#1a1a2e color:#FFFFFF'])\n"
        "  4. sheets(['data A2',\n"
        "             '| Month | Revenue | COGS | Gross Profit |',\n"
        "             '|-------|---------|------|--------------|',\n"
        "             '| Jan | 500000 | =B3*0.22 | =B3-C3 |',\n"
        "             '| Feb | 600000 | =B4*0.22 | =B4-C4 |',\n"
        "             '| Total | =SUM(B3:B4) | =SUM(C3:C4) | =SUM(D3:D4) |',\n"
        "             'data end'])\n"
        "  5. sheets(['style A2:D2 bold fill:#4472C4 color:#FFFFFF',\n"
        "             'style B3:D6 fmt:$#,##0',\n"
        "             'freeze A3',\n"
        "             'width A 14', 'width B:D 16'])\n"
        "  6. sheets(['chart add column title:\"Revenue\" data:B3:B4 categories:A3:A4'])\n"
        "  7. sheets_session('save as:./report.xlsx')"
    ),
    "Conventions": (
        "- Use data blocks for tables/grids — never set cells one-by-one\n"
        "  - Batch multiple ops in one sheets() call for efficiency\n"
        "  - Values beginning with = are formulas\n"
        "  - Quoted strings are text; bare numbers are numeric\n"
        "  - Active sheet is implicit target; use sheet:NAME for cross-sheet\n"
        "  - Call sheets_help after context truncation for full reference"
    ),
}
