# fcp-sheets Benchmark: PE Portfolio Review

A 3-phase, 84-check benchmark comparing **fcp-sheets** (verb DSL via MCP) against **raw openpyxl** (Python scripts) for building a complex Excel workbook.

| Metric | FCP | Raw | Delta |
|---|---|---|---|
| **Audit Score** | 84/84 (100%) | 84/84 (100%) | Tie |
| **Total Time** | 559s (9.3 min) | 1,360s (22.7 min) | FCP 2.4x faster |
| **Total Cost** | $3.37 | $4.11 | FCP 18% cheaper |
| **Output Tokens** | 29,065 | 101,909 | FCP 3.5x fewer |

**Task:** Build a 6-sheet PE portfolio workbook for an LP meeting — fund overview, performance benchmarks, portfolio holdings, quarterly cash flows, scenario analysis, and sector allocation. Then apply CFO feedback (cold-open edits) and LP-meeting polish. Each phase runs in a fresh session with no prior context.

**Model:** Claude Sonnet 4.6 for both contestants. FCP uses the `sheets` MCP tool; Raw writes standalone Python scripts.

[Full prompt spec](prompt.md) &#183; [FCP output](v2-fcp-result.xlsx) &#183; [Raw output](v2-raw-result.xlsx)

---

## How FCP Builds a Spreadsheet

An agent using fcp-sheets sends batched verb operations through the `sheets()` MCP tool. Each call is a list of ops that execute atomically. Here are representative calls from building the workbook.

### Session lifecycle

```python
# Start a new workbook
sheets_session('new "Meridian Capital Partners — Q4 2025"')

# At the end, save to disk
sheets_session('save as:./v2-fcp-result.xlsx')
```

### Data blocks — the core pattern

The `data` block enters tabular data in a single operation. Bare numbers become numeric cells, `=expressions` become formulas, `"quoted strings"` stay as text.

```python
sheets([
    'data A5',
    '| Fund Size | 500000000 |',
    '| Vintage Year | 2021 |',
    '| Investment Period | "2021–2024" |',
    '| Invested Capital | 385000000 |',
    '| Realized Value | 142000000 |',
    '| Unrealized Value | 498000000 |',
    '| Total Value | =B9+B10 |',
    'data end'
])
```

This single call creates 14 cells (7 rows x 2 columns) with proper types. The equivalent in openpyxl is 7 iterations of `ws.cell(row=i, column=1, value=label)` + `ws.cell(row=i, column=2, value=val)` with manual format assignment.

### Styling — merges, fills, fonts, borders

Multiple style operations batch into one call:

```python
sheets([
    'merge A1:F1',
    'set A1 "Meridian Capital Partners"',
    'style A1:F1 bold size:20 fill:#1a1a2e color:#FFFFFF',
    'style A5:A11 bold',
    'style B5:B11 fmt:$#,##0,,"M"',
    'border A5:B11 outline line:medium',
    'border A5:B11 inner line:thin'
])
```

### Formulas — including cross-sheet references

Cross-sheet formulas work naturally inside data blocks:

```python
sheets([
    'sheet activate "Scenario Analysis"',
    'data B13',
    '| =Holdings!G13*(1+B5) | =Holdings!G13*(1+C5) | =Holdings!G13*(1+D5) |',
    '| =B13*(Holdings!I13+B6) | =C13*(Holdings!I13+C6) | =D13*(Holdings!I13+D6) |',
    '| =B14*12 | =C14*12 | =D14*12 |',
    '| =B15*(-B9) | =C15*(-C9) | =D15*(-D9) |',
    '| =B15+B16 | =C15+C16 | =D15+D16 |',
    "| =B17/'Executive Summary'!B8 | =C17/'Executive Summary'!B8 | =D17/'Executive Summary'!B8 |",
    'data end'
])
```

### Charts

One line per chart, with data ranges and options inline:

```python
sheets([
    'chart add stacked-column title:"Annual Cash Flows" data:B3:C7 categories:A3:A7',
    'chart add line title:"IRR vs Benchmark" data:F3:F7 categories:A3:A7 at:A28 size:700x300'
])
```

### Conditional formatting

```python
sheets([
    'cond-fmt B14 cell-is gt 0.15 fill:#C6EFCE',
    'cond-fmt B14 cell-is between 0.08 0.15 fill:#FFF2CC',
    'cond-fmt B14 cell-is lt 0.08 fill:#FFC7CE',
    'cond-fmt F3:F12 color-scale min-color:#F8696B mid-color:#FFEB84 max-color:#63BE7B',
    'cond-fmt J3:J12 data-bar color:#70AD47'
])
```

### Named ranges, validation, page setup

```python
sheets([
    'name define FundSize range:"Executive Summary"!B5',
    'name define InvestedCapital range:"Executive Summary"!B8',
    'validate K3:K12 list items:"Active,Exited,Written Off"',
    'page-setup orient:landscape fit-width:1 print-title-rows:1:2',
    'protect'
])
```

---

## How Raw Builds the Same Spreadsheet

The Raw approach generates a standalone Python script using openpyxl. Phase 1 produced a single 788-line script. Here's a representative excerpt — entering the same Fund Performance data and creating a chart:

```python
# Data entry — explicit cell-by-cell with format assignment (lines 209–231)
rows = [
    (2021, -125000000, 0,         "=B3+C3", "=D3",    -0.08, 0.92, -0.05),
    (2022, -140000000, 12000000,  "=B4+C4", "=E3+D4", -0.03, 0.95,  0.02),
    (2023,  -85000000, 58000000,  "=B5+C5", "=E4+D5",  0.11, 1.18,  0.09),
    (2024,  -35000000, 72000000,  "=B6+C6", "=E5+D6",  0.19, 1.48,  0.14),
    (2025,          0,         0, "=B7+C7", "=E6+D7", 0.218, 1.66,  0.16),
]
money_fmt = '$#,##0,,"M"'
for i, row_data in enumerate(rows, start=3):
    yr, contrib, distrib, ncf, cum, irr, moic, bench = row_data
    ws.cell(row=i, column=1, value=yr).font = Font(bold=True)
    ws.cell(row=i, column=2, value=contrib).number_format = money_fmt
    ws.cell(row=i, column=3, value=distrib).number_format = money_fmt
    ws.cell(row=i, column=4, value=ncf).number_format = money_fmt
    ws.cell(row=i, column=5, value=cum).number_format = money_fmt
    ws.cell(row=i, column=6, value=irr).number_format = "0.0%"
    ws.cell(row=i, column=7, value=moic).number_format = '0.00"x"'
    ws.cell(row=i, column=8, value=bench).number_format = "0.0%"

# Chart creation — 15 lines for one stacked bar chart (lines 274–292)
chart1 = BarChart()
chart1.type     = "col"
chart1.grouping = "stacked"
chart1.title    = "Annual Cash Flows"
chart1.y_axis.title = "Amount"
chart1.x_axis.title = "Year"
contrib_ref = Reference(ws, min_col=2, min_row=2, max_row=7)
distrib_ref = Reference(ws, min_col=3, min_row=2, max_row=7)
cats        = Reference(ws, min_col=1, min_row=3, max_row=7)
s1 = Series(contrib_ref, title_from_data=True)
s2 = Series(distrib_ref, title_from_data=True)
chart1.append(s1)
chart1.append(s2)
chart1.set_categories(cats)
chart1.width  = 700 / 96 * 2.54
chart1.height = 350 / 96 * 2.54
ws.add_chart(chart1, "A11")
```

The data entry requires unpacking each tuple, assigning values cell-by-cell, and setting number formats individually. The chart requires constructing Reference objects, Series objects, configuring axes, computing pixel-to-cm conversions, and anchoring. The full Phase 1 script is 788 lines; the complete raw suite across all 3 phases is 1,160 lines.

---

## Benchmark Results

### Phase-by-Phase

| Phase | FCP Time | Raw Time | FCP Cost | Raw Cost | FCP Tokens (out) | Raw Tokens (out) |
|---|---|---|---|---|---|---|
| **1: Build** | 183s (3.1 min) | 982s (16.4 min) | $1.25 | $2.39 | 9,259 | 76,772 |
| **2: Update** | 158s (2.6 min) | 229s (3.8 min) | $0.92 | $0.98 | 8,405 | 15,675 |
| **3: Polish** | 218s (3.6 min) | 149s (2.5 min) | $1.20 | $0.74 | 11,401 | 9,462 |
| **Total** | **559s** | **1,360s** | **$3.37** | **$4.11** | **29,065** | **101,909** |

### Why FCP is faster

FCP's speed advantage comes from **token density**. A single `data` block with markdown table syntax replaces dozens of `ws.cell()` calls. The 5.4x Phase 1 gap reflects this directly — FCP emits 9K output tokens vs Raw's 77K for the same workbook.

### Why FCP is cheaper

Despite 2.5x more tool calls (82 vs 33), FCP is 18% cheaper because output tokens cost more than cached input tokens. FCP produces 3.5x fewer output tokens, which dominates the cost equation.

### Where Raw wins

**Phase 3 (Polish):** Raw wins 149s vs 218s. Simple one-liner tasks (hyperlinks, footers, data bars) map directly to openpyxl properties. A 114-line script handles all 6 polish items cleanly.

**Correctness:** Both achieve 84/84 (100%). The original v2 audit showed FCP at 82/84 due to missing `print-title-rows` — fixed in v0.1.9.

**Reusable artifacts:** Raw produces version-controllable Python scripts (1,160 lines total). FCP's operations are conversational and ephemeral.

### Token Economics

| Metric | FCP | Raw | Ratio |
|---|---|---|---|
| Output tokens | 29,065 | 101,909 | FCP 3.5x fewer |
| Cache read tokens | 3,779,625 | 1,442,774 | FCP 2.6x more |
| Cache create tokens | 121,207 | 127,299 | ~equal |
| Total cost | $3.37 | $4.11 | FCP 18% cheaper |

---

## Cold-Open Modifications (Phase 2)

Phase 2 tests the realistic scenario: open an existing workbook with no prior context and apply targeted edits. The agent sees the CFO's feedback and must figure out the workbook structure before modifying it.

### FCP's open → query → modify → save flow

```python
# Open the existing file
sheets_session('open ./v2-fcp-result.xlsx')

# Inspect structure (returns sheet names, cell values, ranges)
sheets_query('list')
sheets_query('describe "Cash Flows"')

# Add a Total Return column to Cash Flows
sheets([
    'set G2 "Total Return"',
    'style G2 bold fill:#70AD47 color:#FFFFFF',
    'set G3 =E3+F3',
    'fill G3 dir:down to:G22',
    'set G23 =G22',
    'style G3:G23 fmt:$#,##0,,"M"'
])

# Rename a sheet
sheets(['sheet rename "Holdings" "Portfolio Detail"'])

# Save
sheets_session('save')
```

FCP handles cold-open modification in 25 turns (158s). The `query` tool lets the agent inspect before editing — no guessing at cell positions. Raw wrote a 258-line update script in 14 turns (229s).

---

## Correctness Audit

The [audit script](audit.py) runs 84 checks across all 3 phases, covering:

- **Structure:** Sheet names, no default "Sheet" tab
- **Data values:** All hardcoded numbers match the spec
- **Formulas:** SUM, cross-sheet references, calculated MOIC/margins/cumulative
- **Number formats:** $M display, percentages, MOIC with "x" suffix, dates
- **Styling:** Title bars, header fills, bold, frozen panes, merges
- **Conditional formatting:** IRR thresholds, MOIC color scales, data bars
- **Charts:** 7 charts across 5 sheets (stacked bar, line, bubble, area, clustered bar, doughnut)
- **Named ranges:** FundSize, InvestedCapital, TotalValue, LatestNAV, LatestMOIC
- **Data validation, borders, alternating rows, filters, sheet protection**
- **Phase 2 additions:** Total Return column, renamed sheet, Investment Thesis, Watchlist with conditional formatting, Sector Allocation sheet
- **Phase 3 polish:** Navigation hyperlinks, header consistency, data bars, footers, print titles

### Results

| | FCP | Raw |
|---|---|---|
| **Score** | 84/84 (100%) | 84/84 (100%) |
| **Failures** | None | None |

Both implementations achieve a perfect score. The original v2 audit showed FCP at 82/84 due to missing `print-title-rows` support — this was fixed in v0.1.9 by adding `print-title-rows` and `print-title-cols` parameters to the `page-setup` verb.

---

## Files

| File | Description |
|---|---|
| [README.md](README.md) | This document |
| [prompt.md](prompt.md) | Full 3-phase benchmark specification |
| [audit.py](audit.py) | 84-check audit script (requires openpyxl) |
| [v2-fcp-result.xlsx](v2-fcp-result.xlsx) | FCP output workbook (25KB) |
| [v2-raw-result.xlsx](v2-raw-result.xlsx) | Raw output workbook (25KB) |
