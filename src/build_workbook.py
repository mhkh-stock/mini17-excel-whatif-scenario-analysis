"""Build the Mini #17 workbook: Excel What-If & Scenario Analysis Pack.

Sheets: Cover · 1_Model · 2_DataTable_1Var · 3_DataTable_2Var ·
4_Scenarios · 5_GoalSeek · 6_BreakEven.

Every derived cell is a live formula — no hard-coded totals anywhere.
Inputs come only from data/inputs.csv (single source of truth).

Two honest Excel notes, documented in docs/WALKTHROUGH.md and on the
sheets themselves:
- Excel's native Data Table writes {=TABLE(...)} array formulas.
  LibreOffice cannot recalculate those, so the two data-table grids
  are shipped as explicit live formulas that compute exactly what a
  Data Table computes (each cell = (price - cost) x volume - fixed),
  and the walkthrough gives the exact Data -> What-If -> Data Table
  steps to rebuild them natively in Excel.
- MINIFS / MAXIFS / AGGREGATE are post-2007 functions: openpyxl
  writes them with the _xlfn. prefix they need in the file format
  (plain names show #NAME? in LibreOffice — verified 17 Nov 2026).
"""

import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "outputs" / "excel_whatif_scenario_analysis_SAMPLE.xlsx"
OUT.parent.mkdir(parents=True, exist_ok=True)

inp = {r["key"]: int(r["value"]) for r in csv.DictReader(open(ROOT / "data" / "inputs.csv"))}

PRICE = inp["model_unit_price_pkr"]      # 4500
COST = inp["model_unit_cost_pkr"]        # 2700
VOLUME = inp["model_volume_units"]       # 1200
FIXED = inp["model_fixed_costs_pkr"]     # 1500000
TARGET = inp["goalseek_target_profit_pkr"]  # 1000000

# Model-sheet cell addresses (single place, every other sheet refs these)
M_PRICE, M_COST, M_VOL, M_FIXED = "B5", "B6", "B7", "B8"
M_PROFIT = "B16"  # monthly profit on 1_Model
M = "'1_Model'!"

TITLE = Font(bold=True, size=14)
SECTION = Font(bold=True, size=11)
HDR_FILL = PatternFill("solid", fgColor="D9E2F1")
INPUT_FILL = PatternFill("solid", fgColor="BDD7EE")
INPUT_FONT = Font(bold=True)
BOLD = Font(bold=True)
MONEY = "#,##0"
MONEY2 = "#,##0.00"
PCT = "0.00%"


def title_row(ws, text, note):
    ws["A1"] = text
    ws["A1"].font = TITLE
    ws["A2"] = note
    ws["A2"].font = Font(italic=True)
    ws.merge_cells("A2:F2")


def section(ws, row, text):
    ws.cell(row=row, column=1, value=text).font = SECTION
    for c in range(1, 7):
        ws.cell(row=row, column=c).fill = HDR_FILL


def label_value(ws, row, label, value, fmt=None, bold=False, is_input=False):
    ws.cell(row=row, column=1, value=label)
    cell = ws.cell(row=row, column=2, value=value)
    if fmt:
        cell.number_format = fmt
    if bold:
        cell.font = BOLD
    if is_input:
        cell.fill = INPUT_FILL
        cell.font = INPUT_FONT
    return cell


wb = Workbook()

# ---------------------------------------------------------------- Cover
cover = wb.active
cover.title = "Cover"
cover["A1"] = "Excel What-If & Scenario Analysis Pack — SAMPLE"
cover["A1"].font = TITLE
cover["A2"] = ("SAMPLE / DEMO data only. Noor Electronics is a fictional shop; "
               "every number here is invented for teaching. Never claimed as "
               "real client or real business data.")
cover["A2"].font = Font(italic=True)
cover.merge_cells("A2:F2")
cover["A4"] = "The question: what price should Noor Electronics charge for its SoundBuds Pro earbuds?"
cover["A4"].font = BOLD
cover.merge_cells("A4:F4")
for i, (sheet, what) in enumerate([
    ("1_Model", "The pricing model — price, cost, volume, fixed costs → monthly profit. Change the blue cells and everything recalculates."),
    ("2_DataTable_1Var", "One-variable data table — profit at 11 prices, volume held at base. Plus MINIFS / MAXIFS / AGGREGATE answers."),
    ("3_DataTable_2Var", "Two-variable data table — profit at 6 prices × 5 volumes (30 cells)."),
    ("4_Scenarios", "Three named scenarios — Pessimistic / Base / Optimistic — with a scenario summary and chart."),
    ("5_GoalSeek", "Goal Seek walkthrough — what price hits a PKR 1,000,000 monthly profit? Answer verified by formula."),
    ("6_BreakEven", "Break-even — units, revenue, price, and margin of safety."),
], start=6):
    cover.cell(row=i, column=1, value=sheet).font = BOLD
    cover.cell(row=i, column=2, value=what)
    cover.merge_cells(start_row=i, start_column=2, end_row=i, end_column=6)
cover["A13"] = "Headline answers (live formulas — they read the other sheets):"
cover["A13"].font = SECTION
label_value(cover, 14, "Base monthly profit (PKR)", f"={M}{M_PROFIT}", MONEY, bold=True)
label_value(cover, 15, "Break-even units", "='6_BreakEven'!B5", MONEY2, bold=True)
label_value(cover, 16, "Goal Seek price for PKR 1,000,000 profit", "='5_GoalSeek'!B10", MONEY2, bold=True)
label_value(cover, 17, "Best scenario profit — Optimistic (PKR)", "='4_Scenarios'!B10", MONEY, bold=True)
cover["A19"] = "How to use: start on 1_Model, change a blue input, watch every sheet move. Full guide: docs/WALKTHROUGH.md"
cover["A19"].font = Font(italic=True)
cover.merge_cells("A19:F19")
cover.column_dimensions["A"].width = 46
cover.column_dimensions["B"].width = 110

# ---------------------------------------------------------------- 1_Model
ws = wb.create_sheet("1_Model")
title_row(ws, "1 — Pricing model (Noor Electronics, SAMPLE)",
          "Monthly profit = (unit price − unit cost) × volume − fixed costs. "
          "Blue cells are the only inputs — everything else is a formula.")
section(ws, 4, "INPUTS — change the blue cells")
label_value(ws, 5, "Unit price (PKR)", PRICE, MONEY, is_input=True)
label_value(ws, 6, "Unit cost (PKR)", COST, MONEY, is_input=True)
label_value(ws, 7, "Volume (units / month)", VOLUME, MONEY, is_input=True)
label_value(ws, 8, "Fixed costs (PKR / month)", FIXED, MONEY, is_input=True)
section(ws, 10, "RESULTS — all live formulas")
label_value(ws, 11, "Revenue (PKR / month)", f"={M_PRICE}*{M_VOL}", MONEY)
label_value(ws, 12, "Variable cost (PKR / month)", f"={M_COST}*{M_VOL}", MONEY)
label_value(ws, 13, "Contribution (PKR / month)", "=B11-B12", MONEY)
label_value(ws, 14, "Contribution per unit (PKR)", f"={M_PRICE}-{M_COST}", MONEY)
label_value(ws, 15, "Fixed costs (PKR / month)", f"={M_FIXED}", MONEY)
label_value(ws, 16, "Monthly profit (PKR)", "=B13-B15", MONEY, bold=True)
label_value(ws, 17, "Profit margin", "=B16/B11", PCT)
label_value(ws, 18, "Contribution margin %", "=B14/B5", PCT)
ws["A20"] = "Read it: at PKR 4,500 the shop keeps PKR 1,800 per earbud; 1,200 earbuds cover the PKR 1,500,000 fixed costs and leave PKR 660,000."
ws["A20"].font = Font(italic=True)
ws.merge_cells("A20:F20")
ws.column_dimensions["A"].width = 34
ws.column_dimensions["B"].width = 22

# ------------------------------------------------------- 2_DataTable_1Var
ws = wb.create_sheet("2_DataTable_1Var")
title_row(ws, "2 — One-variable data table: profit vs unit price (SAMPLE)",
          "Volume held at the base 1,200 units. Each profit cell is a live formula — "
          "exactly what Excel's Data Table computes. To rebuild it as a native Data Table in Excel, "
          "see docs/WALKTHROUGH.md §2 (Data → What-If Analysis → Data Table, column input cell = 1_Model!B5).")
ws["A4"] = "Unit price (PKR)"
ws["B4"] = "Monthly profit (PKR)"
for c in ("A", "B"):
    ws[f"{c}4"].font = BOLD
    ws[f"{c}4"].fill = HDR_FILL
prices1 = [inp["table1_price_first"] + i * inp["table1_price_step"]
           for i in range(inp["table1_price_count"])]
for i, p in enumerate(prices1):
    r = 5 + i
    ws.cell(row=r, column=1, value=p).number_format = MONEY
    ws.cell(row=r, column=2,
           value=f"=($A{r}-{M}{M_COST})*{M}{M_VOL}-{M}{M_FIXED}").number_format = MONEY
last1 = 4 + len(prices1)  # 15
section(ws, 17, "ANSWERS FROM THE TABLE — live formulas")
label_value(ws, 18, "Highest profit in table (PKR)", f"=MAX(B5:B{last1})", MONEY)
label_value(ws, 19, "Lowest profit in table (PKR)", f"=MIN(B5:B{last1})", MONEY)
label_value(ws, 20, "First price with profit ≥ 0 (PKR)",
            f"=_xlfn.MINIFS($A$5:$A${last1},$B$5:$B${last1},\">=0\")", MONEY)
label_value(ws, 21, "Highest profit at price ≥ 4,500 (PKR)",
            f"=_xlfn.MAXIFS($B$5:$B${last1},$A$5:$A${last1},\">=4500\")", MONEY)
label_value(ws, 22, "Largest profit — AGGREGATE (PKR)",
            f"=_xlfn.AGGREGATE(14,6,$B$5:$B${last1},1)", MONEY)
label_value(ws, 23, "Smallest profit — AGGREGATE (PKR)",
            f"=_xlfn.AGGREGATE(15,6,$B$5:$B${last1},1)", MONEY)
ws["A25"] = "Read it: profit turns positive at PKR 4,000 (PKR 60,000) — the first price above the PKR 3,950 break-even price."
ws["A25"].font = Font(italic=True)
ws.merge_cells("A25:F25")
chart = LineChart()
chart.title = "Monthly profit vs unit price"
chart.y_axis.title = "Profit (PKR)"
chart.x_axis.title = "Unit price (PKR)"
chart.height, chart.width = 7.5, 15
data = Reference(ws, min_col=2, min_row=4, max_row=last1)
cats = Reference(ws, min_col=1, min_row=5, max_row=last1)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "D4")
ws.column_dimensions["A"].width = 40
ws.column_dimensions["B"].width = 24
ws.freeze_panes = "A5"

# ------------------------------------------------------- 3_DataTable_2Var
ws = wb.create_sheet("3_DataTable_2Var")
title_row(ws, "3 — Two-variable data table: profit vs price × volume (SAMPLE)",
          "Rows = unit price, columns = monthly volume. Each cell is a live formula — "
          "exactly what Excel's Data Table computes with a row and a column input cell. "
          "Native rebuild steps: docs/WALKTHROUGH.md §3.")
ws["A4"] = "Price \\ Volume"
ws["A4"].font = BOLD
ws["A4"].fill = HDR_FILL
vols = [inp["table2_volume_first"] + i * inp["table2_volume_step"]
        for i in range(inp["table2_volume_count"])]
prices2 = [inp["table2_price_first"] + i * inp["table2_price_step"]
           for i in range(inp["table2_price_count"])]
for j, v in enumerate(vols):
    c = ws.cell(row=4, column=2 + j, value=v)
    c.font = BOLD
    c.fill = HDR_FILL
    c.number_format = MONEY
for i, p in enumerate(prices2):
    r = 5 + i
    ws.cell(row=r, column=1, value=p).number_format = MONEY
    for j in range(len(vols)):
        col = chr(ord("B") + j)
        ws.cell(row=r, column=2 + j,
               value=f"=($A{r}-{M}{M_COST})*{col}$4-{M}{M_FIXED}").number_format = MONEY
last2 = 4 + len(prices2)  # 10
lastcol = chr(ord("B") + len(vols) - 1)  # F
ws.conditional_formatting.add(f"B5:{lastcol}{last2}",
                              ColorScaleRule(start_type="min", start_color="F8696B",
                                             mid_type="percentile", mid_value=50, mid_color="FFEB84",
                                             end_type="max", end_color="63BE7B"))
section(ws, 12, "ANSWERS FROM THE TABLE — live formulas")
label_value(ws, 13, "Best cell — max profit (PKR)", f"=MAX(B5:{lastcol}{last2})", MONEY)
label_value(ws, 14, "Worst cell — min profit (PKR)", f"=MIN(B5:{lastcol}{last2})", MONEY)
label_value(ws, 15, "Profit at base price & base volume (PKR)",
            f"=INDEX(B5:{lastcol}{last2},MATCH({M}{M_PRICE},A5:A{last2},0),"
            f"MATCH({M}{M_VOL},B4:{lastcol}4,0))", MONEY)
label_value(ws, 16, "Cells making a profit (count)", f"=COUNTIF(B5:{lastcol}{last2},\">=0\")", MONEY)
label_value(ws, 17, "Cells losing money (count)", f"=COUNTIF(B5:{lastcol}{last2},\"<0\")", MONEY)
ws["A19"] = "Read it: at the lowest price (PKR 3,500) every volume loses money; at PKR 4,500 every volume makes money. Price, not volume, decides this shop's fate."
ws["A19"].font = Font(italic=True)
ws.merge_cells("A19:F19")
ws.column_dimensions["A"].width = 40
ws.column_dimensions["B"].width = 18
for col in "CDEF":
    ws.column_dimensions[col].width = 18
ws.freeze_panes = "B5"

# ------------------------------------------------------------ 4_Scenarios
ws = wb.create_sheet("4_Scenarios")
title_row(ws, "4 — Scenarios: Pessimistic / Base / Optimistic (SAMPLE)",
          "In Excel these are Data → What-If Analysis → Scenario Manager entries. "
          "The summary below is what Scenario Manager's 'Summary' report produces — as live formulas.")
headers = ["Scenario", "Unit price (PKR)", "Unit cost (PKR)",
           "Volume (units)", "Fixed costs (PKR)", "Monthly profit (PKR)"]
for j, h in enumerate(headers, start=1):
    c = ws.cell(row=4, column=j, value=h)
    c.font = BOLD
    c.fill = HDR_FILL
scenarios = [
    ("Pessimistic", inp["scenario_pessimistic_price"], inp["scenario_pessimistic_cost"],
     inp["scenario_pessimistic_volume"], inp["scenario_pessimistic_fixed"]),
    ("Base", inp["scenario_base_price"], inp["scenario_base_cost"],
     inp["scenario_base_volume"], inp["scenario_base_fixed"]),
    ("Optimistic", inp["scenario_optimistic_price"], inp["scenario_optimistic_cost"],
     inp["scenario_optimistic_volume"], inp["scenario_optimistic_fixed"]),
]
for i, (name, p, co, v, f) in enumerate(scenarios):
    r = 5 + i
    ws.cell(row=r, column=1, value=name).font = BOLD
    ws.cell(row=r, column=2, value=p).number_format = MONEY
    ws.cell(row=r, column=3, value=co).number_format = MONEY
    ws.cell(row=r, column=4, value=v).number_format = MONEY
    ws.cell(row=r, column=5, value=f).number_format = MONEY
    ws.cell(row=r, column=6, value=f"=(B{r}-C{r})*D{r}-E{r}").number_format = MONEY
section(ws, 9, "SCENARIO SUMMARY — live formulas")
label_value(ws, 10, "Best scenario profit (PKR)", "=MAX(F5:F7)", MONEY, bold=True)
label_value(ws, 11, "Worst scenario profit (PKR)", "=MIN(F5:F7)", MONEY, bold=True)
label_value(ws, 12, "Spread — best minus worst (PKR)", "=B10-B11", MONEY)
label_value(ws, 13, "Optimistic minus Base (PKR)", "=F7-F6", MONEY)
label_value(ws, 14, "Pessimistic minus Base (PKR)", "=F5-F6", MONEY)
ws["A16"] = "Read it: the range is PKR 2,960,000 wide — from losing PKR 610,000 to making PKR 2,350,000. The decision is not 'which price?' but 'which world are we in?'"
ws["A16"].font = Font(italic=True)
ws.merge_cells("A16:F16")
bchart = BarChart()
bchart.type = "col"
bchart.title = "Monthly profit by scenario"
bchart.y_axis.title = "Profit (PKR)"
bchart.height, bchart.width = 7.5, 15
bchart.add_data(Reference(ws, min_col=6, min_row=4, max_row=7), titles_from_data=True)
bchart.set_categories(Reference(ws, min_col=1, min_row=5, max_row=7))
bchart.shape = 4
ws.add_chart(bchart, "A18")
for col, w in zip("ABCDEF", (16, 18, 18, 18, 20, 22)):
    ws.column_dimensions[col].width = w

# ------------------------------------------------------------- 5_GoalSeek
ws = wb.create_sheet("5_GoalSeek")
title_row(ws, "5 — Goal Seek: what price hits the target profit? (SAMPLE)",
          "Goal Seek works backwards: pick the answer you want, Excel finds the input that gives it. "
          "Data → What-If Analysis → Goal Seek. Full click-steps: docs/WALKTHROUGH.md §4.")
section(ws, 4, "SET-UP")
label_value(ws, 5, "Target monthly profit (PKR)", TARGET, MONEY, is_input=True)
ws.cell(row=6, column=1, value="Set cell")
ws.cell(row=6, column=2, value="Monthly profit — 1_Model!B16")
ws.cell(row=7, column=1, value="By changing cell")
ws.cell(row=7, column=2, value="Unit price — 1_Model!B5")
section(ws, 9, "ANSWER — a formula, so it stays true if the inputs change")
label_value(ws, 10, "Goal Seek price (PKR)",
            f"={M}{M_COST}+($B$5+{M}{M_FIXED})/{M}{M_VOL}", MONEY2, bold=True)
label_value(ws, 11, "Profit at this price — check (PKR)",
            f"=(B10-{M}{M_COST})*{M}{M_VOL}-{M}{M_FIXED}", MONEY)
label_value(ws, 12, "Difference vs target (PKR) — must be 0", "=B11-B5", MONEY)
ws["A13"] = "Why this formula: profit = (price − cost) × volume − fixed, so price = cost + (target + fixed) ÷ volume. That is the algebra Goal Seek solves by trial — the formula just writes the answer down."
ws["A13"].font = Font(italic=True)
ws.merge_cells("A13:F13")
section(ws, 15, "SECOND EXAMPLE — break-even price (target profit 0)")
label_value(ws, 16, "Break-even price (PKR)",
            f"={M}{M_COST}+{M}{M_FIXED}/{M}{M_VOL}", MONEY2, bold=True)
label_value(ws, 17, "Profit at break-even price — check (PKR)",
            f"=(B16-{M}{M_COST})*{M}{M_VOL}-{M}{M_FIXED}", MONEY)
ws["A19"] = "Goal Seek changes ONE input to hit ONE target. It cannot optimise, it cannot handle two changing cells, and it has no memory — re-run it after any input change."
ws["A19"].font = Font(italic=True)
ws.merge_cells("A19:F19")
ws.column_dimensions["A"].width = 40
ws.column_dimensions["B"].width = 34

# ------------------------------------------------------------ 6_BreakEven
ws = wb.create_sheet("6_BreakEven")
title_row(ws, "6 — Break-even: how much must the shop sell? (SAMPLE)",
          "Break-even is where profit = 0: contribution covers exactly the fixed costs. All live formulas.")
section(ws, 4, "BREAK-EVEN — all live formulas")
label_value(ws, 5, "Break-even units (exact)", f"={M}{M_FIXED}/({M}{M_PRICE}-{M}{M_COST})", MONEY2, bold=True)
label_value(ws, 6, "Break-even units (whole units)", "=ROUNDUP(B5,0)", MONEY, bold=True)
label_value(ws, 7, "Break-even revenue — exact units (PKR)", f"=B5*{M}{M_PRICE}", MONEY)
label_value(ws, 8, "Break-even revenue — whole units (PKR)", f"=B6*{M}{M_PRICE}", MONEY)
label_value(ws, 9, "Break-even price (PKR)", f"={M}{M_COST}+{M}{M_FIXED}/{M}{M_VOL}", MONEY2)
label_value(ws, 10, "Margin of safety (units)", f"={M}{M_VOL}-B5", MONEY2)
label_value(ws, 11, "Margin of safety (%)", f"=B10/{M}{M_VOL}", PCT)
label_value(ws, 12, "Current monthly profit — cross-check (PKR)", f"={M}{M_PROFIT}", MONEY)
ws["A14"] = "Read it: the shop must sell 834 earbuds a month to break even. It sells 1,200 — a 367-unit (30.56%) cushion before losses start."
ws["A14"].font = Font(italic=True)
ws.merge_cells("A14:F14")
ws.column_dimensions["A"].width = 40
ws.column_dimensions["B"].width = 22

for sh in wb.worksheets:
    sh.sheet_properties.pageSetUpPr.fitToPage = True

wb.save(OUT)
print(f"built {OUT}")
