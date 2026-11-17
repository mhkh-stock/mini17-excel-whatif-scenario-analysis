"""Verify Mini #17 end to end — 100+ independent checks.

Recomputes EVERY workbook number in Python from data/inputs.csv and
compares it against the workbook in two ways:
  1. formula checks — the cell holds a live formula, not a typed-in total
     (openpyxl, data_only=False);
  2. value checks — the formula actually evaluates to the recomputed
     number (LibreOffice headless recalc, then data_only=True).

Also writes the verified outputs (outputs/*.csv + outputs/SUMMARY.md).
Exit code 0 only when every check passes (value checks are skipped,
not faked, if LibreOffice is not installed).

Usage:  python3 src/checks.py        (from the project root)
        python3 -m src.checks        (same)
"""

import csv
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "outputs" / "excel_whatif_scenario_analysis_SAMPLE.xlsx"
OUT = ROOT / "outputs"

inp = {r["key"]: int(r["value"]) for r in csv.DictReader(open(ROOT / "data" / "inputs.csv"))}
PRICE, COST, VOL, FIXED = (inp["model_unit_price_pkr"], inp["model_unit_cost_pkr"],
                           inp["model_volume_units"], inp["model_fixed_costs_pkr"])
TARGET = inp["goalseek_target_profit_pkr"]


def profit(price, cost=COST, vol=VOL, fixed=FIXED):
    return (price - cost) * vol - fixed


# ------------------------------------------------------------------ expected
prices1 = [inp["table1_price_first"] + i * inp["table1_price_step"]
           for i in range(inp["table1_price_count"])]
prices2 = [inp["table2_price_first"] + i * inp["table2_price_step"]
           for i in range(inp["table2_price_count"])]
vols = [inp["table2_volume_first"] + i * inp["table2_volume_step"]
        for i in range(inp["table2_volume_count"])]
table1 = [(p, profit(p)) for p in prices1]
table2 = [[profit(p, vol=v) for v in vols] for p in prices2]
scen = {
    "Pessimistic": (inp["scenario_pessimistic_price"], inp["scenario_pessimistic_cost"],
                    inp["scenario_pessimistic_volume"], inp["scenario_pessimistic_fixed"]),
    "Base": (inp["scenario_base_price"], inp["scenario_base_cost"],
             inp["scenario_base_volume"], inp["scenario_base_fixed"]),
    "Optimistic": (inp["scenario_optimistic_price"], inp["scenario_optimistic_cost"],
                   inp["scenario_optimistic_volume"], inp["scenario_optimistic_fixed"]),
}
scen_profit = {k: profit(*v) for k, v in scen.items()}
goalseek_price = COST + (TARGET + FIXED) / VOL          # 4783.333...
be_price = COST + FIXED / VOL                           # 3950
be_units = FIXED / (PRICE - COST)                       # 833.333...
be_units_whole = 834
be_revenue = be_units * PRICE                           # 3750000
be_revenue_whole = be_units_whole * PRICE               # 3753000
mos_units = VOL - be_units
mos_pct = mos_units / VOL

passed, failed, skipped = [], [], []


def check(name, got, want, tol=0.5):
    if got is None:
        failed.append(f"{name}: no value (got None, want {want})")
    elif abs(got - want) <= tol:
        passed.append(name)
    else:
        failed.append(f"{name}: got {got}, want {want}")


def check_formula(ws, cell, name):
    c = ws[cell]
    if c.data_type == "f":
        passed.append(name)
    else:
        failed.append(f"{name}: {cell} is not a formula (value {c.value!r})")


# ------------------------------------------------------------------ formulas
wf = load_workbook(XLSX, data_only=False)
for name in ["Cover", "1_Model", "2_DataTable_1Var", "3_DataTable_2Var",
             "4_Scenarios", "5_GoalSeek", "6_BreakEven"]:
    if name in wf.sheetnames:
        passed.append(f"sheet present: {name}")
    else:
        failed.append(f"sheet missing: {name}")

cover = wf["Cover"]
if "SAMPLE" in str(cover["A2"].value):
    passed.append("Cover carries the SAMPLE label")
else:
    failed.append("Cover missing SAMPLE label")

m = wf["1_Model"]
for cell in [f"B{r}" for r in range(11, 19)]:
    check_formula(m, cell, f"model formula {cell}")
t1 = wf["2_DataTable_1Var"]
for r in range(5, 16):
    check_formula(t1, f"B{r}", f"1-var profit formula B{r}")
for r in range(18, 24):
    check_formula(t1, f"B{r}", f"1-var summary formula B{r}")
t2 = wf["3_DataTable_2Var"]
for r in range(5, 11):
    for col in "BCDEF":
        check_formula(t2, f"{col}{r}", f"2-var profit formula {col}{r}")
for r in range(13, 18):
    check_formula(t2, f"B{r}", f"2-var summary formula B{r}")
s = wf["4_Scenarios"]
for r in range(5, 8):
    check_formula(s, f"F{r}", f"scenario profit formula F{r}")
for r in range(10, 15):
    check_formula(s, f"B{r}", f"scenario summary formula B{r}")
g = wf["5_GoalSeek"]
for cell in ["B10", "B11", "B12", "B16", "B17"]:
    check_formula(g, cell, f"goalseek formula {cell}")
b = wf["6_BreakEven"]
for r in range(5, 13):
    check_formula(b, f"B{r}", f"breakeven formula B{r}")
for cell in ["B14", "B15", "B16", "B17"]:
    check_formula(cover, cell, f"cover headline formula {cell}")
# inputs are values, not formulas (the one place numbers are typed in)
for cell, want in [("B5", PRICE), ("B6", COST), ("B7", VOL), ("B8", FIXED)]:
    if m[cell].value == want and m[cell].data_type != "f":
        passed.append(f"model input {cell} = {want}")
    else:
        failed.append(f"model input {cell}: got {m[cell].value!r}, want {want}")

# ------------------------------------------------------------------ recalc
soffice = shutil.which("soffice") or shutil.which("libreoffice")
wv = None
if soffice:
    with tempfile.TemporaryDirectory() as td:
        r = subprocess.run([soffice, "--headless", "--calc", "--convert-to",
                            "xlsx", "--outdir", td, str(XLSX)],
                           capture_output=True, text=True, timeout=300)
        recalc = Path(td) / XLSX.name
        if r.returncode == 0 and recalc.exists():
            wv = load_workbook(recalc, data_only=True)
        else:
            failed.append(f"LibreOffice recalc failed: {r.stderr.strip()[-300:]}")
else:
    skipped.append("LibreOffice not installed — value checks skipped")

if wv is not None:
    mv, t1v, t2v, sv, gv, bv, cv = (wv[n] for n in
        ["1_Model", "2_DataTable_1Var", "3_DataTable_2Var", "4_Scenarios",
         "5_GoalSeek", "6_BreakEven", "Cover"])
    # model
    check("model revenue", mv["B11"].value, PRICE * VOL)
    check("model variable cost", mv["B12"].value, COST * VOL)
    check("model contribution", mv["B13"].value, (PRICE - COST) * VOL)
    check("model contribution per unit", mv["B14"].value, PRICE - COST)
    check("model profit", mv["B16"].value, profit(PRICE))
    check("model profit margin", mv["B17"].value, profit(PRICE) / (PRICE * VOL), tol=1e-6)
    check("model contribution margin %", mv["B18"].value, (PRICE - COST) / PRICE, tol=1e-6)
    # 1-var table (11 values) + summary
    for i, (p, pr) in enumerate(table1):
        check(f"1-var profit at price {p}", t1v.cell(row=5 + i, column=2).value, pr)
    check("1-var max profit", t1v["B18"].value, max(pr for _, pr in table1))
    check("1-var min profit", t1v["B19"].value, min(pr for _, pr in table1))
    check("MINIFS first profitable price", t1v["B20"].value,
          min(p for p, pr in table1 if pr >= 0))
    check("MAXIFS profit at price>=4500", t1v["B21"].value,
          max(pr for p, pr in table1 if p >= 4500))
    check("AGGREGATE largest profit", t1v["B22"].value, max(pr for _, pr in table1))
    check("AGGREGATE smallest profit", t1v["B23"].value, min(pr for _, pr in table1))
    # 2-var table (30 values) + summary
    for i, p in enumerate(prices2):
        for j, v in enumerate(vols):
            check(f"2-var profit price {p} vol {v}",
                  t2v.cell(row=5 + i, column=2 + j).value, table2[i][j])
    flat = [x for row in table2 for x in row]
    check("2-var max profit", t2v["B13"].value, max(flat))
    check("2-var min profit", t2v["B14"].value, min(flat))
    check("2-var base cell", t2v["B15"].value, profit(PRICE))
    check("2-var profitable cells", t2v["B16"].value, sum(1 for x in flat if x >= 0))
    check("2-var losing cells", t2v["B17"].value, sum(1 for x in flat if x < 0))
    # scenarios
    for i, name in enumerate(["Pessimistic", "Base", "Optimistic"]):
        check(f"scenario profit {name}", sv.cell(row=5 + i, column=6).value, scen_profit[name])
    check("scenario best", sv["B10"].value, max(scen_profit.values()))
    check("scenario worst", sv["B11"].value, min(scen_profit.values()))
    check("scenario spread", sv["B12"].value, max(scen_profit.values()) - min(scen_profit.values()))
    check("scenario optimistic-base", sv["B13"].value, scen_profit["Optimistic"] - scen_profit["Base"])
    check("scenario pessimistic-base", sv["B14"].value, scen_profit["Pessimistic"] - scen_profit["Base"])
    # goal seek
    check("goalseek price", gv["B10"].value, goalseek_price, tol=0.01)
    check("goalseek profit at price", gv["B11"].value, TARGET, tol=1.0)
    check("goalseek difference vs target", gv["B12"].value, 0, tol=1.0)
    check("breakeven price (goalseek sheet)", gv["B16"].value, be_price, tol=0.01)
    check("profit at breakeven price", gv["B17"].value, 0, tol=1.0)
    # break-even
    check("breakeven units exact", bv["B5"].value, be_units, tol=0.01)
    check("breakeven units whole", bv["B6"].value, be_units_whole)
    check("breakeven revenue exact", bv["B7"].value, be_revenue)
    check("breakeven revenue whole", bv["B8"].value, be_revenue_whole)
    check("breakeven price", bv["B9"].value, be_price, tol=0.01)
    check("margin of safety units", bv["B10"].value, mos_units, tol=0.01)
    check("margin of safety %", bv["B11"].value, mos_pct, tol=1e-6)
    check("breakeven cross-check profit", bv["B12"].value, profit(PRICE))
    # cover headlines
    check("cover base profit", cv["B14"].value, profit(PRICE))
    check("cover breakeven units", cv["B15"].value, be_units, tol=0.01)
    check("cover goalseek price", cv["B16"].value, goalseek_price, tol=0.01)
    check("cover best scenario profit", cv["B17"].value, scen_profit["Optimistic"])

# ------------------------------------------------------------------ outputs
with open(OUT / "01_one_variable_data_table.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["unit_price_pkr", "monthly_profit_pkr"])
    w.writerows(table1)
with open(OUT / "02_two_variable_data_table.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["unit_price_pkr"] + [f"volume_{v}" for v in vols])
    for p, row in zip(prices2, table2):
        w.writerow([p] + row)
with open(OUT / "03_scenario_summary.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["scenario", "unit_price_pkr", "unit_cost_pkr", "volume_units",
                "fixed_costs_pkr", "monthly_profit_pkr"])
    for name in ["Pessimistic", "Base", "Optimistic"]:
        w.writerow([name, *scen[name], scen_profit[name]])
with open(OUT / "04_goalseek_breakeven.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["measure", "value"])
    w.writerow(["goalseek_target_profit_pkr", TARGET])
    w.writerow(["goalseek_price_pkr", round(goalseek_price, 2)])
    w.writerow(["profit_at_goalseek_price_pkr", TARGET])
    w.writerow(["breakeven_price_pkr", round(be_price, 2)])
    w.writerow(["breakeven_units_exact", round(be_units, 2)])
    w.writerow(["breakeven_units_whole", be_units_whole])
    w.writerow(["breakeven_revenue_exact_pkr", round(be_revenue)])
    w.writerow(["breakeven_revenue_whole_pkr", be_revenue_whole])
    w.writerow(["margin_of_safety_units", round(mos_units, 2)])
    w.writerow(["margin_of_safety_pct", round(mos_pct * 100, 2)])

summary = f"""# Mini #17 — Verification summary (SAMPLE data)

Workbook: `outputs/excel_whatif_scenario_analysis_SAMPLE.xlsx`
Checks: **{len(passed)} passed, {len(failed)} failed, {len(skipped)} skipped**
Method: every number recomputed independently in Python from
`data/inputs.csv`; formulas checked cell by cell; values checked after
a LibreOffice headless recalculation of the shipped workbook.

## Headline numbers

| Measure | Value |
|---|---|
| Base monthly profit (PKR) | {profit(PRICE):,} |
| Base revenue / variable cost / contribution (PKR) | {PRICE * VOL:,} / {COST * VOL:,} / {(PRICE - COST) * VOL:,} |
| Contribution per unit (PKR) | {PRICE - COST:,} |
| Break-even units (exact / whole) | {be_units:.2f} / {be_units_whole} |
| Break-even revenue, exact / whole units (PKR) | {be_revenue:,.0f} / {be_revenue_whole:,} |
| Break-even price (PKR) | {be_price:,.2f} |
| Margin of safety | {mos_units:.2f} units ({mos_pct:.2%}) |
| Goal Seek price for PKR {TARGET:,} profit | PKR {goalseek_price:,.2f} (verified: profit at that price = PKR {TARGET:,}) |
| Pessimistic / Base / Optimistic profit (PKR) | {scen_profit['Pessimistic']:,} / {scen_profit['Base']:,} / {scen_profit['Optimistic']:,} |
| Scenario spread (PKR) | {max(scen_profit.values()) - min(scen_profit.values()):,} |
| 1-var table: profit at PKR 3,500 → 6,000 | {table1[0][1]:,} → {table1[-1][1]:,} (first profitable price PKR {min(p for p, pr in table1 if pr >= 0):,}) |
| 2-var table: worst → best cell (PKR) | {min(flat):,} → {max(flat):,} ({sum(1 for x in flat if x >= 0)} of 30 cells profitable) |

## What the checks cover

- Model sheet: revenue, variable cost, contribution, profit, both margins.
- One-variable table: all 11 profit values + MAX / MIN / MINIFS / MAXIFS / AGGREGATE answers.
- Two-variable table: all 30 profit values + max / min / base-cell / profitable-cell counts.
- Scenarios: all 3 profits + best / worst / spread / both deltas vs Base.
- Goal Seek: price, profit-at-price, difference vs target (0), break-even price example.
- Break-even: exact + whole units, both revenues, price, margin of safety, cross-check profit.
- Structure: all 7 sheets present, SAMPLE label on Cover, every derived cell a live formula.
"""
(OUT / "SUMMARY.md").write_text(summary)

print(f"checks: {len(passed)} passed, {len(failed)} failed, {len(skipped)} skipped")
for s in skipped:
    print("SKIPPED:", s)
for f in failed:
    print("FAILED:", f)
sys.exit(1 if failed else 0)
