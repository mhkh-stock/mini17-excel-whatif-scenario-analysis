# Mini #17 — Excel What-If & Scenario Analysis Pack

**Status:** built, verified and published (17 Nov 2026) — https://github.com/mhkh-stock/mini17-excel-whatif-scenario-analysis

**SAMPLE / DEMO data only.** Fictional shop, fictional numbers.
Never claimed as real client or real business data.

The Excel pack in the minis library: the shop from Mini #2's budget
tracker and Mini #6's formulas, now making a *decision*. One
pricing model — unit price, unit cost, volume, fixed costs →
monthly profit — then all four of Excel's What-If tools aimed at
it: a one-variable data table, a two-variable data table, three
named scenarios, and a Goal Seek walkthrough, plus a break-even
calculation that cross-checks the Goal Seek answer. Every workbook
number is a live formula, and every one of them is recomputed
independently in Python and asserted.

## The shop

Noor Electronics — the same fictional shop family as Mini #12 /
Mini #13 / Mini #16. One product is priced here: **SoundBuds Pro
wireless earbuds**, sold online in Pakistan, monthly figures.
Inputs: `data/inputs.csv` (the single source of truth) · Dictionary:
`data/README.md` · Guide: `docs/WALKTHROUGH.md`.

## Headline numbers

| Measure | Value |
|---|---|
| Base monthly profit (PKR) | **660,000** (revenue 5,400,000 − variable cost 3,240,000 − fixed 1,500,000) |
| Contribution per unit (PKR) | 1,800 (price 4,500 − cost 2,700) |
| Break-even units | **833.33** (834 whole units) |
| Break-even revenue (PKR) | 3,750,000 (3,753,000 at 834 units) |
| Break-even price (PKR) | 3,950 |
| Margin of safety | 366.67 units (**30.56%**) |
| Goal Seek price for PKR 1,000,000 profit | **PKR 4,783.33** (verified: profit at that price = PKR 1,000,000 exactly) |
| Pessimistic / Base / Optimistic profit (PKR) | −610,000 / 660,000 / **2,350,000** (spread 2,960,000) |
| 1-var table: profit at PKR 3,500 → 6,000 | −540,000 → 2,460,000 (first profitable price PKR 4,000) |
| 2-var table: worst → best cell (PKR) | −860,000 → 3,780,000 (22 of 30 cells profitable) |

## What the pack teaches

| Technique | Sheet | What it answers |
|---|---|---|
| Pricing model | `1_Model` | What is the profit at today's inputs? |
| One-variable data table | `2_DataTable_1Var` | Profit at 11 prices, volume held at base |
| Two-variable data table | `3_DataTable_2Var` | Profit at 6 prices × 5 volumes (30 cells) |
| Scenarios + summary | `4_Scenarios` | The business in a bad / normal / good world |
| Goal Seek | `5_GoalSeek` | What price hits a PKR 1,000,000 profit? |
| Break-even | `6_BreakEven` | How much must the shop sell to not lose money? |

## How to run (2 steps)

1. `cd mini17-excel-whatif-scenario-analysis`
2. `python3 src/build_workbook.py` — builds
   `outputs/excel_whatif_scenario_analysis_SAMPLE.xlsx` from
   `data/inputs.csv`
3. `python3 src/checks.py` — recomputes every workbook number in
   Python, checks every derived cell is a live formula, recalculates
   the workbook in LibreOffice and compares, then writes the verified
   `outputs/*.csv` and `outputs/SUMMARY.md`

Requires Python 3 + openpyxl (`pip install -r requirements.txt`).
LibreOffice is used for the value checks; without it, `checks.py`
skips the value checks and says so instead of faking them.

## Verification (done before delivery)

- **181 checks, 0 failures** (`python3 src/checks.py`, re-run from a
  clean `outputs/` — delete everything, rebuild, re-verify).
- Model: revenue, variable cost, contribution, profit, both margins.
- One-variable table: all 11 profit values + MAX / MIN / MINIFS /
  MAXIFS / AGGREGATE answers.
- Two-variable table: all 30 profit values + max / min / base-cell /
  profitable-cell counts.
- Scenarios: all 3 profits + best / worst / spread / both deltas vs Base.
- Goal Seek: price, profit-at-price, difference vs target (0),
  break-even-price example.
- Break-even: exact + whole units, both revenues, price, margin of
  safety, cross-check profit.
- Structure: all 7 sheets present, SAMPLE label on Cover, every
  derived cell a live formula (checked cell by cell, not by eye).
- Cross-checks that agree: Goal Seek's break-even price (PKR 3,950)
  = break-even sheet's price; the 2-var base cell (PKR 660,000) =
  the model profit = the Base scenario profit; the 1-var table's
  first profitable price (PKR 4,000) is the first table price above
  the PKR 3,950 break-even price.

## Files

| File | What it is |
|---|---|
| `outputs/excel_whatif_scenario_analysis_SAMPLE.xlsx` | The deliverable — 7 sheets, 2 charts, all live formulas |
| `src/build_workbook.py` | Builds the workbook from `data/inputs.csv` |
| `src/checks.py` | 181 independent checks + writes the verified CSVs and SUMMARY |
| `data/inputs.csv` | Single source of truth for every number |
| `data/README.md` | Data dictionary + why these numbers |
| `docs/WALKTHROUGH.md` | Technique guide: what each tool does, when to use it, exact Excel steps |
| `outputs/01_one_variable_data_table.csv` | Verified 1-var table |
| `outputs/02_two_variable_data_table.csv` | Verified 2-var table |
| `outputs/03_scenario_summary.csv` | Verified scenario table |
| `outputs/04_goalseek_breakeven.csv` | Verified Goal Seek + break-even measures |
| `outputs/SUMMARY.md` | Headline numbers + check count |

## Limits

- SAMPLE / DEMO data only — fictional shop, invented numbers.
- The data tables hold volume fixed (one-variable) or vary it
  independently of price (two-variable). Real demand falls as price
  rises; the pack does not model a demand curve.
- Native Excel `{=TABLE(...)}` Data Table array formulas are
  documented in the walkthrough, not shipped: LibreOffice cannot
  recalculate them (they come back `#VALUE!` — verified 17 Nov
  2026), so the shipped grids are explicit live formulas computing
  exactly what a Data Table computes. The walkthrough gives the
  exact Data → What-If → Data Table steps to rebuild the native
  version in Excel.
- `MINIFS` / `MAXIFS` / `AGGREGATE` are written with the `_xlfn.`
  prefix the file format requires — plain names show `#NAME?` in
  LibreOffice (verified 17 Nov 2026).
- Goal Seek's answer is exact for this linear model; with a curved
  model Goal Seek finds *a* solution, not necessarily the only one.

## Published

Live at https://github.com/mhkh-stock/mini17-excel-whatif-scenario-analysis (17 Nov 2026).
