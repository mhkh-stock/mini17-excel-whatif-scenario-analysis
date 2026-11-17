# Mini #17 — Verification summary (SAMPLE data)

Workbook: `outputs/excel_whatif_scenario_analysis_SAMPLE.xlsx`
Checks: **181 passed, 0 failed, 0 skipped**
Method: every number recomputed independently in Python from
`data/inputs.csv`; formulas checked cell by cell; values checked after
a LibreOffice headless recalculation of the shipped workbook.

## Headline numbers

| Measure | Value |
|---|---|
| Base monthly profit (PKR) | 660,000 |
| Base revenue / variable cost / contribution (PKR) | 5,400,000 / 3,240,000 / 2,160,000 |
| Contribution per unit (PKR) | 1,800 |
| Break-even units (exact / whole) | 833.33 / 834 |
| Break-even revenue, exact / whole units (PKR) | 3,750,000 / 3,753,000 |
| Break-even price (PKR) | 3,950.00 |
| Margin of safety | 366.67 units (30.56%) |
| Goal Seek price for PKR 1,000,000 profit | PKR 4,783.33 (verified: profit at that price = PKR 1,000,000) |
| Pessimistic / Base / Optimistic profit (PKR) | -610,000 / 660,000 / 2,350,000 |
| Scenario spread (PKR) | 2,960,000 |
| 1-var table: profit at PKR 3,500 → 6,000 | -540,000 → 2,460,000 (first profitable price PKR 4,000) |
| 2-var table: worst → best cell (PKR) | -860,000 → 3,780,000 (22 of 30 cells profitable) |

## What the checks cover

- Model sheet: revenue, variable cost, contribution, profit, both margins.
- One-variable table: all 11 profit values + MAX / MIN / MINIFS / MAXIFS / AGGREGATE answers.
- Two-variable table: all 30 profit values + max / min / base-cell / profitable-cell counts.
- Scenarios: all 3 profits + best / worst / spread / both deltas vs Base.
- Goal Seek: price, profit-at-price, difference vs target (0), break-even price example.
- Break-even: exact + whole units, both revenues, price, margin of safety, cross-check profit.
- Structure: all 7 sheets present, SAMPLE label on Cover, every derived cell a live formula.
