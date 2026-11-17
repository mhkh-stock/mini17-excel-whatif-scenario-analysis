# Data — Mini #17 Excel What-If & Scenario Analysis Pack

**SAMPLE / DEMO data only.** Fictional shop, fictional numbers.
Never claimed as real client or real business data.

## Source file

| File | What it is |
|---|---|
| `inputs.csv` | The single source of truth: the pricing-model inputs, the two data-table axes, the three scenario definitions, and the Goal Seek target. `src/build_workbook.py` reads this file and nothing else to build the workbook, so the workbook and `src/checks.py` can never drift apart. |

## The shop

**Noor Electronics** — the same fictional shop family as Mini #12,
Mini #13 and Mini #16. One product is priced here: **SoundBuds Pro
wireless earbuds**, sold online in Pakistan. Everything is monthly.

## Data dictionary — `inputs.csv`

| Key | Meaning | Value |
|---|---|---|
| `model_unit_price_pkr` | Base selling price per unit (PKR) | 4,500 |
| `model_unit_cost_pkr` | Base cost per unit (PKR) | 2,700 |
| `model_volume_units` | Base sales volume (units / month) | 1,200 |
| `model_fixed_costs_pkr` | Base fixed costs: rent, salaries, ads (PKR / month) | 1,500,000 |
| `table1_price_first` / `_step` / `_count` | One-variable table: prices swept | 3,500 → 6,000, step 250, 11 values |
| `table2_price_first` / `_step` / `_count` | Two-variable table: price axis | 3,500 → 6,000, step 500, 6 values |
| `table2_volume_first` / `_step` / `_count` | Two-variable table: volume axis | 800 → 1,600, step 200, 5 values |
| `goalseek_target_profit_pkr` | Goal Seek target monthly profit (PKR) | 1,000,000 |
| `scenario_pessimistic_*` | Pessimistic: price 4,000, cost 2,900, volume 900, fixed 1,600,000 | Profit −610,000 |
| `scenario_base_*` | Base: the model inputs above | Profit 660,000 |
| `scenario_optimistic_*` | Optimistic: price 5,000, cost 2,500, volume 1,500, fixed 1,400,000 | Profit 2,350,000 |

## The model (one formula)

Monthly profit = (unit price − unit cost) × volume − fixed costs.

Every number in the workbook that comes from this model is a live
Excel formula, never a typed-in total. Change any blue input cell on
the `1_Model` sheet and the whole workbook recalculates.

## Why these numbers

- Base contribution per unit is PKR 1,800 (4,500 − 2,700), so base
  contribution is PKR 2,160,000 and base profit is PKR 660,000.
- Break-even is 833.33 units — deliberately not a round number, so
  the workbook has to show the exact value (833.33) and the
  whole-unit value (834) separately.
- The Goal Seek target (PKR 1,000,000) needs a price of
  PKR 4,783.33 — deliberately not a round number, so the answer has
  to be verified, not guessed.
- The Pessimistic scenario deliberately loses money (−610,000):
  a scenario pack that only shows good news teaches nothing.
