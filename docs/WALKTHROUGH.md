# Mini #17 — What-If & Scenario Analysis: the walkthrough

**SAMPLE / DEMO data only.** Noor Electronics is a fictional shop;
the SoundBuds Pro earbuds, the prices and the volumes are invented
for teaching. Never claimed as real client or real business data.

This guide explains the four What-If techniques in the workbook,
when to use each one, and the exact Excel steps — so you can rebuild
any of them on your own numbers. The workbook is
`outputs/excel_whatif_scenario_analysis_SAMPLE.xlsx`.

## The model everything hangs on

Monthly profit = (unit price − unit cost) × volume − fixed costs.

Base case: price PKR 4,500, cost PKR 2,700, volume 1,200 units,
fixed costs PKR 1,500,000 → contribution PKR 1,800/unit →
**profit PKR 660,000/month**. Every sheet in the workbook is a
different way of asking "what if?" about this one formula. Change
any blue input cell on `1_Model` and all six sheets recalculate.

---

## 1. The pricing model — `1_Model`

**What it is:** the inputs and the profit calculation, each step on
its own row so you can see where the money goes: revenue
PKR 5,400,000 → variable cost PKR 3,240,000 → contribution
PKR 2,160,000 → minus fixed PKR 1,500,000 → profit PKR 660,000.

**When to use it:** always, first. A What-If analysis without a
clean model is just guessing in a nicer font. If you cannot point
at the one cell that is "the answer" (here `1_Model!B16`, monthly
profit), none of the four techniques below has anything to work on.

**Rule:** inputs are typed numbers, results are formulas. If you
ever type a total by hand, the model is broken — change a price
later and the total will lie to you.

---

## 2. One-variable data table — `2_DataTable_1Var`

**What it is:** profit at 11 prices (PKR 3,500 → 6,000, step 250),
volume held at the base 1,200 units.

**What it answers:** "If I change *only the price*, what happens
to profit?" Profit runs from −PKR 540,000 (at 3,500) to
+PKR 2,460,000 (at 6,000) and turns positive at **PKR 4,000**.

**When to use it:** one input is uncertain and you want the whole
curve, not one point. The classic use is pricing — and the classic
mistake is reading the table as a recommendation: the table assumes
volume stays at 1,200 at every price, which is false in real life
(higher prices usually cut volume). That is exactly why the
two-variable table exists next.

**Exact Excel steps (native Data Table):**
1. Build the model on `1_Model` (done — profit is `B16`).
2. On a new sheet, list the prices down column A (3,500 … 6,000).
3. In the cell above the first profit cell (B4), type `=1_Model!B16`.
4. Select the whole block (A4:B15).
5. **Data → What-If Analysis → Data Table…** → leave *Row input
   cell* empty → *Column input cell* = `1_Model!$B$5` (unit price).
6. Excel fills the profit column with `{=TABLE(,B5)}` array formulas.

**Why this workbook ships formulas instead:** Excel's native Data
Table writes `{=TABLE(...)}` array formulas, and LibreOffice cannot
recalculate them (verified 17 Nov 2026 — they come back `#VALUE!`).
So each profit cell here is the explicit live formula
`=(price − cost) × volume − fixed` — the exact calculation a Data
Table performs, recalculating in both Excel and LibreOffice. The
numbers are identical either way; the walkthrough above is how you
rebuild the native version in Excel.

**Reading the answers:** the summary block uses the functions that
make a table useful instead of just long: `MINIFS` finds the first
profitable price (PKR 4,000), `MAXIFS` the best profit at prices
≥ 4,500, and `AGGREGATE` the largest/smallest profit. These three
are post-2007 functions — in the file they are written with the
`_xlfn.` prefix the format requires (plain `MINIFS` shows `#NAME?`
in LibreOffice; verified 17 Nov 2026).

---

## 3. Two-variable data table — `3_DataTable_2Var`

**What it is:** profit at 6 prices × 5 volumes — 30 cells, colour
scaled green (profit) to red (loss).

**What it answers:** "What if price *and* volume move together?"
The worst cell is −PKR 860,000 (price 3,500, volume 800); the best
is +PKR 3,780,000 (price 6,000, volume 1,600). **22 of 30 cells
make a profit.** The diagonal tells the story: at PKR 3,500 every
volume loses money; at PKR 4,500 every volume makes money. For this
shop, price — not volume — decides whether the month is won or lost.

**When to use it:** two inputs move together and you need the
joint picture — price × volume, price × discount, ad spend ×
conversion rate. One-variable tables hide interactions; this one
shows them.

**Exact Excel steps (native Data Table):**
1. Prices down column A, volumes across row 4, `=1_Model!B16` in
   the corner cell (A4).
2. Select the whole block (A4:F10).
3. **Data → What-If Analysis → Data Table…** → *Row input cell* =
   `1_Model!$B$7` (volume) → *Column input cell* = `1_Model!$B$5`
   (unit price).
4. Same LibreOffice note as §2: the shipped grid is explicit live
   formulas computing exactly what the Data Table computes.

---

## 4. Scenarios — `4_Scenarios`

**What it is:** three named futures, each changing *four* inputs
at once:

| Scenario | Price | Cost | Volume | Fixed | Profit |
|---|---|---|---|---|---|
| Pessimistic | 4,000 | 2,900 | 900 | 1,600,000 | **−610,000** |
| Base | 4,500 | 2,700 | 1,200 | 1,500,000 | **660,000** |
| Optimistic | 5,000 | 2,500 | 1,500 | 1,400,000 | **2,350,000** |

**What it answers:** "What does the whole business look like in a
bad world, a normal world, a good world?" The spread is
PKR 2,960,000 wide — the decision is not "which price?" but
"which world are we planning for?"

**When to use it:** several inputs move together in a story (a
supplier raises costs *and* demand falls — that is one scenario,
not two separate tables). Data tables vary one or two inputs
mechanically; scenarios bundle a judgement about the world.

**Exact Excel steps (Scenario Manager):**
1. **Data → What-If Analysis → Scenario Manager… → Add…**
2. Name it `Pessimistic`, changing cells
   `1_Model!$B$5:$B$8`, values 4000, 2900, 900, 1600000.
3. Repeat for `Base` and `Optimistic`.
4. **Summary** → result cell `1_Model!$B$16` → Excel builds the
   summary table this sheet reproduces as live formulas.

**Honest note:** the Pessimistic scenario deliberately loses money.
A scenario pack that only shows good news teaches nothing — and a
manager who only sees the Optimistic case makes the wrong decision.

---

## 5. Goal Seek — `5_GoalSeek`

**What it is:** working backwards. Instead of "price X gives profit
Y", ask "what price gives profit **PKR 1,000,000**?" Answer:
**PKR 4,783.33**.

**Exact Excel steps:**
1. **Data → What-If Analysis → Goal Seek…**
2. *Set cell* = `1_Model!$B$16` (monthly profit).
3. *To value* = `1000000`.
4. *By changing cell* = `1_Model!$B$5` (unit price).
5. Excel iterates and reports **4,783.33** — click OK to keep it.

**How the workbook verifies it:** the answer cell is a formula,
not a typed number — price = cost + (target + fixed) ÷ volume —
and the check row recomputes profit at that price: exactly
PKR 1,000,000, difference 0. That formula *is* the algebra Goal
Seek solves by trial, written down so the answer stays true if
any input changes. A second example does the same for target
profit 0: the break-even price, **PKR 3,950**.

**When to use it:** you know the outcome you need (a profit
target, a break-even, a loan you must afford) and one input is
free to move. **Limits:** Goal Seek changes exactly ONE input to
hit exactly ONE target. It cannot optimise ("best price"), it
cannot move two cells at once (use Solver for that), and it has no
memory — change an input afterwards and you must re-run it. If
Goal Seek says "may not have found a solution", the target is
unreachable with that input (e.g. no price makes a profit if the
cost is above every price you will try).

---

## 6. Break-even — `6_BreakEven`

**What it is:** the point where profit = 0 — contribution exactly
covers fixed costs.

- Break-even units = fixed ÷ (price − cost) = **833.33 units**
  (834 whole units — you cannot sell a third of an earbud).
- Break-even revenue = **PKR 3,750,000** (PKR 3,753,000 at 834 units).
- Break-even price = **PKR 3,950** (same number Goal Seek finds
  for target profit 0 — the two sheets cross-check each other).
- Margin of safety = **366.67 units, 30.56%** — sales can fall by
  almost a third before the shop starts losing money.

**When to use it:** before any pricing decision — it is the floor
under every other technique. The data table says profit turns
positive at PKR 4,000; break-even says why (the exact line is
PKR 3,950, and 4,000 is the first table price above it).

---

## Which technique, when?

| Question | Technique |
|---|---|
| What is the profit at today's inputs? | The model (§1) |
| What if I change one input? | One-variable data table (§2) |
| What if two inputs move together? | Two-variable data table (§3) |
| What if the whole world changes? | Scenarios (§4) |
| What input do I need to hit a target? | Goal Seek (§5) |
| How much must I sell to not lose money? | Break-even (§6) |

## Limits

- SAMPLE data only — fictional shop, invented numbers.
- The data tables hold volume fixed (one-variable) or vary it
  independently of price (two-variable). Real demand falls as
  price rises; a demand curve would fix that, and is the natural
  next step, not part of this pack.
- Goal Seek's answer is exact only for this linear model; with a
  curved model it finds *a* solution, not necessarily the only one.
- Native `{=TABLE(...)}` Data Tables are documented, not shipped
  as array formulas — LibreOffice cannot recalculate them (see §2).
