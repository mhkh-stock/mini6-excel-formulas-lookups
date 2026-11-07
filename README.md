# Mini #6 — Excel Advanced Formulas & Lookup Pack

**What it teaches:** the ten formulas that turn an Excel user from
a data-typist into a data analyst — lookups that fill themselves in,
conditional sums that answer questions without pivot tables, and
tiered logic that replaces hand-applied discounts.
**Status: SAMPLE data, real formulas**
**Tools:** Excel (Microsoft 365 for XLOOKUP; INDEX+MATCH and
VLOOKUP work in every version) · Python (openpyxl builder)
**Deliverable:** `outputs/excel_formulas_lookups_SAMPLE.xlsx`
(Guide + Products + Orders + Summary + Formula Lab)

> ⚠️ **What this is and is not.** "DailyNeeds Store" is fictional
> and the 20 products / 30 orders are generated SAMPLE data
> (`src/generate_sample_data.py`, seed 42). What is real is every
> formula: Orders columns G–L, every Summary KPI and table cell,
> and every Formula Lab example are live formulas, not typed values.
> Change a ProductID or a quantity and the workbook recalculates.

## The 10 techniques (Formula Lab sheet)

| # | Technique | Live example in the pack | Expected result |
|---|---|---|---|
| 1 | XLOOKUP | Product name from ID `PR-04` | Washing Powder 3kg |
| 2 | INDEX + MATCH | Price from ID `PR-04` | PKR 1,250 |
| 3 | VLOOKUP | Supplier from ID `PR-04` (exact match) | CleanHome Ltd |
| 4 | SUMIFS | Gross sales, Groceries only | PKR 46,020 |
| 5 | COUNTIFS | Orders from Clifton only | 10 |
| 6 | IFS | Discount rate for quantity 12 | 10% |
| 7 | IF (nested) | Size label for quantity 4 | Small |
| 8 | AVERAGEIFS | Average net value, Groceries orders | PKR 5,004.44 |
| 9 | TEXT | Date formatted as text | 01-Oct-2026 |
| 10 | CONCAT | Order ID + product name joined | ORD-001 - Washing Powder 3kg |

## What is inside

| File | What it does |
|---|---|
| `outputs/excel_formulas_lookups_SAMPLE.xlsx` | The deliverable workbook (5 sheets) |
| `data/products_SAMPLE.csv` | The 20 SAMPLE products the lookups read |
| `data/orders_SAMPLE.csv` | The 30 SAMPLE Oct 2026 orders |
| `src/generate_sample_data.py` | Generates the sample CSVs (deterministic, seed 42) |
| `src/build_workbook.py` | Builds the workbook with live formulas throughout |
| `src/checks.py` | Checks as code — 31 checks, all passing |
| `requirements.txt` | Pinned Python dependency (`openpyxl==3.1.2`) |

## Verified numbers

| KPI | Value |
|---|---|
| Products / orders | 20 / 30 (Oct 2026, SAMPLE) |
| Gross line total | PKR 125,180 |
| Net total after IFS discounts | PKR 119,596.50 |
| Orders with a discount | 12 (5 at 10%, 7 at 5%) |
| Top category by gross | Groceries, PKR 46,020 |
| Top region by gross | Clifton, PKR 40,100 |
| Checks | 31 passing (`python3 src/checks.py`) |

Verification: `python3 src/checks.py` (31 checks pass) plus an
independent reopen of the workbook — 5 sheets in order, all 30
Orders rows carrying all 6 formula columns, Summary KPIs and both
tables as formulas, all 10 Lab examples live with expected results
matching Python-recomputed values.

## How to run it

1. Install the one dependency: `pip install -r requirements.txt`.
2. Open `outputs/excel_formulas_lookups_SAMPLE.xlsx` — that is the
   deliverable.
3. To regenerate the sample data: `python3 src/generate_sample_data.py`.
4. To rebuild the workbook: `python3 src/build_workbook.py`.
5. To check it: `python3 src/checks.py`.
6. To practise on your own data: replace the two CSVs (same columns)
   and re-run steps 4–5 only — do NOT re-run step 3, it regenerates
   the sample data and would overwrite your files.

## Data dictionary (SAMPLE CSVs)

| File | Column | Meaning |
|---|---|---|
| `products_SAMPLE.csv` | `product_id` | Lookup key, e.g. `PR-04` |
|  | `product` | Product name |
|  | `category` | Groceries / Household / Electronics / Stationery / Personal Care |
|  | `unit_price` | Price in PKR (whole rupees) |
|  | `supplier` | Supplier name |
| `orders_SAMPLE.csv` | `order_id` | Order key, e.g. `ORD-001` |
|  | `date` | Order date, Oct 2026 (`yyyy-mm-dd`) |
|  | `product_id` | Must match a `product_id` in the products file |
|  | `qty` | Units ordered; discount tiers: ≥10 → 10%, 5–9 → 5%, else 0% |
|  | `customer` | Customer name |
|  | `region` | Gulshan / Saddar / Clifton / North Nazimabad |

## Limits

- **SAMPLE data only.** "DailyNeeds Store", the 20 products and
  30 orders are fictional; the formulas are real, the numbers are
  not a real shop's.
- **XLOOKUP needs Microsoft 365 / Excel 2021+.** INDEX+MATCH and
  VLOOKUP work in every version — that is why all three are taught.
- **20 products / 30 orders, Oct 2026.** The Summary tables list
  exactly the 5 categories and 4 regions in the sample data; new
  categories or regions need a row added to the Summary table.
- **No cached values.** openpyxl writes formulas without results,
  so a viewer that does not calculate may show 0 until Excel or
  Sheets recalculates on open (see the note below).

## A note on values (read before opening in a viewer)

openpyxl writes formulas without cached results, so a viewer that
does not calculate (some web previews) may show 0 until Excel or
Sheets recalculates on open. The formulas themselves are verified by
`src/checks.py`, and every expected result is listed in the Formula
Lab sheet.
