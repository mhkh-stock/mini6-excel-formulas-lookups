"""Build the Mini #6 workbook: every derived cell is a live formula.

Sheets: Guide · Products (20 SAMPLE products) · Orders (30 SAMPLE
orders, columns G–L are formulas) · Summary (COUNTIFS/SUMIFS tables
+ KPI formulas) · Formula Lab (10 techniques, each a live formula
with its expected result beside it, recomputed here from the CSVs).

openpyxl writes formulas, it does not evaluate them — so checks.py
recomputes every expected value in Python and verifies the formula
strings, rather than trusting cached values that do not exist.
"""

import csv
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "outputs" / "excel_formulas_lookups_SAMPLE.xlsx"
OUT.parent.mkdir(parents=True, exist_ok=True)

products = list(csv.DictReader(open(ROOT / "data" / "products_SAMPLE.csv")))
orders = list(csv.DictReader(open(ROOT / "data" / "orders_SAMPLE.csv")))
price = {p["product_id"]: int(p["unit_price"]) for p in products}
cat_of = {p["product_id"]: p["category"] for p in products}
name_of = {p["product_id"]: p["product"] for p in products}


def disc(q):
    return 0.10 if q >= 10 else 0.05 if q >= 5 else 0.0


wb = Workbook()
guide = wb.active
guide.title = "Guide"
guide["A1"] = "Excel Advanced Formulas & Lookup Pack — SAMPLE"
guide["A1"].font = Font(bold=True, size=14)
guide_rows = [
    ("Sheet", "What it teaches"),
    ("Products", "The lookup table: 20 SAMPLE products with price, category, supplier."),
    ("Orders", "30 SAMPLE Oct 2026 orders. Columns G–L are LIVE formulas — change a ProductID or Qty and the row recalculates."),
    ("Summary", "COUNTIFS / SUMIFS / AVERAGEIFS tables by category and region, plus KPI formulas over the Orders sheet."),
    ("Formula Lab", "10 techniques, each as a live formula with its expected result beside it. XLOOKUP is shown first; INDEX+MATCH and VLOOKUP are shown as the compatible alternatives."),
    ("Note", "XLOOKUP needs Microsoft 365 / Excel 2021+. INDEX+MATCH and VLOOKUP work in every version — that is why all three are taught."),
]
for i, (a, b) in enumerate(guide_rows, start=3):
    guide.cell(row=i, column=1, value=a).font = Font(bold=True)
    guide.cell(row=i, column=2, value=b)
guide.column_dimensions["A"].width = 16
guide.column_dimensions["B"].width = 110

prod = wb.create_sheet("Products")
prod.append(["ProductID", "Product", "Category", "Unit Price (PKR)", "Supplier"])
for p in products:
    prod.append([p["product_id"], p["product"], p["category"],
                 int(p["unit_price"]), p["supplier"]])
for col, w in zip("ABCDE", (12, 30, 16, 18, 18)):
    prod.column_dimensions[col].width = w

ord_ws = wb.create_sheet("Orders")
ord_ws.append(["OrderID", "Date", "ProductID", "Qty", "Customer", "Region",
               "Product (XLOOKUP)", "Unit Price (XLOOKUP)", "Line Total",
               "Category (INDEX+MATCH)", "Discount % (IFS)", "Net Total"])
for i, o in enumerate(orders, start=2):
    ord_ws.cell(row=i, column=1, value=o["order_id"])
    y, m, d = (int(x) for x in o["date"].split("-"))
    date_cell = ord_ws.cell(row=i, column=2, value=date(y, m, d))
    date_cell.number_format = "yyyy-mm-dd"
    ord_ws.cell(row=i, column=3, value=o["product_id"])
    ord_ws.cell(row=i, column=4, value=int(o["qty"]))
    ord_ws.cell(row=i, column=5, value=o["customer"])
    ord_ws.cell(row=i, column=6, value=o["region"])
    ord_ws.cell(row=i, column=7, value=f'=XLOOKUP(C{i},Products!$A:$A,Products!$B:$B)')
    ord_ws.cell(row=i, column=8, value=f'=XLOOKUP(C{i},Products!$A:$A,Products!$D:$D)')
    ord_ws.cell(row=i, column=9, value=f'=D{i}*H{i}')
    ord_ws.cell(row=i, column=10, value=f'=INDEX(Products!$C:$C,MATCH(C{i},Products!$A:$A,0))')
    ord_ws.cell(row=i, column=11, value=f'=IFS(D{i}>=10,0.1,D{i}>=5,0.05,TRUE,0)')
    ord_ws.cell(row=i, column=12, value=f'=I{i}*(1-K{i})')
for col, w in zip("ABCDEFGHIJKL", (10, 12, 11, 6, 16, 16, 28, 18, 12, 18, 14, 12)):
    ord_ws.column_dimensions[col].width = w

summ = wb.create_sheet("Summary")
summ["A1"] = "KPIs (all live formulas over Orders)"
summ["A1"].font = Font(bold=True)
kpis = [
    ("Total orders", "=COUNTA(Orders!A:A)-1"),
    ("Gross line total (PKR)", "=SUM(Orders!I:I)"),
    ("Net total after discounts (PKR)", "=SUM(Orders!L:L)"),
    ("Average net order value (PKR)", "=AVERAGE(Orders!L:L)"),
    ("Largest single line (PKR)", "=MAX(Orders!I:I)"),
    ("Orders with a discount", '=COUNTIF(Orders!K:K,">0")'),
]
for i, (label, formula) in enumerate(kpis, start=2):
    summ.cell(row=i, column=1, value=label)
    summ.cell(row=i, column=2, value=formula)
r = 9
summ.cell(row=r, column=1, value="By category").font = Font(bold=True)
summ.cell(row=r, column=2, value="Orders")
summ.cell(row=r, column=3, value="Gross (PKR)")
summ.cell(row=r, column=4, value="Net (PKR)")
for j, c in enumerate(["Groceries", "Household", "Electronics", "Stationery", "Personal Care"]):
    row = r + 1 + j
    summ.cell(row=row, column=1, value=c)
    summ.cell(row=row, column=2, value=f'=COUNTIFS(Orders!J:J,A{row})')
    summ.cell(row=row, column=3, value=f'=SUMIFS(Orders!I:I,Orders!J:J,A{row})')
    summ.cell(row=row, column=4, value=f'=SUMIFS(Orders!L:L,Orders!J:J,A{row})')
r2 = r + 7
summ.cell(row=r2, column=1, value="By region").font = Font(bold=True)
summ.cell(row=r2, column=2, value="Orders")
summ.cell(row=r2, column=3, value="Gross (PKR)")
for j, reg in enumerate(["Gulshan", "Saddar", "Clifton", "North Nazimabad"]):
    row = r2 + 1 + j
    summ.cell(row=row, column=1, value=reg)
    summ.cell(row=row, column=2, value=f'=COUNTIFS(Orders!F:F,A{row})')
    summ.cell(row=row, column=3, value=f'=SUMIFS(Orders!I:I,Orders!F:F,A{row})')
summ.column_dimensions["A"].width = 34
summ.column_dimensions["B"].width = 14
summ.column_dimensions["C"].width = 16
summ.column_dimensions["D"].width = 16

lab = wb.create_sheet("Formula Lab")
lab.append(["Technique", "What it does", "Live example (formula)", "Expected result"])
gross_groc = sum(price[o["product_id"]] * int(o["qty"]) for o in orders
               if cat_of[o["product_id"]] == "Groceries")
clifton_n = sum(1 for o in orders if o["region"] == "Clifton")
avg_groc_net = sum(price[o["product_id"]] * int(o["qty"]) * (1 - disc(int(o["qty"])))
                     for o in orders if cat_of[o["product_id"]] == "Groceries") / \
    sum(1 for o in orders if cat_of[o["product_id"]] == "Groceries")
lab_rows = [
    ("XLOOKUP", "Find a product name from its ID",
     '=XLOOKUP("PR-04",Products!$A:$A,Products!$B:$B)', "Washing Powder 3kg"),
    ("INDEX + MATCH", "Find a price from its ID (works in old Excel too)",
     '=INDEX(Products!$D:$D,MATCH("PR-04",Products!$A:$A,0))', 1250),
    ("VLOOKUP", "Find a supplier from its ID (exact match)",
     '=VLOOKUP("PR-04",Products!$A:$E,5,FALSE)', "CleanHome Ltd"),
    ("SUMIFS", "Total gross sales for one category only",
     '=SUMIFS(Orders!$I:$I,Orders!$J:$J,"Groceries")', gross_groc),
    ("COUNTIFS", "Count orders from one region only",
     '=COUNTIFS(Orders!$F:$F,"Clifton")', clifton_n),
    ("IFS", "Discount rate for a bulk quantity of 12",
     "=IFS(12>=10,0.1,12>=5,0.05,TRUE,0)", 0.1),
    ("IF (nested)", "Size label for the first order's quantity (4)",
     '=IF(Orders!D2>=10,"Bulk",IF(Orders!D2>=5,"Medium","Small"))', "Small"),
    ("AVERAGEIFS", "Average net value of Groceries orders",
     '=AVERAGEIFS(Orders!$L:$L,Orders!$J:$J,"Groceries")', round(avg_groc_net, 2)),
    ("TEXT", "Format a date as text",
     '=TEXT(DATE(2026,10,1),"dd-mmm-yyyy")', "01-Oct-2026"),
    ("CONCAT", "Join order ID and product name",
     '=CONCAT(Orders!A2," - ",Orders!G2)', "ORD-001 - Washing Powder 3kg"),
]
for row in lab_rows:
    lab.append(list(row))
lab.column_dimensions["A"].width = 16
lab.column_dimensions["B"].width = 44
lab.column_dimensions["C"].width = 62
lab.column_dimensions["D"].width = 34

wb.save(OUT)
print(f"Wrote {OUT} ({OUT.stat().st_size:,} bytes)")
print(f"Gross {sum(price[o['product_id']] * int(o['qty']) for o in orders):,} | "
      f"Groceries gross {gross_groc:,} | Clifton orders {clifton_n} | "
      f"Avg Groceries net {round(avg_groc_net, 2)}")
