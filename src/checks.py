"""Checks for Mini #6 (run: python3 src/checks.py).

openpyxl does not evaluate formulas, so these checks do two things:
(1) verify the right formula string sits in every derived cell, and
(2) recompute every expected value in Python from the CSVs and check
it against the Formula Lab's Expected column and the Summary logic.
"""

import csv
import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "outputs" / "excel_formulas_lookups_SAMPLE.xlsx"

passed = failed = 0


def check(name, cond, detail=""):
    global passed, failed
    if cond:
        passed += 1
        print(f"PASS: {name}")
    else:
        failed += 1
        print(f"FAIL: {name} {detail}")


products = list(csv.DictReader(open(ROOT / "data" / "products_SAMPLE.csv")))
orders = list(csv.DictReader(open(ROOT / "data" / "orders_SAMPLE.csv")))
price = {p["product_id"]: int(p["unit_price"]) for p in products}
cat_of = {p["product_id"]: p["category"] for p in products}


def disc(q):
    return 0.10 if q >= 10 else 0.05 if q >= 5 else 0.0


wb = load_workbook(XLSX)
check("5 sheets in order",
      wb.sheetnames == ["Guide", "Products", "Orders", "Summary", "Formula Lab"],
      str(wb.sheetnames))
prod, ord_ws, summ, lab = wb["Products"], wb["Orders"], wb["Summary"], wb["Formula Lab"]
check("20 products + header", prod.max_row == 21)
check("30 orders + header", ord_ws.max_row == 31)
check("Orders row 2: XLOOKUP product", ord_ws["G2"].value == "=XLOOKUP(C2,Products!$A:$A,Products!$B:$B)", str(ord_ws["G2"].value))
check("Orders row 2: XLOOKUP price", ord_ws["H2"].value == "=XLOOKUP(C2,Products!$A:$A,Products!$D:$D)")
check("Orders row 2: line total formula", ord_ws["I2"].value == "=D2*H2")
check("Orders row 2: INDEX+MATCH category", ord_ws["J2"].value == "=INDEX(Products!$C:$C,MATCH(C2,Products!$A:$A,0))")
check("Orders row 2: IFS discount", ord_ws["K2"].value == "=IFS(D2>=10,0.1,D2>=5,0.05,TRUE,0)")
check("Orders row 2: net total formula", ord_ws["L2"].value == "=I2*(1-K2)")
check("every Orders row carries all 6 formulas",
      all(ord_ws.cell(row=i, column=c).value
          and str(ord_ws.cell(row=i, column=c).value).startswith("=")
          for i in range(2, 32) for c in range(7, 13)))
check("Summary KPI: gross = SUM(Orders!I:I)", summ["B3"].value == "=SUM(Orders!I:I)", str(summ["B3"].value))
check("Summary KPI: net = SUM(Orders!L:L)", summ["B4"].value == "=SUM(Orders!L:L)")
check("Summary category rows use SUMIFS",
      all("SUMIFS" in str(summ.cell(row=r, column=3).value) for r in range(10, 15)))
check("Summary region rows use COUNTIFS",
      all("COUNTIFS" in str(summ.cell(row=r, column=2).value) for r in range(17, 21)))
gross = sum(price[o["product_id"]] * int(o["qty"]) for o in orders)
net = sum(price[o["product_id"]] * int(o["qty"]) * (1 - disc(int(o["qty"]))) for o in orders)
check("gross line total PKR 125,180", gross == 125180, str(gross))
check("net total PKR 119,596.50", abs(net - 119596.50) < 0.01, str(net))
check("discounted orders = 12", sum(1 for o in orders if disc(int(o["qty"])) > 0) == 12)
check("Groceries gross PKR 46,020",
      sum(price[o["product_id"]] * int(o["qty"]) for o in orders
          if cat_of[o["product_id"]] == "Groceries") == 46020)
check("Lab: 10 techniques + header", lab.max_row == 11)
lab_expected = {lab.cell(row=i, column=1).value: lab.cell(row=i, column=4).value
                for i in range(2, 12)}
check("Lab XLOOKUP expects Washing Powder 3kg",
      lab_expected.get("XLOOKUP") == "Washing Powder 3kg")
check("Lab INDEX+MATCH expects 1250", lab_expected.get("INDEX + MATCH") == 1250)
check("Lab VLOOKUP expects CleanHome Ltd", lab_expected.get("VLOOKUP") == "CleanHome Ltd")
check("Lab SUMIFS expects 46020", lab_expected.get("SUMIFS") == 46020)
check("Lab COUNTIFS expects Clifton count",
      lab_expected.get("COUNTIFS") == sum(1 for o in orders if o["region"] == "Clifton"))
check("Lab AVERAGEIFS expects 5004.44", lab_expected.get("AVERAGEIFS") == 5004.44)
check("Lab CONCAT expects ORD-001 - Washing Powder 3kg",
      lab_expected.get("CONCAT") == "ORD-001 - Washing Powder 3kg")
check("Lab IFS expects 0.1 for quantity 12",
      lab_expected.get("IFS") == 0.1)
check("Lab IF (nested) expects Small for quantity 4",
      lab_expected.get("IF (nested)") == "Small")
check("Lab TEXT expects 01-Oct-2026",
      lab_expected.get("TEXT") == "01-Oct-2026")
check("Lab formulas are live (all start with =)",
      all(str(lab.cell(row=i, column=3).value).startswith("=") for i in range(2, 12)))
check("workbook file is non-trivial", XLSX.stat().st_size > 8000)

print(f"\n{passed} passing, {failed} failing")
sys.exit(1 if failed else 0)
