"""Generate SAMPLE shop data for Mini #6 (seed 42, deterministic).

Fictional shop: DailyNeeds Store, Karachi. 20 products, 30 orders
(Oct 2026). The workbook's formulas look up and summarise THIS data,
so re-running this script and the builder reproduces the identical
workbook — and checks.py recomputes every expected total from here.
"""

import csv
import random
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
random.seed(42)

PRODUCTS = [
    ("PR-01", "Basmati Rice 5kg", "Groceries", 1850, "Karachi Foods"),
    ("PR-02", "Cooking Oil 5L", "Groceries", 2450, "Karachi Foods"),
    ("PR-03", "Tea Leaves 500g", "Groceries", 890, "Lahore Tea Co"),
    ("PR-04", "Washing Powder 3kg", "Household", 1250, "CleanHome Ltd"),
    ("PR-05", "Bath Soap (Pack of 4)", "Personal Care", 640, "CleanHome Ltd"),
    ("PR-06", "Shampoo 400ml", "Personal Care", 780, "GlowCare"),
    ("PR-07", "Toothpaste 150g", "Personal Care", 420, "GlowCare"),
    ("PR-08", "LED Bulb 12W", "Electronics", 550, "BrightElectric"),
    ("PR-09", "USB Cable 1m", "Electronics", 350, "BrightElectric"),
    ("PR-10", "Notebook A4 (Pack of 3)", "Stationery", 480, "PaperWorks"),
    ("PR-11", "Ball Pen (Box of 10)", "Stationery", 320, "PaperWorks"),
    ("PR-12", "Sugar 5kg", "Groceries", 1150, "Karachi Foods"),
    ("PR-13", "Floor Cleaner 2L", "Household", 690, "CleanHome Ltd"),
    ("PR-14", "Face Wash 100ml", "Personal Care", 920, "GlowCare"),
    ("PR-15", "Extension Board 4-way", "Electronics", 1450, "BrightElectric"),
    ("PR-16", "Colour Pencils (Set of 12)", "Stationery", 560, "PaperWorks"),
    ("PR-17", "Atta Flour 10kg", "Groceries", 1980, "Karachi Foods"),
    ("PR-18", "Dishwash Liquid 1L", "Household", 540, "CleanHome Ltd"),
    ("PR-19", "Hair Oil 200ml", "Personal Care", 610, "GlowCare"),
    ("PR-20", "Geometry Box", "Stationery", 380, "PaperWorks"),
]

CUSTOMERS = ["Ahmed Raza", "Fatima Khan", "Bilal Hussain", "Ayesha Malik",
             "Usman Tariq", "Zainab Ali", "Hassan Nawaz", "Maryam Siddiqui"]
REGIONS = ["Gulshan", "Saddar", "Clifton", "North Nazimabad"]

orders = []
for i in range(1, 31):
    pid, _, _, _, _ = random.choice(PRODUCTS)
    orders.append({
        "order_id": f"ORD-{i:03d}",
        "date": f"2026-10-{random.randint(1, 28):02d}",
        "product_id": pid,
        "qty": random.choice([1, 2, 2, 3, 4, 5, 6, 8, 10, 12]),
        "customer": random.choice(CUSTOMERS),
        "region": random.choice(REGIONS),
    })

(ROOT / "data").mkdir(parents=True, exist_ok=True)
with open(ROOT / "data" / "products_SAMPLE.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["product_id", "product", "category", "unit_price", "supplier"])
    w.writerows(PRODUCTS)
with open(ROOT / "data" / "orders_SAMPLE.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=list(orders[0].keys()))
    w.writeheader()
    w.writerows(orders)

price = {p[0]: p[3] for p in PRODUCTS}
total = sum(price[o["product_id"]] * o["qty"] for o in orders)
print(f"20 products, 30 orders, gross line total PKR {total:,}")
