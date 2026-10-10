"""Regenerate examples/quarterly_budget.xlsx — a realistic inherited workbook.

It deliberately contains the four bugs xl-audit exists to catch (see README).
Run from the repo root:  py scripts/make_example.py
"""
from pathlib import Path

from openpyxl import Workbook

ROWS = [
    # month, category, item, amount
    ("Jan", "Software", "IDE licence", 120.00),
    ("Jan", "Software", "Cloud storage", 29.99),
    ("Jan", "Hardware", "Keyboard", 49.50),
    ("Jan", "Hardware", "Monitor", 189.00),
    ("Feb", "Software", "IDE licence", 120.00),
    ("Feb", "Software", "Vector editor", 15.00),
    ("Feb", "Hardware", "USB hub", 24.99),
    ("Feb", "Hardware", "Webcam", 59.99),
    ("Mar", "Software", "IDE licence", 120.00),
    ("Mar", "Software", "Cloud storage", 29.99),
    ("Mar", "Hardware", "Desk lamp", 34.50),
    ("Mar", "Hardware", "Office chair", 149.00),
]

wb = Workbook()
ws = wb.active
ws.title = "Budget"
ws.append(["Month", "Category", "Item", "Amount", "With tax", "Pct of total"])

TOTAL_ROW = 15
for i, (month, cat, item, amount) in enumerate(ROWS, start=2):
    ws.append([month, cat, item, amount])
    r = ws.max_row
    ws.cell(row=r, column=5, value=f"=D{r}*1.18")          # hardcoded constant
    if r == 8:                                             # one copy-paste slip:
        ws.cell(row=r, column=6, value=f"=D{r}*0.05")      # wrong formula in column F
    else:
        ws.cell(row=r, column=6, value=f"=D{r}/$D${TOTAL_ROW}")

ws.cell(row=14, column=1, value="TOTAL")
ws.cell(row=14, column=4, value="=SUM(D2:D12)")    # data really runs to row 13:
                                                   # row 13 (office chair) is excluded
ws.cell(row=2, column=7, value="=#REF!*2")         # pointed at a deleted sheet

notes = wb.create_sheet("Notes")
notes["A1"] = "Owner: previous analyst, no longer at the company."
notes["A2"] = "Edited by at least four people since 2024. Trust nothing."

out = Path(__file__).resolve().parent.parent / "examples" / "quarterly_budget.xlsx"
out.parent.mkdir(exist_ok=True)
wb.save(out)
print(f"wrote {out}")
