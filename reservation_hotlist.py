import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from datetime import timedelta

# Example: load reservation data
df = pd.read_csv("reservations.csv", sep=";")# columns: Party, StartDate, EndDate

# Expand to daily
records = []
for _, row in df.iterrows():
    start, end = pd.to_datetime(row["StartDate"]), pd.to_datetime(row["EndDate"])
    for d in pd.date_range(start, end):
        records.append({"Party": row["Party"], "Date": d})
days = pd.DataFrame(records)

# Aggregate counts
pivot = days.groupby(["Date","Party"]).size().unstack(fill_value=0)

# Ensure A,B,C columns
for p in ["A","B","C"]:
    if p not in pivot.columns: pivot[p]=0

pivot["Total"] = pivot.sum(axis=1)
pivot["Week"]  = pivot.index.isocalendar().week
pivot["Day"]   = pivot.index.day_name()

# Make output table: weeks (rows), days (cols)
weeks = range(18,41)
days_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]

# Start Excel
wb = openpyxl.Workbook()
ws = wb.active

# Write headers
ws.append(["Week"] + days_order)

# Colors for A,B,C
colors = {"A":"9933FF","B":"996633","C":"FF6600"}

for w in weeks:
    row = [w]
    for d in days_order:
        sel = pivot[(pivot["Week"]==w) & (pivot["Day"]==d)]
        if sel.empty:
            row.append("")
            continue
        A,B,C,Total = sel[["A","B","C","Total"]].sum()
        cell_str = f"A:{A} B:{B} C:{C}"  # for debug
        row.append(cell_str)
    ws.append(row)

# TODO: apply fills proportional to A,B,C in each cell
# You’d need to draw stacked shapes or gradient fills.
# Simplest: just write text numbers for now.
# More advanced: use `openpyxl.drawing` to insert colored rectangles.

wb.save("reservation_table.xlsx")
