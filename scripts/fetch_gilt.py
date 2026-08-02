import requests, zipfile, io, json
from openpyxl import load_workbook
from datetime import date

URL = "https://www.bankofengland.co.uk/-/media/boe/files/statistics/yield-curves/latest-yield-curve-data.zip"

resp = requests.get(URL, timeout=30)
resp.raise_for_status()

with zipfile.ZipFile(io.BytesIO(resp.content)) as z:
    xlsx_name = next(n for n in z.namelist() if 'nominal' in n.lower() and n.endswith('.xlsx'))
    with z.open(xlsx_name) as f:
        wb = load_workbook(io.BytesIO(f.read()), data_only=True)

sheet = wb['4. spot curve']

# Find the 2-year column by scanning row 4
col_2y = None
for c in range(2, sheet.max_column + 1):
    if sheet.cell(row=4, column=c).value == 2:
        col_2y = c
        break

# Walk backwards from the last row to find the most recent non-empty date
for r in range(sheet.max_row, 6, -1):
    d = sheet.cell(row=r, column=1).value
    v = sheet.cell(row=r, column=col_2y).value
    if d is not None and v is not None:
        out = {"date": d.strftime("%Y-%m-%d"), "yield_2y": v}
        break

with open("data/gilt-2y.json", "w") as f:
    json.dump(out, f)
