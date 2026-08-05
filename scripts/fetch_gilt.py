import io
import json
import os
import zipfile
from datetime import date

import requests
from openpyxl import load_workbook

URL = "https://www.bankofengland.co.uk/-/media/boe/files/statistics/yield-curves/latest-yield-curve-data.zip"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}


def main():
    resp = requests.get(URL, headers=HEADERS, timeout=30)
    resp.raise_for_status()

    with zipfile.ZipFile(io.BytesIO(resp.content)) as z:
        xlsx_name = next(
            n for n in z.namelist()
            if n.lower().endswith(".xlsx") and "nominal" in n.lower()
        )
        with z.open(xlsx_name) as f:
            wb = load_workbook(io.BytesIO(f.read()), data_only=True)

    sheet = wb["4. spot curve"]

    # Find the 2-year column by scanning the maturity header row
    col_2y = None
    for c in range(2, sheet.max_column + 1):
        if sheet.cell(row=4, column=c).value == 2:
            col_2y = c
            break

    if col_2y is None:
        raise RuntimeError("Could not find 2-year maturity column")

   # Walk backwards from the last row to find the most recent complete entry
out = None
for r in range(sheet.max_row, 1, -1):          # go all the way down
    d = sheet.cell(row=r, column=1).value
    v = sheet.cell(row=r, column=col_2y).value

    # Skip header / empty rows
    if d is None or v is None:
        continue

    # Skip the header text rows (just in case)
    if isinstance(d, str) and "maturity" in d.lower():
        continue

    # Found a valid data row
    if hasattr(d, "strftime"):
        date_str = d.strftime("%Y-%m-%d")
    else:
        date_str = str(d)[:10]

    out = {"date": date_str, "yield_2y": float(v)}
    break

    if out is None:
        raise RuntimeError("Could not find a valid data row")

    os.makedirs("data", exist_ok=True)
    with open("data/gilt-2y.json", "w") as f:
        json.dump(out, f)

    print("Wrote:", out)


if __name__ == "__main__":
    main()
