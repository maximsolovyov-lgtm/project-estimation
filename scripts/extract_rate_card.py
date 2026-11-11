#!/usr/bin/env python3
"""Extract the Rate Card (Fix-Price / A-la-Carte tiers) and the holidays list
from Project_Estimation_Template.xlsx into worker/rateCard.json.

Usage: python3 -I scripts/extract_rate_card.py path/to/Project_Estimation_Template.xlsx

The JSON is bundled into the Worker only (never into the static frontend), so
the full rate card is never shipped to the browser.
"""
import json, sys, os, datetime
import openpyxl

def base_role(name, grade):
    # Mirrors Lists!A: strip a trailing " <Grade>" unless grade is "Architect"
    n = (name or "").strip(); g = (grade or "").strip()
    if not n:
        return ""
    if g and g != "Architect" and len(n) > len(g) + 1 and n.endswith(" " + g):
        return n[: len(n) - len(g) - 1]
    return n

def num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("$", "").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None

def main(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    rc = wb["Rate Card"]
    tiers = [rc.cell(3, c).value for c in range(14, 19)]
    rows = []
    for r in range(4, rc.max_row + 1):
        name = rc.cell(r, 1).value
        if not name or not str(name).strip():
            continue
        grade = str(rc.cell(r, 3).value or "").strip()
        rows.append({
            "name": str(name).strip(),
            "role": base_role(str(name), grade),
            "code": rc.cell(r, 2).value,
            "grade": grade,
            "rates": [num(rc.cell(r, c).value) for c in range(14, 19)],
        })
    hol = wb["Holidays"]
    holidays = []
    for r in range(1, hol.max_row + 1):
        v = hol.cell(r, 1).value
        if isinstance(v, (datetime.date, datetime.datetime)):
            holidays.append(v.strftime("%Y-%m-%d"))
    out = {"title": rc["A1"].value, "tiers": tiers, "rows": rows, "defaultHolidays": holidays}
    dst = os.path.join(os.path.dirname(__file__), "..", "worker", "rateCard.json")
    with open(dst, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f"wrote {len(rows)} rate rows, {len(holidays)} holidays -> {os.path.normpath(dst)}")

if __name__ == "__main__":
    main(sys.argv[1])
