"""
Step 9 - Check the two dimensions the model generates in M.

Dim_Date and Dim_Hour are authored in Power Query rather than loaded from a
file, so they depend on no CSV at all. That independence costs a second
implementation of the same logic, which could quietly drift from 03_clean.py.

This re-implements that M with M's own semantics - ISO week taken from the
Thursday of the Monday-week, day parts on 03_clean.py's cut points - and
compares it column-by-column against the CSVs. It catches drift without paying
for a five-minute Power BI refresh.

Run:  python 09_check_generated_dims.py
"""
import datetime as dt
import pathlib
import sys

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
CLEAN = ROOT / "02_clean"

# ------------------------------------------------------------------ Dim_Date
START, END = dt.date(2009, 12, 1), dt.date(2011, 12, 9)
days = (END - START).days + 1
rows = []
for i in range(days):
    d = START + dt.timedelta(days=i)
    quarter = (d.month - 1) // 3 + 1
    weekday_no = d.weekday()                      # Date.DayOfWeek(_, Day.Monday)
    # ISO week, exactly as the M computes it
    monday = d - dt.timedelta(days=weekday_no)    # Date.StartOfWeek(_, Day.Monday)
    thu = monday + dt.timedelta(days=3)
    iso_week = (thu - dt.date(thu.year, 1, 1)).days // 7 + 1
    rows.append({
        "DateKey": pd.Timestamp(d),
        "Year": d.year,
        "Quarter": quarter,
        "QuarterLabel": f"Q{quarter} {d.year}",
        "Month": d.month,
        "MonthName": d.strftime("%b"),            # Date.ToText(_, "MMM")
        "MonthYear": d.strftime("%Y-%m"),
        "MonthYearLabel": d.strftime("%b %Y"),
        "Weekday": d.strftime("%A"),              # Date.DayOfWeekName
        "WeekdayNo": weekday_no,
        "IsWeekend": weekday_no >= 5,
        "WeekOfYear": iso_week,
        "DayOfYear": d.timetuple().tm_yday,
    })
gen_date = pd.DataFrame(rows)

# ------------------------------------------------------------------ Dim_Hour
gen_hour = pd.DataFrame({"Hour": range(24)})
gen_hour["HourLabel"] = gen_hour.Hour.map(lambda h: f"{h:02d}:00")
gen_hour["DayPart"] = gen_hour.Hour.map(
    lambda h: "Morning" if h <= 11 else "Afternoon" if h <= 16
    else "Evening" if h <= 20 else "Night")
gen_hour["DayPartNo"] = gen_hour.Hour.map(
    lambda h: 0 if h <= 11 else 1 if h <= 16 else 2 if h <= 20 else 3)

# ------------------------------------------------------------------- compare
failures = 0
for name, gen, parse in (("Dim_Date", gen_date, ["DateKey"]),
                         ("Dim_Hour", gen_hour, None)):
    csv = pd.read_csv(CLEAN / f"{name}.csv", parse_dates=parse,
                      encoding="utf-8-sig")
    print(f"\n=== {name}: M-generated {len(gen)} rows vs CSV {len(csv)} rows")
    if len(gen) != len(csv):
        print("  [FAIL] row count differs")
        failures += 1
        continue
    for col in csv.columns:
        a, b = gen[col], csv[col]
        if a.dtype == bool or b.dtype == bool:
            a, b = a.astype(str).str.lower(), b.astype(str).str.lower()
        same = (a.reset_index(drop=True) == b.reset_index(drop=True))
        if same.all():
            print(f"  [PASS] {col}")
        else:
            bad = (~same).sum()
            print(f"  [FAIL] {col} -- {bad} of {len(csv)} differ, "
                  f"first: M={a[~same].iloc[0]!r} CSV={b[~same].iloc[0]!r}")
            failures += 1

if failures:
    print(f"\n{failures} MISMATCHES - the M and 03_clean.py have drifted.")
    sys.exit(1)
print("\nALL COLUMNS MATCH - the generated M reproduces 03_clean.py exactly.")
