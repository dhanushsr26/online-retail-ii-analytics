"""
Step 10 - Build the preprocessing walkthrough notebook.

The dashboard needs a star schema, so 03_clean.py writes eleven CSVs. That is
the right shape for Power BI and the wrong shape to hand a person: nobody wants
to be shown eleven files and told they are one dataset.

This generates a single notebook that walks through the same eighteen steps one
cell at a time, each with its explanation above it and its result printed below
it, and finishes by writing ONE Excel workbook holding the cleaned data.

The logic is copied from 03_clean.py deliberately - same order, same constants,
same decisions - so the notebook, the PDFs and the dashboard cannot disagree.
Where 03_clean.py is terse, this is spelled out longhand, because the notebook
is read by people rather than run by a pipeline.

Run:  python 10_build_notebook.py        (writes the .ipynb)
      python 10_build_notebook.py --run  (writes it, then executes it)
"""
import pathlib
import subprocess
import sys

import nbformat as nbf

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "06_deliverable"
NB_PATH = OUT_DIR / "Preprocessing_Walkthrough.ipynb"

cells = []


def md(text: str) -> None:
    cells.append(nbf.v4.new_markdown_cell(text.strip("\n")))


def code(text: str) -> None:
    cells.append(nbf.v4.new_code_cell(text.strip("\n")))


# ============================================================ title
md(r"""
# Online Retail II — Preprocessing Walkthrough

**M.Sc Business Analytics · Semester 3**

This notebook takes the **raw** `online_retail_II.xlsx` workbook and applies every
preprocessing step, one step per cell, finishing with a **single cleaned Excel
file**.

### How to read it

Each step below has three parts:

| | |
|---|---|
| **A heading** | which technique from *BA-6 Preprocessing Techniques.pdf* it applies |
| **A short explanation** | why the step is needed and what decision was made |
| **A code cell** | the code, with its result printed directly underneath |

Nothing is hidden in an imported module — every transformation is visible on
this page.

### The dataset

[UCI Online Retail II](https://archive.ics.uci.edu/dataset/502/online+retail+ii) —
transactions from a UK online gift retailer, **1 December 2009 to 9 December 2011**,
spread across two Excel sheets.

### Where the eleven CSV files come from

The Power BI dashboard reads a *star schema* — one fact table and several
dimension tables — because that is the shape Power BI models efficiently. Those
files are a **presentation** of the cleaned data, not a different dataset. The
cleaned data itself is one table, and this notebook writes it as one file at the
end.
""")

# ============================================================ setup
md(r"""
---

## Step 0 · Setting up

Nothing is cleaned yet. This cell only loads the libraries and points at the raw
workbook.

`pandas` is the table library — think of a `DataFrame` as a spreadsheet held in
memory.
""")

code(r'''
import pathlib
import warnings

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# Show wide tables without pandas hiding the middle columns.
pd.set_option("display.max_columns", 60)
pd.set_option("display.width", 200)

# The notebook lives in 06_deliverable, so the project root is one level up.
ROOT = pathlib.Path.cwd().parent if pathlib.Path.cwd().name == "06_deliverable" \
    else pathlib.Path.cwd()

RAW_XLSX = ROOT / "00_source" / "online_retail_II.xlsx"
OUT_XLSX = ROOT / "06_deliverable" / "OnlineRetailII_Cleaned.xlsx"

# Constants used later. They are named here so every number in this notebook
# has one obvious place it came from.
Z_THRESHOLD = 3.0      # |z| beyond this is *flagged* as an outlier, never deleted
IQR_MULTIPLIER = 3.0   # and it must also fall beyond 3 x IQR
MA_WINDOW = 7          # moving-average window, in days
ALPHA = 0.3            # exponential smoothing constant
LIFESPAN_YEARS = 3     # assumed customer lifespan for the CLV formula

print("pandas version :", pd.__version__)
print("raw workbook   :", RAW_XLSX)
print("exists         :", RAW_XLSX.exists())
''')

# ------------------------------------------------------------ audit helper
md(r"""
### The audit trail

Every step records how many rows it removed and why. That record is what turns
"the data is clean" into something an examiner can check, and it is written into
the final Excel file as its own sheet.

One distinction matters. Only steps marked **Row filter** actually delete rows.
Steps marked *Derivation*, *Schema* or *Aggregation* add columns or reshape the
table — counting those as deletions would make it look like a million rows were
thrown away to build a 739-row summary. So the row counts are left blank for
them, and their effect is described in words instead.
""")

code(r'''
audit_rows = []


def record(technique, pdf_section, rows_before, rows_after,
           rationale, step_type="Row filter", note=""):
    """Save one line of evidence about a step, and print it so we can see it."""

    if step_type == "Row filter":
        removed = rows_before - rows_after
        before, after = rows_before, rows_after
    else:
        # Not a deletion - leave the counts blank so it cannot be misread.
        removed = 0
        before = after = None

    audit_rows.append({
        "Step": len(audit_rows),
        "Technique": technique,
        "PDF_Section": pdf_section,
        "Step_Type": step_type,
        "Rows_Before": before,
        "Rows_After": after,
        "Rows_Removed": removed,
        "Pct_Removed": round(removed / before * 100, 4) if before else None,
        "Note": note,
        "Rationale": rationale,
    })

    detail = note if note else f"{rows_before:,} -> {rows_after:,}  ({-removed:+,})"
    print(f"[{len(audit_rows) - 1:>2}] {technique:<44} {detail}")


print("Audit trail ready.")
''')

# ============================================================ step 1
md(r"""
---

## Step 1 · Data Integration — combining the two sheets

*PDF section: Data Integration*

The workbook stores the two years on separate sheets. Analysis needs one table,
so the sheets are stacked on top of each other.

A `SourceSheet` column is added first. It costs one column and it means that
after combining, we can still tell which sheet any row came from — which the
very next step depends on.

> Reading a million rows out of Excel takes a minute or two. This is the slowest
> cell in the notebook.
""")

code(r'''
sheet_2009_2010 = pd.read_excel(
    RAW_XLSX, sheet_name="Year 2009-2010",
    dtype={"Invoice": str, "StockCode": str})

sheet_2010_2011 = pd.read_excel(
    RAW_XLSX, sheet_name="Year 2010-2011",
    dtype={"Invoice": str, "StockCode": str})

# Remember where each row came from, before the two are mixed together.
sheet_2009_2010["SourceSheet"] = "Year 2009-2010"
sheet_2010_2011["SourceSheet"] = "Year 2010-2011"

df = pd.concat([sheet_2009_2010, sheet_2010_2011], ignore_index=True)
df["Description"] = df["Description"].astype("string")

rows_at_start = len(df)

record("Raw load (both sheets unioned)", "Data Integration",
       rows_at_start, rows_at_start,
       "Year 2009-2010 and Year 2010-2011 stacked into one table.")

print(f"\nSheet 1 : {len(sheet_2009_2010):>9,} rows")
print(f"Sheet 2 : {len(sheet_2010_2011):>9,} rows")
print(f"Combined: {rows_at_start:>9,} rows, {df.shape[1]} columns")

df.head()
''')

# ============================================================ step 2
md(r"""
---

## Step 2 · Data Integration — removing the overlap between the sheets

*PDF section: Data Integration*

The two sheets are **not** a clean split of time. They overlap between
1 and 9 December 2010: the same invoices were exported twice, once into each
sheet.

Stacking them therefore counts that week's sales twice. This is the single most
valuable correction in the whole pipeline, and it is invisible unless you look
for it.

The test distinguishes two different things:

- a row that is a duplicate **across** the two sheets → the export overlap, remove it
- a row that is a duplicate **inside** one sheet → a different problem, handled in Step 3

`keep="first"` keeps the earliest copy and marks the later one for removal.
""")

code(r'''
# The eight real data columns - everything except the marker we added.
value_columns = [c for c in df.columns if c != "SourceSheet"]

# A row is a cross-sheet copy if it repeats on the data columns, but does NOT
# repeat once the sheet name is included - i.e. the twin lives in the other sheet.
repeats_on_data = df.duplicated(subset=value_columns, keep="first")
repeats_within_sheet = df.duplicated(subset=value_columns + ["SourceSheet"], keep="first")
is_cross_sheet_copy = repeats_on_data & ~repeats_within_sheet

double_counted_revenue = (df.loc[is_cross_sheet_copy, "Quantity"]
                          * df.loc[is_cross_sheet_copy, "Price"]).sum()

rows_before = len(df)
df = df[~is_cross_sheet_copy].reset_index(drop=True)

record("Cross-sheet overlap resolution", "Data Integration",
       rows_before, len(df),
       f"Sheets overlap 1-9 Dec 2010; the same invoices appear twice. "
       f"Removing the second copy corrects GBP {double_counted_revenue:,.2f} "
       f"of double-counted revenue.")

print(f"\nRows removed          : {rows_before - len(df):,}")
print(f"Double-counted revenue: GBP {double_counted_revenue:,.2f}")
''')

# ============================================================ step 3
md(r"""
---

## Step 3 · Duplicate Removal — exact repeats inside a sheet

*PDF section: Duplicate Removal*

Now the ordinary case: the identical line recorded twice within the same sheet.
Same invoice, same product, same quantity, same price, same timestamp — a data
entry repeat rather than two real sales.

`.duplicated()` marks every copy after the first, so the original survives.
""")

code(r'''
is_duplicate = df.duplicated(subset=value_columns, keep="first")

duplicate_revenue = (df.loc[is_duplicate, "Quantity"]
                     * df.loc[is_duplicate, "Price"]).sum()

rows_before = len(df)
df = df[~is_duplicate].reset_index(drop=True)

record("Duplicate removal (within sheet)", "Duplicate Removal",
       rows_before, len(df),
       f"Exact repeats of all 8 fields inside one sheet. "
       f"Removed GBP {duplicate_revenue:,.2f} of double-counted revenue.")

print(f"\nRows removed     : {rows_before - len(df):,}")
print(f"Revenue corrected: GBP {duplicate_revenue:,.2f}")
print(f"Rows remaining   : {len(df):,}")
''')

# ============================================================ step 4 (country map)
md(r"""
---

## Step 4 · Unit Conversion — one spelling for everything

*PDF section: Unit Conversion / format standardisation*

The same thing is written several ways in this data. `EIRE` and `Ireland` are one
country. `" white hanging heart "` and `"White Hanging Heart"` are one product.
Left alone, each spelling becomes its own row in every chart.

Four fixes are applied:

1. **Whitespace and case** — trim, collapse repeated spaces, title-case descriptions, upper-case stock codes
2. **Country names** — map every raw label to one canonical name, and attach an ISO-3 code
3. **Region** — group countries into regions (this is *Data Generalization*: a detailed attribute replaced by a broader one)
4. **Identifiers as text** — an invoice number is a label, not a quantity. Storing it as a number invites Excel to average it.

The mapping table below is the standardisation, written out in full so it can be
checked.
""")

code(r'''
# Canonical name, ISO-3 code, and region for every country label in the raw data.
COUNTRY_MAP = {
    "United Kingdom":       ("United Kingdom", "GBR", "United Kingdom"),
    "EIRE":                 ("Ireland", "IRL", "Europe"),
    "Germany":              ("Germany", "DEU", "Europe"),
    "France":               ("France", "FRA", "Europe"),
    "Netherlands":          ("Netherlands", "NLD", "Europe"),
    "Spain":                ("Spain", "ESP", "Europe"),
    "Switzerland":          ("Switzerland", "CHE", "Europe"),
    "Belgium":              ("Belgium", "BEL", "Europe"),
    "Portugal":             ("Portugal", "PRT", "Europe"),
    "Australia":            ("Australia", "AUS", "Asia-Pacific"),
    "Channel Islands":      ("Channel Islands", "JEY", "Europe"),
    "Italy":                ("Italy", "ITA", "Europe"),
    "Norway":               ("Norway", "NOR", "Europe"),
    "Sweden":               ("Sweden", "SWE", "Europe"),
    "Cyprus":               ("Cyprus", "CYP", "Europe"),
    "Finland":              ("Finland", "FIN", "Europe"),
    "Austria":              ("Austria", "AUT", "Europe"),
    "Denmark":              ("Denmark", "DNK", "Europe"),
    "Greece":               ("Greece", "GRC", "Europe"),
    "Japan":                ("Japan", "JPN", "Asia-Pacific"),
    "USA":                  ("United States", "USA", "Americas"),
    "Poland":               ("Poland", "POL", "Europe"),
    "United Arab Emirates": ("United Arab Emirates", "ARE", "Middle East & Africa"),
    "Israel":               ("Israel", "ISR", "Middle East & Africa"),
    "Hong Kong":            ("Hong Kong", "HKG", "Asia-Pacific"),
    "Singapore":            ("Singapore", "SGP", "Asia-Pacific"),
    "Malta":                ("Malta", "MLT", "Europe"),
    "Iceland":              ("Iceland", "ISL", "Europe"),
    "Canada":               ("Canada", "CAN", "Americas"),
    "Lithuania":            ("Lithuania", "LTU", "Europe"),
    "RSA":                  ("South Africa", "ZAF", "Middle East & Africa"),
    "Bahrain":              ("Bahrain", "BHR", "Middle East & Africa"),
    "Brazil":               ("Brazil", "BRA", "Americas"),
    "Thailand":             ("Thailand", "THA", "Asia-Pacific"),
    "Korea":                ("South Korea", "KOR", "Asia-Pacific"),
    "Lebanon":              ("Lebanon", "LBN", "Middle East & Africa"),
    "West Indies":          ("West Indies", "", "Americas"),
    "Bermuda":              ("Bermuda", "BMU", "Americas"),
    "Nigeria":              ("Nigeria", "NGA", "Middle East & Africa"),
    "Czech Republic":       ("Czech Republic", "CZE", "Europe"),
    "Saudi Arabia":         ("Saudi Arabia", "SAU", "Middle East & Africa"),
    "Unspecified":          ("Unknown", "", "Unknown"),
    "European Community":   ("Unknown (EU, unallocated)", "", "Europe"),
}

print(f"{len(COUNTRY_MAP)} country labels mapped.")
print("\nExamples of labels that change:")
for raw in ["EIRE", "RSA", "USA", "Korea", "Unspecified"]:
    print(f"  {raw:<14} ->  {COUNTRY_MAP[raw][0]:<26} ({COUNTRY_MAP[raw][2]})")
''')

code(r'''
# --- 1. text tidy-up ------------------------------------------------------
df["Invoice"] = df["Invoice"].astype("string").str.strip()
df["StockCode"] = df["StockCode"].astype("string").str.strip().str.upper()
df["Description"] = (df["Description"].astype("string")
                     .str.strip()
                     .str.replace(r"\s+", " ", regex=True)   # collapse double spaces
                     .str.title())

# --- 2 & 3. country, ISO code, region ------------------------------------
df["CountryRaw"] = df["Country"]


def look_up_country(label, position):
    """Return part of the mapping for a country label, or a safe default."""
    default = (label, "", "Unknown")
    return COUNTRY_MAP.get(label, default)[position]


rows_renamed = int((df["CountryRaw"].map(lambda c: look_up_country(c, 0)) != df["CountryRaw"]).sum())

df["Country"] = df["CountryRaw"].map(lambda c: look_up_country(c, 0))
df["CountryISO"] = df["CountryRaw"].map(lambda c: look_up_country(c, 1))
df["Region"] = df["CountryRaw"].map(lambda c: look_up_country(c, 2))

# --- 4. identifiers and dates --------------------------------------------
df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"])

# A missing Customer ID becomes -1 rather than blank, so the column stays a
# whole number and the rows stay visible instead of quietly vanishing.
df["CustomerKey"] = (df["Customer ID"].astype("Float64").astype("Int64")
                     .fillna(-1).astype("int64"))

record("Unit conversion & format standardisation", "Unit Conversion",
       len(df), len(df),
       f"Trimmed and title-cased Description, upper-cased StockCode, mapped "
       f"{rows_renamed:,} rows to canonical country names with ISO-3 codes, "
       f"cast Invoice and Customer ID to text (identifiers, not numbers).",
       step_type="Derivation",
       note=f"{rows_renamed:,} rows remapped to canonical country names + ISO-3 codes")

print(f"\nRows given a new country name: {rows_renamed:,}")
print(f"Rows with no Customer ID     : {(df['CustomerKey'] == -1).sum():,}")

df[["Invoice", "StockCode", "Description", "CountryRaw",
    "Country", "CountryISO", "Region", "CustomerKey"]].head()
''')

# ============================================================ step 5
md(r"""
---

## Step 5 · Duplicate Removal — the business key

*PDF section: Duplicate Removal*

Step 3 removed rows where **all eight** columns matched. That misses a real case:
the same sale entered twice under two different spellings of the product name.
The descriptions differ, so an exact match does not see them.

The fix is to decide what actually *identifies a transaction*. A description is
an attribute of the product, not part of the sale's identity, so it is excluded
from the key:

> `Invoice + StockCode + Quantity + Price + InvoiceDate + CustomerKey + Country`

This is also the reason the product table is keyed on `StockCode` and never on
`Description`.
""")

code(r'''
business_key = ["Invoice", "StockCode", "Quantity", "Price",
                "InvoiceDate", "CustomerKey", "Country"]

# How many stock codes are sold under more than one description?
codes_with_many_names = int((df.groupby("StockCode")["Description"].nunique() > 1).sum())

is_business_duplicate = df.duplicated(subset=business_key, keep="first")
business_duplicate_revenue = (df.loc[is_business_duplicate, "Quantity"]
                              * df.loc[is_business_duplicate, "Price"]).sum()

rows_before = len(df)
df = df[~is_business_duplicate].reset_index(drop=True)

record("Duplicate removal (transaction business key)", "Duplicate Removal",
       rows_before, len(df),
       f"{codes_with_many_names:,} StockCodes carry more than one Description, so an "
       f"exact 8-column match misses lines duplicated under a second spelling. "
       f"De-duplicating on the business key caught {rows_before - len(df)} further "
       f"rows worth GBP {business_duplicate_revenue:,.2f}.")

print(f"\nStock codes with more than one description: {codes_with_many_names:,}")
print(f"Extra duplicate rows found                : {rows_before - len(df):,}")
print(f"Revenue corrected                         : GBP {business_duplicate_revenue:,.2f}")
''')

# ============================================================ step 6
md(r"""
---

## Step 6 · Missing Values — deciding *why* they are missing

*PDF section: Missing Value Types (MCAR / MAR / MNAR)*

Missing values are not all the same, and the type decides the treatment:

| Type | Meaning | Treatment |
|---|---|---|
| **MCAR** | Missing completely at random | Safe to drop; nothing is biased |
| **MAR** | Missingness is explained by other columns | Drop or impute, with justification |
| **MNAR** | Missingness depends on the missing value itself | Dangerous to drop |

Here the rows with no `Description` are **MAR**, and the evidence is in the data:
they also have `Price = 0`, and they also have no customer. The missingness is
fully predicted by columns we can see. They are not sales — they carry zero
revenue by construction — so removing them loses nothing.

The cell below checks that claim before acting on it, rather than assuming it.
""")

code(r'''
missing_description = df["Description"].isna()

# Evidence for the MAR claim - check it before deleting anything.
print(f"Rows with no Description        : {missing_description.sum():,}")
if missing_description.sum() > 0:
    share_zero_price = (df.loc[missing_description, "Price"] == 0).mean() * 100
    share_no_customer = (df.loc[missing_description, "CustomerKey"] == -1).mean() * 100
    print(f"  ... of those, Price = 0       : {share_zero_price:.1f}%")
    print(f"  ... of those, no Customer ID  : {share_no_customer:.1f}%")

is_non_transaction = missing_description & (df["Price"] == 0)
lost_revenue = (df.loc[is_non_transaction, "Quantity"]
                * df.loc[is_non_transaction, "Price"]).sum()

rows_before = len(df)
df = df[~is_non_transaction].reset_index(drop=True)

record("Missing-value removal (MAR, zero-value)", "Missing Value Types (MAR)",
       rows_before, len(df),
       f"Rows with no Description AND Price = 0. Missingness is fully predicted "
       f"by observed fields, so MAR not MCAR. Revenue impact GBP {lost_revenue:,.2f} "
       f"(zero by construction) - deleting them loses nothing.")

print(f"\nRows removed  : {rows_before - len(df):,}")
print(f"Revenue lost  : GBP {lost_revenue:,.2f}")
''')

# ============================================================ step 7
md(r"""
---

## Step 7 · Missing Value Handling — impute rather than delete

*PDF section: Missing Value Handling*

Step 6 deleted rows that were worth nothing. Any description still missing after
that belongs to a row carrying **real revenue**, and deleting those would destroy
money that genuinely happened. So they are filled with a placeholder instead.

In this dataset the count comes out at zero — which is not a wasted step, it is
the *proof* of the MAR claim made in Step 6. Every missing description really was
a zero-value record. The safeguard stays in the code because the rule should not
depend on the dataset happening to be tidy.
""")

code(r'''
still_missing = int(df["Description"].isna().sum())
df["Description"] = df["Description"].fillna("Unknown Item")

record("Missing Description imputation (placeholder)", "Missing Value Handling",
       len(df), len(df),
       f"Safeguard for any Description still missing after the previous step: fill "
       f"with 'Unknown Item' rather than delete, since such rows carry real revenue. "
       f"{still_missing:,} rows needed it - every missing Description was a "
       f"zero-value record already removed, which is itself the evidence for the "
       f"MAR claim.",
       step_type="Derivation", note=f"{still_missing:,} rows imputed")

print(f"\nRows needing a placeholder: {still_missing:,}")
print(f"Missing descriptions now  : {df['Description'].isna().sum():,}")
''')

# ============================================================ step 8
md(r"""
---

## Step 8 · Anomaly removal — test records

*PDF section: Data Cleaning (anomalies)*

Stock codes `TEST001` and `TEST002` are somebody checking the till worked. They
are system events, not business events, so they leave.
""")

code(r'''
TEST_CODES = {"TEST001", "TEST002"}

is_test_record = df["StockCode"].isin(TEST_CODES)

rows_before = len(df)
df = df[~is_test_record].reset_index(drop=True)

record("Test-record removal", "Data Cleaning (anomalies)",
       rows_before, len(df),
       "StockCode TEST001/TEST002 are system tests, not transactions.")

print(f"\nTest rows removed: {rows_before - len(df):,}")
''')

# ============================================================ step 9
md(r"""
---

## Step 9 · Attribute Construction — building new columns

*PDF section: Attribute Construction*

The raw data does not contain revenue. It contains quantity and price, and
revenue is the two multiplied — the same shape as the PDF's
`Defect Percentage = Defective / Produced × 100`.

Fourteen columns are built here: `Revenue`, four flags (`IsCancelled`,
`IsReturn`, `IsProduct`, `HasCustomer`), two labels (`ItemType`,
`TransactionType`) and seven calendar columns. The most important is
`ItemType`, because **not every row is a product**. Postage, bank charges and manual adjustments all sit
in the same column as gift wrap, and totalling them together silently overstates
product sales. The classifier below separates them.
""")

code(r'''
FEE_CODES = {"POST", "DOT", "C2", "C3"}
ADJUST_CODES = {"M", "D", "B", "S", "BANK CHARGES", "AMAZONFEE", "ADJUST",
                "ADJUST2", "CRUK", "PADS"}


def classify_item(code):
    """Decide what kind of thing a stock code refers to."""
    upper = code.upper()
    if upper in TEST_CODES:
        return "Test Record"
    if upper in FEE_CODES:
        return "Postage & Carriage"
    if upper in ADJUST_CODES:
        return "Fee / Adjustment"
    if upper.startswith("GIFT"):
        return "Gift Voucher"
    if upper.startswith("DCGS"):
        return "Discount / Sample"
    if code[:5].isdigit():          # a real catalogue number looks like 85123A
        return "Product"
    return "Other Non-Product"


# --- money ----------------------------------------------------------------
df["Revenue"] = (df["Quantity"] * df["Price"]).round(2)

# --- flags ----------------------------------------------------------------
df["IsCancelled"] = df["Invoice"].str.startswith("C").fillna(False)
df["IsReturn"] = df["Quantity"] < 0
df["ItemType"] = df["StockCode"].map(classify_item)
df["IsProduct"] = df["ItemType"] == "Product"
df["HasCustomer"] = df["CustomerKey"] != -1

# --- one label describing what the row is ---------------------------------
# Assigned weakest-first, so the strongest description ends up winning.
df["TransactionType"] = "Sale"
df.loc[df["Price"] == 0, "TransactionType"] = "Zero-price (promo/adjustment)"
df.loc[df["IsReturn"], "TransactionType"] = "Write-off / Return"
df.loc[df["IsCancelled"], "TransactionType"] = "Cancellation"

# --- calendar parts, so charts can group by them --------------------------
timestamp = df["InvoiceDate"]
df["DateKey"] = timestamp.dt.normalize()     # midnight of that day
df["Year"] = timestamp.dt.year
df["Month"] = timestamp.dt.month
df["Quarter"] = timestamp.dt.quarter
df["Hour"] = timestamp.dt.hour
df["Weekday"] = timestamp.dt.day_name()
df["WeekdayNo"] = timestamp.dt.dayofweek

record("Attribute construction", "Attribute Construction",
       len(df), len(df),
       "Derived Revenue = Quantity x Price, plus the IsCancelled, IsReturn, "
       "IsProduct and HasCustomer flags, the ItemType and TransactionType "
       "labels, and seven calendar attributes. Mirrors the PDF's "
       "Defect Percentage = Defective / Produced x 100.",
       step_type="Derivation", note="14 new attributes added")

print("\nRows by item type:")
print(df.groupby("ItemType").agg(Rows=("Invoice", "size"),
                                 Revenue=("Revenue", "sum")).round(2))

print("\nRows by transaction type:")
print(df["TransactionType"].value_counts())
''')

# ============================================================ step 10
md(r"""
---

## Step 10 · Outlier Detection — flag, do not delete

*PDF section: Z-score Normalization / outlier detection*

This is the step where most projects quietly destroy their own data.

Two tests are used, and a row must **fail both** to be flagged:

- **Z-score** — how many standard deviations from the mean, flagged beyond 3
- **IQR** — beyond 3× the interquartile range past the quartiles

Requiring both is deliberate. A plain 1.5×IQR rule alone would discard around
**11% of all rows** in this dataset.

And flagged rows are **kept**. The 80,995-unit order is a genuine wholesale
purchase. The £38,970 manual adjustment is a genuine correction. Deleting them
would remove real money from the totals to make a chart look tidier. Flagging
them lets the dashboard show or exclude them as the question requires.
""")

code(r'''
df["Is_Outlier"] = False
df["Outlier_Reason"] = ""

for column in ["Quantity", "Price"]:
    values = df[column].astype(float)

    # Test 1: z-score - distance from the mean, in standard deviations.
    z_score = (values - values.mean()) / values.std()

    # Test 2: IQR - distance beyond the middle 50% of the data.
    q1 = values.quantile(0.25)
    q3 = values.quantile(0.75)
    iqr = q3 - q1
    lower_fence = q1 - IQR_MULTIPLIER * iqr
    upper_fence = q3 + IQR_MULTIPLIER * iqr

    # A row must fail BOTH tests, not either one.
    is_extreme = ((z_score.abs() > Z_THRESHOLD)
                  & ((values < lower_fence) | (values > upper_fence)))

    df.loc[is_extreme, "Is_Outlier"] = True
    df.loc[is_extreme, "Outlier_Reason"] = (
        df.loc[is_extreme, "Outlier_Reason"] + " + Extreme " + column)

    print(f"{column:<9}  z-limit +/-{Z_THRESHOLD}   "
          f"IQR fences [{lower_fence:,.2f}, {upper_fence:,.2f}]   "
          f"flagged {is_extreme.sum():,}")

df["Outlier_Reason"] = df["Outlier_Reason"].str.strip(" +")
df.loc[df["Outlier_Reason"] == "", "Outlier_Reason"] = "Within normal range"

flagged = int(df["Is_Outlier"].sum())
flagged_revenue = df.loc[df["Is_Outlier"], "Revenue"].sum()

record("Outlier detection (z-score + IQR)", "Z-score Normalization",
       len(df), len(df),
       f"{flagged:,} rows flagged (both |z| > 3 AND beyond 3xIQR), carrying "
       f"GBP {flagged_revenue:,.2f}. FLAGGED, NOT DELETED - the 80,995-unit order "
       f"and the GBP 38,970 Manual adjustment are real events.",
       step_type="Derivation", note=f"{flagged:,} rows flagged, 0 deleted")

print(f"\nRows flagged as outliers: {flagged:,}  ({flagged / len(df) * 100:.2f}% of the data)")
print(f"Rows deleted            : 0")
print("\nThe five most extreme rows, all retained:")
df.nlargest(5, "Quantity")[["Invoice", "StockCode", "Description",
                            "Quantity", "Price", "Revenue", "Outlier_Reason"]]
''')

# ============================================================ cleaning done
md(r"""
---

## Checkpoint · Cleaning is finished

Everything above was **data cleaning** — the table is now correct. Everything
below is **transformation and reduction**: reshaping correct data into forms that
answer questions.

The row count from here on does not change.
""")

code(r'''
print(f"Rows in  : {rows_at_start:,}")
print(f"Rows out : {len(df):,}   ({len(df) - rows_at_start:+,},"
      f" {(len(df) - rows_at_start) / rows_at_start * 100:+.2f}%)")
print(f"Revenue  : GBP {df['Revenue'].sum():,.2f}")
print(f"Date span: {df['DateKey'].min().date()} to {df['DateKey'].max().date()}")
print(f"Customers: {df.loc[df['HasCustomer'], 'CustomerKey'].nunique():,} identified")
''')

# ============================================================ step 11
md(r"""
---

## Step 11 · Listwise Deletion — applied to one table only

*PDF section: Listwise Deletion*

Listwise deletion means dropping any row with a missing value. Applied to the
whole dataset here it would be a serious mistake: the rows with no `Customer ID`
carry several million pounds of real revenue.

So the deletion is applied **only where it is needed**. You cannot analyse
customer behaviour for a customer you cannot identify — but you can still count
their money in the sales totals.

The result is a **dual grain**: the sales table keeps every row, and the customer
table is built from the identified subset. This is the single most important
judgement call in the project.
""")

code(r'''
rows_in_fact = len(df)

customers_only = df[df["HasCustomer"]].copy()

revenue_without_customer = df.loc[~df["HasCustomer"], "Revenue"].sum()

record("Listwise deletion (customer table only)", "Listwise Deletion",
       rows_in_fact, len(customers_only),
       f"{rows_in_fact - len(customers_only):,} rows have no Customer ID but carry "
       f"GBP {revenue_without_customer:,.2f}. They stay in the sales table and are "
       f"dropped only for customer analysis - you cannot impute an identity.",
       step_type="Parallel table",
       note=f"sales table keeps {rows_in_fact:,}; customer table built from "
            f"{len(customers_only):,}")

print(f"\nRows with no Customer ID : {rows_in_fact - len(customers_only):,}")
print(f"Revenue they carry       : GBP {revenue_without_customer:,.2f}")
print(f"Kept in the sales table  : yes - all of them")
''')

# ============================================================ step 12
md(r"""
---

## Step 12 · Normalization — Min-Max and Z-score

*PDF section: Data Normalization*

Before customers can be scored, the three RFM measures have to be put on a
comparable footing:

- **Recency** — days since last purchase (roughly 0–740)
- **Frequency** — number of orders (roughly 1–200)
- **Monetary** — total spend (roughly £1–£350,000)

Monetary would dominate any combined score purely because pounds are bigger
numbers than days. Both scalings from the PDF are applied:

$$\text{Min-Max} = \frac{x - \min}{\max - \min} \qquad\qquad Z = \frac{x - \mu}{\sigma}$$

Min-Max squeezes everything into 0–1. Z-score re-expresses each value as standard
deviations from the mean.
""")

code(r'''
# Build the RFM table first: one row per customer.
# Cancellations are excluded - a cancelled order is not a purchase.
snapshot_date = df["InvoiceDate"].max() + pd.Timedelta(days=1)
real_sales = customers_only[~customers_only["IsCancelled"]]

rfm = real_sales.groupby("CustomerKey").agg(
    Recency=("InvoiceDate", lambda s: (snapshot_date - s.max()).days),
    Frequency=("Invoice", "nunique"),
    Monetary=("Revenue", "sum"),
    FirstPurchase=("InvoiceDate", "min"),
    LastPurchase=("InvoiceDate", "max"),
    UnitsBought=("Quantity", "sum"),
    LinesBought=("Invoice", "size"),
).reset_index()

# Customers whose net spend is zero or negative (returns only) cannot be scored
# on Monetary, so they are set aside and added back at the end - never dropped,
# or their sales rows would point at a customer that no longer exists.
no_net_value = rfm[rfm["Monetary"] <= 0].copy()
rfm = rfm[rfm["Monetary"] > 0].copy()
scored_customers = len(rfm)

rfm["TenureDays"] = (snapshot_date - rfm["FirstPurchase"]).dt.days
rfm["AvgOrderValue"] = (rfm["Monetary"] / rfm["Frequency"]).round(2)

# --- the two normalizations ----------------------------------------------
for column in ["Recency", "Frequency", "Monetary"]:
    values = rfm[column].astype(float)
    rfm[f"{column}_MinMax"] = ((values - values.min())
                               / (values.max() - values.min())).round(4)
    rfm[f"{column}_Z"] = ((values - values.mean()) / values.std()).round(4)

record("Normalization (Min-Max + Z-score)", "Data Normalization",
       len(rfm), len(rfm),
       "Recency, Frequency and Monetary each scaled to 0-1 (Min-Max) and to "
       "standard-deviation units (Z-score) before scoring, so no variable "
       "dominates on units alone.",
       step_type="Derivation",
       note=f"{scored_customers:,} scored customers scaled, 6 new columns")

print(f"\nCustomers scored: {scored_customers:,}\n")
print("Before scaling - notice the wildly different ranges:")
print(rfm[["Recency", "Frequency", "Monetary"]].describe().round(2))

print("\nAfter Min-Max - everything now sits between 0 and 1:")
print(rfm[["Recency_MinMax", "Frequency_MinMax", "Monetary_MinMax"]].describe().round(4))
''')

# ============================================================ step 13
md(r"""
---

## Step 13 · Discretization and Generalization — from numbers to names

*PDF section: Discretization / Binning, Data Generalization*

**Discretization** turns a continuous number into a small number of ordered
buckets. `qcut` cuts into five equal-sized groups, so each score from 1 to 5
holds a fifth of the customers.

Note the direction: for Recency, **lower is better** — a customer who bought
yesterday scores 5, not 1.

**Generalization** then replaces the three scores with a single name a human can
act on: *Champions*, *At Risk*, *Hibernating*. This is the PDF's
*Young / Middle-aged / Senior* pattern — detail traded for meaning.
""")

code(r'''
# --- discretization into quintiles ---------------------------------------
# Recency is reversed: fewer days since purchase = better score.
rfm["R_Score"] = pd.qcut(rfm["Recency"], 5, labels=[5, 4, 3, 2, 1]).astype(int)

# .rank(method="first") breaks ties so the five buckets come out equal-sized.
rfm["F_Score"] = pd.qcut(rfm["Frequency"].rank(method="first"), 5,
                         labels=[1, 2, 3, 4, 5]).astype(int)
rfm["M_Score"] = pd.qcut(rfm["Monetary"].rank(method="first"), 5,
                         labels=[1, 2, 3, 4, 5]).astype(int)

rfm["RFM_Score"] = rfm["R_Score"] * 100 + rfm["F_Score"] * 10 + rfm["M_Score"]
rfm["RFM_Sum"] = rfm[["R_Score", "F_Score", "M_Score"]].sum(axis=1)


# --- generalization into named segments ----------------------------------
def name_the_segment(row):
    """Turn three scores into one word a marketer can act on."""
    recency = row["R_Score"]
    value = (row["F_Score"] + row["M_Score"]) / 2      # how much they are worth

    if recency >= 4 and value >= 4:
        return "Champions"
    if recency >= 3 and value >= 3:
        return "Loyal Customers"
    if recency >= 4 and value < 3:
        return "New / Promising"
    if recency == 3 and value < 3:
        return "Needs Attention"
    if recency == 2 and value >= 3:
        return "At Risk"
    if recency == 1 and value >= 4:
        return "Cannot Lose Them"
    if recency <= 2 and value < 3:
        return "Hibernating / Lost"
    return "Others"


rfm["Segment"] = rfm.apply(name_the_segment, axis=1)

# Two more generalizations: recency bands and value tiers.
rfm["RecencyBand"] = pd.cut(
    rfm["Recency"], [-1, 30, 90, 180, 365, 10 ** 6],
    labels=["0-30 days", "31-90 days", "91-180 days", "181-365 days", "365+ days"])
rfm["ValueTier"] = pd.qcut(rfm["Monetary"].rank(method="first"), 3,
                           labels=["Low", "Medium", "High"])

print("Customers per segment:\n")
print(rfm["Segment"].value_counts())
print("\nRevenue per segment:\n")
print(rfm.groupby("Segment")["Monetary"].sum().sort_values(ascending=False).round(2))
''')

# ============================================================ step 14
md(r"""
---

## Step 14 · CRM strategy, churn and Customer Lifetime Value

*PDF section: BA-2 CRM.pdf — CRM strategies, churn, CLV*

A segment name is a label. The dashboard needs to say what to **do** about it, so
each segment is mapped to one of the eight CRM strategies from the CRM PDF, plus
a concrete action.

**Churn** is inferred from recency, because a gift shop has no contract to
cancel — unlike the telecom example in the PDF, where churn is an event you can
observe.

**CLV** uses the PDF's own formula:

$$CLV = \text{Average Purchase Value} \times \text{Purchase Frequency} \times \text{Customer Lifespan}$$

Purchase frequency is annualised over each customer's observed tenure. Lifespan
is a stated **assumption** of 3 years — the data only covers two, so this is a
forward-looking estimate, not a measurement.
""")

code(r'''
CRM_STRATEGY = {
    "Champions":          "Loyalty Programs",
    "Loyal Customers":    "Cross-Selling and Up-Selling",
    "Cannot Lose Them":   "Customer Retention",
    "At Risk":            "Customer Retention",
    "Needs Attention":    "Personalization",
    "New / Promising":    "Customer Acquisition",
    "Hibernating / Lost": "Customer Acquisition (win-back)",
    "Others":             "Customer Feedback Management",
}

CRM_ACTION = {
    "Loyalty Programs":                "Reward repeat purchase; recruit as advocates",
    "Cross-Selling and Up-Selling":    "Recommend related and higher-value lines",
    "Customer Retention":              "Personal outreach before they lapse",
    "Personalization":                 "Tailor offers to past purchases",
    "Customer Acquisition":            "Onboard well; secure the second order",
    "Customer Acquisition (win-back)": "Reactivation campaign or let go",
    "Customer Feedback Management":    "Survey to find out what went wrong",
    "Not applicable":                  "Identify the customer at point of sale",
}

rfm["CRM_Strategy"] = rfm["Segment"].map(CRM_STRATEGY).fillna("Not applicable")
rfm["CRM_Action"] = rfm["CRM_Strategy"].map(CRM_ACTION)

# --- churn, inferred from recency ----------------------------------------
rfm["ChurnStatus"] = pd.cut(rfm["Recency"], [-1, 90, 180, 10 ** 6],
                            labels=["Active", "At Risk", "Churned"])
rfm["IsChurned"] = rfm["Recency"] > 180

# --- customer lifetime value ---------------------------------------------
# Tenure is floored at a quarter of a year so a customer who bought once last
# week does not appear to buy 52 times a year.
tenure_years = (rfm["TenureDays"] / 365).clip(lower=0.25)
rfm["PurchaseFrequencyPerYear"] = (rfm["Frequency"] / tenure_years).round(2)
rfm["CLV"] = (rfm["AvgOrderValue"] * rfm["PurchaseFrequencyPerYear"]
              * LIFESPAN_YEARS).round(2)

record("CRM segmentation & CLV", "CRM Strategies (BA-2)",
       len(rfm), len(rfm),
       f"Each RFM segment mapped to one of the eight CRM strategies with a "
       f"recommended action; churn status derived from recency; CLV computed "
       f"with the PDF's formula assuming a {LIFESPAN_YEARS}-year lifespan.",
       step_type="Derivation",
       note=f"{rfm['CRM_Strategy'].nunique()} CRM strategies, 3 churn states, CLV per customer")

record("Discretization & generalization (RFM)", "Discretization / Generalization",
       len(rfm), len(rfm),
       f"Continuous R, F, M cut into quintile scores 1-5, then generalized into "
       f"{rfm['Segment'].nunique()} named segments, 5 recency bands and 3 value tiers.",
       step_type="Derivation",
       note=f"{rfm['Segment'].nunique()} segments, 5 recency bands, 3 value tiers")

print("\nChurn status:\n")
print(rfm["ChurnStatus"].value_counts())

print("\nWhat to do about each segment:\n")
print(rfm.groupby(["Segment", "CRM_Strategy", "CRM_Action"])
      .agg(Customers=("CustomerKey", "size"), Revenue=("Monetary", "sum"))
      .round(0).sort_values("Revenue", ascending=False))
''')

# ============================================================ step 15
md(r"""
---

## Step 15 · Dimensionality Reduction — the star schema

*PDF section: Dimensionality Reduction*

The cleaned table is wide, and most of its width repeats. "United Kingdom" is
stored a million times; the description of stock code 85123A is stored every time
it sold.

Dimensionality reduction stores each descriptive attribute **once per key**
instead of once per row: a narrow sales table of measurements, plus small lookup
tables of attributes.

**This is the reason the Power BI folder contains several CSV files.** They are
not several datasets — they are one dataset, stored in the shape Power BI reads
efficiently. The single Excel file written at the end of this notebook is the
same data, joined back into one table.

> One deliberate redundancy: `Revenue = Quantity × Price` is stored even though it
> could be recomputed. It saves the dashboard doing a million multiplications on
> every click. It also means those three columns must never be fed into one
> correlation, because two of them contain the third.
""")

code(r'''
sales_columns = ["Invoice", "StockCode", "CustomerKey", "Country",
                 "InvoiceDate", "Hour", "Quantity", "Price", "Revenue",
                 "IsCancelled", "IsReturn", "IsProduct", "HasCustomer",
                 "TransactionType", "Is_Outlier", "Outlier_Reason", "SourceSheet"]

columns_before = df.shape[1]
sales = df[sales_columns].copy()

record("Dimensionality reduction (star schema)", "Dimensionality Reduction",
       columns_before, len(sales_columns),
       f"One flat {columns_before}-column table split into a {len(sales_columns)}-column "
       f"sales table plus 5 lookup tables. Descriptive attributes stored once per "
       f"key instead of once per row.",
       step_type="Schema",
       note=f"{columns_before} columns -> {len(sales_columns)} sales columns + 6 lookups")

print(f"\nWide table  : {columns_before} columns x {len(df):,} rows")
print(f"Sales table : {len(sales_columns)} columns x {len(sales):,} rows")
print(f"Customers   : {len(rfm):,} rows")
print(f"Products    : {df['StockCode'].nunique():,} rows")
print(f"Countries   : {df['Country'].nunique():,} rows")
''')

# ============================================================ step 16
md(r"""
---

## Step 16 · Noise Reduction — smoothing the daily series

*PDF section: Noise Reduction*

Daily revenue is spiky. A single large wholesale order makes one day look like a
trend. Two smoothing methods from the PDF are applied so the underlying shape is
visible:

**Moving average** — the mean of the last 7 days.

**Exponential smoothing** — each new value is a blend of what actually happened
and what was previously expected:

$$\text{New} = \alpha \times \text{Actual} + (1 - \alpha) \times \text{Previous}$$

With α = 0.3, each day is 30% today and 70% history, so the line reacts to change
without chasing every spike.
""")

code(r'''
daily = (sales.groupby(sales["InvoiceDate"].dt.normalize().rename("DateKey"))
         .agg(Revenue=("Revenue", "sum"),
              Orders=("Invoice", "nunique"),
              Units=("Quantity", "sum"),
              Lines=("Invoice", "size"),
              Customers=("CustomerKey", "nunique"))
         .reset_index().sort_values("DateKey"))
daily["Revenue"] = daily["Revenue"].round(2)

# --- moving average -------------------------------------------------------
daily[f"Revenue_MA{MA_WINDOW}"] = (daily["Revenue"]
                                   .rolling(MA_WINDOW, min_periods=1)
                                   .mean().round(2))

# --- exponential smoothing, written out step by step ----------------------
smoothed_values = []
previous = daily["Revenue"].iloc[0]          # seed with the first real value
for actual in daily["Revenue"]:
    previous = ALPHA * actual + (1 - ALPHA) * previous
    smoothed_values.append(round(previous, 2))
daily["Revenue_ExpSmooth"] = smoothed_values

record(f"Noise reduction (MA{MA_WINDOW} + exponential smoothing)", "Noise Reduction",
       len(sales), len(daily),
       f"Daily revenue is spiky. Added a {MA_WINDOW}-point moving average and an "
       f"exponentially smoothed series with alpha = {ALPHA}, exactly the PDF's "
       f"New = a*Actual + (1-a)*Previous formula.",
       step_type="Derivation",
       note=f"2 smoothed series over {len(daily):,} daily points")

print(f"\n{len(daily):,} trading days\n")
daily.head(10)
''')

# ============================================================ step 17
md(r"""
---

## Step 17 · Data Cube Aggregation — pre-summarising

*PDF section: Data Cube Aggregation*

A chart of monthly revenue by country does not need a million rows; it needs a few
thousand pre-computed totals. Three cubes are built so the dashboard scans
thousands of rows instead of a million on every click.
""")

code(r'''
month = sales["InvoiceDate"].dt.strftime("%Y-%m")

month_country = (sales.assign(MonthYear=month)
                 .groupby(["MonthYear", "Country"])
                 .agg(Revenue=("Revenue", "sum"), Orders=("Invoice", "nunique"),
                      Units=("Quantity", "sum"),
                      Customers=("CustomerKey", "nunique"))
                 .reset_index())

month_product = (sales.assign(MonthYear=month)
                 .groupby(["MonthYear", "StockCode"])
                 .agg(Revenue=("Revenue", "sum"), Units=("Quantity", "sum"),
                      Orders=("Invoice", "nunique"))
                 .reset_index())

total_summary_rows = len(daily) + len(month_country) + len(month_product)

record("Data cube aggregation", "Data Cube Aggregation",
       len(sales), total_summary_rows,
       f"Line-level rows pre-summarised into daily ({len(daily):,}), "
       f"month x country ({len(month_country):,}) and month x product "
       f"({len(month_product):,}) cubes.",
       step_type="Aggregation",
       note=f"{len(sales):,} lines -> {total_summary_rows:,} summary rows "
            f"({len(sales) / total_summary_rows:.0f}x smaller)")

print(f"\nDaily           : {len(daily):>8,} rows")
print(f"Month x Country : {len(month_country):>8,} rows")
print(f"Month x Product : {len(month_product):>8,} rows")
print(f"Total           : {total_summary_rows:>8,} rows "
      f"({len(sales) / total_summary_rows:.0f}x smaller than {len(sales):,})")
''')

# ============================================================ step 18
md(r"""
---

## Step 18 · Histogram Reduction — keeping the shape, dropping the values

*PDF section: Histogram Reduction*

To draw a distribution you do not need every value — you need how many values fell
into each band. Three distributions are stored as bin counts instead of raw
numbers, preserving the shape at a tiny fraction of the size.
""")

code(r'''
basket_value = sales.groupby("Invoice")["Revenue"].sum()

histogram_specs = [
    ("Quantity", sales.loc[sales["IsProduct"], "Quantity"],
     [-10 ** 6, -1, 0, 2, 5, 10, 25, 50, 100, 10 ** 6],
     ["Negative", "0", "1-2", "3-5", "6-10", "11-25", "26-50", "51-100", "100+"]),
    ("Unit Price", sales.loc[sales["IsProduct"], "Price"],
     [-10 ** 6, 0, 1, 2, 5, 10, 20, 10 ** 6],
     ["<= 0", "0-1", "1-2", "2-5", "5-10", "10-20", "20+"]),
    ("Basket Value", basket_value,
     [-10 ** 9, 0, 50, 150, 300, 600, 1500, 10 ** 9],
     ["Negative", "0-50", "50-150", "150-300", "300-600", "600-1500", "1500+"]),
]

histogram_frames = []
for metric, series, bin_edges, bin_labels in histogram_specs:
    binned = pd.cut(series, bins=bin_edges, labels=bin_labels)
    counts = binned.value_counts().reindex(bin_labels, fill_value=0)
    histogram_frames.append(pd.DataFrame({
        "Metric": metric,
        "Bin": bin_labels,
        "BinOrder": range(1, len(bin_labels) + 1),
        "Frequency": counts.values,
    }))

histograms = pd.concat(histogram_frames, ignore_index=True)

raw_values = len(sales) + len(basket_value)

record("Histogram reduction", "Histogram Reduction",
       raw_values, len(histograms),
       f"Distributions of Quantity, Unit Price and Basket Value stored as "
       f"{len(histograms)} bin frequencies instead of {raw_values:,} raw values, "
       f"preserving shape at a fraction of the size.",
       step_type="Aggregation",
       note=f"{raw_values:,} values -> {len(histograms)} bin frequencies")

print(f"\n{raw_values:,} raw values  ->  {len(histograms)} bin counts\n")
histograms
''')

# ============================================================ audit trail
md(r"""
---

## The complete audit trail

Every step, in order, with what it did and why. This table is written into the
Excel file as its own sheet, so the preprocessing travels with the data.
""")

code(r'''
audit = pd.DataFrame(audit_rows)

print(f"{len(audit)} steps, "
      f"{audit['PDF_Section'].nunique()} distinct techniques from the PDF\n")

audit[["Step", "Technique", "PDF_Section", "Step_Type",
       "Rows_Before", "Rows_After", "Rows_Removed", "Note"]]
''')

md(r"""
### Reconciliation

The numbers below are the check. They must match the dashboard and both PDF
reports exactly — if any one of them drifted, this is where it would show.
""")

code(r'''
gross_revenue = sales["Revenue"].sum()
sale_rows = sales[sales["TransactionType"] == "Sale"]
returns = -sales.loc[sales["Quantity"] < 0, "Revenue"].sum()

summary_table = pd.DataFrame({
    "Measure": [
        "Rows in (raw)", "Rows out (cleaned)", "Rows removed",
        "Net revenue", "Orders", "Identified customers",
        "Units sold", "Average order value", "Return rate",
        "Outlier rows flagged (kept)", "Rows with no customer (kept)",
    ],
    "Value": [
        f"{rows_at_start:,}",
        f"{len(sales):,}",
        f"{rows_at_start - len(sales):,} "
        f"({(rows_at_start - len(sales)) / rows_at_start * 100:.2f}%)",
        f"GBP {gross_revenue:,.2f}",
        f"{sales['Invoice'].nunique():,}",
        f"{sales.loc[sales['HasCustomer'], 'CustomerKey'].nunique():,}",
        f"{sales['Quantity'].sum():,}",
        f"GBP {gross_revenue / sales['Invoice'].nunique():,.2f}",
        f"{returns / sale_rows['Revenue'].sum() * 100:.1f}%",
        f"{int(sales['Is_Outlier'].sum()):,}",
        f"{int((~sales['HasCustomer']).sum()):,}",
    ],
})

summary_table
''')

# ============================================================ excel output
md(r"""
---

## The result · one cleaned Excel file

Everything above produced one corrected table. This writes it out as a **single
workbook** with three sheets:

| Sheet | What it holds |
|---|---|
| `Cleaned_Data` | Every cleaned transaction line, with the columns built along the way |
| `Audit_Trail` | The eighteen steps, what each removed, and why |
| `Summary` | The reconciliation figures above |

**On the row limit.** An `.xlsx` sheet holds at most **1,048,576** rows. The
cleaned data has just over a million, so it fits — with only about 20,000 rows to
spare. Had the cleaning removed less, this file could not have existed, which is
worth knowing before promising anyone a single Excel file of a large dataset.

Writing a million rows takes a couple of minutes and produces a large file, so
Excel will also take a moment to open it.
""")

code(r'''
FINAL_COLUMNS = [
    # what was sold
    "Invoice", "StockCode", "Description", "ItemType", "IsProduct",
    # the money
    "Quantity", "Price", "Revenue",
    # when
    "InvoiceDate", "Year", "Quarter", "Month", "Weekday", "Hour",
    # who
    "CustomerKey", "HasCustomer",
    # where
    "Country", "CountryISO", "Region",
    # what kind of row it is
    "TransactionType", "IsCancelled", "IsReturn",
    "Is_Outlier", "Outlier_Reason", "SourceSheet",
]

cleaned = df[FINAL_COLUMNS].copy()

# Excel shows True/False for a real boolean but 1/0 for pandas' own boolean
# type, so the flag columns are converted to plain Python booleans first.
for column in ["IsProduct", "HasCustomer", "IsCancelled", "IsReturn", "Is_Outlier"]:
    cleaned[column] = cleaned[column].astype(bool).astype(object)

EXCEL_ROW_LIMIT = 1_048_576
print(f"Rows to write : {len(cleaned):,}")
print(f"Excel limit   : {EXCEL_ROW_LIMIT:,}")
print(f"Headroom      : {EXCEL_ROW_LIMIT - len(cleaned) - 1:,} rows")
print(f"Missing values: {int(cleaned.isna().sum().sum()):,}")

if len(cleaned) + 1 > EXCEL_ROW_LIMIT:
    raise ValueError("Cleaned data does not fit in one Excel sheet.")

cleaned.head()
''')

md(r"""
### Why this is not just `.to_excel()`

A million rows will not fit in memory as a workbook, so `xlsxwriter` is put in
**constant memory** mode: it writes each row straight to disk and then forgets
it.

That mode has one condition — rows must be written **in order, top to bottom** —
and pandas' `.to_excel()` fills a sheet *column by column*. Every row therefore
gets flushed to disk before its later cells arrive, and the file comes out with
only its first column populated. There is no error message; the file simply
comes out wrong.

So the rows are written directly below, in order. `itertuples()` hands them over
one at a time rather than building a copy of the whole table.
""")

code(r'''
import xlsxwriter

OUT_XLSX.parent.mkdir(parents=True, exist_ok=True)

workbook = xlsxwriter.Workbook(str(OUT_XLSX), {
    "constant_memory": True,                      # stream rows, do not hold them
    "default_date_format": "yyyy-mm-dd hh:mm:ss",
})

header_style = workbook.add_format(
    {"bold": True, "bg_color": "#DCE6F1", "border": 1})


def write_sheet(sheet_name, frame):
    """Write one table to its own sheet, one row at a time and in order."""
    sheet = workbook.add_worksheet(sheet_name)

    # The audit trail leaves row counts blank for steps that are not deletions.
    # Those blanks arrive here as NaN, which xlsxwriter refuses to write as a
    # number, so they are turned into None - which it writes as an empty cell.
    if frame.isna().any().any():
        frame = frame.astype(object).where(frame.notna(), None)

    sheet.write_row(0, 0, [str(c) for c in frame.columns], header_style)

    for row_number, values in enumerate(frame.itertuples(index=False), start=1):
        sheet.write_row(row_number, 0, values)

    sheet.freeze_panes(1, 0)                      # keep headers visible
    sheet.autofilter(0, 0, len(frame), len(frame.columns) - 1)
    print(f"  {sheet_name:<14} {len(frame):>9,} rows x {len(frame.columns):>2} columns")


print("Writing sheets:")
write_sheet("Cleaned_Data", cleaned)
write_sheet("Audit_Trail", audit)
write_sheet("Summary", summary_table)

workbook.close()

size_mb = OUT_XLSX.stat().st_size / 1024 / 1024
print(f"\nWritten: {OUT_XLSX}")
print(f"Size   : {size_mb:,.1f} MB")
''')

md(r"""
### Checking the file rather than trusting it

The failure described above produces a file that looks fine until you open it, so
the workbook is read back and checked against the numbers computed earlier.
""")

code(r'''
check = pd.read_excel(OUT_XLSX, sheet_name="Cleaned_Data")

print(f"Rows read back : {len(check):,}   (expected {len(cleaned):,})")
print(f"Columns        : {check.shape[1]}   (expected {len(FINAL_COLUMNS)})")
print(f"Revenue        : GBP {check['Revenue'].sum():,.2f}   "
      f"(expected GBP {cleaned['Revenue'].sum():,.2f})")
print(f"Orders         : {check['Invoice'].nunique():,}   "
      f"(expected {cleaned['Invoice'].nunique():,})")

assert len(check) == len(cleaned), "row count does not match"
assert check.shape[1] == len(FINAL_COLUMNS), "column count does not match"
assert round(check["Revenue"].sum(), 2) == round(cleaned["Revenue"].sum(), 2), \
    "revenue does not match"

# Every column must be complete except CountryISO. Three labels - "Unknown",
# "Unknown (EU, unallocated)" and "West Indies" - have no ISO-3 code to give,
# so those cells are left genuinely blank rather than filled with an invented
# code. Anything else being blank would mean the write went wrong.
blank_counts = check.isna().sum()
unexpected_blanks = blank_counts.drop("CountryISO")

print(f"\nBlank cells outside CountryISO: {int(unexpected_blanks.sum()):,}   (expected 0)")
print(f"Blank CountryISO (no ISO-3 exists): {int(blank_counts['CountryISO']):,}")
print(check.loc[check["CountryISO"].isna(), "Country"].value_counts().to_string())

assert int(unexpected_blanks.sum()) == 0, "the file has unexpected empty cells"

print("\nAll checks passed - the workbook holds the complete cleaned dataset.")
check.head()
''')

md(r"""
---

## Summary

Starting from the raw workbook, **eighteen steps** applied **thirteen techniques**
from *BA-6 Preprocessing Techniques.pdf*.

The two corrections that mattered most were both invisible without looking for
them: the **overlap between the two sheets**, which was double-counting a week of
December 2010, and the decision to **flag outliers rather than delete them**,
which kept genuine wholesale orders in the totals.

The one judgement call worth defending is the **dual grain** — rows with no
customer stay in the sales table and are excluded only from customer analysis.
Dropping them would have been simpler and would have quietly deleted several
million pounds of real revenue.

The result is `OnlineRetailII_Cleaned.xlsx`: one file, one cleaned dataset, with
its own audit trail attached.
""")

# ============================================================ write
nb = nbf.v4.new_notebook(cells=cells)
nb.metadata = {
    "kernelspec": {"display_name": "Python 3", "language": "python", "name": "python3"},
    "language_info": {"name": "python", "version": "3.12.10"},
}

OUT_DIR.mkdir(parents=True, exist_ok=True)
nbf.write(nb, NB_PATH)

n_md = sum(1 for c in cells if c.cell_type == "markdown")
n_code = sum(1 for c in cells if c.cell_type == "code")
print(f"Written: {NB_PATH}")
print(f"  {len(cells)} cells  ({n_md} markdown, {n_code} code)")

if "--run" in sys.argv:
    print("\nExecuting the notebook (this takes several minutes)...")
    result = subprocess.run(
        [sys.executable, "-m", "jupyter", "nbconvert", "--to", "notebook",
         "--execute", "--inplace",
         "--ExecutePreprocessor.timeout=2400",
         str(NB_PATH)],
        cwd=str(OUT_DIR))
    if result.returncode != 0:
        sys.exit("Notebook execution FAILED - see the error above.")
    print("Executed. Outputs are saved inside the notebook.")
