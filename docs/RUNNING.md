# Online Retail II — Preprocessing & Power BI Dashboard

Dataset: [UCI Online Retail II](https://archive.ics.uci.edu/dataset/502/online+retail+ii) —
1,067,371 raw transaction lines, Dec 2009 – Dec 2011, a UK online gift retailer.

## Opening the dashboard

Open `03_powerbi/RetailAnalytics.pbip` in **Power BI Desktop**.

The model loads from the raw workbook, so a first refresh takes about
**5–6 minutes** and peaks near 900 MB. That is the point: `Transform data`
shows the whole cleaning pipeline as named Applied Steps rather than a
finished table appearing from nowhere.

### If it cannot find the data

Every source path in the model is built from one value, the `SourceFolder`
parameter, recorded as `D:\Semester 3\Business Analytics\Project`. If the
project now lives anywhere else, point it there once:

> **Home → Transform data → Manage parameters → `SourceFolder`**

Set it to the folder containing this README, click **Close & Apply**, and
refresh. Nothing else needs editing — the workbook and every CSV path is built
from that one value.

(Power Query has no way to ask where the `.pbip` file itself is, so an absolute
path has to live somewhere. Auto-detecting it by probing the drives was tried
and removed: it works, but it makes Power BI demand a privacy level for every
drive root in a modal dialog the moment the file opens, which is a worse thing
to hand someone than a single parameter.)

**Keep the folder whole.** `03_powerbi` alone will not open — the model reads
the workbook in `00_source` and three dimensions from `02_clean`.

## The Tableau version

`07_tableau/RetailAnalytics.twbx` is the same four dashboards in Tableau, as
one 10 MB file that opens anywhere in **Tableau Desktop** or **Tableau Public**.
The data is packaged inside it as a Hyper extract, so there is no path to edit.

| File | What it is |
|---|---|
| `RetailAnalytics.twbx` | The one to open, send or submit |
| `RetailAnalytics.twb` | Same workbook, reading `data/Sales.hyper` by absolute path (this machine only) |
| `data/Sales.csv`, `data/Sales.hyper` | The star schema flattened: fact rows with their dimension attributes joined on |
| `expected_values.json` | The numbers every dashboard must show, computed in pandas |

Where the Tableau and Power BI versions differ:

- **Star schema vs one table.** Tableau Public reads one flat table best, and
  every dimension is many-to-one from the fact, so joining them on adds columns
  without adding or losing a row (`11_build_tableau_data.py` asserts this).
  Customer-grain values such as CLV repeat on every line a customer bought, so
  the workbook divides each by that customer's line count before summing. That
  counts each customer once, which is what `AVERAGEX(Dim_Customer, ...)` does.
- **YoY growth** is 2011 up to 9 December against 2010 over the same days
  (+3.5%). Power BI's `SAMEPERIODLASTYEAR` card, with no date filter, compares
  the whole two years against one shifted year and shows 90.9%.
- **Year and Region** are parameters driving a context filter, so the Top 10
  products re-ranks within the selection.

`11_build_tableau_data.py` needs `pip install tableauhyperapi`.

## What is in here

| Folder | Contents |
|---|---|
| `00_source/` | The raw `.xlsx` as downloaded, plus a Parquet copy for the fast path |
| `01_scripts/` | The pipeline, numbered in run order |
| `02_clean/` | Cleaned fact and dimensions |
| `03_powerbi/` | The PBIP project, plus the M and DAX as readable text |
| `04_report/` | `Preprocessing_Report.pdf` and `Storytelling_Script.pdf` |
| `05_evidence/` | Data profiles before and after cleaning, and screenshots |
| `06_deliverable/` | The notebook walkthrough and the single cleaned Excel file |
| `07_tableau/` | The Tableau version of the dashboard (`.twbx`) and its flat data |

## Showing the preprocessing to a person

`06_deliverable/Preprocessing_Walkthrough.ipynb` is the version to open in front
of someone. It takes the raw workbook through all eighteen steps, one step per
cell, each with its explanation above it and its result printed below it. The
outputs are already saved in the file, so it can be read without running
anything.

```
jupyter notebook "06_deliverable/Preprocessing_Walkthrough.ipynb"
```

Running it end to end takes roughly ten minutes — most of it reading a million
rows out of Excel and writing them back again — and rewrites
`06_deliverable/OnlineRetailII_Cleaned.xlsx`.

That workbook is the answer to "where is the cleaned data": **one file**, three
sheets — `Cleaned_Data`, `Audit_Trail`, `Summary`. The eleven CSVs in `02_clean/`
are the same data reshaped into a star schema because that is what Power BI
models efficiently, not a different dataset.

One thing worth knowing before promising anyone a single Excel of a large
dataset: a sheet holds at most **1,048,576** rows. The cleaned data is 1,028,743,
so it fits with about 20,000 rows to spare.

## Rebuilding from scratch

```
cd 01_scripts
python 01_download.py                  # fetch the dataset
python 02_profile_raw.py               # profile before cleaning
python 03_clean.py                     # 18 steps, 13 techniques
python 04_validate.py                  # 35 assertions - must pass before building
python 05_build_pbip.py --excel        # the model, reading the raw workbook
python 06_build_report.py              # the 4 report pages, 48 visuals
python 07_build_preprocessing_report.py
python 08_build_storytelling_script.py
python 09_check_generated_dims.py      # the M-generated dimensions vs the CSVs
python 10_build_notebook.py --run      # the walkthrough notebook + single Excel
python 11_build_tableau_data.py         # Tableau: flat table + Hyper extract
python 12_build_tableau_workbook.py     # Tableau: the 4 dashboards, .twb + .twbx
```

`05_build_pbip.py` resets the report folder, so **always run `06` after it**.

Dropping `--excel` on step 5 points the model at Parquet instead: refresh
becomes near-instant, which is easier while editing visuals, but Applied Steps
then shows a single load rather than the pipeline. Same data either way; build
with `--excel` for submission.

## Where each table comes from

| Table | Source | Why |
|---|---|---|
| `Fact_Sales` | raw `.xlsx` | Carries the full cleaning pipeline as named Applied Steps |
| `Dim_Date` | **generated in M** | Pure calendar logic — no file needed |
| `Dim_Hour` | **generated in M** | Pure banding logic — no file needed |
| `Dim_Customer` | `02_clean/*.csv` | An aggregation of the cleaned fact (RFM, CLV, CRM) |
| `Dim_Product` | `02_clean/*.csv` | An aggregation of the cleaned fact (ABC, price bands) |
| `Dim_Country` | `02_clean/*.csv` | An aggregation of the cleaned fact |

The three aggregate dimensions load from CSV deliberately. Power Query
evaluates each table independently with no shared cache, so a dimension that
referenced `Fact_Sales` would re-run the entire million-row workbook pipeline
again — six passes over the `.xlsx` instead of one, roughly half an hour of
refresh. Each of their queries opens with a comment recording what the table is
and which techniques produced it.

`09_check_generated_dims.py` re-implements the M for the two generated
dimensions and compares all 16 columns against `03_clean.py`'s output, so the
two implementations cannot drift apart unnoticed.

## Headline numbers

Every figure in both PDFs is computed from the cleaned data at build time, so
the reports and the dashboard cannot disagree.

| | |
|---|---|
| Rows after cleaning | 1,028,743 |
| Revenue | £18,854,778.33 |
| Orders | 49,337 |
| Identified customers | 5,940 |
| Units sold | 10,483,663 |
| Average order value | £382.16 |
| Return rate | 7.2% |
| Double-counted revenue removed | £431,716.87 |
