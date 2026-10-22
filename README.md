# Online Retail II: Data Preprocessing and BI Dashboards (Power BI + Tableau)

M.Sc Business Analytics project (Amrita Vishwa Vidyapeetham, Coimbatore, 2026).

A UK online gift retailer's two years of transactions, cleaned with **13 documented preprocessing
techniques**, modelled as a **star schema**, and presented as a **4-page Power BI dashboard** that was
then rebuilt in **Tableau** with every KPI checked against pandas.

**Data:** [UCI Online Retail II](https://archive.ics.uci.edu/dataset/502/online+retail+ii),
1,067,371 transaction lines, Dec 2009 to Dec 2011 (not included here; `scripts/01_download.py` fetches it).

## What was done

| Stage | Detail |
|---|---|
| Profiling | Missingness, duplicates, overlaps and outliers measured before any change (`reports/profile_before.md`) |
| Cleaning | 18 steps / 13 techniques: cross-sheet overlap removal, duplicate removal (exact and business-key), MAR missing-value analysis, test-record removal, z-score + IQR outlier **flagging** (1,751 rows flagged, none deleted), attribute construction, normalisation |
| Audit trail | Every step logs rows before/after and its rationale into `Dim_DataQuality` (`data/Dim_DataQuality.csv`), so each dashboard number traces back to a cleaning step |
| Customer analytics | RFM quintile scoring, 10 named segments, CRM strategy per segment, churn state and CLV per customer |
| Modelling | Flat 29-column table -> 17-column fact + date, hour, customer, product and country dimensions; data-cube aggregates for fast pages |
| Power BI | 4 pages: Executive Overview, Customer Intelligence, Product & Returns, Time & Seasonality (PBIP project, DAX measures, named Power Query steps) |
| Tableau | Same four dashboards in one self-contained `.twbx` (Hyper extract); every KPI matched the pandas values in `tableau/expected_values.json` |

**Result:** 1,028,743 clean rows (-3.62%), net revenue GBP 18.85 M, 5,941 customers, 4,776 stock codes.

## Repository layout

```
scripts/    01..12 numbered Python pipeline (download, profile, clean, validate, build PBIP, build Tableau)
powerbi/    Power BI project (.pbip), DAX measures, Power Query M
tableau/    RetailAnalytics.twbx (open in Tableau Desktop / Public)
notebook/   Preprocessing_Walkthrough.ipynb: every cleaning step, one cell each, with output
reports/    Preprocessing report (PDF) and before/after data profiles
data/       Dim_DataQuality.csv: the audit trail
docs/       RUNNING.md: how to open and refresh the dashboards
```

## Tools
Python (pandas, NumPy), Power BI (Power Query M, DAX), Tableau (Hyper extracts), Excel.

Author: Dhanush S R
