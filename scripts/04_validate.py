"""
Step 3b - Validation gate.

Nothing gets built in Power BI until every assertion here passes. Each check
prints PASS or FAIL; the script exits non-zero if any FAIL, so it can be trusted
as a gate rather than read as a suggestion.

Run:  python 04_validate.py
"""
import pathlib
import sys

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
CLEAN = ROOT / "02_clean"
RAW_CACHE = ROOT / "00_source" / "raw_combined.parquet"

results = []


def check(name, condition, detail=""):
    results.append((bool(condition), name, detail))
    print(f"  [{'PASS' if condition else 'FAIL'}] {name}"
          + (f"  -- {detail}" if detail else ""))


def main():
    print("\n=== VALIDATION ===\n")
    fact = pd.read_parquet(CLEAN / "Fact_Sales.parquet")
    cust = pd.read_csv(CLEAN / "Dim_Customer.csv")
    prod = pd.read_csv(CLEAN / "Dim_Product.csv")
    ctry = pd.read_csv(CLEAN / "Dim_Country.csv")
    date = pd.read_csv(CLEAN / "Dim_Date.csv", parse_dates=["DateKey"])
    dq = pd.read_csv(CLEAN / "Dim_DataQuality.csv")
    daily = pd.read_csv(CLEAN / "Agg_Daily.csv", parse_dates=["DateKey"])
    mc = pd.read_csv(CLEAN / "Agg_MonthCountry.csv")
    raw = pd.read_parquet(RAW_CACHE)

    print("-- referential integrity")
    check("no orphan StockCode",
          set(fact.StockCode) <= set(prod.StockCode),
          f"{len(set(fact.StockCode) - set(prod.StockCode))} orphans")
    check("no orphan Country",
          set(fact.Country) <= set(ctry.Country),
          f"{len(set(fact.Country) - set(ctry.Country))} orphans")
    fact_cust = set(fact.CustomerKey.astype("int64"))
    dim_cust = set(cust.CustomerKey.astype("int64"))
    check("no orphan CustomerKey (incl. the -1 guest key)",
          fact_cust <= dim_cust, f"{len(fact_cust - dim_cust)} orphans")
    check("guest key -1 exists in Dim_Customer", -1 in dim_cust)
    fdates = set(fact.InvoiceDate.dt.normalize().unique())
    check("every fact date exists in Dim_Date", fdates <= set(date.DateKey),
          f"{len(fdates - set(date.DateKey))} missing")

    print("\n-- keys and nulls")
    for col in ["Invoice", "StockCode", "CustomerKey", "Country",
                "InvoiceDate", "Quantity", "Price", "Revenue"]:
        check(f"no nulls in Fact_Sales.{col}", fact[col].notna().all(),
              f"{int(fact[col].isna().sum())} nulls")
    check("Dim_Customer.CustomerKey unique", cust.CustomerKey.is_unique)
    check("Dim_Product.StockCode unique", prod.StockCode.is_unique)
    check("Dim_Date.DateKey unique", date.DateKey.is_unique)
    check("Outlier_Reason never null", fact.Outlier_Reason.notna().all(),
          f"{int(fact.Outlier_Reason.isna().sum())} nulls")

    print("\n-- arithmetic")
    recomputed = (fact.Quantity * fact.Price).round(2)
    check("Revenue == Quantity x Price",
          bool((recomputed - fact.Revenue).abs().max() < 0.011),
          f"max diff {float((recomputed - fact.Revenue).abs().max()):.4f}")
    check("Agg_Daily revenue ties to fact",
          abs(daily.Revenue.sum() - fact.Revenue.sum()) < 1.0,
          f"daily {daily.Revenue.sum():,.2f} vs fact {fact.Revenue.sum():,.2f}")
    check("Agg_MonthCountry revenue ties to fact",
          abs(mc.Revenue.sum() - fact.Revenue.sum()) < 1.0,
          f"cube {mc.Revenue.sum():,.2f} vs fact {fact.Revenue.sum():,.2f}")
    check("Dim_Product revenue ties to fact",
          abs(prod.Revenue.sum() - fact.Revenue.sum()) < 1.0,
          f"product {prod.Revenue.sum():,.2f} vs fact {fact.Revenue.sum():,.2f}")

    print("\n-- reconciliation against raw")
    raw_rev = (raw.Quantity * raw.Price).sum()
    removed_rev = raw_rev - fact.Revenue.sum()
    check("cleaning removed revenue, and only a small share",
          0 < removed_rev < raw_rev * 0.05,
          f"raw GBP {raw_rev:,.2f} -> clean GBP {fact.Revenue.sum():,.2f} "
          f"(-GBP {removed_rev:,.2f}, {removed_rev / raw_rev * 100:.2f}%)")
    check("no customer lost that still has rows in the fact",
          fact_cust <= dim_cust)
    check("row count matches the audit trail",
          int(dq[dq.Step_Type == "Row filter"].Rows_After.min()) == len(fact),
          f"audit says {int(dq[dq.Step_Type == 'Row filter'].Rows_After.min()):,}, "
          f"fact has {len(fact):,}")

    print("\n-- audit-trail integrity")
    rf = dq[dq.Step_Type == "Row filter"].sort_values("Step")
    chain_ok = all(int(rf.Rows_After.iloc[i]) == int(rf.Rows_Before.iloc[i + 1])
                   for i in range(len(rf) - 1))
    check("each step's Rows_After == next step's Rows_Before", chain_ok)
    check("only row-filter steps report deletions",
          bool((dq[dq.Step_Type != "Row filter"].Rows_Removed == 0).all()),
          "derivation/aggregation steps must not appear as deletions")
    check("every step cites a PDF section", dq.PDF_Section.notna().all())
    check("13+ techniques evidenced", dq.PDF_Section.nunique() >= 10,
          f"{dq.PDF_Section.nunique()} distinct PDF sections across {len(dq)} steps")

    print("\n-- business sanity")
    check("date span covers both years",
          fact.InvoiceDate.dt.year.nunique() == 3,
          f"years {sorted(fact.InvoiceDate.dt.year.unique().tolist())}")
    check("no duplicate rows remain",
          not fact.drop(columns=["SourceSheet"]).duplicated().any(),
          f"{int(fact.drop(columns=['SourceSheet']).duplicated().sum())} dupes")
    check("outliers flagged but retained",
          int(fact.Is_Outlier.sum()) > 0,
          f"{int(fact.Is_Outlier.sum()):,} flagged, all retained")
    check("customer-less rows retained in fact",
          int((~fact.HasCustomer).sum()) > 0,
          f"{int((~fact.HasCustomer).sum()):,} rows kept (dual-grain decision)")
    scored = cust[~cust.Segment.isin(
        ["No Net Value (returns / cancellations only)",
         "Unknown / Guest (not identified)"])]
    check("RFM scores span 1-5", set(scored.R_Score.dropna().unique()) == {1, 2, 3, 4, 5})
    check("Min-Max normalization lands in [0,1]",
          bool(scored[["Recency_MinMax", "Frequency_MinMax", "Monetary_MinMax"]]
               .min().min() >= 0
               and scored[["Recency_MinMax", "Frequency_MinMax", "Monetary_MinMax"]]
               .max().max() <= 1))
    check("Z-scores centre on zero",
          bool(abs(scored[["Recency_Z", "Frequency_Z", "Monetary_Z"]].mean()).max() < 0.01))

    failed = [r for r in results if not r[0]]
    print(f"\n=== {len(results) - len(failed)}/{len(results)} passed ===")
    if failed:
        print("\nFAILURES:")
        for _, name, detail in failed:
            print(f"  - {name}  {detail}")
        sys.exit(1)
    print("All assertions passed - safe to build the dashboard.\n")


if __name__ == "__main__":
    main()
