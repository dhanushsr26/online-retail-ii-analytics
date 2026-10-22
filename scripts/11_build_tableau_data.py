"""
Step 11a - Tableau data: flatten the star schema into one table.

Power BI keeps Fact_Sales and five dimensions as separate tables joined by
relationships. Tableau Public reads flat files best, and every dimension here
is many-to-one from the fact, so a left join onto the fact adds attribute
columns without adding or losing a single row. The result is the same star
schema, pre-joined:

    Fact_Sales  (1,028,743 rows, the grain)
      + Dim_Product   on StockCode    -> Description, ABC_Class, PriceBand
      + Dim_Country   on Country      -> Region, CountryISO (for the map)
      + Dim_Customer  on CustomerKey  -> Segment, CRM_*, ChurnStatus, RFM, CLV
      + Dim_Hour      on Hour         -> HourLabel, DayPart

Dim_Date is not joined: Tableau derives Year/Quarter/Month/Weekday from
InvoiceDate itself, which is what a date dimension is for in Power BI.

Customer-grain values (CLV, Frequency, ...) repeat on every line a customer
bought. The workbook therefore never averages them over rows: it reduces them
to one value per customer first with {FIXED [CustomerKey] : MIN(...)}.

Also writes 07_tableau/expected_values.json - the numbers each dashboard must
show, computed here in pandas, so the workbook can be checked against them.

Run:  python 11_build_tableau_data.py
"""
import json
import pathlib

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
CLEAN = ROOT / "02_clean"
OUT = ROOT / "07_tableau"
DATA = OUT / "data"


def read(name, **kw):
    # The cleaned CSVs carry a UTF-8 byte-order mark; utf-8-sig strips it.
    return pd.read_csv(CLEAN / f"{name}.csv", encoding="utf-8-sig", **kw)


def write_hyper(df, csv_path, hyper_path):
    """
    Tableau Public refuses a live connection to a file: every data source has
    to be an extract. A .hyper file is that extract - Tableau's own columnar
    database. Hyper bulk-loads the CSV just written, so the extract and the CSV
    are the same rows by construction.
    """
    from tableauhyperapi import (Connection, CreateMode, HyperProcess,
                                 SqlType, TableDefinition, TableName,
                                 Telemetry, escape_string_literal)

    def sql_type(col):
        if col == "InvoiceDate":
            return SqlType.timestamp()
        kind = df[col].dtype.kind
        return {"b": SqlType.bool(), "i": SqlType.big_int(),
                "f": SqlType.double()}.get(kind, SqlType.text())

    # Schema "Extract", table "Extract" is the layout Tableau itself writes for
    # a single-table extract, and what the workbook's <extract> points at.
    table = TableDefinition(TableName("Extract", "Extract"),
                            [TableDefinition.Column(c, sql_type(c))
                             for c in df.columns])
    with HyperProcess(Telemetry.DO_NOT_SEND_USAGE_DATA_TO_TABLEAU) as hyper:
        with Connection(hyper.endpoint, str(hyper_path),
                        CreateMode.CREATE_AND_REPLACE) as conn:
            conn.catalog.create_schema("Extract")
            conn.catalog.create_table(table)
            n = conn.execute_command(
                f"COPY {table.table_name} FROM "
                f"{escape_string_literal(str(csv_path))} "
                f"WITH (format csv, header true, NULL '')")
    assert n == len(df), f"extract has {n} rows, expected {len(df)}"
    print(f"Wrote {hyper_path}  {n:,} rows "
          f"({hyper_path.stat().st_size / 1e6:.0f} MB)")


def main():
    DATA.mkdir(parents=True, exist_ok=True)

    fact = read("Fact_Sales", dtype={"Invoice": str, "StockCode": str},
                usecols=["Invoice", "StockCode", "CustomerKey", "Country",
                         "InvoiceDate", "Hour", "Quantity", "Price", "Revenue",
                         "IsCancelled", "IsProduct", "TransactionType",
                         "Is_Outlier"])
    prod = read("Dim_Product", dtype={"StockCode": str},
                usecols=["StockCode", "Description", "ABC_Class", "PriceBand"])
    ctry = read("Dim_Country", usecols=["Country", "CountryISO", "Region"])
    cust = read("Dim_Customer",
                usecols=["CustomerKey", "Segment", "CRM_Strategy", "CRM_Action",
                         "ChurnStatus", "Frequency", "Monetary", "CLV",
                         "PurchaseFrequencyPerYear"])
    hour = read("Dim_Hour", usecols=["Hour", "HourLabel", "DayPart", "DayPartNo"])

    n = len(fact)
    flat = (fact.merge(prod, on="StockCode", how="left", validate="m:1")
                .merge(ctry, on="Country", how="left", validate="m:1")
                .merge(cust, on="CustomerKey", how="left", validate="m:1")
                .merge(hour, on="Hour", how="left", validate="m:1"))
    assert len(flat) == n, "a join changed the row count"
    for col in ["Description", "Region", "Segment", "DayPart"]:
        missing = flat[col].isna().sum()
        assert missing == 0, f"{missing} fact rows found no {col}"

    # Plain UTF-8 without a BOM: Tableau's text reader would otherwise fold the
    # mark into the first column's name.
    path = DATA / "Sales.csv"
    flat.to_csv(path, index=False, encoding="utf-8")
    print(f"Wrote {path}  {len(flat):,} rows x {flat.shape[1]} columns")
    write_hyper(flat, path, DATA / "Sales.hyper")

    # ------------------------------------------------ expected dashboard values
    f = flat
    d = pd.to_datetime(f["InvoiceDate"])
    rev = f["Revenue"].sum()
    orders = f["Invoice"].nunique()
    known = f[f["CustomerKey"] != -1]
    gross = f.loc[f["TransactionType"] == "Sale", "Revenue"].sum()
    returns = -f.loc[f["Quantity"] < 0, "Revenue"].sum()
    c = cust[cust["CustomerKey"] != -1]
    scored = (cust["ChurnStatus"] != "Not scored").sum()
    churned = (cust["ChurnStatus"] == "Churned").sum()

    last = d.max()
    cy = f.loc[(d.dt.year == last.year) & (d <= last), "Revenue"].sum()
    py_end = last - pd.DateOffset(years=1)
    py = f.loc[(d.dt.year == last.year - 1) & (d <= py_end), "Revenue"].sum()

    exp = {
        "Total Revenue": round(rev, 2),
        "Total Orders": int(orders),
        "Total Customers": int(known["CustomerKey"].nunique()),
        "Average Order Value": round(rev / orders, 2),
        "Units Sold": int(f["Quantity"].sum()),
        "Return Rate %": round(returns / gross * 100, 2),
        "Average CLV": round(c["CLV"].mean(), 2),
        "Churn Rate %": round(churned / scored * 100, 2),
        "Retention Rate %": round(100 - churned / scored * 100, 2),
        "Purchase Frequency per Year": round(c["PurchaseFrequencyPerYear"].mean(), 2),
        "Guest Revenue %": round(f.loc[f["CustomerKey"] == -1, "Revenue"].sum() / rev * 100, 2),
        "Products Sold": int(f["StockCode"].nunique()),
        "Class A Revenue %": round(f.loc[f["ABC_Class"] == "A", "Revenue"].sum() / rev * 100, 2),
        "Product Return Rate %": round(-f.loc[f["Quantity"] < 0, "Quantity"].sum()
                                       / f.loc[f["Quantity"] > 0, "Quantity"].sum() * 100, 2),
        "Non-Product Revenue": round(f.loc[~f["IsProduct"], "Revenue"].sum(), 2),
        "Revenue YoY % (latest year to date)": round((cy - py) / py * 100, 2),
        "Average Basket Size": round(f["Quantity"].sum() / orders, 2),
        "Cancellation Rate %": round(f.loc[f["IsCancelled"], "Invoice"].nunique() / orders * 100, 2),
        "Outlier Rows": int(f["Is_Outlier"].sum()),
        "Outlier Revenue Share %": round(f.loc[f["Is_Outlier"], "Revenue"].sum() / rev * 100, 2),
        "Revenue by Region": f.groupby("Region")["Revenue"].sum().round(2).to_dict(),
        "Revenue by Year": f.groupby(d.dt.year)["Revenue"].sum().round(2)
                            .rename(index=str).to_dict(),
        "Revenue by DayPart": f.groupby("DayPart")["Revenue"].sum().round(2).to_dict(),
        "Revenue by Segment": f.groupby("Segment")["Revenue"].sum().round(2).to_dict(),
        "Customers by CRM strategy": c.groupby("CRM_Strategy").size().to_dict(),
        "Top 5 products": f.groupby("Description")["Revenue"].sum()
                           .nlargest(5).round(2).to_dict(),
    }
    (OUT / "expected_values.json").write_text(json.dumps(exp, indent=2),
                                             encoding="utf-8")
    for k, v in exp.items():
        print(f"  {k:<38} {v if not isinstance(v, dict) else '{...}'}")


if __name__ == "__main__":
    main()
