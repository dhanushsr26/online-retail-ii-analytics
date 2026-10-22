"""
Step 6 - Generate a Power BI project (PBIP) from the cleaned data.

Power BI Desktop can open a folder of TMDL + JSON definition files, so the
semantic model, its 50-odd measures and the report pages can be authored as
files rather than clicked together by hand.

Column definitions are read from the actual data files, never hard-coded, so
the model cannot drift from what 03_clean.py produced.

Run:  python 05_build_pbip.py            (fast path: parquet/CSV sources)
      python 05_build_pbip.py --excel    (submission: full M from raw .xlsx)
"""
import json
import pathlib
import shutil
import sys
import uuid

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
CLEAN = ROOT / "02_clean"
PBI = ROOT / "03_powerbi"
NAME = "RetailAnalytics"
MODEL_DIR = PBI / f"{NAME}.SemanticModel"
REPORT_DIR = PBI / f"{NAME}.Report"

USE_EXCEL = "--excel" in sys.argv

# ---------------------------------------------------------------- type mapping
TMDL_TYPE = {
    "int64": "int64", "Int64": "int64", "int32": "int64",
    "float64": "double", "Float64": "double",
    "bool": "boolean", "boolean": "boolean",
    "datetime64[ns]": "dateTime", "datetime64[us]": "dateTime",
    "object": "string", "string": "string", "str": "string",
}
SUMMARIZE_NONE = {"string", "boolean", "dateTime"}

# Text labels sort alphabetically unless told otherwise, which puts "Apr 2010"
# before "Dec 2009" on every month axis. Each entry maps a label column to the
# numeric column that gives its true order.
SORT_BY = {
    ("Dim_Date", "MonthYearLabel"): "MonthYear",
    ("Dim_Date", "MonthName"): "Month",
    ("Dim_Date", "Weekday"): "WeekdayNo",
    ("Dim_Date", "QuarterLabel"): "MonthYear",
    ("Dim_Hour", "DayPart"): "DayPartNo",
    ("Dim_Hour", "HourLabel"): "Hour",
    ("Agg_Histograms", "Bin"): "BinOrder",
}

GBP = r"\£#,0.00;(\£#,0.00);\£#,0.00"
GBP0 = r"\£#,0;(\£#,0);\£#,0"
PCT1 = '0.0"%"'
INT = "#,0"


def m_str(value) -> str:
    """
    Quote a string for M. M has no backslash escape - a Windows path is written
    with single backslashes - so json.dumps would corrupt it into \\\\. Only a
    double quote needs escaping, by doubling it.
    """
    return '"' + str(value).replace('"', '""') + '"'


def tmdl_type(dtype) -> str:
    return TMDL_TYPE.get(str(dtype), "string")


def esc(name: str) -> str:
    """Quote a TMDL identifier if it is not a bare word."""
    return name if name.replace("_", "").isalnum() and not name[0].isdigit() \
        else f"'{name}'"


# ============================================================ table definitions
# Where each CSV-backed dimension actually comes from. These three are
# aggregations of the cleaned fact, and they load from CSV rather than being
# re-derived in M on purpose. Power Query evaluates every table's partition
# independently, with no shared cache between them, so a dimension that
# referenced Fact_Sales would re-run the whole million-row workbook pipeline
# again. Five dimensions doing that is six passes over the .xlsx instead of
# one: roughly half an hour of refresh, and the memory profile that killed
# Power BI once already. The header comment and the named steps below keep the
# lineage readable in Applied Steps even though the immediate source is a file.
DIM_PROVENANCE = {
    "Dim_Customer": (
        "one row per customer, aggregated from the cleaned fact",
        ["RFM scoring - quintile discretization of Recency/Frequency/Monetary",
         "Min-Max normalization and Z-score standardization of the same three",
         "Segment, CRM strategy, churn status and CLV derived from the scores"],
    ),
    "Dim_Product": (
        "one row per stock code, aggregated from the cleaned fact",
        ["Item type classification - product / postage / fee / adjustment",
         "Price banding - discretization of average unit price",
         "ABC classification - cumulative revenue share, Pareto cut-offs"],
    ),
    "Dim_Country": (
        "one row per country, aggregated from the cleaned fact",
        ["Unit Conversion - canonical country name",
         "Data Integration - ISO code and region lookup"],
    ),
}


def csv_partition(table: str) -> str:
    """
    M expression reading one cleaned dimension, headed by a comment recording
    what the file is and how to regenerate it, so Applied Steps shows
    provenance rather than an anonymous CSV read.
    """
    summary, techniques = DIM_PROVENANCE.get(table, ("cleaned dimension", []))
    head = [
        "    // ================================================================",
        f"    // {table} - {summary}.",
        "    // Derived by 01_scripts/03_clean.py in the same cleaning run that",
        "    // produces Fact_Sales, then written out as CSV. Techniques applied",
        "    // upstream, per BA-6 Preprocessing Techniques.pdf:",
    ]
    head += [f"    //   - {t}" for t in techniques]
    head += [
        "    // Regenerate with:  python 01_scripts/03_clean.py",
        "    // Verify with:      python 01_scripts/04_validate.py",
        "    // ================================================================",
        "",
    ]
    return (
        "let\n"
        + "\n".join(head) + "\n"
        + f'    Source = Csv.Document(File.Contents(SourceFolder & '
        f'"\\02_clean\\{table}.csv"), [Delimiter = ",", Encoding = 65001, '
        f'QuoteStyle = QuoteStyle.Csv]),\n'
        '    #"Promoted Headers" = Table.PromoteHeaders(Source, '
        '[PromoteAllScalars = true])\n'
        'in\n'
        '    #"Promoted Headers"'
    )


def parquet_partition(table: str) -> str:
    return (
        f'let\n'
        f'    Source = Parquet.Document(File.Contents(SourceFolder & '
        f'"\\02_clean\\{table}.parquet"))\n'
        f'in\n'
        f'    Source'
    )


def excel_fact_partition() -> str:
    """The full cleaning pipeline in M, straight from the raw workbook."""
    return (ROOT / "01_scripts" / "_fact_m.txt").read_text(encoding="utf-8")


def date_dim_partition(first: str, last: str) -> str:
    """
    Dim_Date, generated entirely in M - no source file at all.

    A calendar table is pure deterministic logic, so authoring it in Power
    Query removes a file dependency and costs nothing at refresh. The span is
    inlined as constants for the same reason the outlier thresholds are: it is
    a settled fact about the data, not something worth recomputing over a
    million rows on every refresh. 03_clean.py derives it, 04_validate.py
    re-checks it, and this builder reads it back from the cleaned dimension so
    the two cannot drift.
    """
    y1, m1, d1 = (int(x) for x in first.split("-"))
    y2, m2, d2 = (int(x) for x in last.split("-"))
    return f"""let
    // ================================================================
    // Dim_Date - generated, not loaded. One row per calendar day across
    // the fact's span, so time intelligence has an unbroken date table.
    // The span below is the min and max InvoiceDate of the cleaned fact,
    // reported by 03_clean.py and re-checked by 04_validate.py.
    // ================================================================

    StartDate = #date({y1}, {m1}, {d1}),
    EndDate = #date({y2}, {m2}, {d2}),
    DayCount = Duration.Days(EndDate - StartDate) + 1,

    #"Generate Calendar" = Table.FromList(
        List.Dates(StartDate, DayCount, #duration(1, 0, 0, 0)),
        Splitter.SplitByNothing(), {{"DateKey"}}),

    // Held as a date while the parts are derived - Date.ToText refuses a
    // datetime - then cast to datetime at the end to match the fact's key.
    #"Set Data Types" = Table.TransformColumnTypes(
        #"Generate Calendar", {{{{"DateKey", type date}}}}),

    #"Derive Year" = Table.AddColumn(
        #"Set Data Types", "Year", each Date.Year([DateKey]), Int64.Type),
    #"Derive Quarter" = Table.AddColumn(
        #"Derive Year", "Quarter", each Date.QuarterOfYear([DateKey]), Int64.Type),
    #"Derive Quarter Label" = Table.AddColumn(
        #"Derive Quarter", "QuarterLabel",
        each "Q" & Text.From([Quarter]) & " " & Text.From([Year]), type text),
    #"Derive Month" = Table.AddColumn(
        #"Derive Quarter Label", "Month", each Date.Month([DateKey]), Int64.Type),
    #"Derive Month Name" = Table.AddColumn(
        #"Derive Month", "MonthName",
        each Date.ToText([DateKey], "MMM", "en-GB"), type text),
    #"Derive Month Year" = Table.AddColumn(
        #"Derive Month Name", "MonthYear",
        each Date.ToText([DateKey], "yyyy-MM", "en-GB"), type text),
    #"Derive Month Year Label" = Table.AddColumn(
        #"Derive Month Year", "MonthYearLabel",
        each Date.ToText([DateKey], "MMM yyyy", "en-GB"), type text),
    #"Derive Weekday" = Table.AddColumn(
        #"Derive Month Year Label", "Weekday",
        each Date.DayOfWeekName([DateKey], "en-GB"), type text),
    #"Derive Weekday Number" = Table.AddColumn(
        #"Derive Weekday", "WeekdayNo",
        each Date.DayOfWeek([DateKey], Day.Monday), Int64.Type),
    #"Flag Weekend" = Table.AddColumn(
        #"Derive Weekday Number", "IsWeekend",
        each [WeekdayNo] >= 5, type logical),

    // ISO 8601 week: the week owning the Thursday of that Monday-week. This
    // reproduces pandas .isocalendar().week, which is what 03_clean.py used.
    #"Derive ISO Week" = Table.AddColumn(
        #"Flag Weekend", "WeekOfYear",
        each let Thu = Date.AddDays(Date.StartOfWeek([DateKey], Day.Monday), 3)
             in Number.RoundDown(
                 Duration.Days(Thu - Date.StartOfYear(Thu)) / 7) + 1,
        Int64.Type),

    #"Derive Day Of Year" = Table.AddColumn(
        #"Derive ISO Week", "DayOfYear",
        each Date.DayOfYear([DateKey]), Int64.Type),

    #"Set Date Key Type" = Table.TransformColumnTypes(
        #"Derive Day Of Year", {{{{"DateKey", type datetime}}}})
in
    #"Set Date Key Type\""""


def hour_dim_partition() -> str:
    """Dim_Hour, generated in M. Bands match the cut points in 03_clean.py."""
    return '''let
    // ================================================================
    // Dim_Hour - generated, not loaded. The 24 hours of the trading day
    // with the day-part banding used on the Time page. The bands match
    // the cut points in 03_clean.py exactly: 0-11, 12-16, 17-20, 21-23.
    // ================================================================

    #"Generate Hours" = Table.FromList(
        List.Numbers(0, 24, 1), Splitter.SplitByNothing(), {"Hour"}),

    #"Set Data Types" = Table.TransformColumnTypes(
        #"Generate Hours", {{"Hour", Int64.Type}}),

    #"Derive Hour Label" = Table.AddColumn(
        #"Set Data Types", "HourLabel",
        each Text.PadStart(Text.From([Hour]), 2, "0") & ":00", type text),

    // Discretization / Binning - BA-6 Preprocessing Techniques.pdf
    #"Binning - Day Part" = Table.AddColumn(
        #"Derive Hour Label", "DayPart",
        each if [Hour] <= 11 then "Morning"
             else if [Hour] <= 16 then "Afternoon"
             else if [Hour] <= 20 then "Evening"
             else "Night", type text),

    // One sort value per band. Sorting DayPart by Hour would make Power BI
    // group the matrix by hour and label every column with its band name.
    #"Day Part Order" = Table.AddColumn(
        #"Binning - Day Part", "DayPartNo",
        each if [Hour] <= 11 then 0
             else if [Hour] <= 16 then 1
             else if [Hour] <= 20 then 2
             else 3, Int64.Type)
in
    #"Day Part Order"'''


def typed_m(table: str, frame: pd.DataFrame, source_m: str) -> str:
    """Append an explicit type-change step so Power BI does not guess."""
    casts = []
    for col, dtype in frame.dtypes.items():
        t = tmdl_type(dtype)
        m_type = {"int64": "Int64.Type", "double": "type number",
                  "boolean": "type logical", "dateTime": "type datetime",
                  "string": "type text"}[t]
        casts.append(f'{{"{col}", {m_type}}}')
    cast_str = ", ".join(casts)
    # The previous step needs a trailing comma before the one being appended,
    # otherwise M reports "Token ',' expected" and the whole project is refused.
    body = source_m.rsplit("in\n", 1)[0].rstrip()
    last = source_m.rsplit("in\n", 1)[1].strip()
    return (f'{body},\n    #"Set Data Types" = Table.TransformColumnTypes({last}, '
            f'{{{cast_str}}})\nin\n    #"Set Data Types"')


def build_table_tmdl(name: str, frame: pd.DataFrame, partition_m: str,
                     hidden_cols=(), formats=None) -> str:
    formats = formats or {}
    lines = [f"table {esc(name)}", ""]
    for col, dtype in frame.dtypes.items():
        t = tmdl_type(dtype)
        lines.append(f"\tcolumn {esc(col)}")
        lines.append(f"\t\tdataType: {t}")
        if col in formats:
            lines.append(f"\t\tformatString: {formats[col]}")
        if t == "dateTime":
            lines.append("\t\tformatString: General Date")
        summarize = "none" if t in SUMMARIZE_NONE else \
            ("sum" if col in ("Revenue", "Quantity", "Frequency", "Monetary",
                              "Units", "Orders", "Lines", "Customers",
                              "UnitsSold", "TimesSold", "Rows", "Frequency")
             else "none")
        lines.append(f"\t\tsummarizeBy: {summarize}")
        lines.append(f"\t\tsourceColumn: {col}")
        if (name, col) in SORT_BY:
            lines.append(f"\t\tsortByColumn: {SORT_BY[(name, col)]}")
        if col in hidden_cols:
            lines.append("\t\tisHidden")
        lines.append("")
    indented = "\n".join("\t\t\t\t" + ln for ln in partition_m.splitlines())
    lines += [f"\tpartition {esc(name)} = m", "\t\tmode: import",
              "\t\tsource =", indented, ""]
    return "\n".join(lines)


# ==================================================================== measures
# Single-line DAX on purpose: TMDL multi-line expressions need exact
# indentation and one wrong tab rejects the whole model. DAX ignores newlines,
# so the readable formatted versions live in DAX_Measures.txt.
MEASURES = [
    ("Total Revenue", "SUM ( Fact_Sales[Revenue] )", GBP0),
    ("Gross Revenue",
     'CALCULATE ( [Total Revenue], Fact_Sales[TransactionType] = "Sale" )', GBP0),
    ("Return Revenue",
     "- CALCULATE ( [Total Revenue], Fact_Sales[Quantity] < 0 )", GBP0),
    ("Total Orders", "DISTINCTCOUNT ( Fact_Sales[Invoice] )", INT),
    ("Total Customers",
     "CALCULATE ( DISTINCTCOUNT ( Fact_Sales[CustomerKey] ), "
     "Fact_Sales[CustomerKey] <> -1 )", INT),
    ("Units Sold", "SUM ( Fact_Sales[Quantity] )", INT),
    ("Total Lines", "COUNTROWS ( Fact_Sales )", INT),
    ("Average Order Value", "DIVIDE ( [Total Revenue], [Total Orders] )", GBP),
    ("Revenue per Customer", "DIVIDE ( [Total Revenue], [Total Customers] )", GBP0),
    ("Average Basket Size", "DIVIDE ( [Units Sold], [Total Orders] )", "#,0.0"),
    # returns
    ("Return Rate %", "DIVIDE ( [Return Revenue], [Gross Revenue] ) * 100", PCT1),
    ("Cancelled Orders",
     "CALCULATE ( DISTINCTCOUNT ( Fact_Sales[Invoice] ), "
     "Fact_Sales[IsCancelled] = TRUE () )", INT),
    ("Cancellation Rate %", "DIVIDE ( [Cancelled Orders], [Total Orders] ) * 100", PCT1),
    ("Returned Units",
     "- CALCULATE ( SUM ( Fact_Sales[Quantity] ), Fact_Sales[Quantity] < 0 )", INT),
    # time intelligence
    ("Revenue LY",
     "CALCULATE ( [Total Revenue], SAMEPERIODLASTYEAR ( Dim_Date[DateKey] ) )", GBP0),
    ("Revenue YoY",
     "VAR P = [Revenue LY] RETURN IF ( NOT ISBLANK ( P ), [Total Revenue] - P )", GBP0),
    ("Revenue YoY %",
     "VAR P = [Revenue LY] RETURN IF ( NOT ISBLANK ( P ), "
     "DIVIDE ( [Total Revenue] - P, P ) * 100 )", PCT1),
    ("Revenue YTD", "TOTALYTD ( [Total Revenue], Dim_Date[DateKey] )", GBP0),
    ("Revenue MTD", "TOTALMTD ( [Total Revenue], Dim_Date[DateKey] )", GBP0),
    ("Revenue Rolling 12M",
     "CALCULATE ( [Total Revenue], DATESINPERIOD ( Dim_Date[DateKey], "
     "MAX ( Dim_Date[DateKey] ), -12, MONTH ) )", GBP0),
    # smoothing - PDF: Noise Reduction
    ("Revenue 7D Moving Average",
     "AVERAGEX ( DATESINPERIOD ( Dim_Date[DateKey], MAX ( Dim_Date[DateKey] ), "
     "-7, DAY ), [Total Revenue] )", GBP0),
    ("Revenue 30D Moving Average",
     "AVERAGEX ( DATESINPERIOD ( Dim_Date[DateKey], MAX ( Dim_Date[DateKey] ), "
     "-30, DAY ), [Total Revenue] )", GBP0),
    # customer
    ("Guest Revenue",
     "CALCULATE ( [Total Revenue], Fact_Sales[CustomerKey] = -1 )", GBP0),
    ("Guest Revenue %", "DIVIDE ( [Guest Revenue], [Total Revenue] ) * 100", PCT1),
    ("Identified Revenue %", "100 - [Guest Revenue %]", PCT1),
    ("Average Customer Lifetime Value",
     "AVERAGEX ( Dim_Customer, Dim_Customer[Monetary] )", GBP),
    ("Average Recency (days)",
     "AVERAGEX ( FILTER ( Dim_Customer, NOT ISBLANK ( Dim_Customer[Recency] ) ), "
     "Dim_Customer[Recency] )", "#,0"),
    ("Average Frequency",
     "AVERAGEX ( FILTER ( Dim_Customer, NOT ISBLANK ( Dim_Customer[Frequency] ) ), "
     "Dim_Customer[Frequency] )", "#,0.0"),
    ("Segment Revenue %",
     "DIVIDE ( [Total Revenue], CALCULATE ( [Total Revenue], "
     "ALLSELECTED ( Dim_Customer[Segment] ) ) ) * 100", PCT1),
    # CRM (BA-2 CRM.pdf)
    ("Average CLV",
     "AVERAGEX ( FILTER ( Dim_Customer, NOT ISBLANK ( Dim_Customer[CLV] ) ), "
     "Dim_Customer[CLV] )", GBP0),
    ("Total CLV", "SUM ( Dim_Customer[CLV] )", GBP0),
    ("Scored Customers",
     'CALCULATE ( COUNTROWS ( Dim_Customer ), Dim_Customer[ChurnStatus] <> "Not scored" )',
     INT),
    ("Churned Customers",
     'CALCULATE ( COUNTROWS ( Dim_Customer ), Dim_Customer[ChurnStatus] = "Churned" )',
     INT),
    ("Active Customers",
     'CALCULATE ( COUNTROWS ( Dim_Customer ), Dim_Customer[ChurnStatus] = "Active" )',
     INT),
    ("At Risk Customers",
     'CALCULATE ( COUNTROWS ( Dim_Customer ), Dim_Customer[ChurnStatus] = "At Risk" )',
     INT),
    ("Churn Rate %", "DIVIDE ( [Churned Customers], [Scored Customers] ) * 100", PCT1),
    ("Retention Rate %", "100 - [Churn Rate %]", PCT1),
    ("Average Purchase Value",
     "AVERAGEX ( FILTER ( Dim_Customer, NOT ISBLANK ( Dim_Customer[AvgOrderValue] ) ), "
     "Dim_Customer[AvgOrderValue] )", GBP),
    ("Purchase Frequency per Year",
     "AVERAGEX ( FILTER ( Dim_Customer, "
     "NOT ISBLANK ( Dim_Customer[PurchaseFrequencyPerYear] ) ), "
     "Dim_Customer[PurchaseFrequencyPerYear] )", "#,0.0"),
    # product
    ("Products Sold", "DISTINCTCOUNT ( Fact_Sales[StockCode] )", INT),
    ("Class A Revenue %",
     'DIVIDE ( CALCULATE ( [Total Revenue], Dim_Product[ABC_Class] = "A" ), '
     "CALCULATE ( [Total Revenue], ALL ( Dim_Product[ABC_Class] ) ) ) * 100", PCT1),
    ("Product Return Rate %",
     "DIVIDE ( - CALCULATE ( SUM ( Fact_Sales[Quantity] ), Fact_Sales[Quantity] < 0 ), "
     "CALCULATE ( SUM ( Fact_Sales[Quantity] ), Fact_Sales[Quantity] > 0 ) ) * 100", PCT1),
    ("Non-Product Revenue",
     "CALCULATE ( [Total Revenue], Fact_Sales[IsProduct] = FALSE () )", GBP0),
    # data quality
    ("Rows Retained", "COUNTROWS ( Fact_Sales )", INT),
    ("Rows in Raw Source", "1067371", INT),
    ("Data Retention %", "DIVIDE ( [Rows Retained], [Rows in Raw Source] ) * 100", PCT1),
    ("Outlier Rows",
     "CALCULATE ( COUNTROWS ( Fact_Sales ), Fact_Sales[Is_Outlier] = TRUE () )", INT),
    ("Outlier Revenue",
     "CALCULATE ( [Total Revenue], Fact_Sales[Is_Outlier] = TRUE () )", GBP0),
    ("Outlier Revenue Share %",
     "DIVIDE ( [Outlier Revenue], [Total Revenue] ) * 100", PCT1),
    ("Revenue excluding Outliers",
     "CALCULATE ( [Total Revenue], Fact_Sales[Is_Outlier] = FALSE () )", GBP0),
    ("Rows Missing Customer",
     "CALCULATE ( COUNTROWS ( Fact_Sales ), Fact_Sales[CustomerKey] = -1 )", INT),
    ("Customer Completeness %",
     "DIVIDE ( COUNTROWS ( Fact_Sales ) - [Rows Missing Customer], "
     "COUNTROWS ( Fact_Sales ) ) * 100", PCT1),
    # labels
    ("Selected Period Label",
     'FORMAT ( MIN ( Dim_Date[DateKey] ), "d MMM yyyy" ) & " to " & '
     'FORMAT ( MAX ( Dim_Date[DateKey] ), "d MMM yyyy" )', None),
    ("YoY Arrow",
     'VAR P = [Revenue YoY %] RETURN IF ( ISBLANK ( P ), "no prior year", '
     'IF ( P >= 0, UNICHAR ( 9650 ) & " ", UNICHAR ( 9660 ) & " " ) & '
     'FORMAT ( ABS ( P ), "0.0" ) & "% vs LY" )', None),
]


# ============================================================ report visuals
VISUAL_SCHEMA = ("https://developer.microsoft.com/json-schemas/fabric/item/"
                 "report/definition/visualContainer/1.0.0/schema.json")


def measure_field(name, table="_Measures"):
    """A projection referencing a measure."""
    return {
        "field": {"Measure": {
            "Expression": {"SourceRef": {"Entity": table}},
            "Property": name}},
        "queryRef": table + "." + name,
        "nativeQueryRef": name,
    }


def column_field(table, name):
    """A projection referencing a column."""
    return {
        "field": {"Column": {
            "Expression": {"SourceRef": {"Entity": table}},
            "Property": name}},
        "queryRef": table + "." + name,
        "nativeQueryRef": name,
    }


def visual_json(vid, vtype, x, y, w, h, roles, title=None, z=0):
    """
    One visual container. `roles` maps a visual role name (Values, Category,
    Y, X ...) to a list of projections.
    """
    visual = {
        "visualType": vtype,
        "query": {"queryState": {
            role: {"projections": proj} for role, proj in roles.items()}},
        "drillFilterOtherVisuals": True,
    }
    if title:
        visual["visualContainerObjects"] = {
            "title": [{"properties": {
                "text": {"expr": {"Literal": {"Value": "'" + title + "'"}}},
                "show": {"expr": {"Literal": {"Value": "true"}}},
            }}]}
    return {
        "$schema": VISUAL_SCHEMA,
        "name": vid,
        "position": {"x": x, "y": y, "z": z, "width": w, "height": h,
                     "tabOrder": z},
        "visual": visual,
    }


def measures_table_tmdl() -> str:
    lines = ["table _Measures", "",
             "\tcolumn _dummy",
             "\t\tdataType: int64",
             "\t\tisHidden",
             "\t\tsummarizeBy: none",
             "\t\tsourceColumn: _dummy",
             ""]
    for name, expr, fmt in MEASURES:
        lines.append(f"\tmeasure '{name}' = {expr}")
        if fmt:
            lines.append(f"\t\tformatString: {fmt}")
        lines.append("")
    lines += ["\tpartition _Measures = calculated",
              "\t\tmode: import",
              '\t\tsource = ROW("_dummy", 1)', ""]
    return "\n".join(lines)


# =============================================================== main generator
def main():
    for d in (MODEL_DIR, REPORT_DIR):
        if d.exists():
            shutil.rmtree(d)
    (MODEL_DIR / "definition" / "tables").mkdir(parents=True)
    (REPORT_DIR / "definition" / "pages").mkdir(parents=True)

    print(f"Building PBIP  ({'raw Excel + full M' if USE_EXCEL else 'fast path'})\n")

    # ---- load the cleaned tables so columns are read, never assumed
    fact = pd.read_parquet(CLEAN / "Fact_Sales.parquet")
    frames = {
        "Fact_Sales": fact,
        "Dim_Customer": pd.read_csv(CLEAN / "Dim_Customer.csv"),
        "Dim_Product": pd.read_csv(CLEAN / "Dim_Product.csv"),
        "Dim_Country": pd.read_csv(CLEAN / "Dim_Country.csv"),
        "Dim_Date": pd.read_csv(CLEAN / "Dim_Date.csv", parse_dates=["DateKey"]),
        "Dim_Hour": pd.read_csv(CLEAN / "Dim_Hour.csv"),
    }

    formats = {
        "Fact_Sales": {"Revenue": GBP, "Price": GBP},
        "Dim_Customer": {"Monetary": GBP, "AvgOrderValue": GBP},
        "Dim_Product": {"Revenue": GBP, "AvgPrice": GBP},
    }

    # The calendar and hour dimensions are pure logic, so they are authored in
    # M and depend on no file at all. Their bounds come from the cleaned data
    # rather than being typed in, so the generated table cannot drift from it.
    date_span = (frames["Dim_Date"]["DateKey"].min().strftime("%Y-%m-%d"),
                 frames["Dim_Date"]["DateKey"].max().strftime("%Y-%m-%d"))
    generated = {
        "Dim_Date": lambda: date_dim_partition(*date_span),
        "Dim_Hour": hour_dim_partition,
    }

    # ---- tables
    for name, frame in frames.items():
        if name == "Fact_Sales":
            m = excel_fact_partition() if USE_EXCEL else parquet_partition(name)
        elif name in generated:
            m = generated[name]()
        else:
            m = typed_m(name, frame, csv_partition(name))
        tmdl = build_table_tmdl(name, frame, m, formats=formats.get(name, {}))
        (MODEL_DIR / "definition" / "tables" / f"{name}.tmdl").write_text(
            tmdl, encoding="utf-8")
        origin = "generated in M" if name in generated else "file"
        print(f"  table  {name:<18} {len(frame.columns):>2} cols, "
              f"{len(frame):>9,} rows  ({origin})")

    (MODEL_DIR / "definition" / "tables" / "_Measures.tmdl").write_text(
        measures_table_tmdl(), encoding="utf-8")
    print(f"  table  {'_Measures':<18} {len(MEASURES)} measures")

    # ---- a Date column on the fact, so the datetime can join a date key
    fact_tmdl = (MODEL_DIR / "definition" / "tables" / "Fact_Sales.tmdl")
    txt = fact_tmdl.read_text(encoding="utf-8")
    calc_col = ("\tcolumn DateOnly = DATEVALUE ( Fact_Sales[InvoiceDate] )\n"
                "\t\tdataType: dateTime\n"
                "\t\tformatString: Long Date\n"
                "\t\tsummarizeBy: none\n\n")
    txt = txt.replace("\tpartition ", calc_col + "\tpartition ", 1)
    fact_tmdl.write_text(txt, encoding="utf-8")

    # ---- relationships
    rels = [
        ("Fact_Sales", "DateOnly", "Dim_Date", "DateKey"),
        ("Fact_Sales", "CustomerKey", "Dim_Customer", "CustomerKey"),
        ("Fact_Sales", "StockCode", "Dim_Product", "StockCode"),
        ("Fact_Sales", "Country", "Dim_Country", "Country"),
        ("Fact_Sales", "Hour", "Dim_Hour", "Hour"),
    ]
    rel_lines = []
    for ft, fc, tt, tc in rels:
        rel_lines.append(f"relationship {uuid.uuid4()}")
        rel_lines.append(f"\tfromColumn: {ft}.{fc}")
        rel_lines.append(f"\ttoColumn: {tt}.{tc}")
        rel_lines.append("")
    (MODEL_DIR / "definition" / "relationships.tmdl").write_text(
        "\n".join(rel_lines), encoding="utf-8")
    print(f"  {len(rels)} relationships")

    # ---- model + database
    refs = "\n".join(f"ref table {esc(t)}" for t in list(frames) + ["_Measures"])
    model_tmdl = (
        "model Model\n"
        "\tculture: en-GB\n"
        "\tdefaultPowerBIDataSourceVersion: powerBI_V3\n"
        "\tdiscourageImplicitMeasures\n"
        "\tsourceQueryCulture: en-GB\n"
        "\n"
        "\tannotation PBI_QueryOrder = "
        + json.dumps(list(frames) + ["_Measures"]) + "\n"
        "\n"
        + refs + "\n"
    )
    (MODEL_DIR / "definition" / "model.tmdl").write_text(model_tmdl, encoding="utf-8")
    (MODEL_DIR / "definition" / "database.tmdl").write_text(
        "database\n\tcompatibilityLevel: 1567\n", encoding="utf-8")

    # ---- expressions (the SourceFolder parameter)
    #
    # Power Query cannot ask where the .pbip lives, so the project folder has
    # to be recorded somewhere, and a plain text parameter is the right place:
    # it shows up under Home > Transform data > Manage parameters, so
    # relocating the project is one visible edit.
    #
    # A resolving version of this was tried and reverted. Making SourceFolder
    # probe c:\ d:\ e:\ for the workbook does survive a change of drive letter,
    # but it also touches three filesystem roots, and Power BI then demands a
    # privacy level for each one in a modal dialog the moment the file opens.
    # Handing an examiner that dialog is a worse failure than the one it
    # avoids, so the path stays a parameter and README.md explains the edit.
    expr = (
        "expression SourceFolder = " + m_str(ROOT) + " meta "
        "[IsParameterQuery=true, Type=\"Text\", IsParameterQueryRequired=true]\n"
        "\tlineageTag: " + str(uuid.uuid4()) + "\n"
        "\tannotation PBI_NavigationStepName = Navigation\n"
    )
    (MODEL_DIR / "definition" / "expressions.tmdl").write_text(expr, encoding="utf-8")
    model_tmdl = model_tmdl.replace("ref table", "ref expression SourceFolder\n\nref table", 1)
    (MODEL_DIR / "definition" / "model.tmdl").write_text(model_tmdl, encoding="utf-8")

    # ---- model wrapper files
    (MODEL_DIR / "definition.pbism").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/"
                   "semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.2", "settings": {},
    }, indent=2), encoding="utf-8")

    # ---- report wrapper files (one blank page for now; visuals come next)
    (REPORT_DIR / "definition.pbir").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/"
                   "report/definitionProperties/1.0.0/schema.json",
        "version": "4.0",
        "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}},
    }, indent=2), encoding="utf-8")
    (REPORT_DIR / "definition" / "report.json").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/"
                   "report/definition/report/1.0.0/schema.json",
        "themeCollection": {},
        "layoutOptimization": "None",
    }, indent=2), encoding="utf-8")

    # Required - removing it gets "Cannot find file 'version.json'" - but it
    # takes no $schema property; supplying one fails the pattern check.
    (REPORT_DIR / "definition" / "version.json").write_text(json.dumps({
        "version": "1.0.0",
    }, indent=2), encoding="utf-8")

    page_id = "execoverview"
    (REPORT_DIR / "definition" / "pages" / "pages.json").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/"
                   "report/definition/pagesMetadata/1.0.0/schema.json",
        "pageOrder": [page_id],
        "activePageName": page_id,
    }, indent=2), encoding="utf-8")
    pdir = REPORT_DIR / "definition" / "pages" / page_id
    (pdir / "visuals").mkdir(parents=True)

    # One card, to prove the visual schema before generating forty of them.
    probe = visual_json("cardRevenue", "card", 40, 40, 300, 160,
                        {"Values": [measure_field("Total Revenue")]},
                        title="Total Revenue")
    vdir = pdir / "visuals" / "cardRevenue"
    vdir.mkdir(parents=True)
    (vdir / "visual.json").write_text(json.dumps(probe, indent=2), encoding="utf-8")
    (pdir / "page.json").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/"
                   "report/definition/page/1.0.0/schema.json",
        "name": page_id,
        "displayName": "Executive Overview",
        "displayOption": "FitToPage",
        "height": 720,
        "width": 1280,
    }, indent=2), encoding="utf-8")

    # ---- the .pbip entry point
    (PBI / f"{NAME}.pbip").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/"
                   "pbipProperties/1.0.0/schema.json",
        "version": "1.0",
        "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
        "settings": {"enableAutoRecovery": True},
    }, indent=2), encoding="utf-8")

    print(f"\nWritten: {PBI / (NAME + '.pbip')}")
    print("Open that file in Power BI Desktop.")


if __name__ == "__main__":
    main()
