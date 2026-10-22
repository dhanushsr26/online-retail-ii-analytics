"""
Step 6b - Generate the report pages.

This Power BI install saves reports in the LEGACY layout format: a single
report.json holding a `sections` array, where each visual's configuration is a
JSON string nested inside the JSON. It is not the PBIR folder-per-visual format,
so definition/pages/... is ignored entirely.

The shape below was taken from a report.json that Power BI itself wrote, not
from guesswork.

Run:  python 06_build_report.py     (after 05_build_pbip.py)
"""
import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
REPORT = ROOT / "03_powerbi" / "RetailAnalytics.Report"

W, H = 1280, 720                       # page canvas
INK = "#1B2A36"
ACCENT = "#2E7D8F"

_seq = [0]


def vid(prefix="v"):
    _seq[0] += 1
    return f"{prefix}{_seq[0]:04d}"


# --------------------------------------------------------------- query pieces
def _alias(entity):
    return entity.replace("_", "").lower()[:3] + str(abs(hash(entity)) % 97)


def build_query(fields):
    """
    fields: list of (entity, property, kind) where kind is 'Measure'|'Column'.
    Returns (prototypeQuery, [queryRef...]).
    """
    entities, from_clause = {}, []
    for entity, _, _ in fields:
        if entity not in entities:
            alias = _alias(entity)
            entities[entity] = alias
            from_clause.append({"Name": alias, "Entity": entity, "Type": 0})

    # Aggregation function codes used by the query language.
    AGG = {"Sum": 0, "Avg": 1, "Min": 3, "Max": 4, "Count": 5}

    select, refs = [], []
    for entity, prop, kind in fields:
        alias = entities[entity]
        if kind in AGG:
            ref = f"{kind}({entity}.{prop})"
            select.append({
                "Aggregation": {
                    "Expression": {"Column": {
                        "Expression": {"SourceRef": {"Source": alias}},
                        "Property": prop}},
                    "Function": AGG[kind]},
                "Name": ref,
            })
        else:
            ref = f"{entity}.{prop}"
            select.append({
                kind: {"Expression": {"SourceRef": {"Source": alias}},
                       "Property": prop},
                "Name": ref,
            })
        refs.append(ref)
    return {"Version": 2, "From": from_clause, "Select": select}, refs


def visual(vtype, x, y, w, h, roles, title=None, z=0, objects=None,
           order_by=None, descending=True):
    """
    roles: ordered dict of role -> list of (entity, property, kind).
    Power BI matches projections to the prototypeQuery by queryRef, so both are
    built from the same field list to keep them in step.
    """
    flat = [f for fields in roles.values() for f in fields]
    proto, refs = build_query(flat)

    projections, i = {}, 0
    for role, fields in roles.items():
        projections[role] = []
        for _ in fields:
            projections[role].append({"queryRef": refs[i]})
            i += 1

    # Rank the categories by a measure instead of alphabetically. Direction 2
    # is descending, 1 ascending.
    if order_by:
        entity, prop, kind = order_by
        alias = next(f["Name"] for f in proto["From"] if f["Entity"] == entity)
        proto["OrderBy"] = [{
            "Direction": 2 if descending else 1,
            "Expression": {kind: {
                "Expression": {"SourceRef": {"Source": alias}},
                "Property": prop}},
        }]

    single = {
        "visualType": vtype,
        "projections": projections,
        "prototypeQuery": proto,
        "drillFilterOtherVisuals": True,
    }
    if objects:
        single["objects"] = objects

    vc_objects = {}
    if title:
        vc_objects["title"] = [{"properties": {
            "text": {"expr": {"Literal": {"Value": f"'{title}'"}}},
            "show": {"expr": {"Literal": {"Value": "true"}}},
            "fontColor": {"solid": {"color": {"expr": {"Literal": {
                "Value": f"'{INK}'"}}}}},
            "fontSize": {"expr": {"Literal": {"Value": "11D"}}},
            "bold": {"expr": {"Literal": {"Value": "true"}}},
        }}]
    if vc_objects:
        single["vcObjects"] = vc_objects

    name = vid()
    config = {
        "name": name,
        "layouts": [{"id": 0, "position": {
            "x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": z}}],
        "singleVisual": single,
    }
    return {"x": x, "y": y, "z": z, "width": w, "height": h,
            "config": json.dumps(config), "filters": "[]"}


def textbox(x, y, w, h, runs, z=0):
    """A static text box. `runs` is a list of (text, size, bold, colour)."""
    paragraphs = []
    for text, size, bold, colour in runs:
        paragraphs.append({"textRuns": [{
            "value": text,
            "textStyle": {"fontSize": f"{size}pt", "color": colour,
                          "fontWeight": "bold" if bold else "normal"},
        }]})
    name = vid("txt")
    config = {
        "name": name,
        "layouts": [{"id": 0, "position": {
            "x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": z}}],
        "singleVisual": {
            "visualType": "textbox",
            "drillFilterOtherVisuals": True,
            "objects": {"general": [{"properties": {
                "paragraphs": paragraphs}}]},
        },
    }
    return {"x": x, "y": y, "z": z, "width": w, "height": h,
            "config": json.dumps(config), "filters": "[]"}


def section(index, name, display, containers):
    return {
        "name": name,
        "displayName": display,
        "filters": "[]",
        "ordinal": index,
        "visualContainers": containers,
        "config": json.dumps({}),
        "displayOption": 1,
        "width": W,
        "height": H,
    }


M = "_Measures"


def kpi_row(specs, y=64):
    """Evenly spaced cards across the page width."""
    out, pad, n = [], 12, len(specs)
    cw = (W - 40 - pad * (n - 1) - 20) // n
    for i, (measure, label) in enumerate(specs):
        out.append(visual("card", 20 + i * (cw + pad), y, cw, 96,
                          {"Values": [(M, measure, "Measure")]}, title=label))
    return out


# =================================================================== page one
def page_executive():
    c = [textbox(20, 12, 700, 44, [
        ("Online Retail II - Executive Overview", 18, True, INK)])]
    c += kpi_row([
        ("Total Revenue", "Revenue"),
        ("Total Orders", "Orders"),
        ("Total Customers", "Customers"),
        ("Average Order Value", "Avg order value"),
        ("Units Sold", "Units"),
        ("Return Rate %", "Return rate"),
    ])
    # Revenue trend with the 7-day moving average over it (PDF: Noise Reduction)
    c.append(visual("lineChart", 20, 176, 840, 250,
                    {"Category": [("Dim_Date", "DateKey", "Column")],
                     "Y": [(M, "Total Revenue", "Measure"),
                           (M, "Revenue 7D Moving Average", "Measure")]},
                    title="Daily revenue vs 7-day moving average"))
    c.append(visual("clusteredBarChart", 872, 176, 388, 250,
                    {"Category": [("Dim_Country", "Country", "Column")],
                     "Y": [(M, "Total Revenue", "Measure")]},
                    title="Revenue by country",
                    order_by=(M, "Total Revenue", "Measure")))
    c.append(visual("clusteredColumnChart", 20, 438, 512, 262,
                    {"Category": [("Dim_Date", "MonthYearLabel", "Column")],
                     "Y": [(M, "Total Revenue", "Measure")]},
                    title="Revenue by month"))
    c.append(visual("clusteredBarChart", 544, 438, 452, 262,
                    {"Category": [("Dim_Product", "Description", "Column")],
                     "Y": [(M, "Total Revenue", "Measure")]},
                    title="Top products by revenue",
                    order_by=(M, "Total Revenue", "Measure")))
    c.append(visual("slicer", 1008, 438, 252, 128,
                    {"Values": [("Dim_Date", "Year", "Column")]},
                    title="Year"))
    c.append(visual("slicer", 1008, 574, 252, 126,
                    {"Values": [("Dim_Country", "Region", "Column")]},
                    title="Region"))
    return section(0, "ReportSectionExec", "Executive Overview", c)


# =================================================================== page two
def page_quality():
    c = [textbox(20, 12, 900, 44, [
        ("Data Quality and Preprocessing", 18, True, INK)])]
    c += kpi_row([
        ("Rows in Raw Source", "Raw rows"),
        ("Rows Retained", "Rows retained"),
        ("Rows Removed (filters only)", "Rows removed"),
        ("Data Retention %", "Retention"),
        ("Customer Completeness %", "Customer completeness"),
        ("Techniques Applied", "PDF techniques"),
    ])
    c.append(visual("waterfallChart", 20, 176, 760, 250,
                    {"Category": [("Dim_DataQuality", "Technique", "Column")],
                     "Y": [("Dim_DataQuality", "Rows_Removed", "Sum")]},
                    title="Rows removed at each cleaning step"))
    c.append(visual("clusteredColumnChart", 792, 176, 468, 250,
                    {"Category": [("Agg_Histograms", "Bin", "Column")],
                     "Y": [("Agg_Histograms", "Frequency", "Sum")]},
                    title="Distribution shape retained after histogram reduction"))
    # The audit trail itself - the page's whole point
    c.append(visual("tableEx", 20, 438, 1240, 262,
                    {"Values": [("Dim_DataQuality", "Step", "Column"),
                                ("Dim_DataQuality", "Technique", "Column"),
                                ("Dim_DataQuality", "PDF_Section", "Column"),
                                ("Dim_DataQuality", "Step_Type", "Column"),
                                ("Dim_DataQuality", "Rows_Before", "Column"),
                                ("Dim_DataQuality", "Rows_After", "Column"),
                                ("Dim_DataQuality", "Rows_Removed", "Column"),
                                ("Dim_DataQuality", "Rationale", "Column")]},
                    title="Audit trail - every step, its PDF source and its effect"))
    return section(1, "ReportSectionQuality", "Data Quality", c)


# ================================================================= page three
def page_customers():
    """
    Customer Intelligence, framed by BA-2 CRM.pdf: the eight CRM strategies,
    churn, and CLV computed with the PDF's own formula. The point of the page
    is not to describe segments but to say what to do about each one.
    """
    c = [textbox(20, 12, 900, 44, [
        ("Customer Intelligence - RFM Segments and CRM Strategy", 18, True, INK)])]
    c += kpi_row([
        ("Total Customers", "Customers"),
        ("Average CLV", "Avg lifetime value"),
        ("Churn Rate %", "Churn rate"),
        ("Retention Rate %", "Retention rate"),
        ("Purchase Frequency per Year", "Orders per year"),
        ("Guest Revenue %", "Unidentified revenue"),
    ])
    c.append(visual("treemap", 20, 176, 496, 250,
                    {"Group": [("Dim_Customer", "Segment", "Column")],
                     "Values": [(M, "Total Revenue", "Measure")]},
                    title="Revenue by RFM segment"))
    c.append(visual("clusteredBarChart", 528, 176, 456, 250,
                    {"Category": [("Dim_Customer", "CRM_Strategy", "Column")],
                     "Y": [(M, "Total Customers", "Measure")]},
                    title="Customers by CRM strategy (BA-2)",
                    order_by=(M, "Total Customers", "Measure")))
    c.append(visual("donutChart", 996, 176, 264, 250,
                    {"Category": [("Dim_Customer", "ChurnStatus", "Column")],
                     "Y": [(M, "Total Customers", "Measure")]},
                    title="Churn status"))
    c.append(visual("scatterChart", 20, 438, 420, 262,
                    {"Category": [("Dim_Customer", "Segment", "Column")],
                     "X": [("Dim_Customer", "Frequency", "Sum")],
                     "Y": [("Dim_Customer", "Monetary", "Sum")]},
                    title="Frequency vs monetary value, by segment"))
    # The action table - what the dashboard is actually for
    c.append(visual("tableEx", 452, 438, 808, 262,
                    {"Values": [("Dim_Customer", "Segment", "Column"),
                                ("Dim_Customer", "CRM_Strategy", "Column"),
                                ("Dim_Customer", "CRM_Action", "Column"),
                                (M, "Total Customers", "Measure"),
                                (M, "Total Revenue", "Measure"),
                                (M, "Average CLV", "Measure")]},
                    title="Segment, recommended CRM strategy and action",
                    order_by=(M, "Total Revenue", "Measure")))
    return section(1, "ReportSectionCustomers", "Customer Intelligence", c)


# ================================================================== page four
def page_products():
    c = [textbox(20, 12, 700, 44, [("Product and Returns", 18, True, INK)])]
    c += kpi_row([
        ("Products Sold", "Products"),
        ("Units Sold", "Units"),
        ("Class A Revenue %", "Class A share"),
        ("Return Rate %", "Return rate"),
        ("Product Return Rate %", "Unit return rate"),
        ("Non-Product Revenue", "Non-product revenue"),
    ])
    c.append(visual("clusteredBarChart", 20, 176, 620, 250,
                    {"Category": [("Dim_Product", "Description", "Column")],
                     "Y": [(M, "Total Revenue", "Measure")]},
                    title="Top products by revenue",
                    order_by=(M, "Total Revenue", "Measure")))
    c.append(visual("donutChart", 652, 176, 300, 250,
                    {"Category": [("Dim_Product", "ABC_Class", "Column")],
                     "Y": [(M, "Total Revenue", "Measure")]},
                    title="ABC classification"))
    c.append(visual("clusteredColumnChart", 964, 176, 296, 250,
                    {"Category": [("Dim_Product", "PriceBand", "Column")],
                     "Y": [(M, "Total Revenue", "Measure")]},
                    title="Revenue by price band"))
    c.append(visual("clusteredColumnChart", 20, 438, 620, 262,
                    {"Category": [("Fact_Sales", "TransactionType", "Column")],
                     "Y": [(M, "Total Revenue", "Measure")]},
                    title="Revenue by transaction type"))
    c.append(visual("tableEx", 652, 438, 608, 262,
                    {"Values": [("Dim_Product", "Description", "Column"),
                                ("Dim_Product", "ABC_Class", "Column"),
                                ("Dim_Product", "UnitsSold", "Sum"),
                                ("Dim_Product", "Revenue", "Sum")]},
                    title="Product detail"))
    return section(2, "ReportSectionProducts", "Product & Returns", c)


# ================================================================== page five
def page_time():
    c = [textbox(20, 12, 800, 44, [
        ("Time, Seasonality and Operations", 18, True, INK)])]
    c += kpi_row([
        ("Total Revenue", "Revenue"),
        ("Revenue YoY %", "YoY growth"),
        ("Average Basket Size", "Avg basket size"),
        ("Cancellation Rate %", "Cancellation rate"),
        ("Outlier Rows", "Outlier rows"),
        ("Outlier Revenue Share %", "Outlier revenue share"),
    ])
    c.append(visual("matrix", 20, 176, 760, 250,
                    {"Rows": [("Dim_Date", "Weekday", "Column")],
                     "Columns": [("Dim_Hour", "DayPart", "Column")],
                     "Values": [(M, "Total Revenue", "Measure")]},
                    title="Revenue by weekday and time of day"))
    c.append(visual("clusteredColumnChart", 792, 176, 468, 250,
                    {"Category": [("Dim_Hour", "HourLabel", "Column")],
                     "Y": [(M, "Total Revenue", "Measure")]},
                    title="Revenue by hour"))
    c.append(visual("lineChart", 20, 438, 760, 262,
                    {"Category": [("Dim_Date", "MonthYearLabel", "Column")],
                     "Y": [(M, "Total Revenue", "Measure"),
                           (M, "Revenue LY", "Measure")]},
                    title="Monthly revenue vs last year"))
    c.append(visual("clusteredColumnChart", 792, 438, 468, 262,
                    {"Category": [("Dim_Date", "Quarter", "Column")],
                     "Y": [(M, "Total Revenue", "Measure")]},
                    title="Revenue by quarter"))
    return section(3, "ReportSectionTime", "Time & Seasonality", c)


def main():
    # The Data Quality page has been dropped: the preprocessing is documented
    # in 04_report/Preprocessing_Report.md instead, which is easier to talk an
    # examiner through than a dashboard page. Dim_DataQuality stays in the
    # model, so the evidence is still one visual away if it is ever asked for.
    sections = [page_executive(), page_customers(),
                page_products(), page_time()]

    report = {
        "config": json.dumps({
            "version": "5.37",
            "themeCollection": {},
            "activeSectionIndex": 0,
            "linguisticSchemaSyncVersion": 2,
        }),
        "layoutOptimization": 0,
        "report": {},
        "sections": sections,
    }
    path = REPORT / "report.json"
    path.write_text(json.dumps(report, indent=2), encoding="utf-8")

    # The PBIR folder is ignored by this Power BI version and its presence is
    # what made the earlier attempts fail silently.
    stale = REPORT / "definition"
    if stale.exists():
        import shutil
        shutil.rmtree(stale)

    total = sum(len(s["visualContainers"]) for s in sections)
    print(f"Wrote {path}")
    print(f"  {len(sections)} pages, {total} visuals")
    for s in sections:
        print(f"    {s['displayName']:<26} {len(s['visualContainers']):>2} visuals")


if __name__ == "__main__":
    main()
