"""
Step 12 - Generate the Tableau workbook (07_tableau/RetailAnalytics.twb).

The Tableau counterpart of 06_build_report.py: the same four dashboards with
the same KPIs and charts, written as Tableau's XML workbook format, reading the
flat table that 11_build_tableau_data.py produced.

How the Power BI pieces translate:

    Power BI                          Tableau
    --------------------------------  ----------------------------------------
    DAX measure                       calculated field (aggregate formula)
    DIVIDE(a, b)                      a / b  (Tableau returns NULL on /0)
    Dim_Date columns                  date parts of [InvoiceDate]
    SAMEPERIODLASTYEAR                LOOKUP(..., -12) table calculation
    7-day moving average              WINDOW_AVG(..., -6, 0) table calculation
    AVERAGEX(Dim_Customer, ...)       per-customer value spread over its lines
    Year / Region slicers             parameters + a boolean filter field

Run:  python 12_build_tableau_workbook.py      (after 11_build_tableau_data.py)
"""
import pathlib
import uuid
from xml.sax.saxutils import quoteattr, escape

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "07_tableau"
DATA_DIR = (OUT / "data").as_posix()
HYPER_PATH = (OUT / "data" / "Sales.hyper").as_posix()

DS = "federated.retail_sales"
DSREF = f"[{DS}]"
INK = "#1B2A36"
ACCENT = "#2E7D8F"
MUTED = "#8A9BA8"

REGIONS = ["United Kingdom", "Europe", "Asia-Pacific", "Middle East & Africa",
           "Americas", "Unknown"]
YEARS = ["2009", "2010", "2011"]

# ------------------------------------------------------------------ source data
# (name, tableau datatype, role)
COLUMNS = [
    ("Invoice", "string", "dimension"),
    ("StockCode", "string", "dimension"),
    ("CustomerKey", "integer", "dimension"),
    ("Country", "string", "dimension"),
    ("InvoiceDate", "datetime", "dimension"),
    ("Hour", "integer", "dimension"),
    ("Quantity", "integer", "measure"),
    ("Price", "real", "measure"),
    ("Revenue", "real", "measure"),
    ("IsCancelled", "boolean", "dimension"),
    ("IsProduct", "boolean", "dimension"),
    ("TransactionType", "string", "dimension"),
    ("Is_Outlier", "boolean", "dimension"),
    ("Description", "string", "dimension"),
    ("PriceBand", "string", "dimension"),
    ("ABC_Class", "string", "dimension"),
    ("CountryISO", "string", "dimension"),
    ("Region", "string", "dimension"),
    ("Frequency", "real", "measure"),
    ("Monetary", "real", "measure"),
    ("Segment", "string", "dimension"),
    ("CRM_Strategy", "string", "dimension"),
    ("CRM_Action", "string", "dimension"),
    ("ChurnStatus", "string", "dimension"),
    ("PurchaseFrequencyPerYear", "real", "measure"),
    ("CLV", "real", "measure"),
    ("HourLabel", "string", "dimension"),
    ("DayPart", "string", "dimension"),
    ("DayPartNo", "integer", "dimension"),
]
COLTYPE = {n: t for n, t, _ in COLUMNS}
# CountryISO carries Tableau's country geographic role (ISO 3166 alpha-3), so
# the map places each country by its code - Channel Islands as JEY - with no
# name matching.
GEO_ROLES = {"CountryISO": "[Country].[ISO3166_3]"}
BAND_COLOURS = [("1: £1M or more", "#0d366b"), ("2: £100k to £1M", "#184f95"),
                ("3: £10k to £100k", "#256abf"), ("4: £1k to £10k", "#3987e5"),
                ("5: under £1k", "#6da7ec")]

# --------------------------------------------------------------------- formats
# Tableau format codes: a type letter (c currency, n number, p percent)
# followed by an Excel-style pattern.
GBP0 = 'c"£"#,##0;-"£"#,##0'
GBP2 = 'c"£"#,##0.00;-"£"#,##0.00'
INT = "n#,##0;-#,##0"
DEC1 = "n#,##0.0;-#,##0.0"
PCT1 = "p0.0%"

# ------------------------------------------------------------ calculated fields
# key -> (caption, formula, datatype, role, format)
# Rates are stored as fractions and formatted as percentages, Tableau's way,
# rather than multiplied by 100 as the DAX does.
CALCS = {
    "TotalRevenue": ("Total Revenue", "SUM([Revenue])", "real", "measure", GBP0),
    "GrossRevenue": ("Gross Revenue",
                     "SUM(IF [TransactionType] = 'Sale' THEN [Revenue] END)",
                     "real", "measure", GBP0),
    "ReturnRevenue": ("Return Revenue",
                      "-SUM(IF [Quantity] < 0 THEN [Revenue] END)",
                      "real", "measure", GBP0),
    "TotalOrders": ("Total Orders", "COUNTD([Invoice])", "integer", "measure", INT),
    "TotalCustomers": ("Total Customers",
                       "COUNTD(IF [CustomerKey] <> -1 THEN [CustomerKey] END)",
                       "integer", "measure", INT),
    "UnitsSold": ("Units Sold", "SUM([Quantity])", "integer", "measure", INT),
    "AOV": ("Average Order Value", "SUM([Revenue]) / COUNTD([Invoice])",
            "real", "measure", GBP2),
    "ReturnRate": ("Return Rate %", "[Calc_ReturnRevenue] / [Calc_GrossRevenue]",
                   "real", "measure", PCT1),
    # ---- customer grain. Customer attributes repeat on every line the customer
    # bought, so each line carries value / (that customer's line count): summed
    # over a customer's lines it gives the value exactly once. This is
    # AVERAGEX(Dim_Customer, ...) rebuilt on a flat table.
    "CustLines": ("Customer Line Count", "{FIXED [CustomerKey] : COUNT([Invoice])}",
                  "integer", "measure", INT),
    "AvgCLV": ("Average CLV",
               "SUM([CLV] / [Calc_CustLines]) / "
               "COUNTD(IF NOT ISNULL([CLV]) THEN [CustomerKey] END)",
               "real", "measure", GBP0),
    "ChurnRate": ("Churn Rate %",
                  "COUNTD(IF [ChurnStatus] = 'Churned' THEN [CustomerKey] END) / "
                  "COUNTD(IF [ChurnStatus] <> 'Not scored' THEN [CustomerKey] END)",
                  "real", "measure", PCT1),
    "RetentionRate": ("Retention Rate %", "1 - [Calc_ChurnRate]",
                      "real", "measure", PCT1),
    "PurchFreq": ("Purchase Frequency per Year",
                  "SUM([PurchaseFrequencyPerYear] / [Calc_CustLines]) / "
                  "COUNTD(IF NOT ISNULL([PurchaseFrequencyPerYear]) "
                  "THEN [CustomerKey] END)",
                  "real", "measure", DEC1),
    "GuestRevenuePct": ("Guest Revenue %",
                        "SUM(IF [CustomerKey] = -1 THEN [Revenue] END) / SUM([Revenue])",
                        "real", "measure", PCT1),
    "SegFrequency": ("Segment Frequency",
                     "SUM([Frequency] / [Calc_CustLines])", "real", "measure", INT),
    "SegMonetary": ("Segment Monetary",
                    "SUM([Monetary] / [Calc_CustLines])", "real", "measure", GBP0),
    # ---- product
    "ProductsSold": ("Products Sold", "COUNTD([StockCode])", "integer", "measure", INT),
    "ClassAPct": ("Class A Revenue %",
                  "SUM(IF [ABC_Class] = 'A' THEN [Revenue] END) / SUM([Revenue])",
                  "real", "measure", PCT1),
    "UnitReturnRate": ("Product Return Rate %",
                       "-SUM(IF [Quantity] < 0 THEN [Quantity] END) / "
                       "SUM(IF [Quantity] > 0 THEN [Quantity] END)",
                       "real", "measure", PCT1),
    "NonProductRevenue": ("Non-Product Revenue",
                          "SUM(IF NOT [IsProduct] THEN [Revenue] END)",
                          "real", "measure", GBP0),
    # ---- time and operations
    "MaxDate": ("Last Invoice Date", "{FIXED : MAX([InvoiceDate])}",
                "datetime", "dimension", None),
    "RevCY": ("Revenue Latest Year to Date",
              "SUM(IF YEAR([InvoiceDate]) = YEAR([Calc_MaxDate]) "
              "AND [InvoiceDate] <= [Calc_MaxDate] THEN [Revenue] END)",
              "real", "measure", GBP0),
    "RevPY": ("Revenue Prior Year to Date",
              "SUM(IF YEAR([InvoiceDate]) = YEAR([Calc_MaxDate]) - 1 "
              "AND [InvoiceDate] <= DATEADD('year', -1, [Calc_MaxDate]) "
              "THEN [Revenue] END)",
              "real", "measure", GBP0),
    "YoY": ("Revenue YoY %", "([Calc_RevCY] - [Calc_RevPY]) / [Calc_RevPY]",
            "real", "measure", PCT1),
    "BasketSize": ("Average Basket Size", "SUM([Quantity]) / COUNTD([Invoice])",
                   "real", "measure", DEC1),
    "CancelRate": ("Cancellation Rate %",
                   "COUNTD(IF [IsCancelled] THEN [Invoice] END) / COUNTD([Invoice])",
                   "real", "measure", PCT1),
    "OutlierRows": ("Outlier Rows", "SUM(IF [Is_Outlier] THEN 1 ELSE 0 END)",
                    "integer", "measure", INT),
    "OutlierShare": ("Outlier Revenue Share %",
                     "SUM(IF [Is_Outlier] THEN [Revenue] END) / SUM([Revenue])",
                     "real", "measure", PCT1),
    # ---- table calculations
    "MA7": ("Revenue 7D Moving Average", "WINDOW_AVG(SUM([Revenue]), -6, 0)",
            "real", "measure", GBP0),
    "RevLY": ("Revenue LY", "LOOKUP(SUM([Revenue]), -12)", "real", "measure", GBP0),
    "Mappable": ("Has Map Location", "NOT ISNULL([CountryISO])",
                 "boolean", "dimension", None),
    # Revenue is extremely skewed (the UK is 85%), so a continuous colour scale
    # paints every other country the same pale shade. Decade bands - each ten
    # times the one below - keep every market distinguishable. FIXED respects
    # the Year/Region context filter, so the bands follow the selection.
    "CountryRevenue": ("Country Revenue", "{FIXED [Country] : SUM([Revenue])}",
                       "real", "measure", GBP0),
    "RevenueBand": ("Country revenue (map)",
                    "IF [Calc_CountryRevenue] >= 1000000 THEN '1: £1M or more' "
                    "ELSEIF [Calc_CountryRevenue] >= 100000 THEN '2: £100k to £1M' "
                    "ELSEIF [Calc_CountryRevenue] >= 10000 THEN '3: £10k to £100k' "
                    "ELSEIF [Calc_CountryRevenue] >= 1000 THEN '4: £1k to £10k' "
                    "ELSE '5: under £1k' END",
                    "string", "dimension", None),
    # ---- the Executive Overview's Year and Region selectors
    "ExecFilter": ("Executive Filter",
                   "([Parameters].[p_Year] = 'All' OR "
                   "STR(YEAR([InvoiceDate])) = [Parameters].[p_Year]) AND "
                   "([Parameters].[p_Region] = 'All' OR "
                   "[Region] = [Parameters].[p_Region])",
                   "boolean", "dimension", None),
}
TABLE_CALCS = {"MA7", "RevLY"}


# ================================================================ field model
class F:
    """A field placed on a shelf: one column-instance of a column."""

    def __init__(self, col, deriv="None", kind=None, calc=False):
        self.col = col                         # bare column name
        self.calc = calc
        self.deriv = deriv                     # None|Sum|User|Year|...
        if kind is None:
            kind = "quantitative" if deriv in ("Sum", "User", "Avg") else "nominal"
        self.kind = kind
        short = {"None": "none", "Sum": "sum", "User": "usr", "Avg": "avg",
                 "Year": "yr", "Quarter": "qr", "Month": "mn", "Weekday": "wd",
                 "Month-Trunc": "tmn", "Day-Trunc": "tdy"}[deriv]
        suffix = {"nominal": "nk", "ordinal": "ok", "quantitative": "qk"}[kind]
        self.name = f"[{short}:{col}:{suffix}]"

    @property
    def ref(self):
        return f"{DSREF}.{self.name}"

    @property
    def colname(self):
        return f"[{col_key(self.col)}]" if self.calc else f"[{self.col}]"


def col_key(name):
    return name


def M(key):
    """A calculated measure."""
    return F(f"Calc_{key}", "User", "quantitative", calc=True)


def D(col, deriv="None", kind=None):
    """A source column used as a dimension (optionally a date part)."""
    if kind is None:
        kind = "ordinal" if deriv not in ("None",) else "nominal"
    return F(col, deriv, kind)


def column_decl(f):
    """The <column> declaration a worksheet needs for each field it uses."""
    if f.calc:
        key = f.col[len("Calc_"):]
        cap, formula, dtype, role, fmt = CALCS[key]
        typ = "quantitative" if role == "measure" else "nominal"
        # The worksheet's copy of the declaration wins over the datasource's,
        # so it has to carry the number format too.
        fmt_attr = f" default-format={quoteattr(fmt)}" if fmt else ""
        return (f"<column caption={quoteattr(cap)} datatype='{dtype}'{fmt_attr} "
                f"name='[{f.col}]' role='{role}' type='{typ}'>"
                f"<calculation class='tableau' formula={quoteattr(formula)} />"
                f"</column>")
    dtype = COLTYPE[f.col]
    role = "measure" if f.deriv in ("Sum", "Avg") else "dimension"
    typ = "quantitative" if role == "measure" else (
        "ordinal" if dtype in ("datetime", "date", "integer") else "nominal")
    geo = f" semantic-role='{GEO_ROLES[f.col]}'" if f.col in GEO_ROLES else ""
    return (f"<column datatype='{dtype}' name='[{f.col}]' role='{role}'{geo} "
            f"type='{typ}' />")


def instance_decl(f):
    inner = ""
    if f.calc and f.col[len("Calc_"):] in TABLE_CALCS:
        # 'Rows' is Tableau's internal name for "Table (across)" - compute
        # along the date axis, not separately for each month.
        inner = "<table-calc ordering-type='Rows' />"
    return (f"<column-instance column='[{f.col}]' derivation='{f.deriv}' "
            f"name='{f.name}' pivot='key' type='{f.kind}'>{inner}"
            f"</column-instance>")


# ================================================================== worksheets
WORKSHEETS = []      # xml strings
SHEET_NAMES = []


def worksheet(name, title, rows=(), cols=(), mark="Automatic", color=None,
              size=None, text=(), wedge=None, detail=(), sort=None,
              top=None, measure_values=None, exec_filter=False, mark_color=ACCENT,
              labels=False, fmt=None, kpi=False, extra_fields=(),
              true_filters=()):
    """
    rows / cols: lists of F (or the strings ':mv' / ':mn' for Measure Values /
    Measure Names). measure_values: list of F shown through Measure Values.
    sort: (dimension F, measure F, 'DESC'|'ASC'). top: (dimension F, n, expr).
    rows=[':lat'], cols=[':lon'] makes a map. true_filters: keys of boolean
    calculations every row shown must satisfy.
    """
    fields = list(rows) + list(cols) + list(text) + list(detail) + list(extra_fields)
    fields += [x for x in (color, size, wedge) if isinstance(x, F)]
    if measure_values:
        fields += measure_values
    if sort:
        fields += [sort[0], sort[1]]
    if exec_filter:
        fields.append(F("Calc_ExecFilter", "None", "nominal", calc=True))
    keep = [F(f"Calc_{k}", "None", "nominal", calc=True) for k in true_filters]
    fields += keep
    fields = [f for f in fields if isinstance(f, F)]

    seen_cols, seen_inst, cdecl, idecl = set(), set(), [], []
    for f in fields:
        if f.col not in seen_cols:
            seen_cols.add(f.col)
            cdecl.append(column_decl(f))
        if f.name not in seen_inst:
            seen_inst.add(f.name)
            idecl.append(instance_decl(f))
    # Calculations that reference other calculations need those declared too.
    for f in list(fields):
        if f.calc:
            formula = CALCS[f.col[5:]][1]
            for dep in CALCS:
                if f"[Calc_{dep}]" in formula and f"Calc_{dep}" not in seen_cols:
                    seen_cols.add(f"Calc_{dep}")
                    cap, frm, dtype, role, fmt = CALCS[dep]
                    typ = "quantitative" if role == "measure" else "nominal"
                    fmt_attr = f" default-format={quoteattr(fmt)}" if fmt else ""
                    cdecl.append(
                        f"<column caption={quoteattr(cap)} datatype='{dtype}'{fmt_attr} "
                        f"name='[Calc_{dep}]' role='{role}' type='{typ}'>"
                        f"<calculation class='tableau' formula={quoteattr(frm)} />"
                        f"</column>")

    filters, slices = [], []
    if measure_values:
        members = "".join(
            f"<groupfilter function='member' level='[:Measure Names]' "
            f"member={quoteattr(chr(34) + f.ref + chr(34))} />"
            for f in measure_values)
        filters.append(
            f"<filter class='categorical' column='{DSREF}.[:Measure Names]'>"
            f"<groupfilter function='union' user:op='manual'>{members}"
            f"</groupfilter></filter>")
        slices.append(f"{DSREF}.[:Measure Names]")
    if exec_filter:
        ef = F("Calc_ExecFilter", "None", "nominal", calc=True)
        filters.append(
            f"<filter class='categorical' column='{ef.ref}' context='true'>"
            f"<groupfilter function='member' level='{ef.name}' member='true' "
            f"user:ui-domain='database' user:ui-enumeration='inclusive' "
            f"user:ui-marker='enumerate' /></filter>")
        slices.append(ef.ref)
    for kf in keep:
        filters.append(
            f"<filter class='categorical' column='{kf.ref}'>"
            f"<groupfilter function='member' level='{kf.name}' member='true' "
            f"user:ui-domain='database' user:ui-enumeration='inclusive' "
            f"user:ui-marker='enumerate' /></filter>")
        slices.append(kf.ref)
    if top:
        dim, n, expr = top
        filters.append(
            f"<filter class='categorical' column='{dim.ref}'>"
            f"<groupfilter count='{n}' end='top' function='end' units='records' "
            f"user:ui-marker='end' user:ui-top-by-field='true'>"
            f"<groupfilter direction='DESC' expression={quoteattr(expr)} "
            f"function='order' user:ui-marker='order'>"
            f"<groupfilter function='level-members' level='{dim.name}' "
            f"user:ui-enumeration='all' user:ui-marker='enumerate' />"
            f"</groupfilter></groupfilter></filter>")
        slices.append(dim.ref)

    sort_xml = ""
    if sort:
        dim, meas, direction = sort
        sort_xml = (f"<sort class='computed' column='{dim.ref}' direction='{direction}' "
                    f"using='{meas.ref}' />")

    def shelf(items):
        parts = []
        for it in items:
            if it == ":mv":
                parts.append(f"{DSREF}.[Multiple Values]")
            elif it == ":mn":
                parts.append(f"{DSREF}.[:Measure Names]")
            elif it == ":lat":
                parts.append(f"{DSREF}.[Latitude (generated)]")
            elif it == ":lon":
                parts.append(f"{DSREF}.[Longitude (generated)]")
            else:
                parts.append(it.ref)
        return " / ".join(parts) if parts else ""

    enc = []
    if color is not None:
        enc.append(f"<color column='{color.ref if isinstance(color, F) else DSREF + '.[:Measure Names]'}' />")
    if size is not None:
        enc.append(f"<size column='{size.ref}' />")
    if wedge is not None:
        enc.append(f"<wedge-size column='{wedge.ref}' />")
    for t in text:
        enc.append(f"<text column='{t.ref if isinstance(t, F) else DSREF + '.[Multiple Values]'}' />")
    for d in detail:
        enc.append(f"<lod column='{d.ref}' />")
    # A filled map draws each country's outline from Tableau's own geocoding.
    is_map = ":lat" in rows
    if is_map:
        enc.append(f"<geometry column='{DSREF}.[Geometry (generated)]' />")

    mark_rules = []
    if mark_color and color is None:
        mark_rules.append(f"<format attr='mark-color' value='{mark_color}' />")
    if labels:
        mark_rules.append("<format attr='mark-labels-show' value='true' />")
        mark_rules.append("<format attr='mark-labels-cull' value='true' />")
    pane_style = (f"<style><style-rule element='mark'>{''.join(mark_rules)}"
                  f"</style-rule></style>") if mark_rules else ""

    # Per-field number formats live on the datasource; KPI cards also get a
    # large centred number.
    style_rules = []
    if kpi:
        style_rules.append(
            "<style-rule element='cell'>"
            "<format attr='text-align' value='center' />"
            "<format attr='vertical-align' value='center' />"
            "</style-rule>")
    table_style = f"<style>{''.join(style_rules)}</style>"

    title_xml = (f"<layout-options><title><formatted-text>"
                 f"<run bold='true' fontcolor='{INK}' fontsize='11'>"
                 f"{escape(title)}</run></formatted-text></title></layout-options>")

    customized = ""
    if kpi:
        t = text[0]
        customized = (
            "<customized-label><formatted-text>"
            f"<run fontcolor='{ACCENT}' fontsize='20' bold='true'>&lt;{escape(t.ref)}&gt;</run>"
            "</formatted-text></customized-label>")

    xml = (
        f"<worksheet name={quoteattr(name)}>{title_xml}<table><view>"
        f"<datasources><datasource caption='Retail Sales' name='{DS}' />"
        + ("<datasource name='Parameters' />" if exec_filter else "")
        + "</datasources>"
        + ("<mapsources><mapsource name='Tableau' /></mapsources>" if is_map else "")
        + f"<datasource-dependencies datasource='{DS}'>{''.join(cdecl)}{''.join(idecl)}"
        f"</datasource-dependencies>"
        + "".join(filters) + sort_xml
        + (f"<slices>{''.join(f'<column>{s}</column>' for s in slices)}</slices>"
           if slices else "")
        + "<aggregation value='true' /></view>"
        f"{table_style}<panes><pane selection-relaxation-option='selection-relaxation-allow'>"
        f"<view><breakdown value='auto' /></view><mark class='{mark}' />"
        + (f"<encodings>{''.join(enc)}</encodings>" if enc else "")
        + customized + pane_style
        + f"</pane></panes><rows>{shelf(rows)}</rows><cols>{shelf(cols)}</cols>"
        f"</table></worksheet>")
    WORKSHEETS.append(xml)
    SHEET_NAMES.append(name)
    return name


def fit_entire(name):
    """Mark a worksheet to fill its dashboard box instead of Standard fit."""
    FIT_ENTIRE.add(name)


def kpi(prefix, key, label, exec_filter=False):
    FIT_ENTIRE.add(f"{prefix} KPI {label}")
    return worksheet(f"{prefix} KPI {label}", label, mark="Text",
                     text=[M(key)], kpi=True, exec_filter=exec_filter,
                     mark_color=None)


# ================================================================== dashboards
DASH_W, DASH_H = 1280, 760
DASHBOARDS = []
DASH_SHEETS = {}     # dashboard name -> worksheet names on it
_zid = [100]


def zid():
    _zid[0] += 1
    return _zid[0]


def U(px, total):
    """Pixels to Tableau's 0-100000 zone units."""
    return int(round(px * 100000 / total))


def zone_sheet(name, x, y, w, h):
    return (f"<zone h='{U(h, DASH_H)}' id='{zid()}' name={quoteattr(name)} "
            f"w='{U(w, DASH_W)}' x='{U(x, DASH_W)}' y='{U(y, DASH_H)}'>"
            f"<zone-style><format attr='border-color' value='#E3E8EC' />"
            f"<format attr='border-style' value='solid' />"
            f"<format attr='border-width' value='1' />"
            f"<format attr='margin' value='4' /></zone-style></zone>")


def zone_text(text, x, y, w, h, size=18):
    return (f"<zone h='{U(h, DASH_H)}' id='{zid()}' type-v2='text' "
            f"w='{U(w, DASH_W)}' x='{U(x, DASH_W)}' y='{U(y, DASH_H)}'>"
            f"<formatted-text><run bold='true' fontcolor='{INK}' "
            f"fontsize='{size}'>{escape(text)}</run></formatted-text></zone>")


def zone_param(param, x, y, w, h):
    return (f"<zone h='{U(h, DASH_H)}' id='{zid()}' mode='compact' "
            f"param='[Parameters].[{param}]' type-v2='paramctrl' "
            f"w='{U(w, DASH_W)}' x='{U(x, DASH_W)}' y='{U(y, DASH_H)}' />")


def zone_legend(sheet, field, x, y, w, h):
    return (f"<zone h='{U(h, DASH_H)}' id='{zid()}' name={quoteattr(sheet)} "
            f"pane-specification-id='0' param='{DSREF}.{field}' type-v2='color' "
            f"w='{U(w, DASH_W)}' x='{U(x, DASH_W)}' y='{U(y, DASH_H)}' />")


def dashboard(name, zones):
    inner = "".join(zones)
    DASH_SHEETS[name] = [n for n in SHEET_NAMES
                         if f"name={quoteattr(n)} " in inner]
    DASHBOARDS.append(
        f"<dashboard name={quoteattr(name)}><style />"
        f"<size maxheight='{DASH_H}' maxwidth='{DASH_W}' minheight='{DASH_H}' "
        f"minwidth='{DASH_W}' sizing-mode='fixed' />"
        + ("<datasources><datasource name='Parameters' /></datasources>"
           if "paramctrl" in inner else "<datasources />")
        + f"<zones><zone h='100000' id='{zid()}' type-v2='layout-basic' w='100000' "
        f"x='0' y='0'>{inner}</zone></zones></dashboard>")


def kpi_zones(sheets, y=56, h=92):
    out, pad, n = [], 10, len(sheets)
    cw = (DASH_W - 20 - pad * (n - 1)) / n
    for i, s in enumerate(sheets):
        out.append(zone_sheet(s, 10 + i * (cw + pad), y, cw, h))
    return out


REV = M("TotalRevenue")
CUST = M("TotalCustomers")
DATE = "InvoiceDate"


def build_executive():
    p = "Exec"
    kp = [kpi(p, "TotalRevenue", "Revenue", True),
          kpi(p, "TotalOrders", "Orders", True),
          kpi(p, "TotalCustomers", "Customers", True),
          kpi(p, "AOV", "Avg order value", True),
          kpi(p, "UnitsSold", "Units", True),
          kpi(p, "ReturnRate", "Return rate", True)]
    day = D(DATE, "Day-Trunc", "quantitative")
    trend = worksheet("Exec Daily Revenue vs 7D MA",
                      "Daily revenue (orange) vs 7-day moving average (blue)",
                      rows=[":mv"], cols=[day], mark="Line", color=":mn",
                      measure_values=[REV, M("MA7")], exec_filter=True)
    # The professor's correction: revenue on a map, by country.
    revenue_map = worksheet("Exec Revenue Map", "Revenue by country (map)",
                            rows=[":lat"], cols=[":lon"], mark="Multipolygon",
                            color=F("Calc_RevenueBand", "None", "nominal", calc=True),
                            detail=[D("CountryISO"), D("Country"), REV],
                            mark_color=None, exec_filter=True,
                            true_filters=["Mappable"])
    country = D("Country")
    by_country = worksheet("Exec Revenue by Country", "Revenue by country",
                           rows=[country], cols=[REV], mark="Bar",
                           sort=(country, REV, "DESC"), exec_filter=True)
    month = D(DATE, "Month-Trunc", "quantitative")
    by_month = worksheet("Exec Revenue by Month", "Revenue by month",
                         rows=[REV], cols=[month], mark="Bar", exec_filter=True)
    desc = D("Description")
    top = worksheet("Exec Top Products", "Top 10 products by revenue",
                    rows=[desc], cols=[REV], mark="Bar",
                    sort=(desc, REV, "DESC"),
                    top=(desc, 10, "SUM([Revenue])"), exec_filter=True)
    z = [zone_text("Online Retail II - Executive Overview", 10, 8, 900, 44)]
    z += kpi_zones(kp)
    z += [zone_sheet(trend, 10, 156, 520, 290),
          zone_sheet(revenue_map, 540, 156, 440, 290),
          zone_sheet(by_country, 990, 156, 280, 290),
          zone_sheet(by_month, 10, 454, 560, 298),
          zone_sheet(top, 580, 454, 480, 298),
          zone_text("Filters", 1070, 456, 200, 28, 11),
          zone_param("p_Year", 1070, 488, 200, 60),
          zone_param("p_Region", 1070, 556, 200, 60),
          zone_legend(revenue_map, "[none:Calc_RevenueBand:nk]",
                      1070, 622, 200, 128)]
    dashboard("Executive Overview", z)


def build_customers():
    p = "Cust"
    kp = [kpi(p, "TotalCustomers", "Customers"),
          kpi(p, "AvgCLV", "Avg lifetime value"),
          kpi(p, "ChurnRate", "Churn rate"),
          kpi(p, "RetentionRate", "Retention rate"),
          kpi(p, "PurchFreq", "Orders per year"),
          kpi(p, "GuestRevenuePct", "Unidentified revenue")]
    seg = D("Segment")
    fit_entire("Cust Revenue by Segment")
    tree = worksheet("Cust Revenue by Segment", "Revenue by RFM segment",
                     mark="Square", color=seg, size=REV, text=[seg, REV],
                     labels=True,
                     mark_color=None)
    strat = D("CRM_Strategy")
    by_strat = worksheet("Cust Customers by CRM Strategy",
                         "Customers by CRM strategy (BA-2)",
                         rows=[strat], cols=[CUST], mark="Bar",
                         sort=(strat, CUST, "DESC"))
    churn = D("ChurnStatus")
    fit_entire("Cust Churn Status")
    pie = worksheet("Cust Churn Status", "Churn status", mark="Pie",
                    color=churn, wedge=CUST, text=[churn, CUST],
                    mark_color=None, labels=True)
    scatter = worksheet("Cust Frequency vs Monetary",
                        "Frequency vs monetary value, by segment",
                        rows=[M("SegMonetary")], cols=[M("SegFrequency")],
                        mark="Circle", color=seg, detail=[seg], mark_color=None)
    action = worksheet("Cust Segment Actions",
                       "Segment, recommended CRM strategy and action",
                       rows=[seg, strat, D("CRM_Action")], cols=[":mn"],
                       mark="Text", text=[":mv"],
                       measure_values=[CUST, REV, M("AvgCLV")],
                       sort=(seg, REV, "DESC"), mark_color=None)
    z = [zone_text("Customer Intelligence - RFM Segments and CRM Strategy",
                   10, 8, 1000, 44)]
    z += kpi_zones(kp)
    z += [zone_sheet(tree, 10, 156, 500, 290),
          zone_sheet(by_strat, 520, 156, 450, 290),
          zone_sheet(pie, 980, 156, 290, 290),
          zone_sheet(scatter, 10, 454, 400, 298),
          zone_sheet(action, 420, 454, 850, 298)]
    dashboard("Customer Intelligence", z)


def build_products():
    p = "Prod"
    kp = [kpi(p, "ProductsSold", "Products"),
          kpi(p, "UnitsSold", "Units"),
          kpi(p, "ClassAPct", "Class A share"),
          kpi(p, "ReturnRate", "Return rate"),
          kpi(p, "UnitReturnRate", "Unit return rate"),
          kpi(p, "NonProductRevenue", "Non-product revenue")]
    desc = D("Description")
    top = worksheet("Prod Top Products", "Top 10 products by revenue",
                    rows=[desc], cols=[REV], mark="Bar",
                    sort=(desc, REV, "DESC"), top=(desc, 10, "SUM([Revenue])"))
    abc = D("ABC_Class")
    fit_entire("Prod ABC Classification")
    pie = worksheet("Prod ABC Classification", "ABC classification",
                    mark="Pie", color=abc, wedge=REV, text=[abc, REV],
                    mark_color=None, labels=True)
    band = D("PriceBand")
    by_band = worksheet("Prod Revenue by Price Band", "Revenue by price band",
                        rows=[REV], cols=[band], mark="Bar",
                        sort=(band, REV, "DESC"))
    tt = D("TransactionType")
    by_tt = worksheet("Prod Revenue by Transaction Type",
                      "Revenue by transaction type",
                      rows=[REV], cols=[tt], mark="Bar")
    detail = worksheet("Prod Product Detail", "Product detail",
                       rows=[desc, abc], cols=[":mn"], mark="Text",
                       text=[":mv"],
                       measure_values=[M("UnitsSold"), REV],
                       sort=(desc, REV, "DESC"), mark_color=None)
    z = [zone_text("Product and Returns", 10, 8, 900, 44)]
    z += kpi_zones(kp)
    z += [zone_sheet(top, 10, 156, 620, 290),
          zone_sheet(pie, 640, 156, 310, 290),
          zone_sheet(by_band, 960, 156, 310, 290),
          zone_sheet(by_tt, 10, 454, 620, 298),
          zone_sheet(detail, 640, 454, 630, 298)]
    dashboard("Product & Returns", z)


def build_time():
    p = "Time"
    kp = [kpi(p, "TotalRevenue", "Revenue"),
          kpi(p, "YoY", "YoY growth (YTD)"),
          kpi(p, "BasketSize", "Avg basket size"),
          kpi(p, "CancelRate", "Cancellation rate"),
          kpi(p, "OutlierRows", "Outlier rows"),
          kpi(p, "OutlierShare", "Outlier revenue share")]
    wd = D(DATE, "Weekday", "ordinal")
    dp = D("DayPart")
    FIT_ENTIRE.add("Time Weekday x Daypart")
    heat = worksheet("Time Weekday x Daypart",
                     "Revenue by weekday and time of day",
                     rows=[wd], cols=[dp], mark="Square", color=REV,
                     text=[REV], mark_color=None, labels=True,
                     sort=(dp, F("DayPartNo", "Avg", "quantitative"), "ASC"))
    hour = D("HourLabel")
    by_hour = worksheet("Time Revenue by Hour", "Revenue by hour",
                        rows=[REV], cols=[hour], mark="Bar")
    month = D(DATE, "Month-Trunc", "quantitative")
    ly = worksheet("Time Monthly Revenue vs LY", "Monthly revenue (orange) vs same month last year (blue)",
                   rows=[":mv"], cols=[month], mark="Line", color=":mn",
                   measure_values=[REV, M("RevLY")])
    qtr = D(DATE, "Quarter", "ordinal")
    by_q = worksheet("Time Revenue by Quarter", "Revenue by quarter",
                     rows=[REV], cols=[qtr], mark="Bar")
    z = [zone_text("Time, Seasonality and Operations", 10, 8, 900, 44)]
    z += kpi_zones(kp)
    z += [zone_sheet(heat, 10, 156, 760, 290),
          zone_sheet(by_hour, 780, 156, 490, 290),
          zone_sheet(ly, 10, 454, 760, 298),
          zone_sheet(by_q, 780, 454, 490, 298)]
    dashboard("Time & Seasonality", z)


# ================================================================== datasource
def datasource_xml():
    cols = "".join(
        f"<column datatype='{t}' name={quoteattr(n)} ordinal='{i}' />"
        for i, (n, t, _) in enumerate(COLUMNS))
    meta = "".join(
        f"<metadata-record class='column'><remote-name>{escape(n)}</remote-name>"
        f"<remote-type>{ {'string': 129, 'integer': 20, 'real': 5, 'datetime': 7, 'boolean': 11}[t] }</remote-type>"
        f"<local-name>[{escape(n)}]</local-name><parent-name>[Sales.csv]</parent-name>"
        f"<remote-alias>{escape(n)}</remote-alias><ordinal>{i}</ordinal>"
        f"<local-type>{t}</local-type><aggregation>"
        f"{'Sum' if r == 'measure' else 'Count'}</aggregation>"
        f"<contains-null>true</contains-null></metadata-record>"
        for i, (n, t, r) in enumerate(COLUMNS))
    src_cols = "".join(
        f"<column datatype='{t}' name='[{n}]' role='{r}' "
        + (f"semantic-role='{GEO_ROLES[n]}' " if n in GEO_ROLES else "")
        + f"type='{'quantitative' if r == 'measure' else ('ordinal' if t in ('datetime', 'integer') else 'nominal')}' />"
        for n, t, r in COLUMNS)
    calc_cols = []
    for key, (cap, formula, dtype, role, fmt) in CALCS.items():
        typ = "quantitative" if role == "measure" else "nominal"
        fmt_attr = f" default-format={quoteattr(fmt)}" if fmt else ""
        calc_cols.append(
            f"<column caption={quoteattr(cap)} datatype='{dtype}'{fmt_attr} "
            f"name='[Calc_{key}]' role='{role}' type='{typ}'>"
            f"<calculation class='tableau' formula={quoteattr(formula)} />"
            f"</column>")
    return (
        f"<datasource caption='Retail Sales' inline='true' name='{DS}' version='18.1'>"
        f"<connection class='federated'><named-connections>"
        f"<named-connection caption='Sales' name='textscan.retail_sales'>"
        f"<connection class='textscan' directory={quoteattr(DATA_DIR)} "
        f"filename='Sales.csv' password='' server='' /></named-connection>"
        f"</named-connections>"
        f"<relation connection='textscan.retail_sales' name='Sales.csv' "
        f"table='[Sales#csv]' type='table'>"
        f"<columns character-set='UTF-8' header='yes' locale='en_GB' separator=','>"
        f"{cols}</columns></relation>"
        f"<metadata-records>{meta}</metadata-records></connection>"
        f"<aliases enabled='yes' />{src_cols}{''.join(calc_cols)}"
        f"<column-instance column='[Calc_RevenueBand]' derivation='None' "
        f"name='[none:Calc_RevenueBand:nk]' pivot='key' type='nominal' />"
        + extract_xml(meta) +
        f"<layout dim-ordering='alphabetic' dim-percentage='0.5' "
        f"measure-ordering='alphabetic' measure-percentage='0.4' "
        f"show-structure='true' />"
        f"<style><style-rule element='mark'><encoding attr='color' "
        f"field='[none:Calc_RevenueBand:nk]' type='palette'>"
        + "".join(f"<map to='{c}'><bucket>{escape(chr(34) + b + chr(34))}</bucket></map>"
                  for b, c in BAND_COLOURS)
        + f"</encoding></style-rule></style>"
        f"<semantic-values><semantic-value key='[Country].[Name]' "
        f"value='&quot;United Kingdom&quot;' /></semantic-values>"
        f"</datasource>")


def extract_xml(meta):
    """
    Tableau Public only accepts extract-backed data sources. The extract is
    07_tableau/data/Sales.hyper, written by 11_build_tableau_data.py with
    Tableau's single-table layout: schema "Extract", table "Extract".
    """
    meta = meta.replace("<parent-name>[Sales.csv]</parent-name>",
                        "<parent-name>[Extract]</parent-name>")
    return (
        "<extract count='-1' enabled='true' units='records'>"
        "<connection access_mode='readonly' author-locale='en_GB' class='hyper' "
        f"dbname={quoteattr(HYPER_PATH)} default-settings='yes' schema='Extract' "
        "sslmode='' tablename='Extract' update-time='10/15/2026 11:00:00 AM' "
        "username='tableau'>"
        "<relation name='Extract' table='[Extract].[Extract]' type='table' />"
        f"<metadata-records>{meta}</metadata-records></connection></extract>")


def parameters_xml():
    def param(name, caption, values):
        members = "".join(f"<member value={quoteattr(chr(34) + v + chr(34))} />"
                          for v in values)
        return (f"<column caption={quoteattr(caption)} datatype='string' "
                f"name='[{name}]' param-domain-type='list' role='measure' "
                f"type='nominal' value='&quot;All&quot;'>"
                f"<calculation class='tableau' formula='&quot;All&quot;' />"
                f"<members>{members}</members></column>")
    return ("<datasource hasconnection='false' inline='true' name='Parameters' "
            "version='18.1'><aliases enabled='yes' />"
            + param("p_Year", "Year", ["All"] + YEARS)
            + param("p_Region", "Region", ["All"] + REGIONS)
            + "</datasource>")


# The side cards a worksheet window shows; required by the schema even for
# worksheets that are hidden because they only live inside dashboards.
CARDS = ("<cards><edge name='left'><strip size='160'><card type='pages' />"
         "<card type='filters' /><card type='marks' /></strip></edge>"
         "<edge name='top'><strip size='2147483647'><card type='columns' />"
         "</strip><strip size='2147483647'><card type='rows' /></strip>"
         "<strip size='31'><card type='title' /></strip></edge></cards>")


FIT_ENTIRE = set()   # worksheets that should fill their dashboard box


def viewpoint(name):
    if name in FIT_ENTIRE:
        return f"<viewpoint name={quoteattr(name)}><zoom type='entire-view' /></viewpoint>"
    return f"<viewpoint name={quoteattr(name)} />"


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--open", default="Executive Overview",
                    help="dashboard the workbook opens on")
    ap.add_argument("--out", default=str(OUT / "RetailAnalytics.twb"))
    ap.add_argument("--no-package", action="store_true",
                    help="skip writing the single-file .twbx")
    args = ap.parse_args()
    build_executive()
    build_customers()
    build_products()
    build_time()

    windows = "".join(
        f"<window class='worksheet' hidden='true' maximized='true' name={quoteattr(n)}>"
        f"{CARDS}</window>"
        for n in SHEET_NAMES)
    # The tabs follow the order of these windows; the one marked maximized is
    # the dashboard the workbook opens on.
    dash_names = list(DASH_SHEETS)
    windows += "".join(
        f"<window class='dashboard'{' maximized=' + chr(39) + 'true' + chr(39) if n == args.open else ''} "
        f"name={quoteattr(n)}>"
        f"<viewpoints>{''.join(viewpoint(v) for v in DASH_SHEETS[n])}"
        f"</viewpoints><active id='-1' /></window>"
        for n in dash_names)

    xml = (
        "<?xml version='1.0' encoding='utf-8' ?>\n"
        "<workbook source-build='2025.1.0 (20251.25.0213.1012)' "
        "source-platform='win' version='18.1' "
        "xmlns:user='http://www.tableausoftware.com/xml/user'>"
        "<preferences><preference name='ui.encoding.shelf.height' value='24' />"
        "<preference name='ui.shelf.height' value='26' /></preferences>"
        "<datasources>" + parameters_xml() + datasource_xml() + "</datasources>"
        "<mapsources><mapsource name='Tableau' /></mapsources>"
        "<worksheets>" + "".join(WORKSHEETS) + "</worksheets>"
        "<dashboards>" + "".join(DASHBOARDS) + "</dashboards>"
        "<windows source-height='30'>" + windows + "</windows>"
        "</workbook>\n")
    path = pathlib.Path(args.out)
    path.write_text(xml, encoding="utf-8")
    print(f"Wrote {path}")
    print(f"  {len(DASHBOARDS)} dashboards, {len(WORKSHEETS)} worksheets")
    if not args.no_package:
        package(xml, path.with_suffix(".twbx"))


def package(xml, twbx):
    """
    The .twb points at the extract by absolute path, so it only opens on this
    machine. A .twbx is a zip of the workbook plus its data, with paths made
    relative to the zip - one file that opens anywhere. The 276 MB CSV is left
    out: the extract holds the same rows, and only a refresh needs the CSV.
    """
    import zipfile
    inner = (xml.replace(HYPER_PATH, "Data/Extracts/Sales.hyper")
                .replace(DATA_DIR, "Data/data"))
    with zipfile.ZipFile(twbx, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(twbx.with_suffix(".twb").name, inner)
        z.write(HYPER_PATH, "Data/Extracts/Sales.hyper")
    print(f"Wrote {twbx}  ({twbx.stat().st_size / 1e6:.0f} MB, opens anywhere)")


if __name__ == "__main__":
    main()
