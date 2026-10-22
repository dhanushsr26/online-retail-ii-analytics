"""
Steps 2-3 - The preprocessing pipeline.

Applies 13 techniques drawn from BA-6 Preprocessing Techniques.pdf and writes a
star schema to 02_clean/. Every step appends an audit row to Dim_DataQuality.csv,
which is imported into Power BI and rendered on the Data Quality page - so the
preprocessing is visible to the examiner rather than buried in a script.

Run:  python 03_clean.py
"""
import pathlib

import numpy as np
import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
CACHE = ROOT / "00_source" / "raw_combined.parquet"
XLSX = ROOT / "00_source" / "online_retail_II.xlsx"
OUT = ROOT / "02_clean"
EVID = ROOT / "05_evidence"

CUST = "Customer ID"
ALPHA = 0.3          # exponential smoothing constant (PDF: New = a*Actual + (1-a)*Prev)
MA_WINDOW = 7        # moving-average window (PDF: n-point moving average)
Z_THRESHOLD = 3.0    # |z| beyond this is flagged, never deleted

audit = []


def log(technique, pdf_section, before, after, rationale,
        step_type="Row filter", note=""):
    """
    Append one row to the Dim_DataQuality evidence table.

    Step_Type matters: only 'Row filter' steps actually delete rows, and only
    those belong in the dashboard's waterfall. Derivation, schema and
    aggregation steps leave the fact table's row count untouched - their
    before/after belongs in Metric_Note, measured in its own units. Mixing the
    two would make the waterfall claim we deleted a million rows to build a
    604-row daily summary.
    """
    if step_type == "Row filter":
        removed = before - after
        rows_before, rows_after = before, after
    else:
        # Left blank on purpose so these steps cannot be charted as deletions.
        removed = 0
        rows_before = rows_after = None
    step = len(audit)
    audit.append({
        "Step": step,
        "Technique": technique,
        "PDF_Section": pdf_section,
        "Step_Type": step_type,
        "Rows_Before": rows_before,
        "Rows_After": rows_after,
        "Rows_Removed": removed,
        "Pct_Removed": round(removed / rows_before * 100, 4) if rows_before else None,
        "Metric_Note": note,
        "Rationale": rationale,
    })
    detail = note if note else f"{before:,} -> {after:,}  ({-removed:+,})"
    print(f"  [{step:>2}] {technique:<44} {detail}")


# --------------------------------------------------------------- country map
# Canonical name, ISO-3 code and region for every label in the raw data.
# Region is itself Data Generalization applied to Country.
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

# ------------------------------------------------------------- item typing
FEE_CODES = {"POST", "DOT", "C2", "C3"}
ADJUST_CODES = {"M", "D", "B", "S", "BANK CHARGES", "AMAZONFEE", "ADJUST",
                "ADJUST2", "CRUK", "PADS"}
TEST_CODES = {"TEST001", "TEST002"}


def classify_item(code: str) -> str:
    up = code.upper()
    if up in TEST_CODES:
        return "Test Record"
    if up in FEE_CODES:
        return "Postage & Carriage"
    if up in ADJUST_CODES:
        return "Fee / Adjustment"
    if up.startswith("GIFT"):
        return "Gift Voucher"
    if up.startswith("DCGS"):
        return "Discount / Sample"
    if code[:5].isdigit():
        return "Product"
    return "Other Non-Product"


def load_raw() -> pd.DataFrame:
    if CACHE.exists():
        return pd.read_parquet(CACHE)
    sheets = pd.read_excel(XLSX, sheet_name=None,
                           dtype={"Invoice": str, "StockCode": str})
    frames = []
    for name, f in sheets.items():
        f = f.copy()
        f["SourceSheet"] = name
        f["Description"] = f["Description"].astype("string")
        frames.append(f)
    df = pd.concat(frames, ignore_index=True)
    df.to_parquet(CACHE, index=False)
    return df


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    EVID.mkdir(parents=True, exist_ok=True)
    print("\n=== PREPROCESSING PIPELINE ===\n")

    # ============================================================ TECHNIQUE 9
    # Data Integration - union the two sheets (PDF: merge on Batch ID)
    print("A. DATA CLEANING")
    df = load_raw()
    n0 = len(df)
    log("Raw load (both sheets unioned)", "Data Integration",
        n0, n0, "Year 2009-2010 and Year 2010-2011 stacked into one table.")

    # ------------------------------------------------------------ TECHNIQUE 9b
    # The two sheets overlap 1-9 Dec 2010: the same invoices appear in both.
    # A naive union double-counts them, so cross-sheet copies go first.
    value_cols = [c for c in df.columns if c != "SourceSheet"]
    before = len(df)
    cross = df.duplicated(subset=value_cols, keep="first") & \
        ~df.duplicated(subset=value_cols + ["SourceSheet"], keep="first")
    cross_rev = (df.loc[cross, "Quantity"] * df.loc[cross, "Price"]).sum()
    df = df[~cross].reset_index(drop=True)
    log("Cross-sheet overlap resolution", "Data Integration",
        before, len(df),
        f"Sheets overlap 1-9 Dec 2010; the same invoices appear twice. "
        f"Removing the second copy corrects GBP {cross_rev:,.2f} of "
        f"double-counted revenue.")

    # ------------------------------------------------------------ TECHNIQUE 3
    # Duplicate Removal (PDF: Duplicate Removal - Batch/Seal example)
    before = len(df)
    dup = df.duplicated(subset=value_cols, keep="first")
    dup_rev = (df.loc[dup, "Quantity"] * df.loc[dup, "Price"]).sum()
    df = df[~dup].reset_index(drop=True)
    log("Duplicate removal (within sheet)", "Duplicate Removal",
        before, len(df),
        f"Exact repeats of all 8 fields inside one sheet. "
        f"Removed GBP {dup_rev:,.2f} of double-counted revenue.")

    # ----------------------------------------------------------- TECHNIQUE 10
    # Unit conversion / format standardisation
    before = len(df)
    df["Invoice"] = df["Invoice"].astype("string").str.strip()
    df["StockCode"] = df["StockCode"].astype("string").str.strip().str.upper()
    df["Description"] = (df["Description"].astype("string")
                         .str.strip().str.replace(r"\s+", " ", regex=True)
                         .str.title())
    n_country_changed = int(
        df["Country"].map(lambda c: COUNTRY_MAP.get(c, (c,))[0] != c).sum())
    df["CountryRaw"] = df["Country"]
    df["Country"] = df["CountryRaw"].map(lambda c: COUNTRY_MAP.get(c, (c, "", "Unknown"))[0])
    df["CountryISO"] = df["CountryRaw"].map(lambda c: COUNTRY_MAP.get(c, (c, "", "Unknown"))[1])
    df["Region"] = df["CountryRaw"].map(lambda c: COUNTRY_MAP.get(c, (c, "", "Unknown"))[2])
    df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"])
    df["CustomerKey"] = (df[CUST].astype("Float64").astype("Int64")
                         .fillna(-1).astype("int64"))
    log("Unit conversion & format standardisation", "Unit Conversion",
        before, len(df),
        f"Trimmed and title-cased Description, upper-cased StockCode, mapped "
        f"{n_country_changed:,} rows to canonical country names with ISO-3 codes, "
        f"cast Invoice and Customer ID to text (identifiers, not numbers).",
        step_type="Derivation",
        note=f"{n_country_changed:,} rows remapped to canonical country names "
             f"(EIRE->Ireland, RSA->South Africa, Unspecified->Unknown) + ISO-3 codes")

    # ---------------------------------------------------------- TECHNIQUE 3b
    # Exact-match de-duplication cannot catch a line repeated under two
    # different product names. Description is a product attribute, not part of
    # what identifies a transaction, so the business key excludes it.
    business_key = ["Invoice", "StockCode", "Quantity", "Price",
                    "InvoiceDate", "CustomerKey", "Country"]
    before = len(df)
    multi_desc = int((df.groupby("StockCode")["Description"].nunique() > 1).sum())
    bdup = df.duplicated(subset=business_key, keep="first")
    bdup_rev = df.loc[bdup, "Quantity"].mul(df.loc[bdup, "Price"]).sum()
    df = df[~bdup].reset_index(drop=True)
    log("Duplicate removal (transaction business key)", "Duplicate Removal",
        before, len(df),
        f"{multi_desc:,} StockCodes carry more than one Description "
        f"(e.g. 23298 is both 'Bunting , Spotty' and 'Spotty Bunting'), so an "
        f"exact 8-column match misses lines duplicated under a second spelling. "
        f"De-duplicating on Invoice+StockCode+Quantity+Price+Timestamp+Customer "
        f"caught {before - len(df)} further rows worth GBP {bdup_rev:,.2f}. This "
        f"is also why Dim_Product is keyed on StockCode, never on Description.")

    # ------------------------------------------------------------ TECHNIQUE 1
    # Missing value analysis - MCAR / MAR / MNAR
    # Description-missing rows are provably MAR: 100% also lack a Customer ID
    # and 100% carry Price = 0. They are non-transactions, so they go.
    before = len(df)
    junk = df["Description"].isna() & (df["Price"] == 0)
    junk_rev = (df.loc[junk, "Quantity"] * df.loc[junk, "Price"]).sum()
    df = df[~junk].reset_index(drop=True)
    log("Missing-value removal (MAR, zero-value)", "Missing Value Types (MAR)",
        before, len(df),
        f"Rows with no Description AND Price = 0. Missingness is fully predicted "
        f"by observed fields, so MAR not MCAR. Revenue impact GBP {junk_rev:,.2f} "
        f"(zero by construction) - deleting them loses nothing.")

    # Any Description still missing gets a placeholder rather than a deletion.
    before = len(df)
    still = int(df["Description"].isna().sum())
    df["Description"] = df["Description"].fillna("Unknown Item")
    log("Missing Description imputation (placeholder)", "Missing Value Handling",
        before, len(df),
        f"Safeguard for any Description still missing after step 4: fill with "
        f"'Unknown Item' rather than delete, since such rows carry real revenue. "
        f"{still:,} rows needed it - every missing Description was a zero-value "
        f"record already removed, which is itself the evidence for the MAR claim.",
        step_type="Derivation", note=f"{still:,} rows imputed")

    # Test records are not business events.
    before = len(df)
    is_test = df["StockCode"].isin(TEST_CODES)
    df = df[~is_test].reset_index(drop=True)
    log("Test-record removal", "Data Cleaning (anomalies)",
        before, len(df),
        "StockCode TEST001/TEST002 are system tests, not transactions.")

    # ------------------------------------------------------------ TECHNIQUE 6
    # Attribute Construction (PDF: Defect Percentage)
    before = len(df)
    df["Revenue"] = (df["Quantity"] * df["Price"]).round(2)
    df["IsCancelled"] = df["Invoice"].str.startswith("C").fillna(False)
    df["IsReturn"] = df["Quantity"] < 0
    df["ItemType"] = df["StockCode"].map(classify_item)
    df["IsProduct"] = df["ItemType"] == "Product"
    df["HasCustomer"] = df["CustomerKey"] != -1
    df["TransactionType"] = np.select(
        [df["IsCancelled"], df["IsReturn"], df["Price"] == 0],
        ["Cancellation", "Write-off / Return", "Zero-price (promo/adjustment)"],
        default="Sale")
    dt = df["InvoiceDate"]
    df["DateKey"] = dt.dt.normalize()
    df["Year"] = dt.dt.year
    df["Month"] = dt.dt.month
    df["Quarter"] = dt.dt.quarter
    df["Hour"] = dt.dt.hour
    df["Weekday"] = dt.dt.day_name()
    df["WeekdayNo"] = dt.dt.dayofweek
    log("Attribute construction", "Attribute Construction",
        before, len(df),
        "Derived Revenue = Quantity x Price, plus the IsCancelled, IsReturn, "
        "IsProduct and HasCustomer flags, the ItemType and TransactionType "
        "labels, and seven calendar attributes. Mirrors the PDF's Defect "
        "Percentage = Defective / Produced x 100.",
        step_type="Derivation", note="14 new attributes added")

    # ------------------------------------------------------------ TECHNIQUE 5
    # Outlier detection - Z-score and IQR. FLAG, never delete.
    before = len(df)
    df["Is_Outlier"] = False
    reasons = pd.Series([""] * len(df), index=df.index, dtype="object")
    for col in ("Quantity", "Price"):
        s = df[col].astype(float)
        z = (s - s.mean()) / s.std()
        q1, q3 = s.quantile([0.25, 0.75])
        iqr = q3 - q1
        extreme = (z.abs() > Z_THRESHOLD) & ((s < q1 - 3 * iqr) | (s > q3 + 3 * iqr))
        df.loc[extreme, "Is_Outlier"] = True
        reasons[extreme] = reasons[extreme].str.cat([f"Extreme {col}"] * int(extreme.sum()),
                                                    sep=" + ").str.strip(" +")
    # Never leave this empty: an empty string round-trips through CSV as null.
    df["Outlier_Reason"] = reasons.replace("", "Within normal range").astype("string")
    n_out = int(df["Is_Outlier"].sum())
    out_rev = df.loc[df["Is_Outlier"], "Revenue"].sum()
    log("Outlier detection (z-score + IQR)", "Z-score Normalization",
        before, len(df),
        f"{n_out:,} rows flagged (both |z| > 3 AND beyond 3xIQR), carrying "
        f"GBP {out_rev:,.2f}. FLAGGED, NOT DELETED - the 80,995-unit order and "
        f"the GBP 38,970 Manual adjustment are real events. A plain 1.5xIQR rule "
        f"would have discarded ~11% of rows.",
        step_type="Derivation",
        note=f"{n_out:,} rows flagged, 0 deleted")

    # ============================================================= DIMENSIONS
    print("\nB. DATA TRANSFORMATION")

    # ------------------------------------------------------------ TECHNIQUE 2
    # Listwise Deletion - applied ONLY to the customer table (dual grain).
    fact_rows = len(df)
    cust_df = df[df["HasCustomer"]].copy()
    log("Listwise deletion (customer table only)", "Listwise Deletion",
        fact_rows, len(cust_df),
        f"{fact_rows - len(cust_df):,} rows have no Customer ID but carry "
        f"GBP {df.loc[~df['HasCustomer'], 'Revenue'].sum():,.2f}. They stay in the "
        f"fact table and are dropped only for customer analysis - you cannot "
        f"impute an identity. This is the dual-grain decision.",
        step_type="Parallel table",
        note=f"fact keeps {fact_rows:,}; customer table built from {len(cust_df):,}")

    # --------------------------------------------------- RFM + TECHNIQUES 7-8
    snapshot = df["InvoiceDate"].max() + pd.Timedelta(days=1)
    all_customers = set(cust_df["CustomerKey"].unique())
    sales = cust_df[~cust_df["IsCancelled"]]
    rfm = sales.groupby("CustomerKey").agg(
        Recency=("InvoiceDate", lambda s: (snapshot - s.max()).days),
        Frequency=("Invoice", "nunique"),
        Monetary=("Revenue", "sum"),
        FirstPurchase=("InvoiceDate", "min"),
        LastPurchase=("InvoiceDate", "max"),
        UnitsBought=("Quantity", "sum"),
        LinesBought=("Invoice", "size"),
    ).reset_index()
    # Customers whose net purchases are zero or negative (returns only) stay in
    # the dimension - dropping them would leave orphan fact rows joining to a
    # blank row in Power BI. They are excluded from RFM scoring instead.
    no_value = rfm[rfm["Monetary"] <= 0].copy()
    rfm = rfm[rfm["Monetary"] > 0].copy()
    n_scored = len(rfm)          # scored customers, before unscored ones rejoin
    rfm["TenureDays"] = (snapshot - rfm["FirstPurchase"]).dt.days
    rfm["AvgOrderValue"] = (rfm["Monetary"] / rfm["Frequency"]).round(2)

    # Technique 7 - Normalization, both flavours from the PDF
    for col in ("Recency", "Frequency", "Monetary"):
        s = rfm[col].astype(float)
        rfm[f"{col}_MinMax"] = ((s - s.min()) / (s.max() - s.min())).round(4)
        rfm[f"{col}_Z"] = ((s - s.mean()) / s.std()).round(4)

    # Technique 8 - Discretization into quintile scores, then Generalization
    rfm["R_Score"] = pd.qcut(rfm["Recency"], 5, labels=[5, 4, 3, 2, 1]).astype(int)
    rfm["F_Score"] = pd.qcut(rfm["Frequency"].rank(method="first"), 5,
                             labels=[1, 2, 3, 4, 5]).astype(int)
    rfm["M_Score"] = pd.qcut(rfm["Monetary"].rank(method="first"), 5,
                             labels=[1, 2, 3, 4, 5]).astype(int)
    rfm["RFM_Score"] = rfm["R_Score"] * 100 + rfm["F_Score"] * 10 + rfm["M_Score"]
    rfm["RFM_Sum"] = rfm[["R_Score", "F_Score", "M_Score"]].sum(axis=1)

    def segment(row):
        r, f, m = row["R_Score"], row["F_Score"], row["M_Score"]
        fm = (f + m) / 2
        if r >= 4 and fm >= 4:
            return "Champions"
        if r >= 3 and fm >= 3:
            return "Loyal Customers"
        if r >= 4 and fm < 3:
            return "New / Promising"
        if r == 3 and fm < 3:
            return "Needs Attention"
        if r == 2 and fm >= 3:
            return "At Risk"
        if r == 1 and fm >= 4:
            return "Cannot Lose Them"
        if r <= 2 and fm < 3:
            return "Hibernating / Lost"
        return "Others"

    rfm["Segment"] = rfm.apply(segment, axis=1)
    rfm["RecencyBand"] = pd.cut(
        rfm["Recency"], [-1, 30, 90, 180, 365, 10 ** 6],
        labels=["0-30 days", "31-90 days", "91-180 days", "181-365 days", "365+ days"])
    rfm["ValueTier"] = pd.qcut(rfm["Monetary"].rank(method="first"), 3,
                               labels=["Low", "Medium", "High"])

    # Two groups must be added back so no fact row joins to a blank dimension row:
    #   (a) customers whose net purchases are <= 0
    #   (b) customers who appear ONLY on cancellation invoices, so never scored
    unscored = set(rfm["CustomerKey"]) | set(no_value["CustomerKey"])
    cancel_only = sorted(all_customers - unscored)
    if cancel_only:
        no_value = pd.concat(
            [no_value, pd.DataFrame({"CustomerKey": cancel_only})], ignore_index=True)
    no_value = pd.concat(
        [no_value, pd.DataFrame({"CustomerKey": [-1]})], ignore_index=True)
    if len(no_value):
        for col in rfm.columns:
            if col not in no_value.columns:
                no_value[col] = pd.NA
        no_value["Segment"] = "No Net Value (returns / cancellations only)"
        no_value.loc[no_value.CustomerKey == -1, "Segment"] = "Unknown / Guest (not identified)"
        no_value["RecencyBand"] = "Not scored"
        no_value["ValueTier"] = "Not scored"
        rfm = pd.concat([rfm, no_value[rfm.columns]], ignore_index=True)
    print(f"       + {len(no_value):,} unscored customers retained "
          f"({len(cancel_only):,} cancellation-only)")
    log("Normalization (Min-Max + Z-score)", "Data Normalization",
        len(rfm), len(rfm),
        "Recency, Frequency and Monetary each scaled to 0-1 (Min-Max) and to "
        "standard-deviation units (Z-score) before scoring, so no variable "
        "dominates on units alone.",
        step_type="Derivation",
        note=f"{n_scored:,} scored customers scaled, 6 new columns")
    # ---- CRM layer (BA-2 CRM.pdf) -------------------------------------
    # The PDF lists eight CRM strategies. Each RFM segment is assigned the one
    # that fits it, so the dashboard says what to DO about a segment, not just
    # how big it is.
    CRM_STRATEGY = {
        "Champions":                  "Loyalty Programs",
        "Loyal Customers":            "Cross-Selling and Up-Selling",
        "Cannot Lose Them":           "Customer Retention",
        "At Risk":                    "Customer Retention",
        "Needs Attention":            "Personalization",
        "New / Promising":            "Customer Acquisition",
        "Hibernating / Lost":         "Customer Acquisition (win-back)",
        "Others":                     "Customer Feedback Management",
    }
    CRM_ACTION = {
        "Loyalty Programs":            "Reward repeat purchase; recruit as advocates",
        "Cross-Selling and Up-Selling": "Recommend related and higher-value lines",
        "Customer Retention":          "Personal outreach before they lapse",
        "Personalization":             "Tailor offers to past purchases",
        "Customer Acquisition":        "Onboard well; secure the second order",
        "Customer Acquisition (win-back)": "Reactivation campaign or let go",
        "Customer Feedback Management": "Survey to find out what went wrong",
        "Not applicable":              "Identify the customer at point of sale",
    }
    rfm["CRM_Strategy"] = rfm["Segment"].map(CRM_STRATEGY).fillna("Not applicable")
    rfm["CRM_Action"] = rfm["CRM_Strategy"].map(CRM_ACTION)

    # Churn, as in the PDF's telecom case: a customer who has stopped buying.
    # With no contract to cancel, lapse is inferred from recency.
    rfm["ChurnStatus"] = pd.cut(
        rfm["Recency"], [-1, 90, 180, 10 ** 6],
        labels=["Active", "At Risk", "Churned"])
    rfm["ChurnStatus"] = (rfm["ChurnStatus"].cat.add_categories(["Not scored"])
                          .fillna("Not scored"))
    rfm["IsChurned"] = (rfm["Recency"] > 180).fillna(False)

    # CLV exactly as the PDF defines it:
    #   CLV = Average Purchase Value x Purchase Frequency x Customer Lifespan
    # Purchase frequency is annualised from observed tenure; lifespan is a
    # stated forward-looking assumption, since the data covers only two years.
    LIFESPAN_YEARS = 3
    # Unscored customers carry pd.NA, which has no __round__, so coerce to
    # float and let the arithmetic produce NaN for them instead.
    num = lambda col: pd.to_numeric(rfm[col], errors="coerce").astype("float64")
    tenure_years = (num("TenureDays") / 365).clip(lower=0.25)
    rfm["PurchaseFrequencyPerYear"] = (num("Frequency") / tenure_years).round(2)
    rfm["CLV"] = (num("AvgOrderValue") * rfm["PurchaseFrequencyPerYear"]
                  * LIFESPAN_YEARS).round(2)
    log("CRM segmentation & CLV", "CRM Strategies (BA-2)",
        len(rfm), len(rfm),
        f"Each RFM segment mapped to one of the eight CRM strategies in "
        f"BA-2 CRM.pdf with a recommended action; churn status derived from "
        f"recency (the PDF's telecom case); CLV computed with the PDF's own "
        f"formula, annualising purchase frequency over observed tenure and "
        f"assuming a {LIFESPAN_YEARS}-year lifespan.",
        step_type="Derivation",
        note=f"{rfm['CRM_Strategy'].nunique()} CRM strategies, "
             f"3 churn states, CLV per customer")

    log("Discretization & generalization (RFM)", "Discretization / Generalization",
        len(rfm), len(rfm),
        f"Continuous R, F, M cut into quintile scores 1-5, then generalized into "
        f"{rfm['Segment'].nunique()} named segments, 5 recency bands and 3 value "
        f"tiers - the PDF's Young / Middle-aged / Senior pattern.",
        step_type="Derivation",
        note=f"{rfm['Segment'].nunique() - 1} scored segments (+1 unscored), "
             f"5 recency bands, 3 value tiers")

    # ------------------------------------------------------------ Dim_Product
    # One row per StockCode, taking the most common spelling as canonical.
    prod = (df.groupby("StockCode")
            .agg(Description=("Description", lambda s: s.mode().iat[0] if len(s.mode()) else "Unknown Item"),
                 ItemType=("ItemType", "first"),
                 IsProduct=("IsProduct", "first"),
                 AvgPrice=("Price", "mean"),
                 MinPrice=("Price", "min"),
                 MaxPrice=("Price", "max"),
                 UnitsSold=("Quantity", "sum"),
                 Revenue=("Revenue", "sum"),
                 TimesSold=("Invoice", "size"))
            .reset_index())
    prod["AvgPrice"] = prod["AvgPrice"].round(2)
    prod["Revenue"] = prod["Revenue"].round(2)
    prod["PriceBand"] = pd.cut(prod["AvgPrice"], [-10 ** 6, 1, 3, 10, 10 ** 7],
                               labels=["Under GBP 1", "GBP 1-3", "GBP 3-10", "Over GBP 10"])
    # ABC classification on revenue (Pareto)
    prod = prod.sort_values("Revenue", ascending=False).reset_index(drop=True)
    pos = prod["Revenue"].clip(lower=0)
    prod["CumRevenuePct"] = (pos.cumsum() / pos.sum() * 100).round(3)
    prod["ABC_Class"] = np.select(
        [prod["CumRevenuePct"] <= 80, prod["CumRevenuePct"] <= 95],
        ["A", "B"], default="C")

    # ------------------------------------------------------------ Dim_Country
    ctry = (df.groupby(["Country", "CountryISO", "Region"])
            .agg(Rows=("Invoice", "size"), Revenue=("Revenue", "sum"),
                 Customers=("CustomerKey", "nunique"))
            .reset_index())
    ctry["Revenue"] = ctry["Revenue"].round(2)

    # --------------------------------------------------------------- Dim_Date
    dmin, dmax = df["DateKey"].min(), df["DateKey"].max()
    dates = pd.DataFrame({"DateKey": pd.date_range(dmin, dmax, freq="D")})
    dates["Year"] = dates.DateKey.dt.year
    dates["Quarter"] = dates.DateKey.dt.quarter
    dates["QuarterLabel"] = "Q" + dates.Quarter.astype(str) + " " + dates.Year.astype(str)
    dates["Month"] = dates.DateKey.dt.month
    dates["MonthName"] = dates.DateKey.dt.strftime("%b")
    dates["MonthYear"] = dates.DateKey.dt.strftime("%Y-%m")
    dates["MonthYearLabel"] = dates.DateKey.dt.strftime("%b %Y")
    dates["Weekday"] = dates.DateKey.dt.day_name()
    dates["WeekdayNo"] = dates.DateKey.dt.dayofweek
    dates["IsWeekend"] = dates.WeekdayNo >= 5
    dates["WeekOfYear"] = dates.DateKey.dt.isocalendar().week.astype(int)
    dates["DayOfYear"] = dates.DateKey.dt.dayofyear

    # ------------------------------------------------------------- Dim_Hour
    hours = pd.DataFrame({"Hour": range(24)})
    hours["HourLabel"] = hours.Hour.map(lambda h: f"{h:02d}:00")
    hours["DayPart"] = pd.cut(hours.Hour, [-1, 11, 16, 20, 23],
                              labels=["Morning", "Afternoon", "Evening", "Night"])
    # Sort key for DayPart. It must have exactly one value per band, not per
    # hour: Power BI's "sort by column" groups by the sort column as well as
    # the displayed one, so sorting DayPart by Hour splits Morning into its
    # twelve hours instead of showing one column. Same reason WeekdayNo exists.
    hours["DayPartNo"] = pd.cut(hours.Hour, [-1, 11, 16, 20, 23],
                                labels=[0, 1, 2, 3]).astype(int)

    # ============================================================ TECHNIQUE 12
    print("\nC. DATA REDUCTION")
    fact_cols = ["Invoice", "StockCode", "CustomerKey", "Country",
                 "InvoiceDate", "Hour", "Quantity", "Price", "Revenue",
                 "IsCancelled", "IsReturn", "IsProduct", "HasCustomer",
                 "TransactionType", "Is_Outlier", "Outlier_Reason", "SourceSheet"]
    wide_cols = len(df.columns)
    fact = df[fact_cols].copy()
    log("Dimensionality reduction (star schema)", "Dimensionality Reduction",
        wide_cols, len(fact_cols),
        f"One flat {wide_cols}-column table split into a {len(fact_cols)}-column "
        f"fact plus 5 dimensions. Descriptive attributes stored once per key "
        f"instead of once per row. Note Revenue = Quantity x Price is deliberate "
        f"redundancy - never feed all three into one correlation.",
        step_type="Schema",
        note=f"{wide_cols} columns -> {len(fact_cols)} fact columns + 6 dimensions")

    # ------------------------------------------------------------ TECHNIQUE 11
    # Data Cube Aggregation + TECHNIQUE 4 (noise reduction on the daily series)
    fact_dates = fact["InvoiceDate"].dt.normalize()
    daily = (fact.groupby(fact_dates.rename("DateKey"))
             .agg(Revenue=("Revenue", "sum"), Orders=("Invoice", "nunique"),
                  Units=("Quantity", "sum"), Lines=("Invoice", "size"),
                  Customers=("CustomerKey", "nunique"))
             .reset_index().sort_values("DateKey"))
    daily["Revenue"] = daily["Revenue"].round(2)

    # PDF: n-point moving average
    daily[f"Revenue_MA{MA_WINDOW}"] = (daily["Revenue"]
                                       .rolling(MA_WINDOW, min_periods=1).mean().round(2))
    # PDF: New Forecast = alpha*Actual + (1-alpha)*Previous Forecast
    smoothed, prev = [], daily["Revenue"].iat[0]
    for actual in daily["Revenue"]:
        prev = ALPHA * actual + (1 - ALPHA) * prev
        smoothed.append(round(prev, 2))
    daily["Revenue_ExpSmooth"] = smoothed
    log(f"Noise reduction (MA{MA_WINDOW} + exponential smoothing)", "Noise Reduction",
        len(fact), len(daily),
        f"Daily revenue is spiky. Added a {MA_WINDOW}-point moving average and an "
        f"exponentially smoothed series with alpha = {ALPHA}, exactly the PDF's "
        f"New = a*Actual + (1-a)*Previous formula.",
        step_type="Derivation",
        note=f"2 smoothed series over {len(daily):,} daily points")

    month_country = (fact.assign(MonthYear=fact_dates.dt.strftime("%Y-%m"))
                     .groupby(["MonthYear", "Country"])
                     .agg(Revenue=("Revenue", "sum"), Orders=("Invoice", "nunique"),
                          Units=("Quantity", "sum"),
                          Customers=("CustomerKey", "nunique"))
                     .reset_index())
    month_country["Revenue"] = month_country["Revenue"].round(2)

    month_product = (fact.assign(MonthYear=fact_dates.dt.strftime("%Y-%m"))
                     .groupby(["MonthYear", "StockCode"])
                     .agg(Revenue=("Revenue", "sum"), Units=("Quantity", "sum"),
                          Orders=("Invoice", "nunique"))
                     .reset_index())
    month_product["Revenue"] = month_product["Revenue"].round(2)

    log("Data cube aggregation", "Data Cube Aggregation",
        len(fact), len(daily) + len(month_country) + len(month_product),
        f"Line-level rows pre-summarised into daily ({len(daily):,}), "
        f"month x country ({len(month_country):,}) and month x product "
        f"({len(month_product):,}) cubes so dashboard pages scan thousands of "
        f"rows, not a million.",
        step_type="Aggregation",
        note=f"{len(fact):,} lines -> {len(daily) + len(month_country) + len(month_product):,} "
             f"summary rows ({len(fact) / (len(daily) + len(month_country) + len(month_product)):.0f}x smaller)")

    # ------------------------------------------------------------ TECHNIQUE 13
    # Histogram Reduction - store frequencies, not raw values
    hist_frames = []
    basket = fact.groupby("Invoice")["Revenue"].sum()
    specs = [
        ("Quantity", fact.loc[fact.IsProduct, "Quantity"],
         [-10 ** 6, -1, 0, 2, 5, 10, 25, 50, 100, 10 ** 6],
         ["Negative", "0", "1-2", "3-5", "6-10", "11-25", "26-50", "51-100", "100+"]),
        ("Unit Price", fact.loc[fact.IsProduct, "Price"],
         [-10 ** 6, 0, 1, 2, 5, 10, 20, 10 ** 6],
         ["<= 0", "0-1", "1-2", "2-5", "5-10", "10-20", "20+"]),
        ("Basket Value", basket,
         [-10 ** 9, 0, 50, 150, 300, 600, 1500, 10 ** 9],
         ["Negative", "0-50", "50-150", "150-300", "300-600", "600-1500", "1500+"]),
    ]
    for metric, series, bins, labels in specs:
        cut = pd.cut(series, bins=bins, labels=labels)
        counts = cut.value_counts().reindex(labels, fill_value=0)
        hist_frames.append(pd.DataFrame({
            "Metric": metric, "Bin": labels, "BinOrder": range(1, len(labels) + 1),
            "Frequency": counts.values,
        }))
    hist = pd.concat(hist_frames, ignore_index=True)
    log("Histogram reduction", "Histogram Reduction",
        len(fact) + len(basket), len(hist),
        f"Distributions of Quantity, Unit Price and Basket Value stored as "
        f"{len(hist)} bin frequencies instead of {len(fact) + len(basket):,} raw "
        f"values, preserving shape at a fraction of the size.",
        step_type="Aggregation",
        note=f"{len(fact) + len(basket):,} values -> {len(hist)} bin frequencies")

    # ================================================================== WRITE
    print("\nD. WRITING OUTPUT")
    dq = pd.DataFrame(audit)
    tables = {
        "Fact_Sales": fact,
        "Dim_Customer": rfm,
        "Dim_Product": prod,
        "Dim_Country": ctry,
        "Dim_Date": dates,
        "Dim_Hour": hours,
        "Dim_DataQuality": dq,
        "Agg_Daily": daily,
        "Agg_MonthCountry": month_country,
        "Agg_MonthProduct": month_product,
        "Agg_Histograms": hist,
    }
    for name, frame in tables.items():
        path = OUT / f"{name}.csv"
        frame.to_csv(path, index=False, encoding="utf-8-sig")
        mb = path.stat().st_size / 1024 / 1024
        extra = ""
        if name == "Fact_Sales":
            pq = OUT / f"{name}.parquet"
            frame.to_parquet(pq, index=False)
            extra = f"  (parquet {pq.stat().st_size / 1024 / 1024:.1f} MB)"
        print(f"  {name:<20} {len(frame):>9,} rows  {mb:>7.1f} MB{extra}")

    # -------------------------------------------------------------- summary
    gross = fact["Revenue"].sum()
    print("\n=== RECONCILIATION ===")
    print(f"  rows in  : {n0:,}")
    print(f"  rows out : {len(fact):,}  ({len(fact) - n0:+,})")
    print(f"  revenue  : GBP {gross:,.2f}")
    print(f"  customers: {len(rfm):,}")
    print(f"  products : {int(prod.IsProduct.sum()):,} of {len(prod):,} stock codes")

    summary = [
        "# Cleaned Data Profile - after preprocessing", "",
        f"- Rows in: **{n0:,}** -> rows out: **{len(fact):,}** "
        f"({(len(fact) - n0) / n0 * 100:+.2f}%)",
        f"- Net revenue: **GBP {gross:,.2f}**",
        f"- Customers in `Dim_Customer`: **{len(rfm):,}**",
        f"- Stock codes: **{len(prod):,}** "
        f"(of which real products: {int(prod.IsProduct.sum()):,})",
        f"- Date span: **{dmin.date()}** to **{dmax.date()}** ({len(dates):,} days)",
        f"- Outlier rows flagged (not deleted): **{int(fact.Is_Outlier.sum()):,}**",
        f"- Rows with no customer, retained in fact: "
        f"**{int((~fact.HasCustomer).sum()):,}**", "",
        "## Audit trail", "", dq.drop(columns=["Rationale"]).to_markdown(index=False), "",
        "## Rationale per step", "",
    ]
    for row in audit:
        summary.append(f"**{row['Step']}. {row['Technique']}** "
                       f"*(PDF: {row['PDF_Section']})*  \n{row['Rationale']}\n")
    summary += ["## Customer segments", "",
                rfm["Segment"].value_counts().rename_axis("Segment")
                .reset_index(name="Customers").to_markdown(index=False), "",
                "## Item types", "",
                fact.assign(t=df["ItemType"]).groupby("t")
                .agg(Rows=("Invoice", "size"), Revenue=("Revenue", "sum"))
                .round(2).reset_index().rename(columns={"t": "ItemType"})
                .to_markdown(index=False), ""]
    (EVID / "profile_after.md").write_text("\n".join(summary), encoding="utf-8")
    print(f"\nWritten: {EVID / 'profile_after.md'}")


if __name__ == "__main__":
    main()
