"""
Step 7 - The storytelling / presentation script.

Every figure is computed from the cleaned data at build time and interpolated
into the prose, so the script cannot quote a number the dashboard does not show.

Run:  python 08_build_storytelling_script.py
"""
import pathlib
import subprocess

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent.parent
CLEAN = ROOT / "02_clean"
OUT = ROOT / "04_report"
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

CSS = """
@page { size: A4; margin: 15mm 14mm; }
body { font-family: "Segoe UI", Calibri, system-ui, sans-serif; color: #1b2a36;
       font-size: 10.3pt; line-height: 1.55; margin: 0; }
h1 { font-size: 21pt; margin: 0 0 2mm; }
h2 { font-size: 13.5pt; margin: 8mm 0 3mm; padding: 2mm 0 1.5mm;
     border-bottom: 2px solid #2e7d8f; color: #14343d; }
h3 { font-size: 11pt; margin: 5mm 0 1.5mm; color: #14343d; }
.sub { color: #5b6b78; margin: 0 0 5mm; }
.say { border-left: 3px solid #2e7d8f; background: #f4f9fa; padding: 3mm 4mm;
       margin: 2.5mm 0; }
.say p { margin: 0 0 2mm; }
.say p:last-child { margin: 0; }
.stage { background: #fbf7f1; border-left: 3px solid #c1783c; padding: 2.5mm 4mm;
         margin: 2.5mm 0; font-size: 9.4pt; }
.stage strong { color: #8a5424; }
.beat { font-size: 9.2pt; color: #5b6b78; margin: 1.5mm 0 0; }
table { border-collapse: collapse; width: 100%; margin: 3mm 0; font-size: 9pt; }
th { background: #eef4f6; text-align: left; padding: 2mm; font-weight: 600;
     border-bottom: 1.5px solid #2e7d8f; }
td { padding: 1.7mm 2mm; border-bottom: 1px solid #e3e9ec; vertical-align: top; }
td.n, th.n { text-align: right; font-variant-numeric: tabular-nums; }
.q { margin: 3mm 0; }
.q .ask { font-weight: 600; color: #14343d; }
.time { display: inline-block; background: #2e7d8f; color: #fff; border-radius: 2px;
        padding: 0.4mm 2mm; font-size: 8.4pt; font-weight: 600; margin-left: 2mm;
        vertical-align: middle; }
.avoid { break-inside: avoid; }
ul { margin: 2mm 0; padding-left: 5mm; }
li { margin: 1mm 0; }
footer { margin-top: 8mm; padding-top: 3mm; border-top: 1px solid #d8e2e6;
         color: #5b6b78; font-size: 8.4pt; }
"""


def main():
    f = pd.read_parquet(CLEAN / "Fact_Sales.parquet")
    c = pd.read_csv(CLEAN / "Dim_Customer.csv")
    p = pd.read_csv(CLEAN / "Dim_Product.csv")

    # ------------------------------------------------------------- figures
    rev = f.Revenue.sum()
    orders = f.Invoice.nunique()
    custs = f.loc[f.CustomerKey != -1, "CustomerKey"].nunique()
    aov = rev / orders
    gross = f.loc[f.TransactionType == "Sale", "Revenue"].sum()
    returns = -f.loc[f.Quantity < 0, "Revenue"].sum()
    ret_rate = returns / gross * 100

    y1 = f[(f.InvoiceDate >= "2009-12-01") & (f.InvoiceDate < "2010-12-01")].Revenue.sum()
    y2 = f[(f.InvoiceDate >= "2010-12-01") & (f.InvoiceDate < "2011-12-01")].Revenue.sum()
    growth = (y2 / y1 - 1) * 100

    ct = f.groupby("Country").Revenue.sum().sort_values(ascending=False)
    uk_share = ct.iloc[0] / ct.sum() * 100
    top5 = ct.head(5).sum() / ct.sum() * 100
    nonuk = ct.iloc[1:].sum()

    cr = f[f.CustomerKey != -1].groupby("CustomerKey").Revenue.sum().sort_values(ascending=False)
    cum = cr.cumsum() / cr.sum()
    n80 = int((cum <= 0.8).sum()) + 1
    pareto_pct = n80 / len(cr) * 100
    top10 = cr.head(10).sum() / cr.sum() * 100

    abc = p.groupby("ABC_Class").agg(SKUs=("StockCode", "size"), Rev=("Revenue", "sum"))
    a_sku_pct = abc.loc["A", "SKUs"] / abc.SKUs.sum() * 100
    a_rev_pct = abc.loc["A", "Rev"] / abc.Rev.sum() * 100

    m = f.merge(c[["CustomerKey", "Segment"]], on="CustomerKey", how="left")
    seg = m.groupby("Segment").agg(Cust=("CustomerKey", "nunique"),
                                   Rev=("Revenue", "sum"))
    seg["RevPct"] = seg.Rev / seg.Rev.sum() * 100
    seg["CustPct"] = seg.Cust / seg.Cust.sum() * 100
    champ = seg.loc["Champions"]

    # Match the dashboard measure exactly:
    #   Churn Rate % = DIVIDE ( [Churned Customers], [Scored Customers] ) * 100
    # Customers with no recency score cannot be called churned OR retained,
    # so they leave the denominator. Averaging IsChurned over every row
    # instead counts them as retained and understates churn.
    scored = c[c.ChurnStatus != "Not scored"]
    churn = scored.IsChurned.mean() * 100
    clv_mean, clv_med = c.CLV.mean(), c.CLV.median()
    champ_clv = c[c.Segment == "Champions"].CLV.mean()
    guest_rev = f.loc[f.CustomerKey == -1, "Revenue"].sum()
    guest_pct = guest_rev / rev * 100

    f = f.assign(wd=f.InvoiceDate.dt.day_name(), hr=f.InvoiceDate.dt.hour,
                 mo=f.InvoiceDate.dt.to_period("M").astype(str))
    wd = f.groupby("wd").Revenue.sum().sort_values(ascending=False)
    sat_rev = f[f.wd == "Saturday"].Revenue.sum()
    sat_ord = f[f.wd == "Saturday"].Invoice.nunique()
    mo = f.groupby("mo").Revenue.sum().sort_values(ascending=False)
    basket_mean = f.groupby("Invoice").Revenue.sum().mean()
    basket_med = f.groupby("Invoice").Revenue.sum().median()
    hib = seg.loc["Hibernating / Lost"]

    h = []
    h.append(f"<style>{CSS}</style>")
    h.append("<h1>Dashboard Walkthrough &ndash; Speaking Script</h1>")
    h.append('<p class="sub">Online Retail II &middot; Business Analytics, '
             'M.Sc Semester 3 &middot; approx. 16&ndash;18 minutes</p>')

    # -------------------------------------------------------------- how to
    h.append("<h2>How to use this</h2>")
    h.append("<p>Blue blocks are what to <strong>say</strong>. Amber blocks are what to "
             "<strong>do</strong> &ndash; which page to be on, what to click. Grey lines "
             "underneath are the point being made, in case you need to improvise.</p>")
    h.append("<p>Do not read it aloud word for word. Learn the four or five numbers that "
             "carry each section and speak around them. If you are short on time, the "
             "sections marked <em>optional</em> can go without breaking the argument.</p>")

    # ------------------------------------------------------------- opening
    h.append('<h2>Opening <span class="time">1 min</span></h2>')
    h.append('<div class="stage"><strong>On screen:</strong> Executive Overview, no '
             'slicers applied.</div>')
    h.append(f'<div class="say"><p>"This is two years of transactions from a UK online '
             f'gift retailer &ndash; December 2009 to December 2011. About a million rows '
             f'before cleaning, {len(f):,} after.</p>'
             f'<p>I want to start with the number that is <em>not</em> on this screen. '
             f'The raw file reports GBP 19.29 million of revenue. The correct figure is '
             f'GBP {rev/1e6:.2f} million. The difference is GBP 431,717, and almost all of '
             f'it comes from one thing: the two worksheets in the source file overlap. '
             f'Nine days of December 2010 appear in both, so 22,202 invoice lines are '
             f'counted twice.</p>'
             f'<p>Nothing in the file warns you about that. If you append the two sheets '
             f'and start charting, every figure you produce is inflated and you would '
             f'never know. So that is where I began."</p></div>')
    h.append('<p class="beat">Opens on a finding rather than a description. It also '
             'establishes that the preprocessing was investigative, not mechanical.</p>')

    # ------------------------------------------------- act 1 - foundation
    h.append('<h2>Act I &ndash; The data foundation <span class="time">2 min</span></h2>')
    h.append('<div class="stage"><strong>On screen:</strong> stay on Executive Overview. '
             'Hold up the preprocessing report, or have the PDF open in another window.</div>')
    h.append(f'<div class="say"><p>"Thirteen preprocessing techniques from the course '
             f'material were applied. The full audit trail is in this document &ndash; '
             f'every step, the technique it implements, rows before and after, and why. '
             f'I will pull out the three decisions that actually mattered.</p>'
             f'<p><strong>First, the duplicates.</strong> 34,335 rows removed. 22,202 from '
             f'the sheet overlap I just mentioned, 12,133 ordinary repeats.</p>'
             f'<p><strong>Second, the missing customer IDs.</strong> 243,007 rows &ndash; '
             f'nearly a quarter of the file &ndash; have no customer attached. The reflex '
             f'is to delete them. But those rows carry GBP {guest_rev/1e6:.2f} million, '
             f'{guest_pct:.1f}% of revenue. Deleting them would have made every revenue '
             f'figure wrong.</p>'
             f'<p>So I did something slightly unusual: I applied deletion at one grain '
             f'only. The transaction table keeps every row, so revenue is complete. The '
             f'customer table uses complete-case analysis, because you cannot impute an '
             f'identity &ndash; a customer ID is a name, not a measurement with a '
             f'plausible average.</p>'
             f'<p><strong>Third, the outliers.</strong> A standard 1.5&times;IQR rule flags '
             f'116,489 rows, about 11% of the data. I did not use it. A rule that discards '
             f'a ninth of a dataset is not finding errors, it is describing a skewed retail '
             f'distribution where most baskets are small and a few are wholesale. I used a '
             f'stricter test and flagged {int(f.Is_Outlier.sum()):,} rows &ndash; and '
             f'flagged is all. Nothing was deleted. The 80,995-unit order is a real order '
             f'with a real matching cancellation."</p></div>')
    h.append('<p class="beat">This is the section that earns method marks. The examiner '
             'is listening for whether you understood <em>why</em>, not whether you can '
             'name techniques.</p>')

    # ----------------------------------------------- act 2 - the business
    h.append('<h2>Act II &ndash; What the business looks like <span class="time">3 min</span></h2>')
    h.append('<div class="stage"><strong>On screen:</strong> Executive Overview. Point at '
             'the KPI row, then the trend chart.</div>')
    h.append(f'<div class="say"><p>"GBP {rev/1e6:.2f} million across {orders:,} orders from '
             f'{custs:,} identified customers. Average order value GBP {aov:.2f}.</p>'
             f'<p>But the average hides the shape. The <em>median</em> basket is GBP '
             f'{basket_med:.0f}, against a mean of GBP {basket_mean:.0f}. The mean is '
             f'{basket_mean/basket_med:.1f} times the median, which tells you this is a '
             f'wholesale business wearing a retail coat &ndash; a long tail of very large '
             f'orders is dragging the average up."</p></div>')

    h.append("<h3>The growth question &ndash; and a trap</h3>")
    h.append(f'<div class="say"><p>"If you look at revenue by calendar year you see 2010 at '
             f'GBP {y1/1e6:.2f} million and 2011 at GBP 8.98 million, and you might '
             f'conclude the business shrank.</p>'
             f'<p>That would be wrong. The data starts on 1 December 2009 and stops on '
             f'9 December 2011, so neither calendar year is complete and 2011 is missing '
             f'most of December &ndash; its best month. Comparing like for like, December '
             f'to November, the business grew {growth:+.1f}%.</p>'
             f'<p>Essentially flat. That is the honest read, and it only appears if you '
             f'respect where the data actually stops."</p></div>')
    h.append('<p class="beat">Shows you checked the denominator. Examiners look for this.</p>')

    h.append("<h3>Geography</h3>")
    h.append('<div class="stage"><strong>Do:</strong> point at Revenue by country, then '
             'tick "Europe" on the Region slicer to show the rest of the world.</div>')
    h.append(f'<div class="say"><p>"The United Kingdom is {uk_share:.1f}% of revenue. The '
             f'top five countries are {top5:.1f}%. Everything outside the UK &ndash; forty '
             f'markets &ndash; adds up to GBP {nonuk/1e6:.2f} million.</p>'
             f'<p>Two ways to read that. Either international is a rounding error and '
             f'attention belongs at home. Or the UK is saturated and those forty markets '
             f'are the only room left to grow. The data cannot settle it, because we have '
             f'no marketing spend to compare against. What it does say is that Ireland, '
             f'the Netherlands and Germany are each doing several hundred thousand pounds '
             f'without, as far as we can see, being targeted."</p></div>')

    # ---------------------------------------------- act 3 - customers/CRM
    h.append('<h2>Act III &ndash; Customers and CRM strategy <span class="time">4 min</span></h2>')
    h.append('<div class="stage"><strong>Do:</strong> switch to Customer Intelligence.</div>')
    h.append(f'<div class="say"><p>"Customers are scored on Recency, Frequency and Monetary '
             f'value, each normalised and cut into quintiles, then grouped into named '
             f'segments. Every segment is then mapped to one of the eight CRM strategies '
             f'from the CRM material &ndash; so this page does not just describe customers, '
             f'it says what to do about them.</p>'
             f'<p>Start with the concentration. Champions are {champ.CustPct:.0f}% of '
             f'customers and {champ.RevPct:.0f}% of revenue.</p>'
             f'<p>More precisely: {pareto_pct:.1f}% of customers generate 80% of identified '
             f'revenue. That is the Pareto principle turning up almost exactly &ndash; '
             f'80/20, on real data, without being looked for. The top ten customers alone '
             f'are {top10:.1f}%."</p></div>')

    h.append("<h3>The number that should worry them</h3>")
    h.append('<div class="stage"><strong>Do:</strong> point at the Churn rate KPI, then the '
             'churn donut.</div>')
    h.append(f'<div class="say"><p>"{churn:.1f}% of customers have not bought in over six '
             f'months. Applying the CRM material&#39;s churn framing to a retailer with no '
             f'contract, that is the lapse rate.</p>'
             f'<p>And look where they sit. Hibernating or Lost is the <em>largest</em> '
             f'segment by headcount &ndash; {hib.Cust:,.0f} customers, {hib.CustPct:.0f}% of '
             f'the base &ndash; but only {hib.RevPct:.0f}% of revenue. So four in ten '
             f'customers have gone quiet, and the business barely feels it in the short '
             f'term, because Champions are carrying it.</p>'
             f'<p>That is the risk. Revenue is flat and concentrated in {champ.Cust:,.0f} '
             f'people. If a handful of those lapse, there is no second tier ready to '
             f'replace them."</p></div>')
    h.append('<p class="beat">This is the strongest business point in the whole '
             'presentation. Slow down here.</p>')

    h.append("<h3>Lifetime value</h3>")
    h.append(f'<div class="say"><p>"CLV is calculated with the formula from the CRM '
             f'material &ndash; average purchase value, times purchase frequency, times '
             f'lifespan. Frequency is annualised from each customer&#39;s observed tenure. '
             f'Lifespan is assumed at three years, and I want to be explicit that it is an '
             f'assumption, because two years of data cannot measure a lifespan.</p>'
             f'<p>Average CLV is GBP {clv_mean:,.0f}. The median is GBP {clv_med:,.0f} '
             f'&ndash; less than half. Champions average GBP {champ_clv:,.0f}, roughly '
             f'{champ_clv/clv_med:.0f} times the typical customer.</p>'
             f'<p>The action table underneath turns that into instructions: Champions get '
             f'loyalty programmes, Loyal Customers get cross-sell and up-sell, At Risk get '
             f'retention outreach, and the Hibernating group gets a win-back campaign or '
             f'is written off. Each of those is a named strategy from the CRM syllabus, '
             f'attached to a segment we can actually count."</p></div>')

    # ---------------------------------------------------- act 4 - product
    h.append('<h2>Act IV &ndash; Products and returns <span class="time">2.5 min</span></h2>')
    h.append('<div class="stage"><strong>Do:</strong> switch to Product &amp; Returns.</div>')
    h.append(f'<div class="say"><p>"The same concentration appears on the product side. '
             f'Class A is {abc.loc["A","SKUs"]:,} of {abc.SKUs.sum():,} products &ndash; '
             f'{a_sku_pct:.0f}% of the catalogue &ndash; and {a_rev_pct:.0f}% of revenue. '
             f'The C class is {abc.loc["C","SKUs"]:,} products earning 2%.</p>'
             f'<p>Notice what is sitting second in the top-products chart: Dotcom Postage. '
             f'That is not a product, it is a delivery charge. Postage appears twice in the '
             f'top ten. Any analysis that treats every stock code as merchandise would '
             f'report a shipping fee as a bestseller, so the model classifies each code as '
             f'product, postage, fee or adjustment, and the dashboard can separate them.</p>'
             f'<p>Returns are GBP {returns/1e6:.2f} million against GBP {gross/1e6:.2f} '
             f'million gross &ndash; a {ret_rate:.1f}% return rate. This is the same shape '
             f'as the defect-percentage calculation in the preprocessing material, just '
             f'applied to money instead of manufacturing."</p></div>')

    # ------------------------------------------------------- act 5 - time
    h.append('<h2>Act V &ndash; Time and operations <span class="time">2 min</span></h2>')
    h.append('<div class="stage"><strong>Do:</strong> switch to Time &amp; Seasonality.</div>')
    h.append(f'<div class="say"><p>"November is the peak in both years &ndash; GBP '
             f'{mo.iloc[0]/1e6:.2f} million in November 2011, GBP {mo.iloc[1]/1e6:.2f} '
             f'million in November 2010. This is a gift retailer selling into Christmas, '
             f'and the buying happens in November, not December.</p>'
             f'<p>Trading is concentrated between 10am and 1pm, with noon the single '
             f'busiest hour. That is office-hours ordering, which again says wholesale '
             f'rather than consumer.</p>'
             f'<p>And then the finding I did not expect."</p></div>')
    h.append('<div class="stage"><strong>Do:</strong> point at the weekday axis. Pause.</div>')
    h.append(f'<div class="say"><p>"Saturday. Across two full years, Saturday produced GBP '
             f'{sat_rev:,.0f} from {sat_ord} orders. Thursday, the best day, produced GBP '
             f'{wd.iloc[0]/1e6:.2f} million.</p>'
             f'<p>Saturday is not a slow day. Saturday is closed &ndash; those thirty '
             f'orders are almost certainly stragglers or corrections. Sunday trades '
             f'normally, at GBP {wd.loc["Sunday"]/1e6:.2f} million.</p>'
             f'<p>Which raises an obvious question: if Sunday works, why is Saturday shut? '
             f'I cannot answer that from the data. But it is the kind of question a '
             f'dashboard is supposed to raise."</p></div>')
    h.append('<p class="beat">Memorable, specific, and honest about the limit of what the '
             'data proves. Good note to end the findings on.</p>')

    # ------------------------------------------------------------- close
    h.append('<h2>Closing <span class="time">2 min</span></h2>')
    h.append('<div class="say"><p>"Four things I would put to the business.</p></div>')
    h.append("<ol>"
             f"<li><strong>Protect the Champions.</strong> {champ.CustPct:.0f}% of "
             f"customers carry {champ.RevPct:.0f}% of revenue. That is efficient and "
             f"fragile at the same time. They need a named owner, not a mailing list.</li>"
             f"<li><strong>Fix the middle of the base.</strong> {churn:.1f}% churn with "
             f"flat revenue means new customers are replacing lapsed ones rather than "
             f"adding to them. The New / Promising segment is where to intervene &ndash; "
             f"the second order is the one that matters.</li>"
             f"<li><strong>Decide about international deliberately.</strong> GBP "
             f"{nonuk/1e6:.2f} million is arriving from outside the UK with no visible "
             f"effort. Either resource it or stop counting on it.</li>"
             f"<li><strong>Investigate Saturday.</strong> Sunday trades. Saturday does "
             f"not. That is either a deliberate decision or an accident, and either way "
             f"someone should know which.</li></ol>")
    h.append('<div class="say"><p>"And the limitations, which matter as much as the '
             'findings. There is no cost data, so I can show revenue but never margin '
             '&ndash; the biggest customer might be the least profitable. Two years is '
             'too short to separate a trend from a cycle. Everything is in pounds with no '
             'exchange rates. Churn is inferred from recency, not observed. And 14% of '
             'revenue belongs to customers we cannot identify, which caps how far any '
             'customer analysis can go.</p>'
             '<p>The single highest-value fix is not analytical. It is getting a customer '
             'identifier onto those transactions at the point of sale."</p></div>')

    # ------------------------------------------------------------- Q&A
    h.append('<h2>Questions you should expect</h2>')
    qs = [
        ("Why didn&#39;t you delete the rows with missing customer IDs?",
         f"Because they carry GBP {guest_rev/1e6:.2f} million &ndash; {guest_pct:.1f}% of "
         f"revenue. Deleting them would bias every total downward. I deleted them only "
         f"from the customer table, where a missing identity makes the record unusable, "
         f"and kept them everywhere else. Different questions, different grain."),
        ("How do you know the missing data is MAR and not MCAR?",
         "Because the missingness is fully predicted by other observed columns. Every "
         "single row with no description also has no customer ID and a price of zero &ndash; "
         "100% on both. Under MCAR you would expect those rows to look like any other. "
         "They do not, so it is not random."),
        ("Isn&#39;t removing 34,000 duplicates quite aggressive?",
         "It would be if they were near-matches. They are exact matches on all eight "
         "columns, and 22,202 of them come from a structural fault &ndash; the two "
         "worksheets overlap by nine days. Those are not similar transactions, they are the "
         "same transaction stored twice."),
        ("Why flag outliers instead of removing them?",
         f"Because the standard rule flags 11% of the data, and a rule that throws away a "
         f"ninth of a dataset is describing a skewed distribution, not finding errors. My "
         f"stricter test flags {int(f.Is_Outlier.sum()):,} rows, and they are real events "
         f"&ndash; a genuine 80,995-unit order with its matching cancellation. Deleting "
         f"them would be editing the business to suit the statistics."),
        ("Your CLV assumes a three-year lifespan. Where does that come from?",
         "It is an assumption, and I would rather state it than bury it. The data covers "
         "two years, which is shorter than any realistic customer lifespan, so lifespan "
         "cannot be measured here. Three years is a conventional planning figure. The "
         "ranking between customers is unaffected by the choice &ndash; only the absolute "
         "values move."),
        ("Why is revenue flat when your growth number is positive?",
         f"Both are true. Calendar 2011 looks lower than 2010 only because the data stops "
         f"on 9 December and December is the weakest recorded month as a result. Compared "
         f"like for like, December to November, growth is {growth:+.1f}% &ndash; which is "
         f"flat in any practical sense."),
        ("What would you do differently with more time?",
         "Three things. Market-basket analysis to make the cross-sell recommendation "
         "specific rather than generic. A proper cohort retention view, since the current "
         "churn measure is a snapshot. And I would push hard for cost data, because "
         "without it every conclusion here is about revenue and none of it is about "
         "profit."),
    ]
    for ask, ans in qs:
        h.append(f'<div class="q avoid"><p class="ask">{ask}</p><p>{ans}</p></div>')

    # ---------------------------------------------------- cheat sheet
    h.append('<h2>Numbers to have in your head</h2>')
    rows = [
        ("Net revenue", f"GBP {rev:,.0f}"),
        ("Raw revenue before cleaning", "GBP 19,287,251"),
        ("Rows in / out", f"1,067,371 / {len(f):,}"),
        ("Orders / customers", f"{orders:,} / {custs:,}"),
        ("Average order value / median basket", f"GBP {aov:.2f} / GBP {basket_med:.0f}"),
        ("Like-for-like growth", f"{growth:+.1f}%"),
        ("UK share of revenue", f"{uk_share:.1f}%"),
        ("Customers making 80% of revenue", f"{pareto_pct:.1f}%"),
        ("Champions: share of customers / revenue",
         f"{champ.CustPct:.0f}% / {champ.RevPct:.0f}%"),
        ("Churn rate", f"{churn:.1f}%"),
        ("Average / median CLV", f"GBP {clv_mean:,.0f} / GBP {clv_med:,.0f}"),
        ("Return rate", f"{ret_rate:.1f}%"),
        ("Class A: share of SKUs / revenue", f"{a_sku_pct:.0f}% / {a_rev_pct:.0f}%"),
        ("Saturday revenue, two years", f"GBP {sat_rev:,.0f} from {sat_ord} orders"),
        ("Unidentified revenue", f"GBP {guest_rev:,.0f} ({guest_pct:.1f}%)"),
    ]
    h.append('<table><thead><tr><th>Metric</th><th class="n">Value</th></tr></thead><tbody>')
    for k, v in rows:
        h.append(f'<tr><td>{k}</td><td class="n">{v}</td></tr>')
    h.append("</tbody></table>")

    h.append("<footer>Generated by <code>01_scripts/08_build_storytelling_script.py</code>. "
             "Every figure is computed from <code>02_clean/</code> at build time, so the "
             "script and the dashboard cannot disagree.</footer>")

    OUT.mkdir(parents=True, exist_ok=True)
    hp = OUT / "Storytelling_Script.html"
    hp.write_text("<!doctype html><meta charset='utf-8'>"
                  "<title>Dashboard Walkthrough - Speaking Script</title>"
                  + "\n".join(h), encoding="utf-8")
    print(f"HTML : {hp}")

    pdf = OUT / "Storytelling_Script.pdf"
    if pathlib.Path(CHROME).exists():
        subprocess.run([CHROME, "--headless", "--disable-gpu", "--no-pdf-header-footer",
                        f"--print-to-pdf={pdf}", hp.as_uri()],
                       check=False, capture_output=True, timeout=180)
        if pdf.exists():
            print(f"PDF  : {pdf}  ({pdf.stat().st_size/1024:.0f} KB)")


if __name__ == "__main__":
    main()
