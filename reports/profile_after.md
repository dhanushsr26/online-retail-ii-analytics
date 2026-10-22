# Cleaned Data Profile - after preprocessing

- Rows in: **1,067,371** -> rows out: **1,028,743** (-3.62%)
- Net revenue: **GBP 18,854,778.33**
- Customers in `Dim_Customer`: **5,941**
- Stock codes: **4,776** (of which real products: 4,737)
- Date span: **2009-12-01** to **2011-12-09** (739 days)
- Outlier rows flagged (not deleted): **1,751**
- Rows with no customer, retained in fact: **230,876**

## Audit trail

|   Step | Technique                                     | PDF_Section                     | Step_Type      |   Rows_Before |    Rows_After |   Rows_Removed |   Pct_Removed | Metric_Note                                                                                                            |
|-------:|:----------------------------------------------|:--------------------------------|:---------------|--------------:|--------------:|---------------:|--------------:|:-----------------------------------------------------------------------------------------------------------------------|
|      0 | Raw load (both sheets unioned)                | Data Integration                | Row filter     |   1.06737e+06 |   1.06737e+06 |              0 |        0      |                                                                                                                        |
|      1 | Cross-sheet overlap resolution                | Data Integration                | Row filter     |   1.06737e+06 |   1.04517e+06 |          22202 |        2.0801 |                                                                                                                        |
|      2 | Duplicate removal (within sheet)              | Duplicate Removal               | Row filter     |   1.04517e+06 |   1.03304e+06 |          12133 |        1.1609 |                                                                                                                        |
|      3 | Unit conversion & format standardisation      | Unit Conversion                 | Derivation     | nan           | nan           |              0 |      nan      | 19,247 rows remapped to canonical country names (EIRE->Ireland, RSA->South Africa, Unspecified->Unknown) + ISO-3 codes |
|      4 | Duplicate removal (transaction business key)  | Duplicate Removal               | Row filter     |   1.03304e+06 |   1.03303e+06 |              2 |        0.0002 |                                                                                                                        |
|      5 | Missing-value removal (MAR, zero-value)       | Missing Value Types (MAR)       | Row filter     |   1.03303e+06 |   1.02876e+06 |           4275 |        0.4138 |                                                                                                                        |
|      6 | Missing Description imputation (placeholder)  | Missing Value Handling          | Derivation     | nan           | nan           |              0 |      nan      | 0 rows imputed                                                                                                         |
|      7 | Test-record removal                           | Data Cleaning (anomalies)       | Row filter     |   1.02876e+06 |   1.02874e+06 |             16 |        0.0016 |                                                                                                                        |
|      8 | Attribute construction                        | Attribute Construction          | Derivation     | nan           | nan           |              0 |      nan      | 12 new attributes added                                                                                                |
|      9 | Outlier detection (z-score + IQR)             | Z-score Normalization           | Derivation     | nan           | nan           |              0 |      nan      | 1,751 rows flagged, 0 deleted                                                                                          |
|     10 | Listwise deletion (customer table only)       | Listwise Deletion               | Parallel table | nan           | nan           |              0 |      nan      | fact keeps 1,028,743; customer table built from 797,867                                                                |
|     11 | Normalization (Min-Max + Z-score)             | Data Normalization              | Derivation     | nan           | nan           |              0 |      nan      | 5,878 scored customers scaled, 6 new columns                                                                           |
|     12 | CRM segmentation & CLV                        | CRM Strategies (BA-2)           | Derivation     | nan           | nan           |              0 |      nan      | 8 CRM strategies, 3 churn states, CLV per customer                                                                     |
|     13 | Discretization & generalization (RFM)         | Discretization / Generalization | Derivation     | nan           | nan           |              0 |      nan      | 9 scored segments (+1 unscored), 5 recency bands, 3 value tiers                                                        |
|     14 | Dimensionality reduction (star schema)        | Dimensionality Reduction        | Schema         | nan           | nan           |              0 |      nan      | 29 columns -> 17 fact columns + 6 dimensions                                                                           |
|     15 | Noise reduction (MA7 + exponential smoothing) | Noise Reduction                 | Derivation     | nan           | nan           |              0 |      nan      | 2 smoothed series over 604 daily points                                                                                |
|     16 | Data cube aggregation                         | Data Cube Aggregation           | Aggregation    | nan           | nan           |              0 |      nan      | 1,028,743 lines -> 67,787 summary rows (15x smaller)                                                                   |
|     17 | Histogram reduction                           | Histogram Reduction             | Aggregation    | nan           | nan           |              0 |      nan      | 1,078,080 values -> 23 bin frequencies                                                                                 |

## Rationale per step

**0. Raw load (both sheets unioned)** *(PDF: Data Integration)*  
Year 2009-2010 and Year 2010-2011 stacked into one table.

**1. Cross-sheet overlap resolution** *(PDF: Data Integration)*  
Sheets overlap 1-9 Dec 2010; the same invoices appear twice. Removing the second copy corrects GBP 376,266.82 of double-counted revenue.

**2. Duplicate removal (within sheet)** *(PDF: Duplicate Removal)*  
Exact repeats of all 8 fields inside one sheet. Removed GBP 55,450.05 of double-counted revenue.

**3. Unit conversion & format standardisation** *(PDF: Unit Conversion)*  
Trimmed and title-cased Description, upper-cased StockCode, mapped 19,247 rows to canonical country names with ISO-3 codes, cast Invoice and Customer ID to text (identifiers, not numbers).

**4. Duplicate removal (transaction business key)** *(PDF: Duplicate Removal)*  
1,188 StockCodes carry more than one Description (e.g. 23298 is both 'Bunting , Spotty' and 'Spotty Bunting'), so an exact 8-column match misses lines duplicated under a second spelling. De-duplicating on Invoice+StockCode+Quantity+Price+Timestamp+Customer caught 2 further rows worth GBP 551.85. This is also why Dim_Product is keyed on StockCode, never on Description.

**5. Missing-value removal (MAR, zero-value)** *(PDF: Missing Value Types (MAR))*  
Rows with no Description AND Price = 0. Missingness is fully predicted by observed fields, so MAR not MCAR. Revenue impact GBP 0.00 (zero by construction) - deleting them loses nothing.

**6. Missing Description imputation (placeholder)** *(PDF: Missing Value Handling)*  
Safeguard for any Description still missing after step 4: fill with 'Unknown Item' rather than delete, since such rows carry real revenue. 0 rows needed it - every missing Description was a zero-value record already removed, which is itself the evidence for the MAR claim.

**7. Test-record removal** *(PDF: Data Cleaning (anomalies))*  
StockCode TEST001/TEST002 are system tests, not transactions.

**8. Attribute construction** *(PDF: Attribute Construction)*  
Derived Revenue = Quantity x Price, plus IsCancelled, IsReturn, ItemType, TransactionType and seven calendar attributes. Mirrors the PDF's Defect Percentage = Defective / Produced x 100.

**9. Outlier detection (z-score + IQR)** *(PDF: Z-score Normalization)*  
1,751 rows flagged (both |z| > 3 AND beyond 3xIQR), carrying GBP 401,840.79. FLAGGED, NOT DELETED - the 80,995-unit order and the GBP 38,970 Manual adjustment are real events. A plain 1.5xIQR rule would have discarded ~11% of rows.

**10. Listwise deletion (customer table only)** *(PDF: Listwise Deletion)*  
230,876 rows have no Customer ID but carry GBP 2,565,542.41. They stay in the fact table and are dropped only for customer analysis - you cannot impute an identity. This is the dual-grain decision.

**11. Normalization (Min-Max + Z-score)** *(PDF: Data Normalization)*  
Recency, Frequency and Monetary each scaled to 0-1 (Min-Max) and to standard-deviation units (Z-score) before scoring, so no variable dominates on units alone.

**12. CRM segmentation & CLV** *(PDF: CRM Strategies (BA-2))*  
Each RFM segment mapped to one of the eight CRM strategies in BA-2 CRM.pdf with a recommended action; churn status derived from recency (the PDF's telecom case); CLV computed with the PDF's own formula, annualising purchase frequency over observed tenure and assuming a 3-year lifespan.

**13. Discretization & generalization (RFM)** *(PDF: Discretization / Generalization)*  
Continuous R, F, M cut into quintile scores 1-5, then generalized into 10 named segments, 5 recency bands and 3 value tiers - the PDF's Young / Middle-aged / Senior pattern.

**14. Dimensionality reduction (star schema)** *(PDF: Dimensionality Reduction)*  
One flat 29-column table split into a 17-column fact plus 5 dimensions. Descriptive attributes stored once per key instead of once per row. Note Revenue = Quantity x Price is deliberate redundancy - never feed all three into one correlation.

**15. Noise reduction (MA7 + exponential smoothing)** *(PDF: Noise Reduction)*  
Daily revenue is spiky. Added a 7-point moving average and an exponentially smoothed series with alpha = 0.3, exactly the PDF's New = a*Actual + (1-a)*Previous formula.

**16. Data cube aggregation** *(PDF: Data Cube Aggregation)*  
Line-level rows pre-summarised into daily (604), month x country (594) and month x product (66,589) cubes so dashboard pages scan thousands of rows, not a million.

**17. Histogram reduction** *(PDF: Histogram Reduction)*  
Distributions of Quantity, Unit Price and Basket Value stored as 23 bin frequencies instead of 1,078,080 raw values, preserving shape at a fraction of the size.

## Customer segments

| Segment                                     |   Customers |
|:--------------------------------------------|------------:|
| Hibernating / Lost                          |        1664 |
| Champions                                   |        1324 |
| Loyal Customers                             |        1206 |
| New / Promising                             |         560 |
| At Risk                                     |         454 |
| Needs Attention                             |         441 |
| Others                                      |         161 |
| Cannot Lose Them                            |          68 |
| No Net Value (returns / cancellations only) |          62 |
| Unknown / Guest (not identified)            |           1 |

## Item types

| ItemType           |    Rows |           Revenue |
|:-------------------|--------:|------------------:|
| Discount / Sample  |     131 |    1154.72        |
| Fee / Adjustment   |    1913 | -507203           |
| Gift Voucher       |      79 |    1686.61        |
| Other Non-Product  |       2 |      14.75        |
| Postage & Carriage |    3776 |  433411           |
| Product            | 1022842 |       1.89257e+07 |
