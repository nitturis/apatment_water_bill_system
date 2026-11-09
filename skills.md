# Navya Nisarga Water Billing Management System
## Skill Guide for MC Members

## 0. Automated Monthly Billing Inputs

Run the canonical `monthly_water_billing.py` from the workspace root with the month as the only required input:

```text
--month "August 2026"
--config config/nn_water.json
```

Defaults are resolved from the month: images from `<year>/<Month>/input`, the workbook from `<year>/NN_WM_water.xlsx`, with `<year>/NN_MM_water.xlsx`, `<year>/NN_<month-number>_water.xlsx`, and `<year>/NN_WM_<year>.xlsx` as fallbacks, outputs from `<year>/<Month>/output`, and policy from `skills.md` in the launch directory. Use the optional path arguments only when overriding these locations.

Use `--config` to support another apartment without changing Python code. The config controls workbook sheet names, vendor charges, tank capacity, solar/common prefixes, faulty solar series, and output names. The default solar fill strategy remains controlled by the explicit `Default recommendation` in `skills.md`.

Annual workbook convention: keep exactly one active workbook per year at `<year>/NN_WM_water.xlsx`. Each monthly run updates or adds that month's reading and cost columns in the annual workbook; it must not create a separate monthly meter workbook.

The automation must stop and report missing dependencies before creating bills. Required dependencies are:
- A meter-reading worksheet with a previous and current month column.
- A `cost` worksheet with the same month column, tanker quantities or vendor counts, and source cost rows.
- An `adjustments` worksheet (configured via `"adjustments_sheet"` in `config/nn_water.json`, default: `"adjustments"`) for tracking member-specific credits/debits, society-wide common cost recoveries or rebates, and running carry-forward balances across months.
- Readings for every identifier in the meter worksheet. OCR uncertainty must be reviewed using the generated `*_input_review.json` report or supplied through `--readings-json`.
- Vendor rates and the faulty-solar policy in this file.

Current vendor policy:
- **KAVERI:** ₹740 per tank, 700 litres per tank.
- **KIRAN:** ₹950 per tank, 700 litres per tank.
- **Borewell litres:** total meter consumption minus tanker litres; borewell cost is ₹0 unless the MC adds a cost policy.

The script updates the selected month in the input workbook and creates:
- `<Month>_<Year>_Final_Society_Bill.xlsx` and PDF (includes Gross Amount, Adjustment, Net Amount, and Carry Forward Balance).
- `<Month>_<Year>_Final_Society_Bill_NoBroker_Hood.xlsx` and PDF (clean resident statement showing billable consumption, Individual Cost, Common Cost, Adjustment, and Net Amount).
- `<Month>_<Year>_input_review.json` when OCR readings need review.
- Automatic updates to the `adjustments` worksheet in `<year>/NN_WM_water.xlsx` recording monthly offset transactions and maintaining active running balances.

The superseded exploratory OCR scripts are stored under `archive/ocr_experiments/`; do not use them for a monthly run.

Environment setup: use `python3 -m venv .venv`, activate it, and install `requirements.txt`. OCR also needs the system `tesseract` executable. A virtual environment is recommended for reproducibility, but the script can use any Python interpreter with the listed packages installed.

---

## 1. System Overview

### Apartment Structure
- **Building Type**: Multi-floor apartment complex
- **Total Flats**: 113 (excluding 3 common facility meters)
- **Organizing Principle**: **Floor-wise** (not series-wise)

**Floor-wise Flat Distribution**:
| Floor | Series Code | Flats Count | Flat Range | Status |
|-------|------------|-------------|-----------|--------|
| **Ground** | G | 21 | G01, G02, G03, G05, G07-G23 | ✅ Missing: G04, G06 |
| **1st Floor** | 1 | 23 | 101-123 | ✅ Complete |
| **2nd Floor** | 2 | 23 | 201-223 | ✅ Complete |
| **3rd Floor** | 3 | 23 | 301-323 | ✅ Complete |
| **4th Floor** | 4 | 23 | 401-423 | ✅ Complete |
| **Common Metering** | - | 3 meters | Common_G01, Common_G02, Common_101 | ✅ Separate tracking |
| **TOTAL** | | **113** | | **✅ Verified** |

*Data Source: NN_WM_2026.xlsx - water_meter_reading_quarter sheet*

### Water Infrastructure
#### Metering Points
- **Individual Flat Meters**: 1 per flat (113 total across 5 floors)
  - Ground: 21 meters (G01-G23, excluding G04, G06)
  - 1st-4th Floors: 23 meters each (X01-X23 where X is floor number)
- **Solar Water Meters**: [To be documented] - captures solar tank water usage
  - **Known non-working/faulty solar series**: 03, 06, 09, 10, 11, 14, 18, 22
  - These faulty solar meters must be handled using the selected monthly fill strategy: **mean**, **median**, or **no correction**
- **Common Facility Meters**: 3 meters for building-wide cleaning/maintenance
  - Common_G01, Common_G02, Common_101

#### Water Sources
| Source | Quantity | Cost Tracking | Status |
|--------|----------|---------------|--------|
| **Tanker** | 2 vendors | Tracked in INR in xlsx | ✅ Active |
| **Borewell** | 3 wells | Currently ignored (~0 cost) | ⏳ Future consideration |
| **STP Plant** | 1 (connected to borewell) | Included in borewell readings | ⏳ Future |

---

## 2. Data Structure (NN_WM_2026.xlsx)

### Tab 1: Meter Readings
**Updated**: Monthly (on designated reading date)
301
| Column | Description | Example |
|--------|-------------|---------|
| Floor | Floor identifier (G, 1, 2, 3, or 4) | G, 1, 2, 3, 4 |
| Flat Number | Floor + Unit number | G01, 101, 201, 301, 401 |
| Individual Meter | Current month reading (cubic meters) | 1234.5 |
| Previous Reading | Last month reading | 1200.0 |
| Consumption (m³) | Current - Previous | 34.5 |
| Solar Meter (Series) | Solar meter reading for the specific flat series | 2000.0 |
| Solar Previous | Last month solar reading | 1950.0 |
| Solar Consumption | Solar Current - Solar Previous | 50.0 |
| Common Meter | Common facility meter for the month | 500.0 |
| Common Previous | Last month common reading | 450.0 |
| Common Consumption | Common Current - Common Previous | 50.0 |

### Tab 2: Cost Data
**Updated**: Monthly (when invoice received)

| Column | Description | Example |
|--------|-------------|---------|
| Month | YYYY-MM format | 2026-05 |
| Tanker Vendor 1 | Cost in INR | 15,000 |
| Tanker Vendor 2 | Cost in INR | 12,000 |
| Total Tanker Cost | Sum of vendors | 27,000 |
| Borewell Cost | Currently 0 | 0 |
| STP Plant Cost | If applicable | 0 |
| **Total Water Cost** | **Sum of all sources** | **27,000** |

### Tab 3: Adjustments & Carry-Forward Ledger (`adjustments`)
**Updated**: As adjustments occur and automatically drawn down during each monthly billing run.

Tracks all member-specific credits/debits, society-wide common recoveries/rebates, and multi-month carry-forward balances across any month or year:

| Column | Description | Example |
|--------|-------------|---------|
| **Date** | Date of transaction entry | `2026-11-07` |
| **Billing Month** | Target month for adjustment | `October 2026` |
| **Flat No** | Flat identifier or `ALL` for society-wide items | `413` or `ALL` |
| **Transaction Type** | `CREDIT`, `DEBIT`, `COMMON_RECOVERY`, `COMMON_REBATE`, `OFFSET` | `CREDIT` |
| **Amount (₹)** | Transaction amount (+ for credit/recovery, - for offset) | `+2,889.00` |
| **Running Balance (₹)** | Remaining unabsorbed balance for that flat | `2,473.00` |
| **Status** | `ACTIVE` (carries forward) or `SETTLED` (balance = 0) | `ACTIVE` |
| **Description** | Human-readable audit explanation | `Sep reading 8180 -> 5180 correction` |

---

## 3. Monthly Billing Workflow

### Phase 0: Initial Setup (One-time - First Month Only)
- [x] **Floor-wise Flat Count Structure Verified**
  - Ground (G): 21 flats (missing G04, G06)
  - 1st Floor: 23 flats (101-123)
  - 2nd Floor: 23 flats (201-223)
  - 3rd Floor: 23 flats (301-323)
  - 4th Floor: 23 flats (401-423)
  - **Total: 113 flats** (verified from NN_WM_2026.xlsx)

### Phase 1: Data Collection & Validation (Days 1-5 of month)
- [ ] Collect meter readings from all 5 floors (verify flat count matches structure)
- [ ] Verify readings are not decreasing (catch meter faults, OCR digit drops)
- [ ] Record tanker vendor invoices (Kaveri, Kiran counts and amounts)
- [ ] Input all data into `NN_WM_water.xlsx`
- [ ] **Review Adjustments Ledger (`adjustments` tab)**:
  - Check Section 1 (Active Member Balances) for unabsorbed credits carrying forward from prior months
  - Verify or record any new Section 2 transactions (e.g., resident overcharge credits, past month corrections, society-wide `COMMON_RECOVERY` or `COMMON_REBATE` amounts) targeted for this billing cycle

### Phase 2: Issue Identification & Root Cause Analysis (Days 5-10)
- [ ] **Review Consumption Trends**: 
  - Compare month-on-month consumption per flat
  - Flag flats with >20% variance
  - Identify meter reading anomalies
  
- [ ] **Screen for OCR & Transcription Anomalies**:
  - **Digit Substitutions**: Check for common OCR confusions such as `8` vs `5` (e.g. Flat 413 read as `8180` instead of `5180`) or `8` vs `3` (e.g. `SOLAR_G08` read as `4372` instead of `4872`)
  - **Extra / Missing Digits**: Check for 4 digits where 3 digits are expected (e.g. Flat 301 read as `5337` instead of `537`)
  - **Negative Consumption**: Any negative consumption indicates a transcription typo, rollover, or meter replacement
  - **Compensating 2-Month Swings**: Check if a large spike in month `N` is followed by a sudden drop in month `N+1` (e.g. Flat 405); verify whether month `N` was over-read
  
- [ ] **Check for Meter Issues**:
  - Non-functioning meters (reading unchanged)
  - Inconsistent meters (abnormal spike/drop)
  - Common meter anomalies (shared resources)
  
- [ ] **Solar Meter Distribution Analysis**:
  - Calculate total solar consumption
  - Identify non-working solar meter series; current known list: **03, 06, 09, 10, 11, 14, 18, 22**
  - Select faulty solar meter fill strategy: **mean**, **median**, or **no correction**
  - Decide: **Distribute to all flats equally (A)** OR **Distribute floor-wise (B)**
  - Document reasoning

### Phase 3: Strategy Decision & Billing Options (Days 10-15)

#### Billing Strategy Selection
Create **2 different Excel files** representing different approaches:

| Aspect | **Strategy A: Flat Distribution** | **Strategy B: Floor-wise Distribution** |
|--------|----------------------------------|------------------------------------------|
| Solar Distribution | Equal across all 113 flats | Per-floor (5 groups: G, 1st, 2nd, 3rd, 4th) |
| Common Meter | Equal across all flats | Equal across all flats |
| Faulty Solar Meters | Apply selected fill strategy before equal split | Apply selected fill strategy before series split |
| Fairness Assumption | All flats equal water rights | Floor-level distribution more equitable |
| Use Case | Smaller variance in consumption | Major meter issues in certain towers |

#### Faulty Solar Meter Fill Strategy
Use this decision before calculating Strategy A or Strategy B.

| Fill Strategy | Meaning | When to Use |
|---------------|---------|-------------|
| **No correction** | Ignore the faulty solar meter reading and bill its solar consumption as 0 for that month | Use when MC decides not to estimate faulty readings; simplest and most transparent |
| **Mean correction** | Replace each faulty solar series consumption with the average consumption of working solar series | Use when working solar meters are broadly similar and no extreme outliers exist |
| **Median correction** | Replace each faulty solar series consumption with the median consumption of working solar series | Use when one or more working solar meters have very high/low outliers |

**Current known faulty solar series**: 03, 06, 09, 10, 11, 14, 18, 22

**Default recommendation**: Use **median correction** for the known faulty solar series. This is the society's established monthly billing strategy unless the MC explicitly changes the policy.

---

## 4. Detailed Billing Formulas

### Billing Components per Flat

#### A. Individual Meter Share
```
Individual Share (₹) = (Flat Consumption m³ / Total Consumption m³) × Total Water Cost
```

#### B. Solar Meter Share - Strategy A (Flat Distribution)
```
Adjusted Solar Consumption = Sum of working solar meters + filled values for faulty solar meters

Solar Share per Flat (₹) = Solar Cost Pool / 113 flats
```

**Note**: Solar Cost Allocation % = Percentage of total water cost attributed to solar
- If solar consumption is 10% of total consumption → allocate 10% of cost

#### C. Solar Meter Share - Strategy B (Series Distribution)
```
Adjusted Series Solar Consumption = Actual working meter consumption OR faulty-meter filled value

Solar Share per Flat (₹) = Series Solar Cost Pool / Flats in Series
```

#### D. Common Facility Share
```
Common Share per Flat (₹) = (Total Common Consumption m³ / 113 flats) 
                            × (Common Cost Allocation %)
```

#### E. Total Bill per Flat
```
Gross Bill = Individual Share + Solar Share + Common Share
```

#### F. Adjustments & Carry-Forward Calculation (Multi-Month Scalable)

When society-wide or member-specific adjustments exist in the `adjustments` tab:

1. **Society-Wide Common Adjustment**:
```
Adjusted Common Share (₹) = Base Common Share + (Common Recovery Amount / 113 flats)
```
*(Added equally to all 113 flats so the society recovers the required funds).*

2. **Member Credit Offset & Carry-Forward Drawdown**:
```
Available Credit = Sum of prior CREDIT transactions - Sum of prior OFFSET transactions

Credit Applied (₹) = MIN(Available Credit, Gross Bill)

Net Bill Payable (₹) = Gross Bill - Credit Applied

Closing Carry-Forward (₹) = Available Credit - Credit Applied
```

* **NoBroker Hood Rule**: A resident's `Net Bill Payable` is capped at **₹0** (no negative bills). Any unabsorbed balance automatically rolls over as `Closing Carry-Forward` to offset subsequent months until the balance reaches ₹0 (`SETTLED`).

---

## 5. Creating Output Excel Files

### File 1: Strategy_A_Distribution_[YYYY-MM].xlsx
**Purpose**: Equal distribution of solar costs across all flats

**Tabs Required**:
1. **Billing Summary** (Executive Summary)
   - Total cost recovered
   - Average bill per flat
   - Range (min-max bills)
   - Distribution method used
   - Faulty solar meter list and fill strategy used

2. **Flat-wise Billing** (Detailed)
   | Column | Formula |
   |--------|---------|
   | Tower | Tower number |
   | Flat | G01, 101, etc. |
   | Individual Consumption (m³) | From meter readings |
   | Individual Share (₹) | (Consumption / Total) × Cost |
   | Solar Share (₹) | Flat allocation × Solar % |
   | Common Share (₹) | Fixed across all |
   | **Total Bill (₹)** | Sum of shares |
   | Notes | Meter status, anomalies, faulty solar fill strategy |

3. **Cost Distribution Breakdown**
   - Total water cost: ₹[X]
   - Allocated to individual: ₹[Y] (X%)
   - Allocated to solar: ₹[Z] (Y%)
   - Allocated to common: ₹[W] (Z%)

---

### File 2: Strategy_B_Distribution_[YYYY-MM].xlsx
**Purpose**: Series-wise distribution of solar (accounts for tower-level usage patterns)

**Tabs Required**:
1. **Billing Summary** (same as Strategy A)

2. **Series-wise Summary** (aggregated by tower)
   | Column | Data |
   |--------|------|
   | Tower | 1-23 |
   | Flats in Series | [Variable - see master list] |
   | Total Individual (m³) | Sum of flat consumption in tower |
   | Solar Consumption (m³) | Actual reading for working meters; filled/0 value for faulty meters |
   | Solar Meter Status | Included, ignored, or filled using mean/median |
   | Individual Cost Share | Allocated per tower flats |
   | Solar Cost Share | Tower-specific allocation |
   | Common Cost per Tower | Fixed share / flats in tower |
   | Tower Total | Sum |
   | Average per Flat | Tower total / Flats in Series |

3. **Flat-wise Billing** (Same as Strategy A but with tower-specific solar allocation)

4. **Variance Analysis** Report
   - Strategy A average bill vs Strategy B average bill
   - Flats benefiting/penalized in each strategy
   - Recommendation for next month

---

## 6. Meter Issue Detection & Root Cause Analysis

### A. Common Meter Faults

| Issue | Detection | Root Cause | Action |
|-------|-----------|-----------|--------|
| **No Change** | Reading same as previous month | Meter stopped/stuck | Notify flatowner, send technician |
| **Sudden Spike** | >50% increase from previous | Leakage/burst, meter reset | Investigate flat, check for leaks |
| **Abnormal Jump** | <-10% (decreasing) | Meter reversed/faulty | Replace meter immediately |
| **Drift** | Gradual 5-10% monthly increase | Small leak, behavioral change | Monitor, compare with neighbors |

### B. Tower-Level Anomalies

| Pattern | Possible Cause | Investigation |
|---------|---|---|
| One tower consumption 3x others | Meter malfunction OR high usage | Verify meter, survey residents |
| All towers spike 20% same month | Seasonal OR common meter issue | Check common meter accuracy |
| Solar consumption < 2% total | Solar system underperforming | Inspect solar tank, pipes |

### C. Known Non-working Solar Meters

| Series | Solar Meter | Status | Default Billing Treatment |
|--------|-------------|--------|---------------------------|
| 03 | SOLAR_G03 | Non-working/faulty | Apply selected fill strategy |
| 06 | SOLAR_106 | Non-working/faulty | Apply selected fill strategy |
| 09 | SOLAR_G09 | Non-working/faulty | Apply selected fill strategy |
| 10 | SOLAR_G10 | Non-working/faulty | Apply selected fill strategy |
| 11 | SOLAR_G11 | Non-working/faulty | Apply selected fill strategy |
| 14 | SOLAR_G14 | Non-working/faulty | Apply selected fill strategy |
| 18 | SOLAR_G18 | Non-working/faulty | Apply selected fill strategy |
| 22 | SOLAR_G22 | Non-working/faulty | Apply selected fill strategy |

For each month, record the selected fill strategy in the output workbook:
- **No correction**: faulty series solar consumption = 0
- **Mean correction**: faulty series solar consumption = mean of working solar series consumption
- **Median correction**: faulty series solar consumption = median of working solar series consumption

### D. OCR & Transcription Anomaly Detection (Outlier Playbook)

During each monthly billing run, compare current readings against previous month readings and historical medians. Actively screen for the following proven failure patterns:

| Failure Pattern | Real-World Example | Detection Rule | Immediate Action |
|-----------------|--------------------|----------------|------------------|
| **Digit Substitution (8 vs 5)** | Flat 413: Sep reading read as `8180` instead of `5180` | Consumption >3x historical median or >1,000 units. Check leading/second digit for OCR loop confusion (`5` ↔ `8`, `1` ↔ `7`, `0` ↔ `6`). | Inspect raw handwritten photo. If already billed in a past month, use Section 10 Scenario 6 (adjustments ledger). |
| **Extra / Repeated Digits** | Flat 301: Oct reading read as `5337` instead of `537` | Reading has 4 digits when meter has 3 digits, or sudden order-of-magnitude jump (>4,000 unit spike). | Re-check handwritten sheet. Remove duplicate digit in `NN_WM_water.xlsx` or pass `--readings-json`. |
| **Negative Consumption (3 vs 8)** | SOLAR_G08: Oct reading read as `4372` vs Sep `4378` (-6 units) | Current reading < Previous reading (`Current - Previous < 0`). Check if `3` was confused with `8` (actual: `4872`, +494 units). | Never allow negative consumption into billing. Verify meter replacement vs OCR misread. |
| **Meter Replacement / Zero Reset** | Flat 113: Sep reading was `12`, Oct reading `938` (+926 units) | Abnormally large jump following a near-zero reading. | Confirm with property manager if meter was replaced or flat became newly occupied. |
| **Compensating 2-Month Swings** | Flat 405: Sep 2,026 units → Oct 488 units (2-mo avg: 1,257) | Large spike in month `N` followed by sharp drop in month `N+1`, while 2-month cumulative total matches historical average. | Total water volume is self-corrected by month `N+1`. Check if resident raised a dispute for month `N`. |

### E. Creating the Trend Report

Include in each month's billing file:
- Comparison with same month last year
- Series-wise trend chart
- Top 5 highest consumers
- Top 5 unusual spikes/drops
- Adjustments summary (credits applied, carry-forward balances)
- Recommendations for next month

---

## 7. Monthly Execution Checklist

### Week 1: Data Entry & Adjustments Setup
- [ ] All 113 flats and 3 common facility meters collected across 5 floors
- [ ] Cost invoices compiled (tanker vendor counts and total expenditure)
- [ ] Data entered into annual workbook (`NN_WM_water.xlsx`)
- [ ] Previous month readings verified against prior billing cycle
- [ ] **Validation Check**: Ensure no negative consumptions ($Current \ge Previous$)
- [ ] **Inspect Adjustments Ledger (`adjustments` tab)**:
  - Check Section 1 for carry-forward credit balances from previous months
  - Log any new member credits/debits or society-wide `COMMON_RECOVERY` / `COMMON_REBATE` entries in Section 2

### Week 2: Analysis & Outlier Screening
- [ ] Calculate consumption per flat
- [ ] **Outlier Screening** (see Section 6.D):
  - Check for digit substitution errors (`8` ↔ `5`, `3` ↔ `8`)
  - Check for extra or missing digits (e.g. 4-digit reading on a 3-digit meter)
  - Verify compensating 2-month swings against historical flat averages
- [ ] Analyze solar meter efficiency and verify working vs non-working series
- [ ] Confirm known non-working solar series list (03, 06, 09, 10, 11, 14, 18, 22) and median fill strategy
- [ ] Document unusual readings in `*_input_review.json` or workbook notes

### Week 3: Billing Execution & Ledger Reconciliation
- [ ] Run automated monthly billing:
  ```bash
  python3 monthly_water_billing.py --month "<Month> <Year>"
  ```
- [ ] Verify rate calculation: Total Cost = Tankers + Borewell
- [ ] Verify Common Share calculation: Base Common Share + (Common Recovery / 113)
- [ ] **Verify Adjustments Application**:
  - Member credits correctly offset gross bills
  - NoBroker Hood rule enforced: `Net Payable` $\ge 0$ (no negative bills)
  - Closing carry-forward balance accurately computed for remaining credits
  - Automatic `OFFSET` transaction recorded in `adjustments` Section 2
  - Active balances updated in `adjustments` Section 1
- [ ] Verify total cost recovery and financial balancing

### Week 4: Distribution & Communication
- [ ] Generate `<Month>_<Year>_Final_Society_Bill.xlsx` and PDF (detailed master bill)
- [ ] Generate `<Month>_<Year>_Final_Society_Bill_NoBroker_Hood.xlsx` and PDF (clean resident bill)
- [ ] Verify NoBroker Hood PDF contains no internal comparison wording (`pdftotext` validation)
- [ ] Verify `Adjustment` and `Total Amount` columns are formatted properly
- [ ] For flats with active carry-forward balances, verify resident statement displays ₹0 payable with unabsorbed balance noted
- [ ] Upload approved bills to NoBroker Hood portal
- [ ] Archive all files in `<year>/<Month>/output/` and commit/push changes to git repository

---

## 8. NoBroker Hood Sharing PDF

When sharing the final monthly bill to NoBroker Hood, remove analysis/comparison information and share only the billing essentials. Keep the original final workbook unchanged and create a separate sharing copy.

### Columns to Keep

| Column | Purpose |
|--------|---------|
| Flat No | Resident identifier |
| Previous Reading | Prior meter reading |
| Current Reading | Current meter reading |
| Consumption (x10 Litres) | Billable consumption |
| Individual Cost | Consumption-based cost |
| Common Cost | Equal common share (including common recovery if applicable) |
| Adjustment | Credit/debit offset applied (shown when adjustments exist) |
| Total Amount | Net amount to collect (capped at ₹0 for credit balances) |

### Columns to Remove

Remove these comparison-only fields before PDF export:
- Last Month Consumption (x10 Litres)
- Change vs Last Month (x10 Litres)
- Change vs Last Month (%)
- Mean Consumption (x10 Litres)
- Consumption vs Mean (x10 Litres)
- Any summary row/cell group used only for mean, last-month, or comparison analysis

### Recommended Output Names

Use month-specific names so the sharing file is clearly separate from the MC/internal workbook:
```
output/<Month>_<Year>/<Month>_<Year>_Final_Society_Bill_NoBroker_Hood.xlsx
output/<Month>_<Year>/<Month>_<Year>_Final_Society_Bill_NoBroker_Hood.pdf
```

### Export Procedure

1. Open the final monthly bill workbook.
2. Duplicate it to a `NoBroker_Hood.xlsx` sharing copy.
3. In the bill sheet, remove comparison columns and comparison-only summary rows.
4. Keep only the final bill sheet unless another sheet is explicitly needed for resident sharing.
5. Set the print area to the remaining bill columns and fit to one page wide.
6. Export the sharing copy to PDF using LibreOffice or Excel.
7. Validate the PDF text does not contain comparison words such as `Last Month`, `Change vs`, `Mean`, `Comparison`, or `Consumption vs`.

Example validation command:
```
pdftotext output/July_2026/July_2026_Final_Society_Bill_NoBroker_Hood.pdf - | rg -n "Last Month|Change vs|Mean|Comparison|Consumption vs"
```

No output from this command means the comparison wording was not found in the PDF.

---

## 9. Key Decision Points for MC Members

### Question 1: Which Strategy to Use This Month?

**Use Strategy A (Equal Distribution) if**:
- Solar system working uniformly across all towers
- Individual consumption patterns are normal
- No major series-wise differences
- Simplicity preferred

**Use Strategy B (Floor-wise Distribution) if**:
- Significant variation between floors (>30%)
- Specific floor has faulty meters
- Solar performance differs by floor location
- More granular fairness needed

### Question 2: How to Handle Faulty Meters?

**Option 1**: Exclude flat from billing, recalculate (not recommended)
**Option 2**: Distribute faulty meter's share to working flats in same tower
**Option 3**: Use estimated based on tower average + previous month (temporary)
**Option 4**: Use neighbor comparison method (if available)

**Recommended**: Option 3 (with Option 2 backup)

### Question 2A: How to Handle Non-working Solar Meters?

Known non-working solar series: **03, 06, 09, 10, 11, 14, 18, 22**

**Option 1: No correction**
- Ignore faulty solar readings for the month
- Faulty series solar consumption is treated as 0
- Most transparent, but may under-bill solar usage if the series actually used solar water

**Option 2: Mean correction**
- Fill faulty solar series using the mean of working solar series consumption
- Good when working solar meters are consistent
- Sensitive to extreme high/low readings

**Option 3: Median correction**
- Fill faulty solar series using the median of working solar series consumption
- Good when working solar meters contain outliers
- Usually the best estimation method if MC wants correction

**Default recommendation**: Use **median correction** for the known faulty solar series unless the MC explicitly changes the policy.

### Question 3: Should We Reduce Solar Cost Share?

Consider reducing if:
- Solar system provides <5% of total water
- Solar efficiency declining (trend)
- Major investment needed on solar

Calculate: 
```
Solar Worth % = (Solar Consumption m³ / Total Consumption m³) × 100%
If < 3% → Reduce solar cost allocation to that % instead of flat share
```

---

## 10. Troubleshooting Guide

### Scenario 1: One Tower Shows 5x Higher Consumption
1. Verify meter reading (ask overseer to recheck)
2. Check for building leaks (sump, pipes)
3. Survey residents for burst/leakage
4. Compare with last year same month
5. If confirmed: Investigate further; if meter fault: exclude from this month's billing

### Scenario 2: Solar Meter Unchanged for Multiple Months
1. Check solar tank water level physically
2. Inspect meter for stuck needle
3. Verify pump is running
4. If meter faulty: Mark as "0 consumption" and reduce solar cost to 0%
5. Get technician quote for repair

### Scenario 2A: Solar Meter Is in Known Faulty List
1. Confirm the series is in the known list: 03, 06, 09, 10, 11, 14, 18, 22
2. Do not use the raw faulty reading for billing
3. Choose monthly fill strategy: no correction, mean, or median
4. Record the chosen fill strategy in Billing Summary and Series Summary
5. If using mean/median, calculate fill value only from working solar meters
6. Keep repair/replacement status open until the meter is fixed

### Scenario 3: Total Bill Doesn't Match Cost
- Recheck formulas in Excel
- Verify total consumption = sum of all sources
- Ensure no rounding errors
- Balance sheet: Total Bill should = Total Cost × Cost Recovery %

### Scenario 4: Large Variance Between Strategy A & B
- Document in Variance Analysis tab
- Identify which towers are benefiting/penalized
- Present to residents with justification
- Use for next month decision-making

### Scenario 5: Strategy B Calculation Totals Don't Match (series-wise)
1. Verify flat counts per tower sum to 113 (use master list)
2. Check series-wise solar allocation formula: (Series Solar m³ / Flats in Series) is correct
3. Ensure "Flats in Series" matches actual documented count (variable, not all 5)
4. Recalculate: Sum of (Average per Flat × Flats in Series) should equal each tower's total
5. If still discrepancy: Tower has incomplete data (missing meters from some flats)

### Scenario 6: Resident Overbilling in Past Settled Month & Next-Month Offsets

**Context**: A resident (e.g. Flat 413) discovers that a past month's reading was misread (e.g. September reading recorded as 8,180 instead of 5,180), resulting in an overcharge (e.g. billed ₹3,428 instead of ₹539, an excess of ₹2,889). The past month's bill was already circulated, collected, and paid by all residents.

**Policy & Decision Rules**:
1. **Never Modify or Re-issue Past Settled Bills**:
   - Re-issuing past bills creates severe accounting discrepancies, breaks historical cash records, and confuses residents who already paid.
   - The archived monthly bills (`<year>/<Month>/output/`) must always match what was actually issued and collected.
2. **Recover the Discrepancy via Common Cost**:
   - Because the past month's total tanker cost was recovered across all members, the resident's excess payment artificially reduced the common burden or offset other costs for that month.
   - Recover the exact excess amount (e.g. ₹2,889.00) in the current month across all 113 flats:
     $$\text{Common Recovery per Flat} = \frac{\text{₹}2,889.00}{113 \text{ flats}} = +\text{₹}25.57 \text{ per flat}$$
   - Add this recovery to the current month's Base Common Cost so the society collects the required refund reserve.
3. **Log Entries in `NN_WM_water.xlsx` (`adjustments` tab)**:
   - Add a `COMMON_RECOVERY` row in Section 2 targeting the current billing month:
     - `Date`: Entry date (e.g. `2026-11-07`)
     - `Billing Month`: Current month (e.g. `October 2026`)
     - `Flat No`: `ALL`
     - `Transaction Type`: `COMMON_RECOVERY`
     - `Amount`: `+2889.00`
     - `Description`: `Society recovery for Flat 413 Sep overcharge`
   - Add a `CREDIT` row in Section 2 for the affected resident:
     - `Billing Month`: Current month (e.g. `October 2026`)
     - `Flat No`: `413`
     - `Transaction Type`: `CREDIT`
     - `Amount`: `+2889.00`
     - `Description`: `Sep overcharge credit (reading 8180 -> 5180)`
   - Add or update the resident in Section 1 (Active Member Balances) with initial balance `2889.00` and status `ACTIVE`.
4. **Automated Monthly Billing Drawdown**:
   - When running `monthly_water_billing.py`, the engine automatically:
     - Calculates the flat's Gross Bill (e.g. ₹416.00).
     - Applies allowable credit: `Credit Applied = MIN(Available Credit, Gross Bill)` (e.g. `MIN(₹2,889, ₹416) = ₹416.00`).
     - Caps `Net Bill Payable` at **₹0** (ensuring compatibility with NoBroker Hood).
     - Calculates closing balance: `Closing Carry Forward = Available Credit - Credit Applied` (e.g. `₹2,889 - ₹416 = ₹2,473.00`).
     - Automatically logs an `OFFSET` transaction (`-₹416.00`) in Section 2 and updates the active balance in Section 1 to `₹2,473.00`.
5. **Multi-Month Carry-Forward Scalability**:
   - In subsequent months (e.g. November, December), no manual intervention is needed.
   - The engine automatically reads the active balance from Section 1, applies the credit against each month's gross bill, and decrements the balance until it reaches ₹0, at which point the status updates to `SETTLED`.

---

## 11. Excel File Template Structure

### Common to Both Files

**Cell Format Standards**:
- Currency: ₹ (Indian Rupee)
- Numbers: 2 decimal places
- Dates: DD-MMM-YYYY
- Headers: Bold, Blue background
- Totals: Bold, Yellow background

**Required Calculations**:
- Sum checks: Total Individual Share + Total Solar + Total Common = Total Cost
- Count checks: 113 flats per billing cycle (verify series-wise breakdown)
- Reading checks: No negative or zero consumptions (except known exclusions)
- Faulty solar checks: Known faulty solar series excluded or filled according to selected strategy

**Archive Naming**:
```
Strategy_A_Distribution_2026-05.xlsx
Strategy_B_Distribution_2026-05.xlsx
NN_WM_2026_Readings_2026-05.xlsx (backup readings)
July_2026_Final_Society_Bill_NoBroker_Hood.xlsx
July_2026_Final_Society_Bill_NoBroker_Hood.pdf
```

---

## 12. Long-term Considerations

### Borewell Cost Integration (Future)
When borewell costs become significant:
1. Track borewell readings from 3 wells
2. Allocate borewell cost as separate line item
3. Decide: Allocate equally OR by tower location proximity
4. STP plant cost: Allocated as fixed amount per flat or percentage

### Solar System Expansion (Future)
If additional solar tanks planned:
- Ensure new meters are installed
- Update billing to include new tanks
- Redistribute solar allocation accordingly

### Meter Replacement Program
- Track meter age and accuracy
- Plan annual replacement cycle
- Budget for new meters in annual MC budget
- Validation: Compare meter readings with water tank level drop

---

## 13. Contact & Escalation

### When to Escalate Issues
- [ ] Meter shows negative consumption → Technician + Replace meter
- [ ] Tower consumption anomaly (>100% increase) → Site visit + Investigation
- [ ] Cost discrepancy (>5% unaccounted) → Review all calculations + Recount
- [ ] Resident disputes bill (>30% variance from average) → Compare with tower peers + History

### Monthly MC Meeting Agenda Items
1. Meter issues identified last month (status)
2. Strategy A vs B comparison analysis
3. Decision on this month's billing approach
4. Any resident complaints/disputes review
5. Solar system performance review
6. Borewell/STP status update
7. Next month priorities

---

## Appendix: Formula Reference Sheet

```
Total Water Cost (₹) = Tanker Vendor 1 + Tanker Vendor 2 + Borewell Cost + STP Cost

Total Consumption (m³) = Sum of all flat individual meters

Solar Consumption (m³) = Sum of all tower solar meters

Adjusted Solar Consumption (m³) = Sum of working solar meters + corrected values for known faulty solar meters

Faulty Solar Fill - No Correction = 0

Faulty Solar Fill - Mean = Average consumption of working solar series

Faulty Solar Fill - Median = Median consumption of working solar series

Common Consumption (m³) = Sum of common facility meters

Individual Share per Flat (₹) = (Flat Consumption m³ / Total Consumption m³) × Total Cost

Solar Share per Flat - A (₹) = (Solar Consumption m³ / 113 flats) × (Solar % of Total Cost)

Solar Share per Flat - B (₹) = (Series Solar m³ / Flats in Series) × (Solar % of Total Cost)

Common Share per Flat (₹) = (Common Consumption m³ / 113 flats) × (Common % of Total Cost)

Gross Bill per Flat (₹) = Individual Share + Solar Share + Common Share

Adjusted Common Share (₹) = Base Common Share + (Common Recovery Amount / 113 flats) - (Common Rebate Amount / 113 flats)

Available Member Credit (₹) = Cumulative Member Credits - Cumulative Prior Offsets

Credit Applied (₹) = MIN(Available Member Credit, Gross Bill)

Net Bill Payable (₹) = MAX(0, Gross Bill - Credit Applied)

Closing Carry-Forward Balance (₹) = Available Member Credit - Credit Applied

Total Society Billed Pool (₹) = Total Water Cost + Common Recovery Amount - Common Rebate Amount

Cost Recovery Check = Sum of All Net Bills + Sum of Credits Applied / Total Society Billed Pool (Should = 1.00 or 100%)
```

---

**Document Version**: 1.2  
**Last Updated**: November 2026  
**Maintained By**: Navya Nisarga MC Water Bill Committee  
**Review Frequency**: Quarterly (or as needed)
