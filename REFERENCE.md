# Outlook RWA Model — Reference Guide

## Table of Contents
1. [How to Run the Pipeline](#1-how-to-run-the-pipeline)
2. [Input Datasets & Fields](#2-input-datasets--fields)
3. [Business Knowledge](#3-business-knowledge)
4. [Key Calculations](#4-key-calculations)
5. [Risk Asset Taxonomy](#5-risk-asset-taxonomy)
6. [Model Exposure Reference](#6-model-exposure-reference)
7. [Skills Required & Recommended](#7-skills-required--recommended)

---

## 1. How to Run the Pipeline

### Prerequisites

| Requirement | Version |
|---|---|
| Python | 3.10+ |
| pandas | 1.5+ |
| numpy | 1.23+ |
| openpyxl | 3.0+ |
| python-dateutil | 2.8+ |

Install dependencies:
```bash
pip install pandas numpy openpyxl python-dateutil
```

### File execution order

The three files form a strict dependency chain. **Never run Step 2 before Step 1.**

```
ml_model_convergence.py   ← constants module, imported by both steps (no direct execution)
        │
        ▼
tmp1_model_convergence.py ← Step 1: balance-sheet × convergence waterfall join
        │
        ▼
tmp2_outlook_rwy.py       ← Step 2: adjustments + PMF mapping + upload template build
```

### Step 0 — Configure parameters

Open each script and update the two parameters at the top before every quarterly run:

```python
Q0       = "Dec_2025"          # actuals quarter — format: Mon_YYYY (e.g. Mar_2026)
DATA_DIR = Path(r'C:\Users\ryanc\projects\outlook_rwa_model\data\2026')
```

`DATA_DIR` must contain:
```
data/
  input/
    outlook_balancesheet_cg.xlsx
    outlook_balancesheet_cbna.xlsx
    aggregator_for_convergence.xlsx
    adjustments.xlsx                  ← Step 2 only
    pmf_frm_mapping.xlsx              ← Step 2 only
  output/                             ← created automatically
```

### Step 1 — Model Convergence

```bash
python tmp1_model_convergence.py
```

**What it does:** Reads the CG and CBNA balance-sheet outlooks and the aggregator
convergence file. Splits convergence rows into credit-risk and addon buckets by PMF
account. Builds five waterfall pivot tables (most-specific key → least-specific key)
and left-joins balance-sheet rows through them in order until every row is matched.
Sums monthly balances to a quarterly USD figure.

**Outputs written to `DATA_DIR/output/`:**

| File | Contents |
|---|---|
| `cg_outlook_tap_env.xlsx` | CG balance sheet enriched with convergence RWA |
| `cbna_com_outlook_tap_env.xlsx` | CBNA balance sheet enriched with convergence RWA |
| `addon_all_cona_tap_env.xlsx` | Non-credit-risk CG rows from convergence |
| `addon_all_cbna_tap_env.xlsx` | Non-credit-risk CBNA rows from convergence |

### Step 2 — Outlook RWA

```bash
python tmp2_outlook_rwy.py
```

**What it does:** Loads the Step 1 outputs alongside manual adjustments and the PMF/FRM
account-number mapping. Renames columns to the upload schema, concatenates adjustments +
outlook + addon rows, strips unknown quarter IDs, tags legacy franchises, joins PMF account
numbers, builds AA/SA/ERWA pivot tables, applies a Markets filter, and writes the final
upload templates plus a multi-sheet control workbook.

**Outputs written to `DATA_DIR/output/`:**

| File | Contents |
|---|---|
| `cg_outlook.xlsx` | Final CG upload template |
| `cbna_outlook.xlsx` | Final CBNA upload template |
| `cg_rwa_data.xlsx` | CG raw data before pivoting (audit trail) |
| `cbna_rwa_data.xlsx` | CBNA raw data before pivoting (audit trail) |
| `control_file.xlsx` | Multi-sheet reconciliation control |

---

## 2. Input Datasets & Fields

### outlook_balancesheet_cg / _cbna

Monthly balance-sheet positions at the managed-segment × geography × PMF grain.

| Field | Type | Description |
|---|---|---|
| `M1 USDOLLAR` | float | Month-1 average balance (raw USD) |
| `M2 USDOLLAR` | float | Month-2 average balance (raw USD) |
| `M3 USDOLLAR` | float | Month-3 average balance (raw USD) |
| `Reportable Entity is CG` | str (Y/N) | Row belongs to Citigroup legal entity |
| `Reportable Entity is CBNA` | str (Y/N) | Row belongs to Citibank N.A. |
| `Finance PMF Level 5 Description` | str | Lowest-level PMF account label |
| `GAP Amount` | float | Gap between plan and actual (USD) |
| `SA RWA Amount` | float | Standardized Approach RWA (USD) |
| `Adv. CG Total RWA Amount with 1.06 Multiplier` | float | Advanced Approach RWA × 1.06 capital floor |
| `Adv. CBNA Total RWA Amount with 1.06 Multiplier` | float | Same for CBNA |
| `Quarter Id` | int/str | Numeric quarter identifier |
| `RWA Cycle` | str | RWA reporting cycle label |
| `Managed Segment Level 4 Description` | str | Most granular business segment |
| `Managed Segment Level 3/2/1 Description` | str | Segment roll-up hierarchy |
| `Managed Geography Level 4/3/2/1 Description` | str | Geographic hierarchy |
| `RWA Exposure Type Description` | str | Banking Book / Trading Book / Off-Balance Sheet |

### aggregator_for_convergence

Convergence bridge file that maps each PMF position to RWA amounts and segment codes.

Key additional fields vs. the balance sheet:

| Field | Description |
|---|---|
| `Managed Segment Level 4/3/2/1 Code` | Short code equivalent of the segment descriptions |
| `Managed Geography Level 4 Description` | Geography at the convergence grain |

### adjustments.xlsx (Step 2)

Manual top-side adjustments, one sheet per entity (`Adjustments - CG`, `Adjustments - CBNA`).
Must share the same column schema as the outlook files after renaming.

### pmf_frm_mapping.xlsx (Step 2)

One-to-one crosswalk from `Finance PMF Level 5 Description` to FRM upload account numbers.

| Field | Description |
|---|---|
| `PMF L5` | PMF level-5 label (renamed to `Finance PMF Level 5 Description` on load) |
| `SA Account #` | SA upload account number |
| `AA Account #` | AA upload account number |
| `Reporting Layer` | Regulatory reporting layer tag |

---

## 3. Business Knowledge

### Entities
- **CG (Citigroup)** — consolidated group-level view across all legal entities.
- **CBNA (Citibank N.A.)** — the primary US bank subsidiary; subset of CG.
  Rows can belong to CG only, CBNA only, or both, governed by the `Reportable Entity` flags.

### PMF — Portfolio Management Framework
A hierarchical taxonomy that classifies every balance-sheet position from Level 1
(broadest) to Level 5 (most granular). Level 5 is the atomic unit used in this pipeline.
Examples: `Loans (l2)`, `Deposits with Banks (l2)`, `Trading Account Assets (l2)`.

### RWA — Risk-Weighted Assets
Capital that a bank must hold against each asset, determined by the asset's risk weight.
Two regulatory approaches coexist in this model:

| Approach | Label in code | Description |
|---|---|---|
| Standardized (SA) | `SA RWA Amount` | Fixed regulatory risk weights per asset class |
| Advanced / IRB | `AA RWA` / `ADV_*_TOTAL_RWA_AMT` | Internal model-based risk weights, subject to Basel floor |

### The 1.06 Multiplier
Under Basel III/IV the Advanced Approach output is subject to a capital floor.
The `Adv. CG Total RWA Amount with 1.06 Multiplier` field pre-applies this scalar so
downstream calculations can compare SA and AA on a level footing.

### Convergence
The process of reconciling the balance-sheet view (monthly average balances) with the
RWA/regulatory view (which uses a different segmentation and timing). The waterfall join
in Step 1 is the mechanism: it tries to match at the most granular key (segment L4 +
geography L4 + PMF L5 + segment L2) and falls back progressively until PMF L5 alone.

### Legacy Franchises & Discontinued Operations
Rows tagged with segment L3 values `Legacy Franchises (L3)` or `Discontinued Ops (L2)`
are flagged separately. They are included in totals but tracked independently because they
represent run-off portfolios under different capital treatment.

### Addon vs. Credit-Risk
Convergence rows are split before the waterfall join:
- **Credit-risk accounts** — PMF accounts on the `CREDIT_RISK_PMF_VALUES` list; subject
  to the full SA/AA RWA calculation.
- **Addon accounts** — all other PMF accounts (trading, commitments, off-balance-sheet);
  reported separately and not run through the waterfall.

---

## 4. Key Calculations

### Quarterly Average Balance
```
QTR_USDOLLAR = (M1_USDOLLAR + M2_USDOLLAR + M3_USDOLLAR) / 1,000,000
```
Result is in **USD millions**. The three months represent the three months of the quarter;
averaging reflects regulatory convention for balance-sheet reporting.

### Advanced RWA with Capital Floor
```
ADV_RWA_WITH_FLOOR = ADV_RWA_RAW × 1.06
```
Pre-applied in the source data. Analysts should use the `with 1.06 Multiplier` field
for all SA vs AA comparisons.

### GAP Amount
```
GAP = Convergence RWA − Balance Sheet RWA
```
A non-zero GAP flags a reconciliation difference that needs investigation before submission.

### Waterfall Join (5-key hierarchy)

```
Level 1 (most specific):  Segment L4 + Segment L3 + Geography L4 + PMF L5 + Segment L2
Level 2:                  Segment L4 + Segment L3 + Geography L4 + PMF L5
Level 3:                  Segment L4 + Segment L3 + PMF L5
Level 4:                  Segment L4 + PMF L5
Level 5 (least specific): PMF L5 only
```

A balance-sheet row is matched at the first level where a convergence pivot row exists.
Unmatched rows receive no RWA and trigger a warning.

### RWA Pivots (Step 2)
```
AA pivot  = SUM(AA RWA)   grouped by [Segment L4, L3, L2, PMF L5, Geo L4, L3, Reporting Layer, SA Account #]
SA pivot  = SUM(SA RWA)   same grouping
ERWA pivot = SUM(ERWA RWA) same grouping (if column present)
```

---

## 5. Risk Asset Taxonomy

### Credit-Risk PMF Accounts
These accounts receive full SA and AA RWA treatment:

| PMF Level 5 Account |
|---|
| Deposits with Banks (l2) |
| Investments (l2) |
| Loans (l2) |
| Trading Account Assets (l2) |
| Commitments (l2) |
| Brokerage Receivables (l2) |
| Federal Funds Purchased and Securities Sold Under Agreements to Repurchase (l2) |
| Securities Sold and Resales (l2) |
| Intra-Liabilities (l2) |
| Other Assets (l2) |

### Non-Credit-Risk / Addon PMF Accounts
Off-balance-sheet or market-risk-driven positions. Excluded from the waterfall join;
reported via separate addon output files:

| PMF Level 5 Account |
|---|
| Commitments to Purchase Forward-Dated Securities (l2) |
| Commitments to Sell Forward-Dated Securities (l2) |
| Trading Account Assets (l2) |
| Unsettled Trading Liabilities (l2) |
| Brokerage Receivables (l2) |
| Federal Funds Purch and Sec Loaned or Sold Under Repurchase Agreements (l2) |
| Securities Borrowed (l2) |
| Intra-Liabilities (l2) |
| Other Liabilities (l2) |
| Indirect Assets (l2) |
| Other Assets l3 |

### RWA Exposure Types
| Value | Meaning |
|---|---|
| Banking Book | Traditional loan/deposit positions; credit-risk RWA |
| Trading Book | Mark-to-market positions; market-risk RWA |
| Off-Balance Sheet | Contingent exposures (commitments, guarantees) |

---

## 6. Model Exposure Reference

This is a summary of what this model does and does not cover — useful for onboarding,
audit, or handoff documentation.

### What the model covers
- Quarterly RWA convergence between the balance-sheet outlook and the regulatory
  aggregator for CG and CBNA entities.
- Credit-risk RWA under both Standardized and Advanced (IRB) approaches.
- A 5-level hierarchical key join to maximise match rate between the two source systems.
- Manual adjustment overlay (top-side entries).
- PMF-to-FRM account-number mapping for regulatory upload files.
- Legacy and discontinued-operations flagging.
- Multi-sheet reconciliation control file for sign-off.

### What the model does NOT cover
- Market risk (trading book VaR / SVaR / FRTB) — separate pipeline.
- Operational risk RWA — separate pipeline.
- CVA (Credit Valuation Adjustment) capital charge — separate pipeline.
- CECL / IFRS 9 allowance calculations.
- Stress testing (CCAR/DFAST) projections.
- FX translation from non-USD positions (assumed pre-translated in source data).
- Power Query / pivot-table source data (only standard tabular Excel inputs).

### Known data-quality checks built in
| Check | Location | Action on failure |
|---|---|---|
| All expected columns present | Step 1 load | Warning; extra columns dropped |
| All credit-risk PMF accounts found in convergence | Step 1 | Print missing set |
| Waterfall unmatched rows | Step 1 | Warning with count |
| PMF mapping 1:1 on L5 | Step 2 | Warning with duplicate count |
| PMF join row-expansion | Step 2 | Warning with before/after counts |
| PMF SA Account # coverage | Step 2 | Warning with unmatched count |
| Input files exist | Both steps | FileNotFoundError / Warning |

---

## 7. Skills Required & Recommended

### Required skills to run the pipeline

| Skill | Why needed |
|---|---|
| Python (intermediate) | Editing `Q0` / `DATA_DIR`, reading tracebacks, re-running cells |
| pandas fundamentals | Understanding merge, groupby, concat operations when debugging |
| Excel / openpyxl basics | Inspecting input files, validating output sheet names |
| Banking / RWA concepts | Interpreting control-file discrepancies; knowing when a GAP is material |
| Basel III/IV familiarity | Understanding SA vs AA distinction and the 1.06 floor |

### Recommended skills to maintain / extend the pipeline

| Skill | Why useful |
|---|---|
| pandas advanced (MultiIndex, ExcelWriter) | Modifying pivot logic or adding new output sheets |
| Python type hints & dataclasses | Refactoring constants into a typed config object |
| pytest / hypothesis | Adding unit tests for the waterfall join and PMF split logic |
| SQL / data warehousing | Translating the pipeline to a database-backed version |
| Regulatory capital (CRR2 / Basel IV) | Evaluating impact of rule changes on the waterfall keys |
| git & version control | Managing quarterly parameter changes safely across runs |
| Jupyter notebooks | Running the pipeline interactively (`.py` files are `# %%` cell-compatible) |
| Power BI / Tableau | Building dashboards off the control-file outputs |

### Domain knowledge reference points

| Topic | Relevance to this model |
|---|---|
| Basel III Pillar 1 capital | SA vs AA RWA, 1.06 output floor |
| Credit-risk standardized approach | Risk weights by asset class (sovereign, corporate, retail, etc.) |
| Internal Ratings-Based (IRB) approach | PD × LGD × EAD → RWA; what "Advanced" means in code |
| Portfolio Management Framework (PMF) | The product taxonomy driving every join key |
| FFIEC / FR Y-9C reporting | Likely upstream source of balance-sheet data |
| Managed-segment hierarchy | Business-line segmentation used for internal P&L and capital allocation |
| Legacy Franchises / Discontinued Ops | Run-off portfolio treatment under Basel transition rules |
