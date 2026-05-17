"""
Step 11: Model Convergence (Simplified Walkthrough)

Merges outlook balance sheets with aggregator convergence data using a 5-key
waterfall join (key1-key5), computes SA/AA Risk Weight factors (RWF), and
produces `cg_outlook.xlsx`, `cbna_outlook.xlsx`, `addon_all_cg.xlsx`,
`addon_all_cbna.xlsx`.
"""
from __future__ import annotations

import warnings
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime
from dateutil.relativedelta import relativedelta

import ml_model_convergence as mc

pd.set_option('display.max_columns', 500)

# ── PARAMETERS — Update these before each run ─────────────────────────────────

# Quarter-0 date (format: Mon_YYYY) — the actuals quarter
Q0 = "Dec_2025"

# Base data directory — contains input/ and output/ subfolders
DATA_DIR = Path(r'C:\Users\r10995\Desktop\...\aa-data\my\2026')

# Input filenames (must exist in DATA_DIR/input)
INPUT_CG_FILENAME          = "outlook_balancesheet_cg.xlsx"
INPUT_CBNA_FILENAME        = "outlook_balancesheet_cbna.xlsx"
INPUT_CONVERGENCE_FILENAME = "aggregator_for_convergence.xlsx"

# Output filenames (written to OUTPUT_DIR)
OUTPUT_CG_OUTLOOK_FILENAME   = "cg_outlook_tap_env.xlsx"
OUTPUT_CBNA_OUTLOOK_FILENAME = "cbna_com_outlook_tap_env.xlsx"
OUTPUT_CG_ADDON_FILENAME     = "addon_all_cona_tap_env.xlsx"
OUTPUT_CBNA_ADDON_FILENAME   = "addon_all_cbna_tap_env.xlsx"

# ── Derived paths (no edits needed) ──────────────────────────────────────────
INPUT_DIR  = Path(DATA_DIR) / "input"
OUTPUT_DIR = Path(DATA_DIR) / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

input_cg_file          = INPUT_DIR / INPUT_CG_FILENAME
input_cbna_file        = INPUT_DIR / INPUT_CBNA_FILENAME
input_convergence_file = INPUT_DIR / INPUT_CONVERGENCE_FILENAME

output_cg_file         = OUTPUT_DIR / OUTPUT_CG_OUTLOOK_FILENAME
output_cbna_file       = OUTPUT_DIR / OUTPUT_CBNA_OUTLOOK_FILENAME
output_cg_addon_file   = OUTPUT_DIR / OUTPUT_CG_ADDON_FILENAME
output_cbna_addon_file = OUTPUT_DIR / OUTPUT_CBNA_ADDON_FILENAME

# ── Validity checks ───────────────────────────────────────────────────────────
for f in [input_cg_file, input_cbna_file, input_convergence_file]:
    if not f.exists():
        raise FileNotFoundError(f"Input file NOT found: {f}")

# %% [1. Base Input Files]
print("CG")
cg_df = pd.read_excel(input_cg_file)
print(f"  {len(cg_df):,} rows")

cbna_df = pd.read_excel(input_cbna_file)
print(f"  {len(cbna_df):,} rows")

convergence = pd.read_excel(input_convergence_file)
print(f"  convergence: {len(convergence):,} rows")

# ── Validate expected columns ─────────────────────────────────────────────────
for df, name, expected in [
    (cg_df,       "CG",          mc.EXPECTED_BALANCE_SHEET_COLS),
    (cbna_df,     "CBNA",        mc.EXPECTED_BALANCE_SHEET_COLS),
    (convergence, "Convergence", mc.EXPECTED_CONVERGENCE_COLS),
]:
    missing = [c for c in expected if c not in df.columns]
    if missing:
        warnings.warn(f"{name} missing expected columns: {missing}")
    else:
        print(f"✓ {name} balance sheet has all expected columns")

# Drop any columns not in the expected set (keeps DataFrames tidy)
deleting_col: list[str] = [
    c for c in cg_df.columns if c not in mc.EXPECTED_BALANCE_SHEET_COLS
]
if deleting_col:
    cg_df.drop(columns=deleting_col, inplace=True)

deleting_col = [
    c for c in cbna_df.columns if c not in mc.EXPECTED_BALANCE_SHEET_COLS
]
if deleting_col:
    cbna_df.drop(columns=deleting_col, inplace=True)

# %% [Local constants — balance-sheet-specific column names]

# balance sheet file
MNGD_SGMT_L4_DESC = "Managed Segment Level 4 Description"
MNGD_SGMT_L3_DESC = "Managed Segment Level 3 Description"
MNGD_SGMT_L2_DESC = "Managed Segment Level 2 Description"
MNGD_SGMT_L1_DESC = "Managed Segment Level 1 Description"
MNGD_GEO_L4_DESC  = "Managed Geography Level 4 Description"
MNGD_GEO_L3_DESC  = "Managed Geography Level 3 Description"
MNGD_GEO_L2_DESC  = "Managed Geography Level 2 Description"
MNGD_GEO_L1_DESC  = "Managed Geography Level 1 Description"

SA_RWA           = "SA RWA"
AA_RWA           = "AA RWA"
ERWA_RWA         = "ERWA RWA"
PMF_ACCT_L5_DESC = "PMF Account L5 Description"

# Non-Credit-risk PMF accounts that get RWA = 0
NON_CREDIT_RISK_PMF = [
    "Commitments to Purchase Forward-Dated Securities (l2)",
    "Commitments to Sell Forward-Dated Securities (l2)",
    "Trading Account Assets (l2)",
    "Unsettled Trading Liabilities (l2)",
    "Brokerage Receivables (l2)",
    "Federal Funds Purch and Sec Loaned or Sold Under Repurchase Agreements (l2)",
    "Securities Borrowed (l2)",
    "Intra-Liabilities (l2)",
    "Other Liabilities (l2)",
    "Indirect Assets (l2)",
    "Other Assets l3",
]

# %% [S1.11  Convergence → Credit-Risk vs Addon]

PMF_ACCOUNTS = mc.CREDIT_RISK_PMF_VALUES

crd_cg_df = convergence[
    (convergence[mc.REPORTABLE_ENTITY_IS_CG] == "Y") &
    (convergence[mc.FINANCE_PMF_LEVEL_5_DESC].isin(PMF_ACCOUNTS))
].copy()

crd_cbna_df = convergence[
    (convergence[mc.REPORTABLE_ENTITY_IS_CBNA] == "Y") &
    (convergence[mc.FINANCE_PMF_LEVEL_5_DESC].isin(PMF_ACCOUNTS))
].copy()

addon_cg = convergence[
    (convergence[mc.REPORTABLE_ENTITY_IS_CG] == "Y") &
    (~convergence[mc.FINANCE_PMF_LEVEL_5_DESC].isin(PMF_ACCOUNTS))
].copy()

addon_cbna = convergence[
    (convergence[mc.REPORTABLE_ENTITY_IS_CBNA] == "Y") &
    (~convergence[mc.FINANCE_PMF_LEVEL_5_DESC].isin(PMF_ACCOUNTS))
].copy()

print(f"credit-risk cg rows:   {len(crd_cg_df):,}")
print(f"credit-risk cbna rows: {len(crd_cbna_df):,}")
print(f"Addon CG:              {len(addon_cg):,}")
print(f"Addon CBNA:            {len(addon_cbna):,}")

# All PMF accounts found in convergence data?
missing_accounts = set(PMF_ACCOUNTS) - set(
    convergence[mc.FINANCE_PMF_LEVEL_5_DESC].dropna().unique()
)
if missing_accounts:
    print(f"  No accounts found in convergence data: {missing_accounts}")

# %% [Create key pivot tables — 5-key waterfall]

# Key column chains, most-specific (5 cols) → least-specific (1 col)
KEY_CHAINS: list[list[str]] = [
    [mc.MNGD_SGMT_L4_CDE, mc.MNGD_SGMT_L3_CDE, mc.MNGD_GEO_L4_DESC,
     mc.FINANCE_PMF_LEVEL_5_DESC, mc.MNGD_SGMT_L2_CDE],
    [mc.MNGD_SGMT_L4_CDE, mc.MNGD_SGMT_L3_CDE, mc.MNGD_GEO_L4_DESC,
     mc.FINANCE_PMF_LEVEL_5_DESC],
    [mc.MNGD_SGMT_L4_CDE, mc.MNGD_SGMT_L3_CDE, mc.FINANCE_PMF_LEVEL_5_DESC],
    [mc.MNGD_SGMT_L4_CDE, mc.FINANCE_PMF_LEVEL_5_DESC],
    [mc.FINANCE_PMF_LEVEL_5_DESC],
]

AGG_COLS = [
    mc.ADV_CG_TOTAL_RWA_AMT,
    mc.ADV_CBNA_TOTAL_RWA_AMT,
    mc.GAP_AMOUNT,
    mc.SA_RWA_AMT,
]


def create_key_pivot(df: pd.DataFrame, key_cols: list[str]) -> pd.DataFrame:
    return df.groupby(key_cols, dropna=False)[AGG_COLS].sum().reset_index()


crd_cg_pivots   = [create_key_pivot(crd_cg_df,   kc) for kc in KEY_CHAINS]
crd_cbna_pivots = [create_key_pivot(crd_cbna_df, kc) for kc in KEY_CHAINS]

for i, piv in enumerate(crd_cg_pivots, 1):
    print(f"  Waterfall [l{i}] CG pivot: {len(piv):,} rows")

# %% [Quarter date handling]

try:
    Q0_dt = datetime.strptime(Q0, "%b_%Y")
except ValueError:
    raise ValueError(
        f"Unknown Quarter format: {Q0!r}. "
        "Expected Mon_YYYY (e.g. Dec_2025)."
    )

Q1_dt = Q0_dt - relativedelta(months=3)
Q2_dt = Q0_dt - relativedelta(months=6)
Q3_dt = Q0_dt - relativedelta(months=9)

q0_label = Q0_dt.strftime("%b_%Y")
q1_label = Q1_dt.strftime("%b_%Y")
q2_label = Q2_dt.strftime("%b_%Y")
q3_label = Q3_dt.strftime("%b_%Y")

print(
    f"Quarter run labels:  Q0={q0_label}  Q-1={q1_label}"
    f"  Q-2={q2_label}  Q-3={q3_label}"
)

_year = str(Q0_dt.year)
_qtr  = Q0_dt.strftime("%b")

# %% [Balance Sheet Data — quarterly sums, both math names]
# Sum M1+M2+M3 to quarterly USD total (scale to millions, 1e6)

for df in [cg_df, cbna_df]:
    df["QTR_USDOLLAR"] = (
        df[mc.M1_USDOLLAR].fillna(0)
        + df[mc.M2_USDOLLAR].fillna(0)
        + df[mc.M3_USDOLLAR].fillna(0)
    ) / 1e6

print(f"CG  QTR_USDOLLAR sum: {cg_df['QTR_USDOLLAR'].sum():,.2f}M")
print(f"CBNA QTR_USDOLLAR sum: {cbna_df['QTR_USDOLLAR'].sum():,.2f}M")

# %% [5-key waterfall join]


def waterfall_join(
    bs_df: pd.DataFrame,
    conv_pivots: list[pd.DataFrame],
    key_chains: list[list[str]],
) -> pd.DataFrame:
    """Left-join balance-sheet rows to convergence pivots, most-specific key first."""
    result = bs_df.copy()
    result["_key_level"] = pd.NA

    # Initialise agg columns so they exist even if nothing matches
    for col in AGG_COLS:
        if col not in result.columns:
            result[col] = np.nan

    for level, (key_cols, pivot) in enumerate(zip(key_chains, conv_pivots), 1):
        unmatched_mask = result["_key_level"].isna()
        if not unmatched_mask.any():
            break

        unmatched = result[unmatched_mask].copy()
        merged = unmatched.merge(
            pivot,
            on=key_cols,
            how="inner",
            suffixes=("", "_conv"),
        )
        if merged.empty:
            continue

        # Propagate matched indices back to result
        idx = unmatched.reset_index().merge(
            pivot[key_cols].drop_duplicates(), on=key_cols, how="inner"
        )["index"]

        result.loc[idx, "_key_level"] = level
        for col in AGG_COLS:
            conv_col = col + "_conv" if col + "_conv" in merged.columns else col
            fill_vals = (
                unmatched.loc[idx]
                .merge(pivot[key_cols + [col]], on=key_cols, how="left")[col]
                .values
            )
            result.loc[idx, col] = fill_vals

        print(f"  Waterfall [l{level}] matched {len(idx):,} rows")

    unmatched_n = result["_key_level"].isna().sum()
    if unmatched_n:
        warnings.warn(f"{unmatched_n:,} balance-sheet rows had no convergence match.")

    return result


cg_output   = waterfall_join(cg_df,   crd_cg_pivots,   KEY_CHAINS)
cbna_output = waterfall_join(cbna_df, crd_cbna_pivots, KEY_CHAINS)

print(f"CG output rows:   {len(cg_output):,}")
print(f"CBNA output rows: {len(cbna_output):,}")

# %% [Calculate priority across key levels]

# Rows matched at a higher (more specific) key level take priority;
# _key_level=1 is highest priority, NaN means no match.
for df in [cg_output, cbna_output]:
    df["_priority"] = df["_key_level"].fillna(len(KEY_CHAINS) + 1).astype(int)

# %% [Addon data — build pivot and attach]

addon_agg_cols = [mc.ADV_CG_TOTAL_RWA_AMT, mc.ADV_CBNA_TOTAL_RWA_AMT, mc.GAP_AMOUNT]
addon_key_cols = [mc.FINANCE_PMF_LEVEL_5_DESC]

addon_cg_pivot   = addon_cg.groupby(addon_key_cols, dropna=False)[addon_agg_cols].sum().reset_index()
addon_cbna_pivot = addon_cbna.groupby(addon_key_cols, dropna=False)[addon_agg_cols].sum().reset_index()

print(f"Addon CG pivot:   {len(addon_cg_pivot):,} rows")
print(f"Addon CBNA pivot: {len(addon_cbna_pivot):,} rows")

# %% [Write output files]

print(f"\nwriting {output_cg_file}")
cg_output.to_excel(output_cg_file, index=False)

print(f"writing {output_cbna_file}")
cbna_output.to_excel(output_cbna_file, index=False)

print(f"writing {output_cg_addon_file}")
addon_cg.to_excel(output_cg_addon_file, index=False)

print(f"writing {output_cbna_addon_file}")
addon_cbna.to_excel(output_cbna_addon_file, index=False)

print("\nDone.")
