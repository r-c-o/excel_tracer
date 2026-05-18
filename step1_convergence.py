"""
Step 1: Model Convergence

Merges outlook balance sheets with aggregator convergence data using a 5-key
waterfall join (most-specific key first), then writes four output files:
  cg_outlook_tap_env.xlsx, cbna_com_outlook_tap_env.xlsx,
  addon_all_cg.xlsx, addon_all_cbna.xlsx

Run before step2_outlook_rwa.py.
"""
from __future__ import annotations

import warnings
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime
from dateutil.relativedelta import relativedelta

import rwa_constants as rc

pd.set_option('display.max_columns', 500)

# ── PARAMETERS — update before each run ───────────────────────────────────────

# Quarter-0 date (format: Mon_YYYY) — the actuals quarter
Q0 = "Dec_2025"

# Base data directory — must contain input/ and output/ subfolders
# Example: Path(r'C:\Users\you\Desktop\aa-data\my\2026')
DATA_DIR = Path(r'C:\Users\r10995\Desktop\aa-data\my\2026')

# Input filenames (must exist in DATA_DIR/input)
INPUT_CG_FILENAME          = "outlook_balancesheet_cg.xlsx"
INPUT_CBNA_FILENAME        = "outlook_balancesheet_cbna.xlsx"
INPUT_CONVERGENCE_FILENAME = "aggregator_for_convergence.xlsx"

# Output filenames (written to DATA_DIR/output — consumed by step2)
OUTPUT_CG_OUTLOOK_FILENAME   = "cg_outlook_tap_env.xlsx"
OUTPUT_CBNA_OUTLOOK_FILENAME = "cbna_com_outlook_tap_env.xlsx"
OUTPUT_CG_ADDON_FILENAME     = "addon_all_cg.xlsx"
OUTPUT_CBNA_ADDON_FILENAME   = "addon_all_cbna.xlsx"

# ── Derived paths (no edits needed) ───────────────────────────────────────────
INPUT_DIR  = DATA_DIR / "input"
OUTPUT_DIR = DATA_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

input_cg_file          = INPUT_DIR / INPUT_CG_FILENAME
input_cbna_file        = INPUT_DIR / INPUT_CBNA_FILENAME
input_convergence_file = INPUT_DIR / INPUT_CONVERGENCE_FILENAME

output_cg_file         = OUTPUT_DIR / OUTPUT_CG_OUTLOOK_FILENAME
output_cbna_file       = OUTPUT_DIR / OUTPUT_CBNA_OUTLOOK_FILENAME
output_cg_addon_file   = OUTPUT_DIR / OUTPUT_CG_ADDON_FILENAME
output_cbna_addon_file = OUTPUT_DIR / OUTPUT_CBNA_ADDON_FILENAME

for f in [input_cg_file, input_cbna_file, input_convergence_file]:
    if not f.exists():
        raise FileNotFoundError(f"Input file not found: {f}")

# %% [1. Load inputs]

print("Loading inputs...")
cg_df       = pd.read_excel(input_cg_file)
cbna_df     = pd.read_excel(input_cbna_file)
convergence = pd.read_excel(input_convergence_file)

print(f"  CG balance sheet:   {len(cg_df):,} rows")
print(f"  CBNA balance sheet: {len(cbna_df):,} rows")
print(f"  Convergence:        {len(convergence):,} rows")

# %% [2. Validate and prune columns]


def _keep_expected_cols(df: pd.DataFrame, expected: list[str], label: str) -> pd.DataFrame:
    missing = [c for c in expected if c not in df.columns]
    if missing:
        warnings.warn(f"{label} missing expected columns: {missing}")
    else:
        print(f"  ✓ {label} has all expected columns")
    extra = [c for c in df.columns if c not in expected]
    return df.drop(columns=extra) if extra else df


cg_df       = _keep_expected_cols(cg_df,       rc.EXPECTED_BALANCE_SHEET_COLS, "CG")
cbna_df     = _keep_expected_cols(cbna_df,     rc.EXPECTED_BALANCE_SHEET_COLS, "CBNA")
convergence = _keep_expected_cols(convergence, rc.EXPECTED_CONVERGENCE_COLS,   "Convergence")

# %% [3. Split convergence → credit-risk vs addon]

PMF_ACCOUNTS = rc.CREDIT_RISK_PMF_VALUES

crd_cg_df = convergence[
    (convergence[rc.REPORTABLE_ENTITY_IS_CG] == "Y") &
    (convergence[rc.FINANCE_PMF_LEVEL_5_DESC].isin(PMF_ACCOUNTS))
].copy()

crd_cbna_df = convergence[
    (convergence[rc.REPORTABLE_ENTITY_IS_CBNA] == "Y") &
    (convergence[rc.FINANCE_PMF_LEVEL_5_DESC].isin(PMF_ACCOUNTS))
].copy()

addon_cg = convergence[
    (convergence[rc.REPORTABLE_ENTITY_IS_CG] == "Y") &
    (~convergence[rc.FINANCE_PMF_LEVEL_5_DESC].isin(PMF_ACCOUNTS))
].copy()

addon_cbna = convergence[
    (convergence[rc.REPORTABLE_ENTITY_IS_CBNA] == "Y") &
    (~convergence[rc.FINANCE_PMF_LEVEL_5_DESC].isin(PMF_ACCOUNTS))
].copy()

print(f"  Credit-risk CG:    {len(crd_cg_df):,} rows")
print(f"  Credit-risk CBNA:  {len(crd_cbna_df):,} rows")
print(f"  Addon CG:          {len(addon_cg):,} rows")
print(f"  Addon CBNA:        {len(addon_cbna):,} rows")

missing_accounts = set(PMF_ACCOUNTS) - set(convergence[rc.FINANCE_PMF_LEVEL_5_DESC].dropna().unique())
if missing_accounts:
    warnings.warn(f"PMF accounts not found in convergence data: {missing_accounts}")

# %% [4. Build waterfall pivot tables]

# Key chains ordered most-specific → least-specific
KEY_CHAINS: list[list[str]] = [
    [rc.MNGD_SGMT_L4_CDE, rc.MNGD_SGMT_L3_CDE, rc.MNGD_GEO_L4_DESC,
     rc.FINANCE_PMF_LEVEL_5_DESC, rc.MNGD_SGMT_L2_CDE],
    [rc.MNGD_SGMT_L4_CDE, rc.MNGD_SGMT_L3_CDE, rc.MNGD_GEO_L4_DESC,
     rc.FINANCE_PMF_LEVEL_5_DESC],
    [rc.MNGD_SGMT_L4_CDE, rc.MNGD_SGMT_L3_CDE, rc.FINANCE_PMF_LEVEL_5_DESC],
    [rc.MNGD_SGMT_L4_CDE, rc.FINANCE_PMF_LEVEL_5_DESC],
    [rc.FINANCE_PMF_LEVEL_5_DESC],
]

AGG_COLS = [
    rc.ADV_CG_TOTAL_RWA_AMT,
    rc.ADV_CBNA_TOTAL_RWA_AMT,
    rc.GAP_AMOUNT,
    rc.SA_RWA_AMT,
]


def _make_pivot(df: pd.DataFrame, key_cols: list[str]) -> pd.DataFrame:
    return df.groupby(key_cols, dropna=False)[AGG_COLS].sum().reset_index()


crd_cg_pivots   = [_make_pivot(crd_cg_df,   kc) for kc in KEY_CHAINS]
crd_cbna_pivots = [_make_pivot(crd_cbna_df, kc) for kc in KEY_CHAINS]

for i, piv in enumerate(crd_cg_pivots, 1):
    print(f"  Waterfall [L{i}] CG pivot: {len(piv):,} rows")

# %% [5. Quarter labels]

try:
    Q0_dt = datetime.strptime(Q0, "%b_%Y")
except ValueError:
    raise ValueError(f"Invalid Q0 format: {Q0!r}. Expected Mon_YYYY (e.g. Dec_2025).")

Q1_dt = Q0_dt - relativedelta(months=3)
Q2_dt = Q0_dt - relativedelta(months=6)
Q3_dt = Q0_dt - relativedelta(months=9)
print(
    f"Quarter labels:  Q0={Q0_dt:%b_%Y}  Q-1={Q1_dt:%b_%Y}"
    f"  Q-2={Q2_dt:%b_%Y}  Q-3={Q3_dt:%b_%Y}"
)

# %% [6. Quarterly USD totals]

for df in [cg_df, cbna_df]:
    df["QTR_USDOLLAR"] = (
        df[rc.M1_USDOLLAR].fillna(0)
        + df[rc.M2_USDOLLAR].fillna(0)
        + df[rc.M3_USDOLLAR].fillna(0)
    ) / 1e6

print(f"  CG  QTR_USDOLLAR total: {cg_df['QTR_USDOLLAR'].sum():,.2f}M")
print(f"  CBNA QTR_USDOLLAR total: {cbna_df['QTR_USDOLLAR'].sum():,.2f}M")

# %% [7. Waterfall join]


def waterfall_join(
    bs_df: pd.DataFrame,
    conv_pivots: list[pd.DataFrame],
    key_chains: list[list[str]],
) -> pd.DataFrame:
    """Left-join balance-sheet rows to convergence pivots, most-specific key first.

    Uses merge indicator to track which rows matched at each level so that
    index reassignment is exact and cannot misalign under gaps or resets.
    """
    result = bs_df.copy()
    result["_key_level"] = pd.NA
    for col in AGG_COLS:
        if col not in result.columns:
            result[col] = np.nan

    for level, (key_cols, pivot) in enumerate(zip(key_chains, conv_pivots), 1):
        unmatched_mask = result["_key_level"].isna()
        if not unmatched_mask.any():
            break

        # reset_index() captures original row labels in an "index" column
        merged = (
            result[unmatched_mask]
            .reset_index()
            .merge(pivot, on=key_cols, how="left", suffixes=("", "_r"), indicator=True)
        )
        hit = merged["_merge"] == "both"
        if not hit.any():
            continue

        orig_idx = merged.loc[hit, "index"].to_numpy()
        result.loc[orig_idx, "_key_level"] = level
        for col in AGG_COLS:
            src = col + "_r" if col + "_r" in merged.columns else col
            result.loc[orig_idx, col] = merged.loc[hit, src].to_numpy()

        print(f"  Waterfall [L{level}] matched {hit.sum():,} rows")

    unmatched_n = result["_key_level"].isna().sum()
    if unmatched_n:
        warnings.warn(f"{unmatched_n:,} balance-sheet rows had no convergence match.")
    return result


cg_output   = waterfall_join(cg_df,   crd_cg_pivots,   KEY_CHAINS)
cbna_output = waterfall_join(cbna_df, crd_cbna_pivots, KEY_CHAINS)

print(f"CG output rows:   {len(cg_output):,}")
print(f"CBNA output rows: {len(cbna_output):,}")

# %% [8. Addon summary pivots (informational)]

addon_agg_cols = [rc.ADV_CG_TOTAL_RWA_AMT, rc.ADV_CBNA_TOTAL_RWA_AMT, rc.GAP_AMOUNT]
addon_key_cols = [rc.FINANCE_PMF_LEVEL_5_DESC]

addon_cg_pivot   = addon_cg.groupby(addon_key_cols, dropna=False)[addon_agg_cols].sum().reset_index()
addon_cbna_pivot = addon_cbna.groupby(addon_key_cols, dropna=False)[addon_agg_cols].sum().reset_index()

print(f"Addon CG pivot:   {len(addon_cg_pivot):,} rows")
print(f"Addon CBNA pivot: {len(addon_cbna_pivot):,} rows")

# %% [9. Write outputs]

print(f"\nWriting {output_cg_file}")
cg_output.to_excel(output_cg_file, index=False)

print(f"Writing {output_cbna_file}")
cbna_output.to_excel(output_cbna_file, index=False)

print(f"Writing {output_cg_addon_file}")
addon_cg.to_excel(output_cg_addon_file, index=False)

print(f"Writing {output_cbna_addon_file}")
addon_cbna.to_excel(output_cbna_addon_file, index=False)

print("\n### Step 1 complete — Model Convergence ###")
print(f"Output directory: {OUTPUT_DIR}")
for name in [OUTPUT_CG_OUTLOOK_FILENAME, OUTPUT_CBNA_OUTLOOK_FILENAME,
             OUTPUT_CG_ADDON_FILENAME, OUTPUT_CBNA_ADDON_FILENAME]:
    print(f"  - {name}")
