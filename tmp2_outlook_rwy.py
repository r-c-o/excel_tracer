"""
Step 2: Outlook RWA (Simplified Walkthrough)

Convergence outputs + adjustment files + PMF/RWF mappings,
applies adjustments, maps PMF codes and PMF account numbers, creates
pivot-based upload templates for CG and CBNA entities.

*Prerequisites*: Run `tmp1_model_convergence.py` first.
"""
from __future__ import annotations

import warnings
import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime

import ml_model_convergence as mc

pd.set_option('display.max_columns', 500)

# ── PARAMETERS — Update these before each run ─────────────────────────────────

Q0 = 'Dec_2025'
DATA_DIR = Path(r'C:\Users\r10995\Desktop\...\aa-data\my\2026')

# Model Convergence output directory (from Step 1)
MODEL_CONVERGENCE_DIR = Path(DATA_DIR) / 'walkthrough' / 'output'

# Input filenames — Step 1 outputs consumed here
INPUT_CG_OUTLOOK_FULL_FILENAME   = 'CG_Upload_template_full.xlsx'
INPUT_CBNA_OUTLOOK_FULL_FILENAME = 'CBNA_Upload_template_full.xlsx'
INPUT_CG_ADDON_FILENAME          = 'addon_all_cg.xlsx'
INPUT_CBNA_ADDON_FILENAME        = 'addon_all_cbna.xlsx'

# Input filenames — adjustments / mappings (in DATA_DIR/input)
INPUT_ADJUSTMENTS_FILENAME      = 'adjustments.xlsx'
INPUT_PMF_RWA_MAPPING_FILENAME  = 'pmf_frm_mapping.xlsx'
INPUT_CONVERGENCE_FILENAME      = 'aggregator_for_convergence.xlsx'

# Output filenames
OUTPUT_CG_OUTLOOK_FULL_FILENAME  = 'cg_outlook.xlsx'
OUTPUT_CBNA_OUTLOOK_FULL_FILENAME = 'cbna_outlook.xlsx'
OUTPUT_CG_RWA_DATA_FILENAME      = 'cg_rwa_data.xlsx'
OUTPUT_CBNA_RWA_DATA_FILENAME    = 'cbna_rwa_data.xlsx'
OUTPUT_CONTROL_FILENAME          = 'control_file.xlsx'

# ── Derived paths (no edits needed) ───────────────────────────────────────────
INPUT_DIR  = Path(DATA_DIR) / 'input'
OUTPUT_DIR = Path(DATA_DIR) / 'output'
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Step-1 output file handles
input_cg_outlook_file   = MODEL_CONVERGENCE_DIR / INPUT_CG_OUTLOOK_FULL_FILENAME
input_cbna_outlook_file = MODEL_CONVERGENCE_DIR / INPUT_CBNA_OUTLOOK_FULL_FILENAME
cg_addon_file           = MODEL_CONVERGENCE_DIR / INPUT_CG_ADDON_FILENAME
cbna_addon_file         = MODEL_CONVERGENCE_DIR / INPUT_CBNA_ADDON_FILENAME

# Adjustment / mapping file handles
adjustments_file  = INPUT_DIR / INPUT_ADJUSTMENTS_FILENAME
pmf_mapping_file  = INPUT_DIR / INPUT_PMF_RWA_MAPPING_FILENAME
convergence_file  = INPUT_DIR / INPUT_CONVERGENCE_FILENAME

# Output file handles
output_cg_file       = OUTPUT_DIR / OUTPUT_CG_OUTLOOK_FULL_FILENAME
output_cbna_file     = OUTPUT_DIR / OUTPUT_CBNA_OUTLOOK_FULL_FILENAME
output_cg_rwa_file   = OUTPUT_DIR / OUTPUT_CG_RWA_DATA_FILENAME
output_cbna_rwa_file = OUTPUT_DIR / OUTPUT_CBNA_RWA_DATA_FILENAME
output_control_file  = OUTPUT_DIR / OUTPUT_CONTROL_FILENAME

print("CG")
print(f"  Model Convergence dir: {MODEL_CONVERGENCE_DIR}")
print(f"  Input dir:             {INPUT_DIR}")
print(f"  Output dir:            {OUTPUT_DIR}")

# %% [1. Base Input Files]

# Data Quality Check: verify files exist
all_input_files = [
    adjustments_file, input_cg_outlook_file, input_cbna_outlook_file,
    cg_addon_file, cbna_addon_file, pmf_mapping_file, convergence_file,
]
for f in all_input_files:
    if not f.exists():
        warnings.warn(f"▲ INPUT FILE NOT FOUND: {f}")
    else:
        print(f"  Found: {f.name}")

# %% [2. Read Input Files]

src_cg_adjustments   = pd.read_excel(adjustments_file, sheet_name='Adjustments - CG')
src_cbna_adjustments = pd.read_excel(adjustments_file, sheet_name='Adjustments - CBNA')
src_cg_outlook       = pd.read_excel(input_cg_outlook_file)
src_cbna_outlook     = pd.read_excel(input_cbna_outlook_file)
src_cg_addon         = pd.read_excel(cg_addon_file)
src_cbna_addon       = pd.read_excel(cbna_addon_file)
src_pmf_mapping      = pd.read_excel(pmf_mapping_file)
convergence          = pd.read_excel(convergence_file)

print(f"CG adjustments rows:   {len(src_cg_adjustments):,}")
print(f"CBNA adjustments rows: {len(src_cbna_adjustments):,}")
print(f"CG outlook rows:       {len(src_cg_outlook):,}")
print(f"CBNA outlook rows:     {len(src_cbna_outlook):,}")
print(f"CG addon rows:         {len(src_cg_addon):,}")
print(f"CBNA addon rows:       {len(src_cbna_addon):,}")
print(f"PMF mapping rows:      {len(src_pmf_mapping):,}")
print(f"Convergence rows:      {len(convergence):,}")

# ── Data Quality Checks ───────────────────────────────────────────────────────

pmf_dupes = src_pmf_mapping.duplicated(
    subset=[mc.FINANCE_PMF_LEVEL_5_DESC], keep=False
).sum()
if pmf_dupes > 0:
    warnings.warn(f"▲ PMF RWA mapping has {pmf_dupes} entries — not 1:1 on PMF L5")
else:
    print("✓ PMF RWA mapping is 1:1 on PMF L5")

pmf_dupes_mngd = src_pmf_mapping.duplicated(
    subset=['Managed Segment L4 Descr'], keep=False
).sum()
if pmf_dupes_mngd > 0:
    warnings.warn(f"▲ PMF RWA mapping has {pmf_dupes_mngd} duplicate Managed Segment L4 Descr entries")
else:
    print("✓ PMF RWA mapping is 1:1 on Managed Segment L4 Descr")

# %% [Local column-name constants]

# These column names appear in the adjustment / outlook files (differ from mc.*)
MNGD_SGMT_L4_DESC  = 'Managed Segment L4 Descr'
MNGD_SGMT_L3_DESC  = 'Managed Segment L3 Descr'
MNGD_SGMT_L2_DESC  = 'Managed Segment L2 Descr'
MNGD_SGMT_L1_DESC  = 'Managed Segment L1 Descr'
MNGD_GEO_L4_DESC   = 'Managed Geography L4 Descr'
MNGD_GEO_L3_DESC   = 'Managed Geography L3 Descr'
MNGD_GEO_L2_DESC   = 'Managed Geography L2 Descr'
MNGD_GEO_L1_DESC   = 'Managed Geography L1 Descr'

SA_RWA          = 'SA RWA'
AA_RWA          = 'AA RWA'
ERWA_RWA        = 'ERWA RWA'
RWA_CALC        = 'RWA Calc'
NVA_CALC        = 'NVA Calc'
REPORTING_LAYER = 'Reporting Layer'
SA_ACCOUNT_NUM  = 'SA Account #'
AA_ACCOUNT_NUM  = 'AA Account #'
QUARTER_ID      = 'Quarter Id'
PMF_ACCT_L5_DESC = 'PMF Account L5 Descr'

# Filter-value strings
MARKETS_FILTER            = 'Markets Filter'
DISCONTINUED_OPS_L3       = 'Discontinued Ops (L2)'
LEGACY_FRANCHISES_L3      = 'Legacy Franchises (L3)'
LEGACY_HOLDINGS_ASSETS_L4 = 'Legacy Holdings Assets (L4)'
LATIN_AMERICA             = 'Latin America'
MARKETS_L2                = 'Markets (L2)'
BANKING_L2                = 'Banking'
SERVICES_L2               = 'Services (L2)'

# %% [3. Format Adjustments]


def format_adjustments(input_df: pd.DataFrame) -> pd.DataFrame:
    for col in input_df.select_dtypes(include=['number']).columns:
        input_df[col] = pd.to_numeric(input_df[col], errors='coerce')
    for col in input_df.select_dtypes(include=['object']).columns:
        input_df[col] = input_df[col].fillna('N/A')
    return input_df


cg_adjustments   = format_adjustments(src_cg_adjustments.copy())
cbna_adjustments = format_adjustments(src_cbna_adjustments.copy())

# %% [4. Rename Addon Columns to Match Outlook Schema]


def rename_addon_columns(input_df: pd.DataFrame, entity: str) -> pd.DataFrame:
    """Rename Addon convergence columns to match Outlook schema."""
    rwa_col = mc.ADV_CG_TOTAL_RWA_AMT if entity == 'CG' else mc.ADV_CBNA_TOTAL_RWA_AMT
    renames: dict[str, str] = {
        mc.MNGD_SGMT_L4_CDE:    MNGD_SGMT_L4_DESC,
        mc.MNGD_SGMT_L3_CDE:    MNGD_SGMT_L3_DESC,
        mc.MNGD_SGMT_L2_CDE:    MNGD_SGMT_L2_DESC,
        mc.MNGD_SGMT_L1_CDE:    MNGD_SGMT_L1_DESC,
        mc.PRJD_SGMT_L4_DESC:   MNGD_SGMT_L4_DESC,
        mc.PRJD_SGMT_L3_DESC:   MNGD_SGMT_L3_DESC,
        mc.MNGD_GEO_L4_DESC:    MNGD_GEO_L4_DESC,
        mc.MNGD_GEO_L3_DESC:    MNGD_GEO_L3_DESC,
        mc.MNGD_GEO_L2_DESC:    MNGD_GEO_L2_DESC,
        mc.MNGD_GEO_L1_DESC:    MNGD_GEO_L1_DESC,
        rwa_col:                 AA_RWA,
        mc.SA_RWA_AMT:           SA_RWA,
        mc.GAP_AMOUNT:           NVA_CALC,
        mc.OPFR_ID:              QUARTER_ID,
    }
    input_df = input_df.rename(
        columns={k: v for k, v in renames.items() if k in input_df.columns}
    )
    input_df['Entity'] = entity
    return input_df


cg_outlook     = src_cg_outlook.copy()
cbna_outlook   = src_cbna_outlook.copy()
addon_all_cg   = rename_addon_columns(src_cg_addon.copy(),   'CG')
addon_all_cbna = rename_addon_columns(src_cbna_addon.copy(), 'CBNA')

print(f"CG outlook rows:     {len(cg_outlook):,}")
print(f"CBNA outlook rows:   {len(cbna_outlook):,}")
print(f"addon_all_cg rows:   {len(addon_all_cg):,}")
print(f"addon_all_cbna rows: {len(addon_all_cbna):,}")

# %% [5. Rename PMF Mapping Column]

src_pmf_mapping = src_pmf_mapping.rename(
    columns={'PMF L5': mc.FINANCE_PMF_LEVEL_5_DESC}
)
pmf_mapping = src_pmf_mapping.copy()

# %% [6. Concatenate Adjustments + Outlook + Addon]

cg_concat = pd.concat(
    [cg_adjustments, cg_outlook, addon_all_cg], ignore_index=True
).copy()
cbna_concat = pd.concat(
    [cbna_adjustments, cbna_outlook, addon_all_cbna], ignore_index=True
).copy()

print(f"CG concatenated rows:   {len(cg_concat):,}")
print(f"CBNA concatenated rows: {len(cbna_concat):,}")

# %% [7. Filter Unknown Quarter IDs and Convert to Numeric]

for df in [cg_concat, cbna_concat]:
    df.drop(df[df[QUARTER_ID] == 'Unknown'].index, inplace=True)
    df[QUARTER_ID] = pd.to_numeric(df[QUARTER_ID], errors='coerce')

cg_concat['Entity']   = 'CG'
cbna_concat['Entity'] = 'CBNA'

# Save raw data before further transformations
cg_raw_data   = cg_concat.copy()
cbna_raw_data = cbna_concat.copy()

print(f"CG after filtering unknowns:   {len(cg_concat):,}")
print(f"CBNA after filtering unknowns: {len(cbna_concat):,}")

# %% [8. Legacy Franchises Breakdown]

legacy_cg = cg_concat[
    cg_concat[MNGD_SGMT_L4_DESC] == LEGACY_HOLDINGS_ASSETS_L4
].copy()
legacy_cbna = cbna_concat[
    cbna_concat[MNGD_SGMT_L4_DESC] == LEGACY_HOLDINGS_ASSETS_L4
].copy()

print(f"Legacy Holdings CG rows:   {len(legacy_cg):,}")
print(f"Legacy Holdings CBNA rows: {len(legacy_cbna):,}")

# Legacy Franchises filter — tag rows for downstream reporting
_legacy_filter_cg = (
    cg_concat[MNGD_SGMT_L3_DESC].isin([LEGACY_FRANCHISES_L3, DISCONTINUED_OPS_L3])
)
_legacy_filter_cbna = (
    cbna_concat[MNGD_SGMT_L3_DESC].isin([LEGACY_FRANCHISES_L3, DISCONTINUED_OPS_L3])
)
cg_concat['_legacy_flag']   = _legacy_filter_cg
cbna_concat['_legacy_flag'] = _legacy_filter_cbna

# %% [9. PMF Join — map SA Account Numbers onto concat data]

pre_cg   = len(cg_concat)
pre_cbna = len(cbna_concat)

frm_cg_output   = cg_concat.merge(
    pmf_mapping, on=mc.FINANCE_PMF_LEVEL_5_DESC, how='left'
)
frm_cbna_output = cbna_concat.merge(
    pmf_mapping, on=mc.FINANCE_PMF_LEVEL_5_DESC, how='left'
)

# Data Quality: row expansion check
if len(frm_cg_output) > pre_cg:
    warnings.warn(
        f"▲ CG PMF join caused row expansion "
        f"({pre_cg:,} → {len(frm_cg_output):,}) — may cause row expansion"
    )
if len(frm_cbna_output) > pre_cbna:
    warnings.warn(
        f"▲ CBNA PMF join caused row expansion "
        f"({pre_cbna:,} → {len(frm_cbna_output):,}) — may cause row expansion"
    )

# Data Quality: PMF coverage
for label, df in [('CG', frm_cg_output), ('CBNA', frm_cbna_output)]:
    unmatched = df[SA_ACCOUNT_NUM].isna().sum() if SA_ACCOUNT_NUM in df.columns else len(df)
    if unmatched > 0:
        warnings.warn(f"▲ {label}: {unmatched:,} rows have no PMF mapping match!")
    else:
        print(f"✓ {label}: PMF mapping joined!")

print(f"CG frm output rows:   {len(frm_cg_output):,}")
print(f"CBNA frm output rows: {len(frm_cbna_output):,}")

# %% [10. Build Key Strings Before Pivot (for reference)]

for label, df in [('CG', frm_cg_output), ('CBNA', frm_cbna_output)]:
    key_cols = [
        MNGD_SGMT_L4_DESC, MNGD_SGMT_L3_DESC,
        mc.FINANCE_PMF_LEVEL_5_DESC, MNGD_GEO_L4_DESC,
    ]
    unique_keys = df[key_cols].drop_duplicates()
    print(f"  {label}: {len(unique_keys):,} unique key combinations")

# %% [11. Format Columns Before Pivots]


def format_columns_before_pivot(input_df: pd.DataFrame) -> pd.DataFrame:
    """Coerce RWA amount columns to numeric and fill SA_ACCOUNT_NUM NaN → 0."""
    for col in [ERWA_RWA, AA_RWA, SA_RWA]:
        if col in input_df.columns:
            input_df[col] = pd.to_numeric(input_df[col], errors='coerce')
    if SA_ACCOUNT_NUM in input_df.columns:
        input_df[SA_ACCOUNT_NUM] = (
            pd.to_numeric(input_df[SA_ACCOUNT_NUM], errors='coerce').fillna(0)
        )
    return input_df


frm_cg_output   = format_columns_before_pivot(frm_cg_output)
frm_cbna_output = format_columns_before_pivot(frm_cbna_output)

# %% [12. Upload Template Pivots (ERWA, AA, SA)]

_PIVOT_INDEX = [
    MNGD_SGMT_L4_DESC, MNGD_SGMT_L3_DESC, MNGD_SGMT_L2_DESC,
    mc.FINANCE_PMF_LEVEL_5_DESC, MNGD_GEO_L4_DESC, MNGD_GEO_L3_DESC,
    REPORTING_LAYER, SA_ACCOUNT_NUM,
]


def _make_rwa_pivot(
    df: pd.DataFrame, value_col: str, label: str
) -> pd.DataFrame:
    valid_keys = [c for c in _PIVOT_INDEX if c in df.columns]
    piv = df.groupby(valid_keys, dropna=False)[[value_col]].sum()
    piv.columns = [label]
    return piv


def create_upload_template_pivots(frm_output: pd.DataFrame) -> pd.DataFrame:
    """Return AA / SA / ERWA pivots merged into one DataFrame."""
    aa_pivot = _make_rwa_pivot(frm_output, AA_RWA, 'AA')
    sa_pivot = _make_rwa_pivot(frm_output, SA_RWA, 'SA')

    erwa_pivot = None
    if ERWA_RWA in frm_output.columns and frm_output[ERWA_RWA].notna().any():
        erwa_pivot = _make_rwa_pivot(frm_output, ERWA_RWA, 'ERWA')

    parts = [aa_pivot, sa_pivot]
    if erwa_pivot is not None:
        parts.append(erwa_pivot)

    pivots = pd.concat(parts, axis=1)
    pivots.columns.name = None
    return pivots.reset_index()


cg_frm_pivots   = create_upload_template_pivots(frm_cg_output)
cbna_frm_pivots = create_upload_template_pivots(frm_cbna_output)

print(f"CG frm pivot rows:   {len(cg_frm_pivots):,}")
print(f"CBNA frm pivot rows: {len(cbna_frm_pivots):,}")

# %% [13. Markets Filter]


def create_markets_filter(input_df: pd.DataFrame) -> pd.DataFrame:
    """Tag rows: 'Keep' if Markets L2 + non-zero RWA exposure, else 'Remove'."""
    nw = np.where(
        (input_df.get(MNGD_SGMT_L2_DESC, '') == MARKETS_L2)
        & (input_df.get(mc.RWA_EXPOSURE_TYPE_DESC, '') != 'Remove'),
        'Keep',
        'Remove',
    )
    input_df[MARKETS_FILTER] = nw
    return input_df


frm_cg_output   = create_markets_filter(frm_cg_output)
frm_cbna_output = create_markets_filter(frm_cbna_output)
print("Markets filter applied")

# %% [14. Format Upload Template (Add Upload Columns)]


def format_upload_template(input_df: pd.DataFrame) -> pd.DataFrame:
    """Fill numeric NaN → 0 and map monthly balance-sheet columns."""
    num_cols = input_df.select_dtypes(include='number').columns
    input_df[num_cols] = input_df[num_cols].fillna(0)

    for mc_col, month_col in [
        (mc.M1_USDOLLAR, 'Month1'),
        (mc.M2_USDOLLAR, 'Month2'),
        (mc.M3_USDOLLAR, 'Month3'),
    ]:
        input_df[month_col] = input_df.get(
            mc_col, pd.Series(0, index=input_df.index)
        ).fillna(0)

    return input_df


frm_cg_output   = format_upload_template(frm_cg_output)
frm_cbna_output = format_upload_template(frm_cbna_output)

# %% [15. Build FRM Raw Data Control]

_CTRL_GROUP = [
    MNGD_SGMT_L2_DESC, MNGD_SGMT_L3_DESC,
    mc.FINANCE_PMF_LEVEL_5_DESC, mc.RWA_EXPOSURE_TYPE_DESC,
]


def build_frm_raw_data(
    frm_df: pd.DataFrame,
    entity: str,
) -> pd.DataFrame:
    """Aggregate by L2/L3 Desc + RWA type → control summary."""
    group_cols = [c for c in _CTRL_GROUP if c in frm_df.columns]
    agg_cols   = [c for c in [AA_RWA, SA_RWA, ERWA_RWA] if c in frm_df.columns]
    result = (
        frm_df
        .groupby(group_cols, dropna=False)[agg_cols]
        .sum()
        .reset_index()
    )
    result['Entity'] = entity
    return result


cg_frm_raw   = build_frm_raw_data(frm_cg_output,   'CG')
cbna_frm_raw = build_frm_raw_data(frm_cbna_output, 'CBNA')

# %% [16. Build Convergence Control Summary Tables]

_CONV_GROUP = [c for c in _CTRL_GROUP if c in convergence.columns]

cg_convergence_control = (
    convergence[convergence[mc.REPORTABLE_ENTITY_IS_CG] == 'Y']
    .groupby(_CONV_GROUP, dropna=False)[[mc.ADV_CG_TOTAL_RWA_AMT, mc.SA_RWA_AMT]]
    .sum()
    .reset_index()
)
cbna_convergence_control = (
    convergence[convergence[mc.REPORTABLE_ENTITY_IS_CBNA] == 'Y']
    .groupby(_CONV_GROUP, dropna=False)[[mc.ADV_CBNA_TOTAL_RWA_AMT, mc.SA_RWA_AMT]]
    .sum()
    .reset_index()
)

cg_frm_control   = cg_frm_raw.copy()
cbna_frm_control = cbna_frm_raw.copy()

_raw_agg_cols = [c for c in [AA_RWA, SA_RWA] if c in cg_raw_data.columns]
_raw_group    = [c for c in _CTRL_GROUP if c in cg_raw_data.columns]

cg_raw_data_control   = (
    cg_raw_data.groupby(_raw_group, dropna=False)[_raw_agg_cols].sum().reset_index()
)
cbna_raw_data_control = (
    cbna_raw_data.groupby(_raw_group, dropna=False)[_raw_agg_cols].sum().reset_index()
)

param_data = pd.DataFrame(
    [('Q0', Q0), ('DATA_DIR', str(DATA_DIR))],
    columns=['Parameter', 'Value'],
)

# %% [## 17. Write Control File and Output Workbooks]

# ── Control workbook ──────────────────────────────────────────────────────────
control_file_path = OUTPUT_DIR / OUTPUT_CONTROL_FILENAME

cg_convergence_control.index.name   = 'CG Convergence Control'
cg_frm_control.index.name           = 'CG FRM Control'
cbna_convergence_control.index.name = 'CBNA Convergence Control'
cbna_frm_control.index.name         = 'CBNA FRM Control'
cg_raw_data_control.index.name      = 'CG Raw Data Control'
cbna_raw_data_control.index.name    = 'CBNA Raw Data Control'

start_row = 0
cg_frm_start_row   = len(cg_convergence_control) + 2
cbna_frm_start_row = len(cbna_convergence_control) + 2

with pd.ExcelWriter(control_file_path) as writer:
    cg_convergence_control.to_excel(
        writer, sheet_name='CG Convergence Control', startrow=start_row, index=False
    )
    cg_frm_control.to_excel(
        writer, sheet_name='CG FRM Control', startrow=start_row, index=False
    )
    cbna_convergence_control.to_excel(
        writer, sheet_name='CBNA Convergence Control', startrow=start_row, index=False
    )
    cbna_frm_control.to_excel(
        writer, sheet_name='CBNA FRM Control', startrow=start_row, index=False
    )
    cg_raw_data_control.to_excel(
        writer, sheet_name='CG Raw Data Control', startrow=start_row, index=False
    )
    cbna_raw_data_control.to_excel(
        writer, sheet_name='CBNA Raw Data Control', startrow=start_row, index=False
    )
    param_data.to_excel(writer, sheet_name='Parameters', index=False)

print(f"Exported: {control_file_path}")

# ── Main output workbooks ─────────────────────────────────────────────────────
print(f"\nwriting {output_cg_file}")
frm_cg_output.to_excel(output_cg_file, index=False)

print(f"writing {output_cbna_file}")
frm_cbna_output.to_excel(output_cbna_file, index=False)

print(f"writing {output_cg_rwa_file}")
cg_raw_data.to_excel(output_cg_rwa_file, index=False)

print(f"writing {output_cbna_rwa_file}")
cbna_raw_data.to_excel(output_cbna_rwa_file, index=False)

# %% [## 17. Summary]
print('\n### Step 2 complete — Outlook RWA ###')
print(f'Output directory: {OUTPUT_DIR}')
for f in [
    OUTPUT_CG_OUTLOOK_FULL_FILENAME,
    OUTPUT_CBNA_OUTLOOK_FULL_FILENAME,
    OUTPUT_CG_RWA_DATA_FILENAME,
    OUTPUT_CBNA_RWA_DATA_FILENAME,
    OUTPUT_CONTROL_FILENAME,
]:
    print(f'  - {f}')
