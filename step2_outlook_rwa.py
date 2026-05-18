"""
Step 2: Outlook RWA

Reads step1_convergence.py outputs + adjustment/mapping files, applies
adjustments, maps PMF codes, builds upload-template pivots, and writes
output workbooks for CG and CBNA.

Prerequisites: run step1_convergence.py first.
"""
from __future__ import annotations

import warnings
import numpy as np
import pandas as pd
from pathlib import Path

import rwa_constants as rc

pd.set_option('display.max_columns', 500)

# ── PARAMETERS — update before each run ───────────────────────────────────────

Q0 = 'Dec_2025'

# Must match DATA_DIR in step1_convergence.py
DATA_DIR = Path(r'C:\Users\r10995\Desktop\aa-data\my\2026')

# Input filenames — step1 outputs (in DATA_DIR/output)
INPUT_CG_OUTLOOK_FILENAME  = 'cg_outlook_tap_env.xlsx'
INPUT_CBNA_OUTLOOK_FILENAME = 'cbna_com_outlook_tap_env.xlsx'
INPUT_CG_ADDON_FILENAME    = 'addon_all_cg.xlsx'
INPUT_CBNA_ADDON_FILENAME  = 'addon_all_cbna.xlsx'

# Input filenames — reference data (in DATA_DIR/input)
INPUT_ADJUSTMENTS_FILENAME     = 'adjustments.xlsx'
INPUT_PMF_RWA_MAPPING_FILENAME = 'pmf_frm_mapping.xlsx'
INPUT_CONVERGENCE_FILENAME     = 'aggregator_for_convergence.xlsx'

# Output filenames (written to DATA_DIR/output)
OUTPUT_CG_OUTLOOK_FILENAME   = 'cg_outlook.xlsx'
OUTPUT_CBNA_OUTLOOK_FILENAME = 'cbna_outlook.xlsx'
OUTPUT_CG_PIVOT_FILENAME     = 'cg_outlook_pivot.xlsx'
OUTPUT_CBNA_PIVOT_FILENAME   = 'cbna_outlook_pivot.xlsx'
OUTPUT_CG_RWA_DATA_FILENAME  = 'cg_rwa_data.xlsx'
OUTPUT_CBNA_RWA_DATA_FILENAME = 'cbna_rwa_data.xlsx'
OUTPUT_CONTROL_FILENAME      = 'control_file.xlsx'

# ── Derived paths (no edits needed) ───────────────────────────────────────────
INPUT_DIR  = DATA_DIR / 'input'
OUTPUT_DIR = DATA_DIR / 'output'
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

input_cg_outlook_file   = OUTPUT_DIR / INPUT_CG_OUTLOOK_FILENAME
input_cbna_outlook_file = OUTPUT_DIR / INPUT_CBNA_OUTLOOK_FILENAME
cg_addon_file           = OUTPUT_DIR / INPUT_CG_ADDON_FILENAME
cbna_addon_file         = OUTPUT_DIR / INPUT_CBNA_ADDON_FILENAME

adjustments_file  = INPUT_DIR / INPUT_ADJUSTMENTS_FILENAME
pmf_mapping_file  = INPUT_DIR / INPUT_PMF_RWA_MAPPING_FILENAME
convergence_file  = INPUT_DIR / INPUT_CONVERGENCE_FILENAME

output_cg_file       = OUTPUT_DIR / OUTPUT_CG_OUTLOOK_FILENAME
output_cbna_file     = OUTPUT_DIR / OUTPUT_CBNA_OUTLOOK_FILENAME
output_cg_pivot_file = OUTPUT_DIR / OUTPUT_CG_PIVOT_FILENAME
output_cbna_pivot_file = OUTPUT_DIR / OUTPUT_CBNA_PIVOT_FILENAME
output_cg_rwa_file   = OUTPUT_DIR / OUTPUT_CG_RWA_DATA_FILENAME
output_cbna_rwa_file = OUTPUT_DIR / OUTPUT_CBNA_RWA_DATA_FILENAME
output_control_file  = OUTPUT_DIR / OUTPUT_CONTROL_FILENAME

print(f"Input dir:  {INPUT_DIR}")
print(f"Output dir: {OUTPUT_DIR}")

# %% [1. Verify input files exist]

all_input_files = [
    adjustments_file, input_cg_outlook_file, input_cbna_outlook_file,
    cg_addon_file, cbna_addon_file, pmf_mapping_file, convergence_file,
]
for f in all_input_files:
    if not f.exists():
        warnings.warn(f"INPUT FILE NOT FOUND: {f}")
    else:
        print(f"  Found: {f.name}")

# %% [2. Load inputs]

src_cg_adjustments   = pd.read_excel(adjustments_file, sheet_name='Adjustments - CG')
src_cbna_adjustments = pd.read_excel(adjustments_file, sheet_name='Adjustments - CBNA')
src_cg_outlook       = pd.read_excel(input_cg_outlook_file)
src_cbna_outlook     = pd.read_excel(input_cbna_outlook_file)
src_cg_addon         = pd.read_excel(cg_addon_file)
src_cbna_addon       = pd.read_excel(cbna_addon_file)
src_pmf_mapping      = pd.read_excel(pmf_mapping_file)
convergence          = pd.read_excel(convergence_file)

print(f"CG adjustments:  {len(src_cg_adjustments):,} rows")
print(f"CBNA adjustments:{len(src_cbna_adjustments):,} rows")
print(f"CG outlook:      {len(src_cg_outlook):,} rows")
print(f"CBNA outlook:    {len(src_cbna_outlook):,} rows")
print(f"CG addon:        {len(src_cg_addon):,} rows")
print(f"CBNA addon:      {len(src_cbna_addon):,} rows")
print(f"PMF mapping:     {len(src_pmf_mapping):,} rows")
print(f"Convergence:     {len(convergence):,} rows")

# %% [3. Data quality — PMF mapping cardinality]

# Upload-template column name used as PMF join key in pmf_mapping file
PMF_L5_COL_IN_MAPPING = rc.FINANCE_PMF_LEVEL_5_DESC

for subset_col, label in [
    (PMF_L5_COL_IN_MAPPING,    'PMF L5'),
    ('Managed Segment L4 Descr', 'Managed Segment L4 Descr'),
]:
    if subset_col not in src_pmf_mapping.columns:
        warnings.warn(f"PMF mapping missing column: {subset_col!r}")
        continue
    dupes = src_pmf_mapping.duplicated(subset=[subset_col], keep=False).sum()
    if dupes:
        warnings.warn(f"PMF mapping has {dupes} duplicate {label} entries — not 1:1")
    else:
        print(f"  ✓ PMF mapping is 1:1 on {label}")

# %% [Local column-name constants — upload-template schema]
# These are the abbreviated column names used in adjustments / outlook template files.
# They differ from rc.* which use the full "Level N Description" form.

MNGD_SGMT_L4_DESC  = 'Managed Segment L4 Descr'
MNGD_SGMT_L3_DESC  = 'Managed Segment L3 Descr'
MNGD_SGMT_L2_DESC  = 'Managed Segment L2 Descr'
MNGD_SGMT_L1_DESC  = 'Managed Segment L1 Descr'
MNGD_GEO_L4_DESC   = 'Managed Geography L4 Descr'
MNGD_GEO_L3_DESC   = 'Managed Geography L3 Descr'
MNGD_GEO_L2_DESC   = 'Managed Geography L2 Descr'
MNGD_GEO_L1_DESC   = 'Managed Geography L1 Descr'

SA_RWA           = 'SA RWA'
AA_RWA           = 'AA RWA'
ERWA_RWA         = 'ERWA RWA'
NVA_CALC         = 'NVA Calc'
REPORTING_LAYER  = 'Reporting Layer'
SA_ACCOUNT_NUM   = 'SA Account #'
QUARTER_ID       = 'Quarter Id'
PMF_ACCT_L5_DESC = 'PMF Account L5 Descr'

MARKETS_FILTER            = 'Markets Filter'
DISCONTINUED_OPS_L3       = 'Discontinued Ops (L2)'
LEGACY_FRANCHISES_L3      = 'Legacy Franchises (L3)'
LEGACY_HOLDINGS_ASSETS_L4 = 'Legacy Holdings Assets (L4)'
MARKETS_L2                = 'Markets (L2)'

# %% [4. Format adjustments]


def format_adjustments(input_df: pd.DataFrame) -> pd.DataFrame:
    for col in input_df.select_dtypes(include=['object']).columns:
        coerced = pd.to_numeric(input_df[col], errors='coerce')
        if coerced.notna().any():
            input_df[col] = coerced
        else:
            input_df[col] = input_df[col].fillna('N/A')
    return input_df


cg_adjustments   = format_adjustments(src_cg_adjustments.copy())
cbna_adjustments = format_adjustments(src_cbna_adjustments.copy())

# %% [5. Rename addon columns to upload-template schema]


def rename_addon_columns(input_df: pd.DataFrame, entity: str) -> pd.DataFrame:
    """Map convergence description columns to upload-template abbreviated names.

    Code columns (MNGD_SGMT_*_CDE) are dropped — description columns are kept
    and renamed. Renaming both would create duplicate column names.
    """
    rwa_col = rc.ADV_CG_TOTAL_RWA_AMT if entity == 'CG' else rc.ADV_CBNA_TOTAL_RWA_AMT

    code_cols = [rc.MNGD_SGMT_L4_CDE, rc.MNGD_SGMT_L3_CDE,
                 rc.MNGD_SGMT_L2_CDE, rc.MNGD_SGMT_L1_CDE]
    input_df = input_df.drop(columns=[c for c in code_cols if c in input_df.columns])

    renames = {
        rc.MNGD_SGMT_L4_DESC: MNGD_SGMT_L4_DESC,
        rc.MNGD_SGMT_L3_DESC: MNGD_SGMT_L3_DESC,
        rc.MNGD_SGMT_L2_DESC: MNGD_SGMT_L2_DESC,
        rc.MNGD_SGMT_L1_DESC: MNGD_SGMT_L1_DESC,
        rc.MNGD_GEO_L4_DESC:  MNGD_GEO_L4_DESC,
        rc.MNGD_GEO_L3_DESC:  MNGD_GEO_L3_DESC,
        rc.MNGD_GEO_L2_DESC:  MNGD_GEO_L2_DESC,
        rc.MNGD_GEO_L1_DESC:  MNGD_GEO_L1_DESC,
        rwa_col:               AA_RWA,
        rc.SA_RWA_AMT:         SA_RWA,
        rc.GAP_AMOUNT:         NVA_CALC,
        rc.OPFR_ID:            QUARTER_ID,
    }
    input_df = input_df.rename(columns={k: v for k, v in renames.items() if k in input_df.columns})
    input_df['Entity'] = entity
    return input_df


cg_outlook     = src_cg_outlook.copy()
cbna_outlook   = src_cbna_outlook.copy()
addon_all_cg   = rename_addon_columns(src_cg_addon.copy(),   'CG')
addon_all_cbna = rename_addon_columns(src_cbna_addon.copy(), 'CBNA')

print(f"CG outlook rows:     {len(cg_outlook):,}")
print(f"CBNA outlook rows:   {len(cbna_outlook):,}")
print(f"Addon CG rows:       {len(addon_all_cg):,}")
print(f"Addon CBNA rows:     {len(addon_all_cbna):,}")

# %% [6. Rename PMF mapping join key]

pmf_mapping = src_pmf_mapping.rename(columns={'PMF L5': rc.FINANCE_PMF_LEVEL_5_DESC}).copy()

# %% [7. Concatenate adjustments + outlook + addon]

cg_concat = pd.concat(
    [cg_adjustments, cg_outlook, addon_all_cg], ignore_index=True
).copy()
cbna_concat = pd.concat(
    [cbna_adjustments, cbna_outlook, addon_all_cbna], ignore_index=True
).copy()

print(f"CG concatenated:   {len(cg_concat):,} rows")
print(f"CBNA concatenated: {len(cbna_concat):,} rows")

# %% [8. Filter unknown quarter IDs]

for df in [cg_concat, cbna_concat]:
    df.drop(df[df[QUARTER_ID] == 'Unknown'].index, inplace=True)
    df[QUARTER_ID] = pd.to_numeric(df[QUARTER_ID], errors='coerce')

cg_concat['Entity']   = 'CG'
cbna_concat['Entity'] = 'CBNA'

cg_raw_data   = cg_concat.copy()
cbna_raw_data = cbna_concat.copy()

print(f"CG after filtering unknowns:   {len(cg_concat):,} rows")
print(f"CBNA after filtering unknowns: {len(cbna_concat):,} rows")

# %% [9. Legacy franchises breakdown]

legacy_cg = cg_concat[cg_concat[MNGD_SGMT_L4_DESC] == LEGACY_HOLDINGS_ASSETS_L4].copy()
legacy_cbna = cbna_concat[cbna_concat[MNGD_SGMT_L4_DESC] == LEGACY_HOLDINGS_ASSETS_L4].copy()

print(f"Legacy Holdings CG rows:   {len(legacy_cg):,}")
print(f"Legacy Holdings CBNA rows: {len(legacy_cbna):,}")

cg_concat['_legacy_flag']   = cg_concat[MNGD_SGMT_L3_DESC].isin([LEGACY_FRANCHISES_L3, DISCONTINUED_OPS_L3])
cbna_concat['_legacy_flag'] = cbna_concat[MNGD_SGMT_L3_DESC].isin([LEGACY_FRANCHISES_L3, DISCONTINUED_OPS_L3])

# %% [10. PMF join — map SA account numbers]

pre_cg   = len(cg_concat)
pre_cbna = len(cbna_concat)

frm_cg_output   = cg_concat.merge(pmf_mapping,   on=rc.FINANCE_PMF_LEVEL_5_DESC, how='left')
frm_cbna_output = cbna_concat.merge(pmf_mapping, on=rc.FINANCE_PMF_LEVEL_5_DESC, how='left')

if len(frm_cg_output) > pre_cg:
    warnings.warn(f"CG PMF join caused row expansion ({pre_cg:,} → {len(frm_cg_output):,})")
if len(frm_cbna_output) > pre_cbna:
    warnings.warn(f"CBNA PMF join caused row expansion ({pre_cbna:,} → {len(frm_cbna_output):,})")

for label, df in [('CG', frm_cg_output), ('CBNA', frm_cbna_output)]:
    unmatched = df[SA_ACCOUNT_NUM].isna().sum() if SA_ACCOUNT_NUM in df.columns else len(df)
    if unmatched:
        warnings.warn(f"{label}: {unmatched:,} rows have no PMF mapping match")
    else:
        print(f"  ✓ {label}: PMF mapping joined successfully")

for label, df in [('CG', frm_cg_output), ('CBNA', frm_cbna_output)]:
    key_cols = [MNGD_SGMT_L4_DESC, MNGD_SGMT_L3_DESC, rc.FINANCE_PMF_LEVEL_5_DESC, MNGD_GEO_L4_DESC]
    valid_keys = [c for c in key_cols if c in df.columns]
    print(f"  {label}: {len(df[valid_keys].drop_duplicates()):,} unique key combinations")

# %% [11. Coerce RWA columns to numeric]


def format_columns_before_pivot(input_df: pd.DataFrame) -> pd.DataFrame:
    for col in [ERWA_RWA, AA_RWA, SA_RWA]:
        if col in input_df.columns:
            input_df[col] = pd.to_numeric(input_df[col], errors='coerce')
    if SA_ACCOUNT_NUM in input_df.columns:
        input_df[SA_ACCOUNT_NUM] = pd.to_numeric(input_df[SA_ACCOUNT_NUM], errors='coerce').fillna(0)
    return input_df


frm_cg_output   = format_columns_before_pivot(frm_cg_output)
frm_cbna_output = format_columns_before_pivot(frm_cbna_output)

# %% [12. Upload template pivots]

_PIVOT_INDEX = [
    MNGD_SGMT_L4_DESC, MNGD_SGMT_L3_DESC, MNGD_SGMT_L2_DESC,
    rc.FINANCE_PMF_LEVEL_5_DESC, MNGD_GEO_L4_DESC, MNGD_GEO_L3_DESC,
    REPORTING_LAYER, SA_ACCOUNT_NUM,
]


def _make_rwa_pivot(df: pd.DataFrame, value_col: str, label: str) -> pd.DataFrame:
    valid_keys = [c for c in _PIVOT_INDEX if c in df.columns]
    piv = df.groupby(valid_keys, dropna=False)[[value_col]].sum()
    piv.columns = [label]
    return piv


def create_upload_template_pivots(frm_output: pd.DataFrame) -> pd.DataFrame:
    """Aggregate AA / SA / ERWA by pivot index into one upload-ready DataFrame."""
    parts = [
        _make_rwa_pivot(frm_output, AA_RWA, 'AA'),
        _make_rwa_pivot(frm_output, SA_RWA, 'SA'),
    ]
    if ERWA_RWA in frm_output.columns and frm_output[ERWA_RWA].notna().any():
        parts.append(_make_rwa_pivot(frm_output, ERWA_RWA, 'ERWA'))
    pivots = pd.concat(parts, axis=1)
    pivots.columns.name = None
    return pivots.reset_index()


cg_frm_pivots   = create_upload_template_pivots(frm_cg_output)
cbna_frm_pivots = create_upload_template_pivots(frm_cbna_output)

print(f"CG pivot rows:   {len(cg_frm_pivots):,}")
print(f"CBNA pivot rows: {len(cbna_frm_pivots):,}")

# %% [13. Markets filter]


def create_markets_filter(input_df: pd.DataFrame) -> pd.DataFrame:
    """Tag rows 'Keep' if Markets L2 segment with non-Remove RWA exposure, else 'Remove'."""
    nw = np.where(
        (input_df.get(MNGD_SGMT_L2_DESC, '') == MARKETS_L2)
        & (input_df.get(rc.RWA_EXPOSURE_TYPE_DESC, '') != 'Remove'),
        'Keep',
        'Remove',
    )
    input_df[MARKETS_FILTER] = nw
    return input_df


frm_cg_output   = create_markets_filter(frm_cg_output)
frm_cbna_output = create_markets_filter(frm_cbna_output)
print("Markets filter applied")

# %% [14. Fill numerics and map monthly balance-sheet columns]


def format_upload_template(input_df: pd.DataFrame) -> pd.DataFrame:
    num_cols = input_df.select_dtypes(include='number').columns
    input_df[num_cols] = input_df[num_cols].fillna(0)
    for rc_col, month_col in [
        (rc.M1_USDOLLAR, 'Month1'),
        (rc.M2_USDOLLAR, 'Month2'),
        (rc.M3_USDOLLAR, 'Month3'),
    ]:
        input_df[month_col] = input_df.get(rc_col, pd.Series(0, index=input_df.index)).fillna(0)
    return input_df


frm_cg_output   = format_upload_template(frm_cg_output)
frm_cbna_output = format_upload_template(frm_cbna_output)

# %% [15. Control summary tables]

_CTRL_GROUP = [MNGD_SGMT_L2_DESC, MNGD_SGMT_L3_DESC,
               rc.FINANCE_PMF_LEVEL_5_DESC, rc.RWA_EXPOSURE_TYPE_DESC]


def build_frm_raw_data(frm_df: pd.DataFrame, entity: str) -> pd.DataFrame:
    group_cols = [c for c in _CTRL_GROUP if c in frm_df.columns]
    agg_cols   = [c for c in [AA_RWA, SA_RWA, ERWA_RWA] if c in frm_df.columns]
    result = frm_df.groupby(group_cols, dropna=False)[agg_cols].sum().reset_index()
    result['Entity'] = entity
    return result


cg_frm_raw   = build_frm_raw_data(frm_cg_output,   'CG')
cbna_frm_raw = build_frm_raw_data(frm_cbna_output, 'CBNA')

_CONV_GROUP = [c for c in _CTRL_GROUP if c in convergence.columns]

cg_convergence_control = (
    convergence[convergence[rc.REPORTABLE_ENTITY_IS_CG] == 'Y']
    .groupby(_CONV_GROUP, dropna=False)[[rc.ADV_CG_TOTAL_RWA_AMT, rc.SA_RWA_AMT]]
    .sum().reset_index()
)
cbna_convergence_control = (
    convergence[convergence[rc.REPORTABLE_ENTITY_IS_CBNA] == 'Y']
    .groupby(_CONV_GROUP, dropna=False)[[rc.ADV_CBNA_TOTAL_RWA_AMT, rc.SA_RWA_AMT]]
    .sum().reset_index()
)

_raw_group    = [c for c in _CTRL_GROUP if c in cg_raw_data.columns]
_raw_agg_cols = [c for c in [AA_RWA, SA_RWA] if c in cg_raw_data.columns]

cg_raw_data_control   = cg_raw_data.groupby(_raw_group, dropna=False)[_raw_agg_cols].sum().reset_index()
cbna_raw_data_control = cbna_raw_data.groupby(_raw_group, dropna=False)[_raw_agg_cols].sum().reset_index()

param_data = pd.DataFrame(
    [('Q0', Q0), ('DATA_DIR', str(DATA_DIR))],
    columns=['Parameter', 'Value'],
)

# %% [16. Write outputs]

print(f"\nWriting {output_control_file}")
with pd.ExcelWriter(output_control_file) as writer:
    cg_convergence_control.to_excel(writer,   sheet_name='CG Convergence Control',   index=False)
    cg_frm_raw.to_excel(writer,               sheet_name='CG FRM Control',           index=False)
    cbna_convergence_control.to_excel(writer, sheet_name='CBNA Convergence Control', index=False)
    cbna_frm_raw.to_excel(writer,             sheet_name='CBNA FRM Control',         index=False)
    cg_raw_data_control.to_excel(writer,      sheet_name='CG Raw Data Control',      index=False)
    cbna_raw_data_control.to_excel(writer,    sheet_name='CBNA Raw Data Control',    index=False)
    param_data.to_excel(writer,               sheet_name='Parameters',               index=False)

for path, df in [
    (output_cg_file,        frm_cg_output),
    (output_cbna_file,      frm_cbna_output),
    (output_cg_pivot_file,  cg_frm_pivots),
    (output_cbna_pivot_file, cbna_frm_pivots),
    (output_cg_rwa_file,    cg_raw_data),
    (output_cbna_rwa_file,  cbna_raw_data),
]:
    print(f"Writing {path}")
    df.to_excel(path, index=False)

print('\n### Step 2 complete — Outlook RWA ###')
print(f'Output directory: {OUTPUT_DIR}')
for name in [
    OUTPUT_CG_OUTLOOK_FILENAME, OUTPUT_CBNA_OUTLOOK_FILENAME,
    OUTPUT_CG_PIVOT_FILENAME,   OUTPUT_CBNA_PIVOT_FILENAME,
    OUTPUT_CG_RWA_DATA_FILENAME, OUTPUT_CBNA_RWA_DATA_FILENAME,
    OUTPUT_CONTROL_FILENAME,
]:
    print(f'  - {name}')
