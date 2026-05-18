"""
Shared constants for the RWA Model Convergence pipeline.

Column-name literals, expected-column lists, and PMF classification values
are centralised here so step1_convergence.py and step2_outlook_rwa.py stay
in sync with each other and with the source data schemas.

Naming convention
-----------------
*_CDE  — code columns (convergence/aggregator table only)
*_DESC — description columns (balance-sheet and convergence)
"""
from __future__ import annotations

# ── Table names ───────────────────────────────────────────────────────────────
INPUT_CG_TABLE            = "outlook_balancesheet_cg"
INPUT_CBNA_TABLE          = "outlook_balancesheet_cbna"
OUTPUT_CG_OUTLOOK_TABLE   = "cg_outlook_tap_env"
OUTPUT_CBNA_OUTLOOK_TABLE = "cbna_com_outlook_tap_env"
OUTPUT_CG_ADDON_TABLE     = "addon_all_cg_tap_env"
OUTPUT_CBNA_ADDON_TABLE   = "addon_all_cbna_tap_env"
INPUT_CONVERGENCE_TABLE   = "aggregator_for_convergence"

# ── Convergence / aggregator columns ─────────────────────────────────────────
REPORTABLE_ENTITY_IS_CG   = "Reportable Entity is CG"
REPORTABLE_ENTITY_IS_CBNA = "Reportable Entity is CBNA"
FINANCE_PMF_LEVEL_5_DESC  = "Finance PMF Level 5 Description"
GAP_AMOUNT                = "GAP Amount"
SA_RWA_AMT                = "SA RWA Amount"
ADV_CG_TOTAL_RWA_AMT      = "Adv. CG Total RWA Amount with 1.06 Multiplier"
ADV_CBNA_TOTAL_RWA_AMT    = "Adv. CBNA Total RWA Amount with 1.06 Multiplier"
OPFR_ID                   = "Quarter Id"
CCAL_CYCLE                = "RWA Cycle"

# Managed Segment — code columns (convergence only)
MNGD_SGMT_L4_CDE = "Managed Segment Level 4 Code"
MNGD_SGMT_L3_CDE = "Managed Segment Level 3 Code"
MNGD_SGMT_L2_CDE = "Managed Segment Level 2 Code"
MNGD_SGMT_L1_CDE = "Managed Segment Level 1 Code"

# Managed Segment — description columns (balance-sheet and convergence)
MNGD_SGMT_L4_DESC = "Managed Segment Level 4 Description"
MNGD_SGMT_L3_DESC = "Managed Segment Level 3 Description"
MNGD_SGMT_L2_DESC = "Managed Segment Level 2 Description"
MNGD_SGMT_L1_DESC = "Managed Segment Level 1 Description"

# Managed Geography — description columns (balance-sheet and convergence)
MNGD_GEO_L4_DESC = "Managed Geography Level 4 Description"
MNGD_GEO_L3_DESC = "Managed Geography Level 3 Description"
MNGD_GEO_L2_DESC = "Managed Geography Level 2 Description"
MNGD_GEO_L1_DESC = "Managed Geography Level 1 Description"

RWA_EXPOSURE_TYPE_DESC = "RWA Exposure Type Description"

# ── Balance sheet columns ─────────────────────────────────────────────────────
M1_USDOLLAR = "M1 USDOLLAR"
M2_USDOLLAR = "M2 USDOLLAR"
M3_USDOLLAR = "M3 USDOLLAR"

# ── Expected-column lists (validation + column pruning) ───────────────────────
EXPECTED_BALANCE_SHEET_COLS = [
    M1_USDOLLAR,
    M2_USDOLLAR,
    M3_USDOLLAR,
    REPORTABLE_ENTITY_IS_CG,
    REPORTABLE_ENTITY_IS_CBNA,
    FINANCE_PMF_LEVEL_5_DESC,
    GAP_AMOUNT,
    SA_RWA_AMT,
    ADV_CG_TOTAL_RWA_AMT,
    ADV_CBNA_TOTAL_RWA_AMT,
    OPFR_ID,
    CCAL_CYCLE,
    MNGD_SGMT_L4_DESC,
    MNGD_SGMT_L3_DESC,
    MNGD_SGMT_L2_DESC,
    MNGD_SGMT_L1_DESC,
    MNGD_GEO_L4_DESC,
    MNGD_GEO_L3_DESC,
    MNGD_GEO_L2_DESC,
    MNGD_GEO_L1_DESC,
    RWA_EXPOSURE_TYPE_DESC,
]

EXPECTED_CONVERGENCE_COLS = [
    REPORTABLE_ENTITY_IS_CG,
    REPORTABLE_ENTITY_IS_CBNA,
    FINANCE_PMF_LEVEL_5_DESC,
    GAP_AMOUNT,
    SA_RWA_AMT,
    ADV_CG_TOTAL_RWA_AMT,
    ADV_CBNA_TOTAL_RWA_AMT,
    OPFR_ID,
    CCAL_CYCLE,
    MNGD_SGMT_L4_CDE,
    MNGD_SGMT_L3_CDE,
    MNGD_SGMT_L2_CDE,
    MNGD_SGMT_L1_CDE,
    MNGD_SGMT_L4_DESC,
    MNGD_SGMT_L3_DESC,
    MNGD_SGMT_L2_DESC,
    MNGD_SGMT_L1_DESC,
    MNGD_GEO_L4_DESC,
    MNGD_GEO_L3_DESC,
    MNGD_GEO_L2_DESC,
    MNGD_GEO_L1_DESC,
    RWA_EXPOSURE_TYPE_DESC,
]

# ── PMF classification ────────────────────────────────────────────────────────
CREDIT_RISK_PMF_MAP     = "credit_risk_pmf_map"
NON_CREDIT_RISK_PMF_MAP = "non_credit_risk_pmf_map"
KEY_COMPONENT_MAP       = "key_component_map"

CREDIT_RISK_PMF_VALUES = [
    "Deposits with Banks (l2)",
    "Investments (l2)",
    "Loans (l2)",
    "Trading Account Assets (l2)",
    "Commitments (l2)",
    "Brokerage Receivables (l2)",
    "Federal Funds Purchased and Securities Sold Under Agreements to Repurchase (l2)",
    "Securities Sold and Resales (l2)",
    "Intra-Liabilities (l2)",
    "Other Assets (l2)",
]

NON_CREDIT_RISK_PMF_VALUES = [
    "Trading Account Assets (l2)",
    "Commitments to Purchase Forward-Dated Securities (l2)",
    "Securities Sold or Sold Under Repurchase Agreements (l2)",
    "Fixed Assets Net of Depreciation and Amortization (l2)",
    "Intra-Liabilities (l2)",
    "Other Assets (l2)",
]
