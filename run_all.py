"""
run_all.py
----------
Orchestrator for the Excel Tracer pipeline:

  1. trace_formulas.py   → output/dependency_graph.json
  2. flatten_formulas.py → output/flattened_formulas.json
  3. generate_sql.py     → output/output.sql

Usage:
    python run_all.py workbook.xlsx
    python run_all.py workbook.xlsx --from flatten
    python run_all.py workbook.xlsx --only sql

Steps: trace | flatten | sql
"""

import argparse
import subprocess
import sys
import time
from pathlib import Path

STEPS = [
    ("trace",   "trace_formulas.py",   "Formula Tracing & Dependency Graph"),
    ("flatten", "flatten_formulas.py", "Formula Flattening & Normalization"),
    ("sql",     "generate_sql.py",     "SQL Generation"),
]
STEP_NAMES = [s[0] for s in STEPS]


def run_script(script: str, label: str, extra_args: list[str]) -> bool:
    print(f"\n{'='*60}")
    print(f"  STEP: {label}")
    print(f"  Script: {script}")
    print(f"{'='*60}")
    start = time.time()
    result = subprocess.run(
        [sys.executable, script] + extra_args,
        cwd=Path(__file__).parent,
    )
    elapsed = time.time() - start
    if result.returncode == 0:
        print(f"\n  OK  {label} completed in {elapsed:.1f}s")
        return True
    print(f"\n  FAIL  {label} failed (exit {result.returncode})")
    return False


def main():
    parser = argparse.ArgumentParser(description="Excel Tracer Pipeline")
    parser.add_argument("workbook", help="Path to the Excel file to analyse")
    parser.add_argument(
        "--from", dest="start_from", choices=STEP_NAMES, default="trace",
        help="Start from this step (default: trace)",
    )
    parser.add_argument(
        "--only", choices=STEP_NAMES, default=None,
        help="Run only this single step",
    )
    args = parser.parse_args()

    extra = [args.workbook]

    if args.only:
        steps_to_run = [s for s in STEPS if s[0] == args.only]
    else:
        start_idx = STEP_NAMES.index(args.start_from)
        steps_to_run = STEPS[start_idx:]

    print("\n" + "=" * 60)
    print("  Excel Tracer Pipeline")
    print(f"  Workbook: {args.workbook}")
    print("=" * 60)

    overall_start = time.time()
    for key, script, label in steps_to_run:
        if not run_script(script, label, extra):
            print(f"\nPipeline aborted at '{key}'. Fix the error and re-run.\n")
            sys.exit(1)

    elapsed = time.time() - overall_start
    print(f"\n{'='*60}")
    print(f"  Pipeline complete in {elapsed:.1f}s")
    print(f"  Outputs written to: output/")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
