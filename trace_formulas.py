"""
trace_formulas.py  —  Step 1
-----------------------------
Open the Excel workbook, identify the output sheet, trace all formula
dependencies, and write output/dependency_graph.json.

Output schema:
{
  "workbook": "path/to/file.xlsx",
  "output_sheet": "Summary",
  "sheets": ["Summary", "RawData", ...],
  "cells": {
    "Summary|1|3": {
      "sheet": "Summary", "row": 1, "col": 3,
      "formula": "=SUM(RawData!B2:B100)",
      "is_data": false,
      "value": 42.0
    },
    ...
  },
  "edges": [["Summary|1|3", "RawData|2|2"], ...]
}
"""

import sys
import json
from pathlib import Path

import yaml

from lib.workbook import Workbook
from lib.dependency_graph import build_graph, sheet_summary, source_cells


def addr_key(addr) -> str:
    return f"{addr[0]}|{addr[1]}|{addr[2]}"


def load_config() -> dict:
    cfg_path = Path(__file__).parent / "config.yaml"
    with open(cfg_path) as f:
        return yaml.safe_load(f)


def find_output_sheet(wb: Workbook, candidates: list[str]) -> str:
    for name in candidates:
        if name in wb.sheet_names:
            return name
    # Fall back to last sheet
    print(f"  Warning: none of {candidates} found. Using last sheet: {wb.sheet_names[-1]}")
    return wb.sheet_names[-1]


def main():
    if len(sys.argv) < 2:
        print("Usage: python trace_formulas.py <workbook.xlsx>")
        sys.exit(1)

    wb_path = Path(sys.argv[1])
    if not wb_path.exists():
        print(f"File not found: {wb_path}")
        sys.exit(1)

    cfg = load_config()
    output_dir = Path(cfg["output_dir"])
    output_dir.mkdir(exist_ok=True)

    print(f"\nLoading workbook: {wb_path}")
    wb = Workbook(wb_path)
    print(f"  Sheets found: {wb.sheet_names}")

    output_sheet = find_output_sheet(wb, cfg["output_sheets"])
    print(f"  Output sheet identified: '{output_sheet}'")

    print("\nBuilding dependency graph...")
    G = build_graph(wb, output_sheet)

    node_count = G.number_of_nodes()
    edge_count = G.number_of_edges()
    print(f"  Nodes (cells): {node_count}")
    print(f"  Edges (dependencies): {edge_count}")
    print(f"  Sheet breakdown: {sheet_summary(G)}")
    src = source_cells(G)
    print(f"  Source (data) cells: {len(src)}")

    # Serialise
    cells_out = {}
    for addr, node_data in G.nodes(data=True):
        node = node_data["data"]
        cells_out[addr_key(addr)] = {
            "sheet": node.sheet,
            "row": node.row,
            "col": node.col,
            "col_letter": wb.col_letter(node.col),
            "formula": node.formula,
            "is_data": node.is_data,
            "value": str(node.value) if node.value is not None else None,
        }

    edges_out = [
        [addr_key(u), addr_key(v)]
        for u, v in G.edges()
    ]

    result = {
        "workbook": str(wb_path),
        "output_sheet": output_sheet,
        "sheets": wb.sheet_names,
        "named_ranges": wb.named_ranges(),
        "cells": cells_out,
        "edges": edges_out,
    }

    out_path = output_dir / "dependency_graph.json"
    with open(out_path, "w") as f:
        json.dump(result, f, indent=2)

    print(f"\n  Written: {out_path}")


if __name__ == "__main__":
    main()
