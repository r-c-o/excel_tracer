"""
flatten_formulas.py  —  Step 2
--------------------------------
Read dependency_graph.json and the workbook, reconstruct the graph,
then flatten every output-sheet formula by substituting all intermediate
references until only source (data) cells remain.

Also runs column-pattern detection to identify vectorisable columns.

Output: output/flattened_formulas.json
{
  "output_sheet": "Summary",
  "flattened_cells": {
    "Summary|1|3": {
      "formula": "=SUM(RawData!B2:B100)",
      "flattened_repr": "SUM(RawData.B)",
      ...
    }
  },
  "column_patterns": {
    "3": { "col": 3, "min_row": 2, "max_row": 101, "abstract_repr": "..." }
  }
}
"""

import sys
import json
from pathlib import Path

import yaml
import networkx as nx

from lib.workbook import Workbook
from lib.dependency_graph import CellNode
from lib.formula_parser import parse_formula
from lib.flattener import flatten_all, detect_column_patterns
from lib.sql_emitter import SQLEmitter


def load_config() -> dict:
    with open(Path(__file__).parent / "config.yaml") as f:
        return yaml.safe_load(f)


def _ast_repr(ast, emitter: SQLEmitter) -> str:
    try:
        return emitter._ast_to_sql(ast)
    except Exception as ex:
        return f"/* error: {ex} */"


def main():
    if len(sys.argv) < 2:
        print("Usage: python flatten_formulas.py <workbook.xlsx>")
        sys.exit(1)

    wb_path = Path(sys.argv[1])
    cfg = load_config()
    output_dir = Path(cfg["output_dir"])
    graph_path = output_dir / "dependency_graph.json"

    if not graph_path.exists():
        print(f"dependency_graph.json not found. Run trace_formulas.py first.")
        sys.exit(1)

    print(f"\nLoading dependency graph: {graph_path}")
    with open(graph_path) as f:
        graph_data = json.load(f)

    output_sheet = graph_data["output_sheet"]
    max_depth = cfg.get("max_substitution_depth", 20)

    # Reconstruct nx.DiGraph with CellNode objects
    G = nx.DiGraph()
    for key, cell in graph_data["cells"].items():
        addr = (cell["sheet"], cell["row"], cell["col"])
        formula = cell["formula"]
        ast = parse_formula(formula) if formula else None
        node = CellNode(
            sheet=cell["sheet"], row=cell["row"], col=cell["col"],
            formula=formula, ast=ast, value=cell["value"],
            is_data=cell["is_data"],
        )
        G.add_node(addr, data=node)

    for src_key, dst_key in graph_data["edges"]:
        def _key_to_addr(k):
            parts = k.split("|")
            return (parts[0], int(parts[1]), int(parts[2]))
        G.add_edge(_key_to_addr(src_key), _key_to_addr(dst_key))

    print(f"  Graph: {G.number_of_nodes()} nodes, {G.number_of_edges()} edges")
    print(f"  Output sheet: {output_sheet}")

    emitter = SQLEmitter(
        dialect=cfg.get("sql_dialect", "oracle"),
        schema=cfg.get("sql_schema", ""),
        overrides=cfg.get("function_overrides", {}),
    )

    print("\nFlattening formulas...")
    flattened = flatten_all(G, output_sheet, max_depth=max_depth)
    print(f"  Flattened {len(flattened)} output cells")

    print("Detecting column patterns (vectorisation)...")
    patterns = detect_column_patterns(flattened, output_sheet)
    print(f"  Vectorisable columns: {len(patterns)}")

    # Serialise
    flattened_out = {}
    for addr, ast in flattened.items():
        key = f"{addr[0]}|{addr[1]}|{addr[2]}"
        orig_formula = G.nodes[addr]["data"].formula or ""
        flattened_out[key] = {
            "sheet": addr[0],
            "row": addr[1],
            "col": addr[2],
            "original_formula": orig_formula,
            "flattened_sql": _ast_repr(ast, emitter),
        }

    patterns_out = {}
    for col_idx, pat in patterns.items():
        patterns_out[str(col_idx)] = {
            "col": pat.col,
            "min_row": pat.min_row,
            "max_row": pat.max_row,
            "row_count": len(pat.rows),
            "abstract_sql": _ast_repr(pat.abstract_ast, emitter),
        }

    result = {
        "workbook": graph_data["workbook"],
        "output_sheet": output_sheet,
        "max_substitution_depth": max_depth,
        "flattened_cells": flattened_out,
        "column_patterns": patterns_out,
    }

    out_path = output_dir / "flattened_formulas.json"
    with open(out_path, "w") as f:
        json.dump(result, f, indent=2)

    print(f"\n  Written: {out_path}")

    # Print a human-readable sample
    print("\n  Sample flattened formulas (first 5):")
    for i, (key, cell) in enumerate(flattened_out.items()):
        if i >= 5:
            break
        print(f"    {key}: {cell['original_formula']!r}")
        print(f"      → {cell['flattened_sql']}")


if __name__ == "__main__":
    main()
