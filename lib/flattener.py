"""
lib/flattener.py
----------------
Substitutes intermediate formula cells into the output cells so that
every formula on the summary sheet references only source (data) cells.

Strategy:
  For each output cell whose AST contains a CellRef or RangeRef that
  itself has a formula, replace that reference with the referenced cell's
  (already-flattened) AST. Recurse until all references reach data cells
  or a configured depth limit is hit.

After substitution, `vectorize()` converts column-of-identical-structure
formulas into a single column-vector description (replacing A1 → col[A],
A2 → col[A], etc.) so that the SQL emitter can produce a single SELECT
column rather than one per row.
"""

from __future__ import annotations
import copy
from typing import Any
import networkx as nx

from .formula_parser import (
    CellRef, RangeRef, NameRef, FuncCall, BinOp, UnaryOp, Literal, ArrayLiteral,
    collect_refs,
)
from .dependency_graph import CellAddr, CellNode


def flatten_cell(
    addr: CellAddr,
    G: nx.DiGraph,
    max_depth: int = 20,
    _depth: int = 0,
) -> Any:
    """
    Return a fully-flattened AST for the cell at `addr`.
    Intermediate references are recursively substituted.
    """
    node: CellNode = G.nodes[addr]["data"]
    if node.is_data or node.ast is None:
        # Source cell — represent as a data reference (keep as CellRef)
        return CellRef(addr[0], _col_letter(addr[2]), str(addr[1]))

    if _depth >= max_depth:
        # Stop recursing; leave reference as-is
        return CellRef(addr[0], _col_letter(addr[2]), str(addr[1]))

    return _substitute(node.ast, addr[0], G, max_depth, _depth + 1)


def _substitute(ast: Any, current_sheet: str, G: nx.DiGraph, max_depth: int, depth: int) -> Any:
    """Walk ast, replacing CellRefs and RangeRefs with flattened sub-ASTs."""
    if isinstance(ast, CellRef):
        sheet = ast.sheet or current_sheet
        from .workbook import Workbook  # avoid circular; col_index via util
        col_idx = _col_index(ast.col_letter)
        dep_addr = (sheet, ast.row_number, col_idx)
        if dep_addr in G.nodes:
            return flatten_cell(dep_addr, G, max_depth, depth)
        return ast  # external or out-of-graph ref, leave as-is

    if isinstance(ast, RangeRef):
        # Ranges stay as ranges (aggregated by functions); don't expand here
        return ast

    if isinstance(ast, FuncCall):
        return FuncCall(ast.name, [_substitute(a, current_sheet, G, max_depth, depth) for a in ast.args])

    if isinstance(ast, BinOp):
        return BinOp(
            ast.op,
            _substitute(ast.left,  current_sheet, G, max_depth, depth),
            _substitute(ast.right, current_sheet, G, max_depth, depth),
        )

    if isinstance(ast, UnaryOp):
        return UnaryOp(ast.op, _substitute(ast.operand, current_sheet, G, max_depth, depth))

    if isinstance(ast, ArrayLiteral):
        return ArrayLiteral([
            [_substitute(c, current_sheet, G, max_depth, depth) for c in row]
            for row in ast.rows
        ])

    return ast  # Literal, NameRef — unchanged


def flatten_all(G: nx.DiGraph, output_sheet: str, max_depth: int = 20) -> dict[CellAddr, Any]:
    """
    Flatten every formula cell on the output sheet.
    Returns {addr: flattened_ast}.
    """
    result = {}
    for addr, node_data in G.nodes(data=True):
        node: CellNode = node_data["data"]
        if addr[0] == output_sheet and node.formula:
            result[addr] = flatten_cell(addr, G, max_depth)
    return result


# ---------------------------------------------------------------------------
# Vectorisation
# ---------------------------------------------------------------------------

def detect_column_patterns(
    flattened: dict[CellAddr, Any],
    output_sheet: str,
) -> dict[int, "ColumnPattern"]:
    """
    Detect columns where all rows share the same formula structure with only
    row-number differences. Returns {col_index: ColumnPattern}.

    A ColumnPattern captures:
      - The abstract formula (with row numbers replaced by a placeholder)
      - Which source columns are referenced
      - The row range
    """
    from collections import defaultdict
    col_rows: dict[int, list[tuple[int, Any]]] = defaultdict(list)

    for addr, ast in flattened.items():
        sheet, row, col = addr
        if sheet == output_sheet:
            col_rows[col].append((row, ast))

    patterns = {}
    for col, row_asts in col_rows.items():
        if len(row_asts) < 2:
            continue
        row_asts.sort(key=lambda x: x[0])
        rows = [r for r, _ in row_asts]
        asts = [a for _, a in row_asts]

        # Normalise: replace row numbers with a sentinel, check structural equality
        normalised = [_normalise_rows(a, rows) for a in asts]
        if all(n == normalised[0] for n in normalised[1:]):
            patterns[col] = ColumnPattern(
                col=col,
                rows=rows,
                abstract_ast=normalised[0],
                concrete_asts=asts,
            )

    return patterns


class ColumnPattern:
    """Represents a vectorised column: same formula structure repeated per row."""

    def __init__(self, col: int, rows: list[int], abstract_ast: Any, concrete_asts: list[Any]):
        self.col = col
        self.rows = rows
        self.abstract_ast = abstract_ast   # row numbers replaced with ROW_PLACEHOLDER
        self.concrete_asts = concrete_asts

    @property
    def min_row(self) -> int:
        return min(self.rows)

    @property
    def max_row(self) -> int:
        return max(self.rows)


ROW_PLACEHOLDER = "__ROW__"


def _normalise_rows(ast: Any, actual_rows: list[int]) -> Any:
    """Replace any integer row reference that matches a data row with ROW_PLACEHOLDER."""
    if isinstance(ast, CellRef):
        row_num = ast.row_number
        if row_num in actual_rows:
            return CellRef(ast.sheet, ast.col, ROW_PLACEHOLDER)
        return ast
    if isinstance(ast, FuncCall):
        return FuncCall(ast.name, [_normalise_rows(a, actual_rows) for a in ast.args])
    if isinstance(ast, BinOp):
        return BinOp(ast.op, _normalise_rows(ast.left, actual_rows), _normalise_rows(ast.right, actual_rows))
    if isinstance(ast, UnaryOp):
        return UnaryOp(ast.op, _normalise_rows(ast.operand, actual_rows))
    return ast


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def _col_letter(col_index: int) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(col_index)


def _col_index(letter: str) -> int:
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter.lstrip("$"))
