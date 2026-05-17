"""
lib/dependency_graph.py
-----------------------
Builds a directed acyclic graph (DAG) where each node is a (sheet, row, col)
cell address and edges point from a formula cell to every cell it references.

Also resolves range references to individual cells and expands named ranges.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any
import networkx as nx

from .formula_parser import (
    parse_formula, collect_refs,
    CellRef, RangeRef, NameRef, FuncCall,
)
from .workbook import Workbook


# Canonical cell address tuple
CellAddr = tuple[str, int, int]   # (sheet_name, row, col)


@dataclass
class CellNode:
    sheet: str
    row: int
    col: int
    formula: str | None = None       # raw formula string (or None for data cells)
    ast: Any = None                  # parsed AST
    value: Any = None                # cached value from workbook
    is_data: bool = False            # True if no formula (source input)

    @property
    def addr(self) -> CellAddr:
        return (self.sheet, self.row, self.col)


def _expand_range(ref: RangeRef, default_sheet: str, wb: Workbook) -> list[CellAddr]:
    """Expand a RangeRef into individual CellAddr tuples."""
    sheet = ref.sheet or default_sheet
    col1 = wb.col_index(ref.start.col_letter)
    col2 = wb.col_index(ref.end.col_letter)
    row1 = ref.start.row_number
    row2 = ref.end.row_number
    addrs = []
    for r in range(row1, row2 + 1):
        for c in range(col1, col2 + 1):
            addrs.append((sheet, r, c))
    return addrs


def build_graph(wb: Workbook, output_sheet: str) -> nx.DiGraph:
    """
    Starting from output_sheet, trace all formula dependencies recursively
    across sheets and return a DiGraph.

    Node attributes: CellNode stored under 'data' key.
    Edges: formula_cell → referenced_cell (dependency direction).
    """
    G = nx.DiGraph()
    named = wb.named_ranges()
    visited: set[CellAddr] = set()
    queue: list[CellAddr] = []

    def ensure_node(sheet: str, row: int, col: int) -> CellNode:
        addr = (sheet, row, col)
        if addr not in G:
            formula = wb.cell_formula(sheet, row, col)
            value = wb.cell_value(sheet, row, col)
            ast = parse_formula(formula) if formula else None
            node = CellNode(
                sheet=sheet, row=row, col=col,
                formula=formula, ast=ast, value=value,
                is_data=(formula is None),
            )
            G.add_node(addr, data=node)
        return G.nodes[addr]["data"]

    def enqueue(addr: CellAddr):
        if addr not in visited:
            visited.add(addr)
            queue.append(addr)

    # Seed from every formula cell in the output sheet
    for row, col, formula in wb.iter_formula_cells(output_sheet):
        addr = (output_sheet, row, col)
        ensure_node(output_sheet, row, col)
        enqueue(addr)

    # Also seed data cells in output sheet
    for row, col, val in wb.iter_all_cells(output_sheet):
        addr = (output_sheet, row, col)
        if addr not in G:
            ensure_node(output_sheet, row, col)

    # BFS
    while queue:
        addr = queue.pop(0)
        node = ensure_node(*addr)
        if not node.formula:
            continue

        refs = collect_refs(node.ast)
        for ref in refs:
            if isinstance(ref, CellRef):
                sheet = ref.sheet or addr[0]
                if sheet not in wb.sheet_names:
                    continue
                col_idx = wb.col_index(ref.col_letter)
                dep_addr = (sheet, ref.row_number, col_idx)
                ensure_node(sheet, ref.row_number, col_idx)
                G.add_edge(addr, dep_addr)
                enqueue(dep_addr)

            elif isinstance(ref, RangeRef):
                sheet = ref.sheet or addr[0]
                if sheet not in wb.sheet_names:
                    continue
                for dep_addr in _expand_range(ref, addr[0], wb):
                    ensure_node(*dep_addr)
                    G.add_edge(addr, dep_addr)
                    enqueue(dep_addr)

            elif isinstance(ref, NameRef):
                # Resolve named ranges
                defn = named.get(ref.name)
                if defn:
                    # Named ranges look like Sheet1!$A$1:$B$10 or Sheet1!$A$1
                    _resolve_named(ref.name, defn, addr, G, wb, ensure_node, enqueue)

    return G


def _resolve_named(name, defn, src_addr, G, wb, ensure_node, enqueue):
    """Best-effort resolution of a named range definition string."""
    import re
    # Strip leading = if present
    defn = defn.lstrip("=")
    # Try to match Sheet!Ref patterns
    parts = defn.split(",")
    for part in parts:
        part = part.strip()
        m = re.match(r"'?([^'!]+)'?!(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?", part, re.I)
        if not m:
            continue
        sheet = m.group(1)
        if sheet not in wb.sheet_names:
            continue
        ref_str = m.group(2) + (":" + m.group(3) if m.group(3) else "")
        from .formula_parser import FormulaParser
        try:
            node = FormulaParser("=" + ref_str).parse()
        except Exception:
            continue
        if isinstance(node, CellRef):
            dep = (sheet, node.row_number, wb.col_index(node.col_letter))
            ensure_node(*dep)
            G.add_edge(src_addr, dep)
            enqueue(dep)
        elif isinstance(node, RangeRef):
            from .formula_parser import RangeRef as RR
            for dep in _expand_range(node, sheet, wb):
                ensure_node(*dep)
                G.add_edge(src_addr, dep)
                enqueue(dep)


def sheet_summary(G: nx.DiGraph) -> dict[str, int]:
    """Return {sheet_name: cell_count} for all sheets in the graph."""
    counts: dict[str, int] = {}
    for addr in G.nodes:
        counts[addr[0]] = counts.get(addr[0], 0) + 1
    return counts


def source_cells(G: nx.DiGraph) -> list[CellAddr]:
    """Return cells with no outgoing edges (i.e., pure data/input cells)."""
    return [n for n in G.nodes if G.out_degree(n) == 0]


def output_cells(G: nx.DiGraph, output_sheet: str) -> list[CellAddr]:
    """Return cells on the output sheet with no incoming edges (final outputs)."""
    return [n for n in G.nodes if n[0] == output_sheet and G.in_degree(n) == 0]
