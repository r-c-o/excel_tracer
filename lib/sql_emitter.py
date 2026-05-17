"""
lib/sql_emitter.py
------------------
Converts flattened formula ASTs into SQL expressions and assembles a
complete SELECT query (with CTEs for each source sheet).

Approach:
  1. Each source sheet becomes a CTE:  sheet_name AS (SELECT ... FROM dual/values)
     - For live Oracle connections, we'd reference the actual table.
     - In standalone mode we emit placeholder table names.
  2. The output sheet's formulas become the SELECT column list.
  3. Column vectors (detected by the flattener) become a single output column.
  4. Scalar (single-cell) formulas appear as a SELECT without FROM.

Excel → SQL function map is in `function_map.py`.
"""

from __future__ import annotations
from typing import Any

from .formula_parser import (
    CellRef, RangeRef, NameRef, FuncCall, BinOp, UnaryOp, Literal, ArrayLiteral,
)
from .flattener import ColumnPattern, ROW_PLACEHOLDER, _col_letter
from .dependency_graph import CellAddr, CellNode
from . import function_map as FM
import networkx as nx


class SQLEmitter:
    def __init__(self, dialect: str = "oracle", schema: str = "", overrides: dict | None = None):
        self.dialect = dialect.lower()
        self.schema = schema
        self.overrides = overrides or {}

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def emit_query(
        self,
        G: nx.DiGraph,
        output_sheet: str,
        flattened: dict[CellAddr, Any],
        column_patterns: dict[int, ColumnPattern],
        header_row: dict[str, int],   # {label: col_index} for output sheet
    ) -> str:
        """Generate the complete SQL query string."""

        # Build reverse header map: col_index → label
        col_to_label = {v: k for k, v in header_row.items()}

        # Identify source sheets (non-output sheets that appear in the graph)
        source_sheets = sorted(
            {addr[0] for addr in G.nodes if addr[0] != output_sheet}
        )

        lines = []

        # CTEs for each source sheet
        if source_sheets:
            lines.append("WITH")
            cte_parts = []
            for sheet in source_sheets:
                cte_parts.append(self._emit_source_cte(sheet, G))
            lines.append(",\n".join(cte_parts))

        # SELECT columns
        select_cols = self._emit_select_columns(
            flattened, column_patterns, col_to_label, output_sheet
        )
        lines.append("SELECT")
        lines.append(",\n".join(f"    {c}" for c in select_cols))

        # FROM clause
        if source_sheets:
            if len(source_sheets) == 1:
                lines.append(f"FROM {self._sheet_table(source_sheets[0])}")
            else:
                from_clause = self._emit_from_clause(source_sheets, G, output_sheet, flattened)
                lines.append(from_clause)

        return "\n".join(lines) + "\n"

    # ------------------------------------------------------------------
    # CTEs
    # ------------------------------------------------------------------

    def _emit_source_cte(self, sheet: str, G: nx.DiGraph) -> str:
        table = self._sheet_table(sheet)
        cte_name = self._safe_name(sheet)
        return f"    {cte_name} AS (\n        SELECT * FROM {table}\n    )"

    # ------------------------------------------------------------------
    # SELECT columns
    # ------------------------------------------------------------------

    def _emit_select_columns(
        self,
        flattened: dict[CellAddr, Any],
        column_patterns: dict[int, ColumnPattern],
        col_to_label: dict[int, str],
        output_sheet: str,
    ) -> list[str]:
        cols = []
        handled_cols: set[int] = set()

        # Emit vectorised columns first
        for col_idx, pattern in sorted(column_patterns.items()):
            label = col_to_label.get(col_idx, _col_letter(col_idx))
            expr = self._ast_to_sql(pattern.abstract_ast)
            # Replace ROW_PLACEHOLDER artefacts with just the column ref
            expr = expr.replace(f'"{ROW_PLACEHOLDER}"', "<<ROW>>").replace(ROW_PLACEHOLDER, "<<ROW>>")
            cols.append(f"{expr} AS {self._safe_name(label)}")
            handled_cols.add(col_idx)

        # Emit remaining scalar cells (sorted by col then row)
        scalar_cells = sorted(
            [(addr, ast) for addr, ast in flattened.items()
             if addr[0] == output_sheet and addr[2] not in handled_cols],
            key=lambda x: (x[0][2], x[0][1]),
        )
        for addr, ast in scalar_cells:
            col_idx = addr[2]
            label = col_to_label.get(col_idx, f"{_col_letter(col_idx)}{addr[1]}")
            expr = self._ast_to_sql(ast)
            cols.append(f"{expr} AS {self._safe_name(label)}")

        return cols if cols else ["*"]

    # ------------------------------------------------------------------
    # FROM clause
    # ------------------------------------------------------------------

    def _emit_from_clause(self, source_sheets, G, output_sheet, flattened) -> str:
        # Simple cross-join for now; a join-inference engine is a future enhancement
        tables = " CROSS JOIN ".join(self._safe_name(s) for s in source_sheets)
        return f"FROM {tables}"

    # ------------------------------------------------------------------
    # AST → SQL expression
    # ------------------------------------------------------------------

    def _ast_to_sql(self, node: Any) -> str:
        if node is None:
            return "NULL"

        if isinstance(node, Literal):
            return self._literal_sql(node)

        if isinstance(node, CellRef):
            return self._cellref_sql(node)

        if isinstance(node, RangeRef):
            return self._rangeref_sql(node)

        if isinstance(node, NameRef):
            return self._safe_name(node.name)

        if isinstance(node, FuncCall):
            return self._funccall_sql(node)

        if isinstance(node, BinOp):
            return self._binop_sql(node)

        if isinstance(node, UnaryOp):
            return self._unaryop_sql(node)

        if isinstance(node, ArrayLiteral):
            # Flatten to first row for scalar context
            flat = [self._ast_to_sql(c) for row in node.rows for c in row]
            return f"({', '.join(flat)})"

        return str(node)

    def _literal_sql(self, node: Literal) -> str:
        if node.kind == "null":
            return "NULL"
        if node.kind == "string":
            escaped = str(node.value).replace("'", "''")
            return f"'{escaped}'"
        if node.kind == "bool":
            if self.dialect == "oracle":
                return "1" if node.value else "0"
            return "TRUE" if node.value else "FALSE"
        if node.kind == "error":
            return "NULL"
        # number
        return str(node.value)

    def _cellref_sql(self, node: CellRef) -> str:
        if node.row == ROW_PLACEHOLDER:
            sheet_prefix = f"{self._safe_name(node.sheet)}." if node.sheet else ""
            return f"{sheet_prefix}{self._safe_name(node.col_letter)}"
        sheet_prefix = f"{self._safe_name(node.sheet)}." if node.sheet else ""
        col = self._safe_name(node.col_letter)
        return f"{sheet_prefix}{col}"

    def _rangeref_sql(self, node: RangeRef) -> str:
        sheet_prefix = f"{self._safe_name(node.sheet)}." if node.sheet else ""
        # Ranges are only meaningful inside aggregate functions; emit as subselect marker
        col_start = node.start.col_letter
        col_end = node.end.col_letter
        if col_start == col_end:
            return f"{sheet_prefix}{self._safe_name(col_start)}"
        return f"/* range {node} */"

    def _funccall_sql(self, node: FuncCall) -> str:
        name = node.name.upper()
        # Check override map first
        if name in self.overrides:
            name = self.overrides[name]
        elif name in FM.FUNCTION_MAP:
            return FM.FUNCTION_MAP[name](self, node.args)

        args_sql = ", ".join(self._ast_to_sql(a) for a in node.args)
        return f"{name}({args_sql})"

    def _binop_sql(self, node: BinOp) -> str:
        op_map = {
            "=": "=", "<>": "<>", "&": "||",
            "+": "+", "-": "-", "*": "*", "/": "/",
            "^": self._power_op(),
            "<": "<", ">": ">", "<=": "<=", ">=": ">=",
        }
        sql_op = op_map.get(node.op, node.op)

        left = self._ast_to_sql(node.left)
        right = self._ast_to_sql(node.right)

        if node.op == "^":
            if self.dialect == "oracle":
                return f"POWER({left}, {right})"
            return f"POWER({left}, {right})"

        return f"({left} {sql_op} {right})"

    def _unaryop_sql(self, node: UnaryOp) -> str:
        return f"(-{self._ast_to_sql(node.operand)})" if node.op == "-" else self._ast_to_sql(node.operand)

    def _power_op(self) -> str:
        return "POWER"

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _sheet_table(self, sheet: str) -> str:
        name = self._safe_name(sheet)
        if self.schema:
            return f"{self.schema}.{name}"
        return name

    def _safe_name(self, name: str) -> str:
        """Quote identifiers that need it."""
        if name.upper() == name and name.replace("_", "").isalnum():
            return name
        return f'"{name}"'
