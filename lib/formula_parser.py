"""
lib/formula_parser.py
---------------------
Recursive-descent parser that turns an Excel formula string into an AST.

Supported constructs:
  - Arithmetic: + - * / ^ (unary minus)
  - Comparison: = <> < > <= >=
  - Concatenation: &
  - Cell references: A1, $A$1, Sheet2!B3, 'My Sheet'!B3
  - Range references: A1:B10, Sheet2!A1:B10
  - Named ranges / identifiers
  - Function calls: SUM(...), IF(...), VLOOKUP(...)
  - Literals: numbers, strings ("..."), booleans TRUE/FALSE, errors #N/A etc.
  - Parenthesised expressions
  - Array literals: {1,2,3}

AST node types (all dataclasses):
  Literal, CellRef, RangeRef, NameRef, FuncCall, BinOp, UnaryOp, ArrayLiteral
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any
import re


# ---------------------------------------------------------------------------
# AST node types
# ---------------------------------------------------------------------------

@dataclass
class Literal:
    value: Any          # int | float | str | bool | None
    kind: str           # "number" | "string" | "bool" | "error" | "null"

@dataclass
class CellRef:
    sheet: str | None   # None → same sheet
    col: str            # e.g. "A", "$A"
    row: str            # e.g. "1", "$1"

    @property
    def abs_col(self) -> bool:
        return self.col.startswith("$")

    @property
    def abs_row(self) -> bool:
        return self.row.startswith("$")

    @property
    def col_letter(self) -> str:
        return self.col.lstrip("$")

    @property
    def row_number(self) -> int:
        return int(self.row.lstrip("$"))

    def __str__(self) -> str:
        sheet_prefix = f"{self.sheet}!" if self.sheet else ""
        return f"{sheet_prefix}{self.col}{self.row}"

@dataclass
class RangeRef:
    sheet: str | None
    start: CellRef
    end: CellRef

    def __str__(self) -> str:
        sheet_prefix = f"{self.sheet}!" if self.sheet else ""
        return f"{sheet_prefix}{self.start.col}{self.start.row}:{self.end.col}{self.end.row}"

@dataclass
class NameRef:
    """A named range or defined name."""
    name: str
    sheet: str | None = None

@dataclass
class FuncCall:
    name: str
    args: list[Any] = field(default_factory=list)

@dataclass
class BinOp:
    op: str     # "+", "-", "*", "/", "^", "&", "=", "<>", "<", ">", "<=", ">="
    left: Any
    right: Any

@dataclass
class UnaryOp:
    op: str     # "-", "+"
    operand: Any

@dataclass
class ArrayLiteral:
    """Inline array {1,2;3,4} — rows separated by ";", cols by ","."""
    rows: list[list[Any]]


# ---------------------------------------------------------------------------
# Tokeniser
# ---------------------------------------------------------------------------

_TOKEN_PATTERNS = [
    ("SPACE",    r"[ \t]+"),
    ("STRING",   r'"(?:[^"]|"")*"'),
    ("ERROR",    r"#(?:N/A|REF!|NAME\?|VALUE!|DIV/0!|NULL!|NUM!)"),
    ("BOOL",     r"\b(?:TRUE|FALSE)\b"),
    ("NUMBER",   r"\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?"),
    ("REF",      r"(?:'[^']+'|[A-Za-z_]\w*)!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?|"
                 r"\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?"),
    ("IDENT",    r"[A-Za-z_][\w.]*"),
    ("OP",       r"<>|<=|>=|[+\-*/^&=<>]"),
    ("LPAREN",   r"\("),
    ("RPAREN",   r"\)"),
    ("LBRACE",   r"\{"),
    ("RBRACE",   r"\}"),
    ("COMMA",    r","),
    ("SEMICOLON",r";"),
    ("COLON",    r":"),
    ("PERCENT",  r"%"),
    ("BANG",     r"!"),
]

_TOKEN_RE = re.compile(
    "|".join(f"(?P<{name}>{pat})" for name, pat in _TOKEN_PATTERNS)
)

_CELL_RE = re.compile(
    r"^(?P<col>\$?[A-Z]+)(?P<row>\$?\d+)$", re.IGNORECASE
)

_RANGE_RE = re.compile(
    r"^(?P<col1>\$?[A-Z]+)(?P<row1>\$?\d+):(?P<col2>\$?[A-Z]+)(?P<row2>\$?\d+)$",
    re.IGNORECASE,
)

_SHEET_REF_RE = re.compile(
    r"^(?P<sheet>'[^']+'|[A-Za-z_]\w*)!(?P<ref>.+)$"
)


def tokenise(formula: str) -> list[tuple[str, str]]:
    """Return list of (type, value) tokens, skipping whitespace."""
    tokens = []
    for m in _TOKEN_RE.finditer(formula):
        kind = m.lastgroup
        if kind == "SPACE":
            continue
        tokens.append((kind, m.group()))
    return tokens


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------

class FormulaParser:
    """Recursive-descent parser. Instantiate per formula string."""

    def __init__(self, formula: str):
        raw = formula.strip()
        if raw.startswith("="):
            raw = raw[1:]
        self._tokens = tokenise(raw)
        self._pos = 0

    # -- token helpers -------------------------------------------------------

    def _peek(self) -> tuple[str, str] | None:
        if self._pos < len(self._tokens):
            return self._tokens[self._pos]
        return None

    def _consume(self) -> tuple[str, str]:
        tok = self._tokens[self._pos]
        self._pos += 1
        return tok

    def _expect(self, kind: str, value: str | None = None) -> tuple[str, str]:
        tok = self._consume()
        if tok[0] != kind:
            raise SyntaxError(f"Expected {kind}, got {tok}")
        if value is not None and tok[1] != value:
            raise SyntaxError(f"Expected '{value}', got '{tok[1]}'")
        return tok

    def _match(self, kind: str, value: str | None = None) -> bool:
        tok = self._peek()
        if tok is None:
            return False
        if tok[0] != kind:
            return False
        if value is not None and tok[1] != value:
            return False
        return True

    # -- public entry --------------------------------------------------------

    def parse(self) -> Any:
        node = self._expr()
        if self._peek() is not None:
            raise SyntaxError(f"Unexpected token: {self._peek()}")
        return node

    # -- grammar (precedence low → high) ------------------------------------

    def _expr(self) -> Any:
        return self._concat()

    def _concat(self) -> Any:
        left = self._comparison()
        while self._match("OP", "&"):
            self._consume()
            left = BinOp("&", left, self._comparison())
        return left

    def _comparison(self) -> Any:
        left = self._additive()
        while self._peek() and self._peek()[0] == "OP" and self._peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self._consume()[1]
            left = BinOp(op, left, self._additive())
        return left

    def _additive(self) -> Any:
        left = self._multiplicative()
        while self._peek() and self._peek()[0] == "OP" and self._peek()[1] in ("+", "-"):
            op = self._consume()[1]
            left = BinOp(op, left, self._multiplicative())
        return left

    def _multiplicative(self) -> Any:
        left = self._exponent()
        while self._peek() and self._peek()[0] == "OP" and self._peek()[1] in ("*", "/"):
            op = self._consume()[1]
            left = BinOp(op, left, self._exponent())
        return left

    def _exponent(self) -> Any:
        left = self._unary()
        if self._match("OP", "^"):
            self._consume()
            left = BinOp("^", left, self._unary())
        return left

    def _unary(self) -> Any:
        if self._match("OP", "-"):
            self._consume()
            return UnaryOp("-", self._percent())
        if self._match("OP", "+"):
            self._consume()
            return self._percent()
        return self._percent()

    def _percent(self) -> Any:
        node = self._primary()
        if self._match("PERCENT"):
            self._consume()
            node = BinOp("/", node, Literal(100, "number"))
        return node

    def _primary(self) -> Any:
        tok = self._peek()
        if tok is None:
            raise SyntaxError("Unexpected end of formula")

        kind, val = tok

        # Parenthesised expression
        if kind == "LPAREN":
            self._consume()
            node = self._expr()
            self._expect("RPAREN")
            return node

        # Array literal {1,2,3}
        if kind == "LBRACE":
            return self._array_literal()

        # String literal
        if kind == "STRING":
            self._consume()
            inner = val[1:-1].replace('""', '"')
            return Literal(inner, "string")

        # Error literal
        if kind == "ERROR":
            self._consume()
            return Literal(val, "error")

        # Boolean
        if kind == "BOOL":
            self._consume()
            return Literal(val == "TRUE", "bool")

        # Number
        if kind == "NUMBER":
            self._consume()
            num = float(val) if "." in val or "e" in val.lower() else int(val)
            return Literal(num, "number")

        # REF token: may be sheet!cell, sheet!range, cell, or range
        if kind == "REF":
            self._consume()
            return self._parse_ref_token(val)

        # IDENT: function call or named range
        if kind == "IDENT":
            self._consume()
            # Look-ahead: if next token is LPAREN → function call
            if self._match("LPAREN"):
                return self._func_call(val)
            # If next is BANG → sheet prefix for a ref we didn't catch in REF
            if self._match("BANG"):
                self._consume()
                ref_tok = self._peek()
                if ref_tok and ref_tok[0] == "REF":
                    self._consume()
                    return self._parse_ref_token(ref_tok[1], sheet=val)
            return NameRef(val)

        raise SyntaxError(f"Unexpected token {tok}")

    # -- helpers -------------------------------------------------------------

    def _parse_ref_token(self, val: str, sheet: str | None = None) -> CellRef | RangeRef | NameRef:
        """Parse a REF-token string into CellRef, RangeRef, or NameRef."""
        # Strip sheet prefix if present in token
        sm = _SHEET_REF_RE.match(val)
        if sm:
            sheet = sm.group("sheet").strip("'")
            val = sm.group("ref")

        rm = _RANGE_RE.match(val)
        if rm:
            start = CellRef(None, rm.group("col1").upper(), rm.group("row1"))
            end   = CellRef(None, rm.group("col2").upper(), rm.group("row2"))
            return RangeRef(sheet, start, end)

        cm = _CELL_RE.match(val)
        if cm:
            return CellRef(sheet, cm.group("col").upper(), cm.group("row"))

        return NameRef(val, sheet)

    def _func_call(self, name: str) -> FuncCall:
        self._expect("LPAREN")
        args = []
        if not self._match("RPAREN"):
            args.append(self._expr())
            while self._match("COMMA"):
                self._consume()
                # Allow trailing comma / empty arg
                if self._match("RPAREN"):
                    args.append(Literal(None, "null"))
                    break
                args.append(self._expr())
        self._expect("RPAREN")
        return FuncCall(name.upper(), args)

    def _array_literal(self) -> ArrayLiteral:
        self._expect("LBRACE")
        rows = []
        current_row = []
        while not self._match("RBRACE"):
            current_row.append(self._expr())
            if self._match("COMMA"):
                self._consume()
            elif self._match("SEMICOLON"):
                self._consume()
                rows.append(current_row)
                current_row = []
        rows.append(current_row)
        self._expect("RBRACE")
        return ArrayLiteral(rows)


def parse_formula(formula: str) -> Any:
    """Parse an Excel formula string into an AST node. Returns None for non-formulas."""
    formula = formula.strip() if formula else ""
    if not formula or not formula.startswith("="):
        return None
    try:
        return FormulaParser(formula).parse()
    except Exception:
        return None


def collect_refs(node: Any) -> list[CellRef | RangeRef | NameRef]:
    """Walk an AST and return all reference nodes."""
    if node is None:
        return []
    refs = []
    if isinstance(node, (CellRef, RangeRef, NameRef)):
        refs.append(node)
    elif isinstance(node, FuncCall):
        for arg in node.args:
            refs.extend(collect_refs(arg))
    elif isinstance(node, BinOp):
        refs.extend(collect_refs(node.left))
        refs.extend(collect_refs(node.right))
    elif isinstance(node, UnaryOp):
        refs.extend(collect_refs(node.operand))
    elif isinstance(node, ArrayLiteral):
        for row in node.rows:
            for cell in row:
                refs.extend(collect_refs(cell))
    return refs
