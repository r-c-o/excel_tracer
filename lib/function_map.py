"""
lib/function_map.py
-------------------
Maps Excel function names to SQL-rendering callables.

Each entry is:  FUNC_NAME → callable(emitter, args) → str

The emitter is passed so renderers can call emitter._ast_to_sql() on sub-args.
"""

from __future__ import annotations
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from .sql_emitter import SQLEmitter


def _args(emitter: "SQLEmitter", args: list[Any]) -> list[str]:
    return [emitter._ast_to_sql(a) for a in args]


# ---------------------------------------------------------------------------
# Aggregates
# ---------------------------------------------------------------------------

def _sum(e, args):
    return f"SUM({e._ast_to_sql(args[0])})" if args else "SUM(*)"

def _avg(e, args):
    return f"AVG({e._ast_to_sql(args[0])})" if args else "AVG(*)"

def _count(e, args):
    return f"COUNT({e._ast_to_sql(args[0])})" if args else "COUNT(*)"

def _counta(e, args):
    col = e._ast_to_sql(args[0]) if args else "*"
    return f"COUNT({col})"

def _min(e, args):    return f"MIN({e._ast_to_sql(args[0])})"
def _max(e, args):    return f"MAX({e._ast_to_sql(args[0])})"

def _sumif(e, args):
    # SUMIF(range, criteria, [sum_range])
    rng = e._ast_to_sql(args[0])
    criteria = e._ast_to_sql(args[1])
    sum_rng = e._ast_to_sql(args[2]) if len(args) > 2 else rng
    return f"SUM(CASE WHEN {rng} = {criteria} THEN {sum_rng} ELSE 0 END)"

def _sumifs(e, args):
    # SUMIFS(sum_range, range1, crit1, range2, crit2, ...)
    sum_rng = e._ast_to_sql(args[0])
    conditions = []
    for i in range(1, len(args) - 1, 2):
        conditions.append(f"{e._ast_to_sql(args[i])} = {e._ast_to_sql(args[i+1])}")
    where = " AND ".join(conditions)
    return f"SUM(CASE WHEN {where} THEN {sum_rng} ELSE 0 END)"

def _countif(e, args):
    rng = e._ast_to_sql(args[0])
    criteria = e._ast_to_sql(args[1])
    return f"COUNT(CASE WHEN {rng} = {criteria} THEN 1 END)"

def _averageif(e, args):
    rng = e._ast_to_sql(args[0])
    criteria = e._ast_to_sql(args[1])
    avg_rng = e._ast_to_sql(args[2]) if len(args) > 2 else rng
    return f"AVG(CASE WHEN {rng} = {criteria} THEN {avg_rng} END)"

# ---------------------------------------------------------------------------
# Conditional
# ---------------------------------------------------------------------------

def _if(e, args):
    cond = e._ast_to_sql(args[0])
    true_val = e._ast_to_sql(args[1]) if len(args) > 1 else "NULL"
    false_val = e._ast_to_sql(args[2]) if len(args) > 2 else "NULL"
    return f"CASE WHEN {cond} THEN {true_val} ELSE {false_val} END"

def _ifs(e, args):
    parts = []
    for i in range(0, len(args) - 1, 2):
        parts.append(f"WHEN {e._ast_to_sql(args[i])} THEN {e._ast_to_sql(args[i+1])}")
    return "CASE " + " ".join(parts) + " END"

def _iferror(e, args):
    val = e._ast_to_sql(args[0])
    fallback = e._ast_to_sql(args[1]) if len(args) > 1 else "NULL"
    return f"COALESCE({val}, {fallback})"

def _ifna(e, args):
    return _iferror(e, args)

def _switch(e, args):
    expr = e._ast_to_sql(args[0])
    parts = [f"CASE {expr}"]
    for i in range(1, len(args) - 1, 2):
        parts.append(f"WHEN {e._ast_to_sql(args[i])} THEN {e._ast_to_sql(args[i+1])}")
    if len(args) % 2 == 0:
        parts.append(f"ELSE {e._ast_to_sql(args[-1])}")
    parts.append("END")
    return " ".join(parts)

# ---------------------------------------------------------------------------
# Lookup
# ---------------------------------------------------------------------------

def _vlookup(e, args):
    # VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
    val = e._ast_to_sql(args[0])
    tbl = e._ast_to_sql(args[1])
    col_idx = e._ast_to_sql(args[2])
    return f"/* VLOOKUP({val}, {tbl}, {col_idx}) — rewrite as JOIN */"

def _hlookup(e, args):
    val = e._ast_to_sql(args[0])
    tbl = e._ast_to_sql(args[1])
    return f"/* HLOOKUP({val}, {tbl}) — rewrite as JOIN */"

def _index(e, args):
    arr = e._ast_to_sql(args[0])
    row = e._ast_to_sql(args[1]) if len(args) > 1 else "NULL"
    col = e._ast_to_sql(args[2]) if len(args) > 2 else "NULL"
    return f"/* INDEX({arr}, {row}, {col}) */"

def _match(e, args):
    val = e._ast_to_sql(args[0])
    arr = e._ast_to_sql(args[1])
    return f"/* MATCH({val}, {arr}) — rewrite as ROW_NUMBER/rank */"

def _xlookup(e, args):
    val = e._ast_to_sql(args[0])
    src = e._ast_to_sql(args[1])
    ret = e._ast_to_sql(args[2])
    return f"/* XLOOKUP({val} IN {src} → {ret}) — rewrite as JOIN */"

def _choose(e, args):
    idx = e._ast_to_sql(args[0])
    opts = ", ".join(e._ast_to_sql(a) for a in args[1:])
    return f"/* CHOOSE({idx}, {opts}) */"

# ---------------------------------------------------------------------------
# Text
# ---------------------------------------------------------------------------

def _concat(e, args):
    return " || ".join(e._ast_to_sql(a) for a in args)

def _concatenate(e, args):
    return _concat(e, args)

def _left(e, args):
    s = e._ast_to_sql(args[0])
    n = e._ast_to_sql(args[1]) if len(args) > 1 else "1"
    return f"SUBSTR({s}, 1, {n})"

def _right(e, args):
    s = e._ast_to_sql(args[0])
    n = e._ast_to_sql(args[1]) if len(args) > 1 else "1"
    return f"SUBSTR({s}, -({n}))" if e.dialect != "oracle" else f"SUBSTR({s}, LENGTH({s}) - {n} + 1, {n})"

def _mid(e, args):
    s = e._ast_to_sql(args[0])
    start = e._ast_to_sql(args[1])
    length = e._ast_to_sql(args[2]) if len(args) > 2 else "NULL"
    return f"SUBSTR({s}, {start}, {length})"

def _len(e, args):
    return f"LENGTH({e._ast_to_sql(args[0])})"

def _trim(e, args):
    return f"TRIM({e._ast_to_sql(args[0])})"

def _upper(e, args):   return f"UPPER({e._ast_to_sql(args[0])})"
def _lower(e, args):   return f"LOWER({e._ast_to_sql(args[0])})"
def _proper(e, args):  return f"INITCAP({e._ast_to_sql(args[0])})"

def _text(e, args):
    val = e._ast_to_sql(args[0])
    fmt = e._ast_to_sql(args[1]) if len(args) > 1 else "''"
    return f"TO_CHAR({val}, {fmt})" if e.dialect == "oracle" else f"FORMAT({val}, {fmt})"

def _find(e, args):
    find_text = e._ast_to_sql(args[0])
    within = e._ast_to_sql(args[1])
    return f"INSTR({within}, {find_text})" if e.dialect == "oracle" else f"CHARINDEX({find_text}, {within})"

def _substitute(e, args):
    s = e._ast_to_sql(args[0])
    old = e._ast_to_sql(args[1])
    new = e._ast_to_sql(args[2])
    return f"REPLACE({s}, {old}, {new})"

def _rept(e, args):
    s = e._ast_to_sql(args[0])
    n = e._ast_to_sql(args[1])
    return f"LPAD('', LENGTH({s}) * {n}, {s})" if e.dialect == "oracle" else f"REPLICATE({s}, {n})"

def _value(e, args):
    return f"TO_NUMBER({e._ast_to_sql(args[0])})" if e.dialect == "oracle" else f"CAST({e._ast_to_sql(args[0])} AS NUMERIC)"

# ---------------------------------------------------------------------------
# Math
# ---------------------------------------------------------------------------

def _abs(e, args):    return f"ABS({e._ast_to_sql(args[0])})"
def _round(e, args):
    n = e._ast_to_sql(args[0])
    d = e._ast_to_sql(args[1]) if len(args) > 1 else "0"
    return f"ROUND({n}, {d})"
def _roundup(e, args):
    n = e._ast_to_sql(args[0])
    d = e._ast_to_sql(args[1]) if len(args) > 1 else "0"
    return f"CEIL({n})" if d == "0" else f"/* ROUNDUP({n}, {d}) */"
def _rounddown(e, args):
    n = e._ast_to_sql(args[0])
    d = e._ast_to_sql(args[1]) if len(args) > 1 else "0"
    return f"FLOOR({n})" if d == "0" else f"TRUNC({n}, {d})" if e.dialect == "oracle" else f"TRUNCATE({n}, {d})"
def _int(e, args):    return f"TRUNC({e._ast_to_sql(args[0])})" if e.dialect == "oracle" else f"FLOOR({e._ast_to_sql(args[0])})"
def _mod(e, args):    return f"MOD({e._ast_to_sql(args[0])}, {e._ast_to_sql(args[1])})"
def _sqrt(e, args):   return f"SQRT({e._ast_to_sql(args[0])})"
def _power(e, args):  return f"POWER({e._ast_to_sql(args[0])}, {e._ast_to_sql(args[1])})"
def _ln(e, args):     return f"LN({e._ast_to_sql(args[0])})"
def _log(e, args):
    n = e._ast_to_sql(args[0])
    base = e._ast_to_sql(args[1]) if len(args) > 1 else "10"
    return f"LOG({base}, {n})" if e.dialect == "oracle" else f"LOG({n}) / LOG({base})"
def _exp(e, args):    return f"EXP({e._ast_to_sql(args[0])})"

# ---------------------------------------------------------------------------
# Date / Time
# ---------------------------------------------------------------------------

def _today(e, args):
    return "TRUNC(SYSDATE)" if e.dialect == "oracle" else "CURRENT_DATE"

def _now(e, args):
    return "SYSDATE" if e.dialect == "oracle" else "NOW()"

def _year(e, args):
    d = e._ast_to_sql(args[0])
    return f"EXTRACT(YEAR FROM {d})"

def _month(e, args):
    d = e._ast_to_sql(args[0])
    return f"EXTRACT(MONTH FROM {d})"

def _day(e, args):
    d = e._ast_to_sql(args[0])
    return f"EXTRACT(DAY FROM {d})"

def _date(e, args):
    y, m, d = (e._ast_to_sql(a) for a in args[:3])
    if e.dialect == "oracle":
        return f"TO_DATE({y} || '-' || LPAD({m},2,'0') || '-' || LPAD({d},2,'0'), 'YYYY-MM-DD')"
    return f"MAKE_DATE({y}, {m}, {d})"

def _datevalue(e, args):
    s = e._ast_to_sql(args[0])
    return f"TO_DATE({s})" if e.dialect == "oracle" else f"CAST({s} AS DATE)"

def _edate(e, args):
    d = e._ast_to_sql(args[0])
    n = e._ast_to_sql(args[1])
    return f"ADD_MONTHS({d}, {n})" if e.dialect == "oracle" else f"({d} + INTERVAL '{n} months')"

def _datedif(e, args):
    start = e._ast_to_sql(args[0])
    end   = e._ast_to_sql(args[1])
    unit  = e._ast_to_sql(args[2])
    return f"/* DATEDIF({start}, {end}, {unit}) — use MONTHS_BETWEEN or date arithmetic */"

def _eomonth(e, args):
    d = e._ast_to_sql(args[0])
    n = e._ast_to_sql(args[1]) if len(args) > 1 else "0"
    if e.dialect == "oracle":
        return f"LAST_DAY(ADD_MONTHS({d}, {n}))"
    return f"(DATE_TRUNC('month', {d}) + INTERVAL '1 month' - INTERVAL '1 day')"

# ---------------------------------------------------------------------------
# Logical / utility
# ---------------------------------------------------------------------------

def _and(e, args):
    return "(" + " AND ".join(e._ast_to_sql(a) for a in args) + ")"

def _or(e, args):
    return "(" + " OR ".join(e._ast_to_sql(a) for a in args) + ")"

def _not(e, args):
    return f"NOT ({e._ast_to_sql(args[0])})"

def _isblank(e, args):
    return f"({e._ast_to_sql(args[0])} IS NULL)"

def _isnumber(e, args):
    val = e._ast_to_sql(args[0])
    if e.dialect == "oracle":
        return f"(REGEXP_LIKE({val}, '^[+-]?\\d+(\\.\\d+)?$'))"
    return f"/* ISNUMBER({val}) */"

def _istext(e, args):
    val = e._ast_to_sql(args[0])
    return f"/* ISTEXT({val}) — check column type */"

def _isna(e, args):
    return f"({e._ast_to_sql(args[0])} IS NULL)"

def _iserror(e, args):
    return f"({e._ast_to_sql(args[0])} IS NULL)"

def _rows(e, args):    return "/* ROWS() — scalar row count */"
def _columns(e, args): return "/* COLUMNS() — scalar column count */"

def _rank(e, args):
    val = e._ast_to_sql(args[0])
    rng = e._ast_to_sql(args[1])
    return f"RANK() OVER (ORDER BY {val})"

def _large(e, args):
    rng = e._ast_to_sql(args[0])
    k   = e._ast_to_sql(args[1])
    return f"/* LARGE({rng}, {k}) — use DENSE_RANK() window */"

def _small(e, args):
    rng = e._ast_to_sql(args[0])
    k   = e._ast_to_sql(args[1])
    return f"/* SMALL({rng}, {k}) — use DENSE_RANK() window */"

# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------

FUNCTION_MAP: dict[str, Any] = {
    # Aggregates
    "SUM": _sum, "AVERAGE": _avg, "AVG": _avg,
    "COUNT": _count, "COUNTA": _counta, "COUNTBLANK": lambda e, a: f"COUNT(*) - COUNT({e._ast_to_sql(a[0])})",
    "MIN": _min, "MAX": _max,
    "SUMIF": _sumif, "SUMIFS": _sumifs,
    "COUNTIF": _countif, "COUNTIFS": _countif,
    "AVERAGEIF": _averageif,
    # Conditional
    "IF": _if, "IFS": _ifs, "IFERROR": _iferror, "IFNA": _ifna, "SWITCH": _switch,
    # Lookup
    "VLOOKUP": _vlookup, "HLOOKUP": _hlookup,
    "INDEX": _index, "MATCH": _match, "XLOOKUP": _xlookup, "CHOOSE": _choose,
    # Text
    "CONCAT": _concat, "CONCATENATE": _concatenate,
    "LEFT": _left, "RIGHT": _right, "MID": _mid, "LEN": _len,
    "TRIM": _trim, "UPPER": _upper, "LOWER": _lower, "PROPER": _proper,
    "TEXT": _text, "FIND": _find, "SEARCH": _find,
    "SUBSTITUTE": _substitute, "REPLACE": _substitute,
    "REPT": _rept, "VALUE": _value,
    # Math
    "ABS": _abs, "ROUND": _round, "ROUNDUP": _roundup, "ROUNDDOWN": _rounddown,
    "INT": _int, "MOD": _mod, "SQRT": _sqrt, "POWER": _power,
    "LN": _ln, "LOG": _log, "LOG10": lambda e, a: f"LOG(10, {e._ast_to_sql(a[0])})",
    "EXP": _exp,
    # Date
    "TODAY": _today, "NOW": _now,
    "YEAR": _year, "MONTH": _month, "DAY": _day,
    "DATE": _date, "DATEVALUE": _datevalue,
    "EDATE": _edate, "DATEDIF": _datedif, "EOMONTH": _eomonth,
    # Logical
    "AND": _and, "OR": _or, "NOT": _not,
    "ISBLANK": _isblank, "ISNUMBER": _isnumber, "ISTEXT": _istext,
    "ISNA": _isna, "ISERROR": _iserror, "ISERR": _iserror,
    # Misc
    "ROWS": _rows, "COLUMNS": _columns,
    "RANK": _rank, "RANK.EQ": _rank, "RANK.AVG": _rank,
    "LARGE": _large, "SMALL": _small,
}
