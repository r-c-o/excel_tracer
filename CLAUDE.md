# CLAUDE.md

## Project: excel_tracer

Decompiles Excel workbooks by tracing formulas from the output/summary sheet back to source data, flattening all intermediate transformations, and emitting an equivalent SQL query.

## Environment

Conda env: `excel_tracer` (Python 3.11). Activate before running:

```bat
conda activate excel_tracer
```

First-time setup: `setup_env.bat`

## Running the Pipeline

```bash
python run_all.py path/to/workbook.xlsx          # all three steps
python run_all.py workbook.xlsx --from flatten   # skip trace, start at flatten
python run_all.py workbook.xlsx --only sql       # single step only
```

## Architecture

Three sequential scripts, each consuming the previous step's output:

```
workbook.xlsx  →  trace_formulas.py   →  output/dependency_graph.json
                                      →  flatten_formulas.py  →  output/flattened_formulas.json
                                                              →  generate_sql.py  →  output/output.sql
```

### Step 1 — `trace_formulas.py`
Opens the workbook with openpyxl (formula mode), identifies the output sheet
(`config.yaml → output_sheets`), then BFS-traces all cell dependencies across
sheets into a directed graph. Writes `dependency_graph.json` with every cell's
formula, value, and dependency edges.

### Step 2 — `flatten_formulas.py`
Reconstructs the graph from JSON, then recursively substitutes every
intermediate formula cell into the output-sheet cells until only source data
cells remain. Runs column-pattern detection to find columns where all rows share
the same formula structure (vectorisable). Writes `flattened_formulas.json`.

### Step 3 — `generate_sql.py`
Re-flattens (for live AST objects), reads the workbook header row for column
labels, and passes everything to `SQLEmitter`. Emits a WITH-clause SQL query
where each source sheet becomes a CTE and each output column becomes a SELECT
expression. Writes `output/output.sql`.

## Library (`lib/`)

| File | Purpose |
|------|---------|
| `workbook.py` | openpyxl wrapper — loads both formula and data-only views |
| `formula_parser.py` | Recursive-descent Excel formula parser → AST |
| `dependency_graph.py` | BFS graph builder; expands ranges and named ranges |
| `flattener.py` | Substitution engine + column vectorisation |
| `sql_emitter.py` | AST → SQL expression tree walker |
| `function_map.py` | ~60 Excel functions mapped to Oracle/PostgreSQL/Snowflake SQL |

## Configuration (`config.yaml`)

| Key | Purpose |
|-----|---------|
| `output_sheets` | Ordered list of candidate output sheet names to search for |
| `sql_dialect` | `oracle` \| `postgresql` \| `snowflake` \| `ansi` |
| `sql_schema` | Optional schema prefix for SQL table names |
| `max_substitution_depth` | Recursion limit to guard against circular refs |
| `function_overrides` | Override specific Excel→SQL function translations |

## Known Limitations

- **VLOOKUP/INDEX-MATCH**: emitted as SQL comments with `-- rewrite as JOIN` hints;
  these require human review to determine the correct join key.
- **Pivot tables / Power Query**: not parsed; only standard cell formulas are traced.
- **Circular references**: detected by depth limit; flagged in output but not resolved.
- **External workbook references**: `[other.xlsx]Sheet1!A1` style refs are left as-is.
- **Array formulas (CSE)**: partially supported; complex dynamic array spill not handled.

---

## Project: RWA Model Convergence

A two-step pandas pipeline that merges Outlook balance sheets with aggregator
convergence data, applies adjustments, and produces upload-ready RWA workbooks.

### Files

| File | Purpose |
|------|---------|
| `rwa_constants.py` | Shared column-name literals, expected-column lists, PMF values |
| `step1_convergence.py` | Step 1 — waterfall join, writes 4 xlsx outputs |
| `step2_outlook_rwa.py` | Step 2 — adjustments, PMF mapping, pivots, control file |

### Running

```bash
# Step 1 — waterfall join balance sheets to convergence
python step1_convergence.py

# Step 2 — adjustments, PMF mapping, pivots, control file
python step2_outlook_rwa.py
```

Update `DATA_DIR` and `Q0` at the top of each script before running.

### Data flow

```
DATA_DIR/input/
  outlook_balancesheet_cg.xlsx   ─┐
  outlook_balancesheet_cbna.xlsx  ├─ step1 ─→ DATA_DIR/output/
  aggregator_for_convergence.xlsx ┘             cg_outlook_tap_env.xlsx
                                                cbna_com_outlook_tap_env.xlsx
                                                addon_all_cg.xlsx
                                                addon_all_cbna.xlsx

DATA_DIR/input/                               DATA_DIR/output/ (step1)
  adjustments.xlsx         ─────────────────────────────────────────┐
  pmf_frm_mapping.xlsx      ─┬─ step2 ─→ DATA_DIR/output/          │
  aggregator_for_convergence ┘              cg_outlook.xlsx         │
                                            cbna_outlook.xlsx       │
                                            cg_outlook_pivot.xlsx   │
                                            cbna_outlook_pivot.xlsx ┘
                                            cg_rwa_data.xlsx
                                            cbna_rwa_data.xlsx
                                            control_file.xlsx
```

### Column-name conventions in `rwa_constants.py`

- `*_CDE` — code columns (convergence/aggregator table only)
- `*_DESC` (`rc.*`) — full "Level N Description" form (balance sheet + convergence)
- Local `*_DESC` in `step2_outlook_rwa.py` — abbreviated "L4 Descr" form used in upload templates
