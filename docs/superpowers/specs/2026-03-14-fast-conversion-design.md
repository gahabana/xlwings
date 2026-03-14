# Fast Conversion Pipeline — Design Spec

## Problem

xlwings processes Excel data element-by-element in Python during read/write operations. For large ranges (10K+ cells), the conversion pipeline is the dominant cost — not the COM/appscript calls themselves. Three hot paths account for most of the overhead:

1. **Data cleaning on read** (`_clean_value_data_element` called per cell) — 3-8 `isinstance()` checks per element
2. **Data cleaning on write** (`prepare_xl_data_element` called per cell) — 6-13 `isinstance()` checks per element
3. **Pandas date parsing** via `df.apply(xlserial_to_datetime)` — the slowest possible way to transform a column

## Solution

A new module `xlwings/conversion/fast.py` containing NumPy-accelerated replacements for the slow pipeline stages. The original stages remain untouched as fallback. A toggle system controls which path is used.

## Scope

### Deliverables

| # | Component | File | Replaces |
|---|-----------|------|----------|
| 1 | `FastCleanDataFromReadStage` | `conversion/fast.py` | `CleanDataFromReadStage` in `standard.py` |
| 2 | `FastCleanDataForWriteStage` | `conversion/fast.py` | `CleanDataForWriteStage` in `standard.py` |
| 3 | `FastTransposeStage` | `conversion/fast.py` | `TransposeStage` in `standard.py` |
| 4 | `fast_xlserial_to_datetime_series()` | `conversion/fast.py` | `df[col].apply(xlserial_to_datetime)` in `pandas_conv.py` |
| 5 | Toggle system | `__init__.py` + `standard.py` + `pandas_conv.py` | N/A |
| 6 | Unit tests | `tests/test_fast_conversion.py` | N/A (new) |

### Out of scope

- Changing the pipeline architecture (stage count, ordering)
- Caching of Range properties or COM call batching
- REST API serialization changes
- Changes to chunking defaults

## Architecture

### Toggle system

Two mechanisms, runtime takes precedence:

```python
# 1. Environment variable (checked at import time)
#    Set XLWINGS_FAST=0 to disable
os.environ.get("XLWINGS_FAST", "1") != "0"

# 2. Python-level runtime toggle
import xlwings as xw
xw.USE_FAST_CONVERSION = False  # overrides env var
```

The toggle is defined in `xlwings/__init__.py` and read by `standard.py` and `pandas_conv.py` at call time (not import time), so it can be changed at runtime.

### `FastCleanDataFromReadStage`

Replaces the nested list comprehension in `clean_value_data()` (`_xlwindows.py:487-494`, `_xlmac.py:126-130`).

**Critical: this stage does NOT replace `c.engine.impl.clean_value_data()` directly.** Instead, it provides a fast implementation that the engine impls can delegate to. The stage still calls `c.engine.impl.clean_value_data()`, but the engine impls gain a `fast=True` code path that uses the vectorized logic below. This preserves the platform dispatch mechanism.

**Platform differences:**
- **Windows:** `cell_errors` is a `dict[int, str]` — error codes are integers. Empty check is `value in ("", None)`.
- **Mac:** `cell_errors` is a `tuple[str]` — error codes are strings. Empty check is `value == "" or value == kw.missing_value` (appscript sentinel). Mac has no error-to-string conversion (errors are already strings).

The fast path must handle both platforms. Strategy: accept `cell_errors`, `empty_sentinels`, and `error_is_int` as parameters so the engine impl configures the fast path for its platform.

**Strategy:**
1. Convert 2D list-of-lists to a NumPy object array: `arr = np.array(data, dtype=object)`
   - Input is guaranteed rectangular by `Ensure2DStage` running earlier in the pipeline
2. Build boolean masks for each type category:
   - `empty_mask`: use `np.vectorize(lambda x: x is None or x == '' or x is missing_value_sentinel)(arr)` — element-wise identity check (not `arr is None` which tests the array itself)
   - `datetime_mask` using `np.vectorize(lambda x: isinstance(x, time_types))(arr)` for platform-appropriate time types
   - `error_mask` (Windows only): vectorize `lambda x: isinstance(x, int) and x in cell_errors`. On Mac, error values are strings that pass through unchanged — the original Mac `_clean_value_data_element` does not check `cell_errors`, so the fast path must also skip error handling on Mac
   - `float_mask` when `number_builder` is set
3. Apply transformations per mask using `arr[mask] = ...`
4. Return `arr.tolist()`

**Key detail:** `np.vectorize` is NOT faster than a loop for scalar isinstance — but it enables us to separate the type-detection pass from the transformation pass. The real win comes from:
- Avoiding Python function call overhead per element (no `_clean_value_data_element()` call per cell)
- Batch operations on masks (`arr[mask] = empty_as`) instead of per-element conditionals
- For the common case where most cells are plain floats/strings, the masks are sparse and the vectorized path skips most elements

**Fallback:** If an unexpected type is encountered, fall back to element-wise processing for that element only.

### `FastCleanDataForWriteStage`

Replaces the nested list comprehension in `CleanDataForWriteStage` (`standard.py:130-133`).

**Critical: this stage still delegates to `c.engine.impl.prepare_xl_data_element()` for platform correctness.** The fast path provides a vectorized wrapper that calls the engine's per-element function only for elements that need transformation, skipping the common case (plain strings/floats that pass through unchanged).

**Platform differences:**
- **Windows:** Handles `time_types`, `pd.isna()`, `np.isnan()`, `np.number`, `None`. Returns COM time objects for datetimes.
- **Mac:** All of the above, plus: `np.datetime64` → datetime conversion, `pd.Timestamp` → pydatetime, `pd.NaT` → None, `bool` must pass through before int check, `int` → `float()` (appscript SInt64 workaround, GH #227).

**Strategy:**
1. Convert to NumPy object array
2. Build a "needs processing" mask: identify elements that are NOT plain strings or plain floats (the common fast-path types that need no transformation)
3. For elements that need processing, apply the engine's `prepare_xl_data_element` — but only on the masked subset, not the full array
4. Mask-based bulk handling for the most common transformations:
   - `none_mask`: `np.vectorize(lambda x: x is None)(arr)` → replace with `""`
   - `nan_mask`: detect float NaN and numpy NaN → replace with `""`
   - `np_number_mask`: numpy numeric types → `float()` via vectorized cast
   - `pd_nat_mask`: `pd.NaT` → `None`
   - `bool_mask`: booleans pass through unchanged (must be checked before int on Mac)
   - `int_mask` (Mac only): `int` → `float()` for appscript compatibility
5. Return `arr.tolist()`

### `FastTransposeStage`

Replaces the double nested list comprehension in `TransposeStage` (`standard.py:192-194`).

```python
class FastTransposeStage:
    def __call__(self, c):
        c.value = np.array(c.value, dtype=object).T.tolist()
```

The `.T` produces a zero-copy view of the array; the allocation cost is in the initial `np.array()` and final `.tolist()`. Still significantly faster than the double nested list comprehension for large arrays.

### `fast_xlserial_to_datetime_series(series)`

Replaces `df[col].apply(xlserial_to_datetime)` in `pandas_conv.py:20-23`.

```python
def fast_xlserial_to_datetime_series(series):
    numeric = pd.to_numeric(series, errors='coerce')
    timestamps = ((numeric - 25569) * 86400).round(3)  # match precision of utils.xlserial_to_datetime
    return pd.to_datetime(timestamps, unit='s', utc=True).dt.tz_localize(None)
```

Fully vectorized — no per-element function calls. Non-numeric values are preserved as NaT. The `.round(3)` matches the existing `round(..., 3)` in `xlserial_to_datetime` to avoid sub-millisecond precision discrepancies in equivalence tests.

## Files modified

| File | Change |
|------|--------|
| `xlwings/__init__.py` | Add `USE_FAST_CONVERSION` toggle (read from env var, overridable) |
| `xlwings/conversion/fast.py` | **New file** — all fast stages |
| `xlwings/conversion/standard.py` | `ValueAccessor.reader()` and `.writer()` swap stages based on toggle |
| `xlwings/conversion/pandas_conv.py` | `_parse_dates()` uses vectorized path based on toggle |
| `tests/test_fast_conversion.py` | **New file** — unit tests |

## Testing strategy

### Unit tests (`tests/test_fast_conversion.py`)

No Excel required. Tests feed synthetic 2D lists through both old and new stages, assert identical output.

**Test matrix for read cleaning:**
- Empty strings and None values → `empty_as` parameter
- `datetime.datetime` objects → `datetime_builder` conversion
- Floats with `number_builder` (e.g., int rounding)
- Int error codes (e.g., -2146826281 for `#DIV/0!`) → None or string based on `err_to_str`
- Mixed-type rows (string, float, datetime, None, error in same row)
- Edge cases: single cell `[[val]]`, single row `[[a, b, c]]`, single column `[[a], [b], [c]]`, empty `[]`

**Test matrix for write cleaning:**
- None → `""`
- `float('nan')` and `np.nan` → `""`
- `np.float64`, `np.int64` → `float()`
- `np.datetime64` → datetime conversion
- `datetime.datetime` → COM time conversion
- `pd.NaT` → `None`
- `pd.Timestamp` → datetime
- `bool` values → passthrough (must NOT be converted to int/float)
- `int` values → `float()` on Mac, passthrough on Windows
- Plain strings and numbers (passthrough)

**Test matrix for transpose:**
- Square matrix, rectangular matrix, single row, single column, empty

**Test matrix for pandas date conversion:**
- Column of Excel serial dates → correct datetimes
- Mixed column (some non-numeric) → NaT for non-numeric
- Known edge cases: date serial 1 (1900-01-01), serial 60 (Excel's fake 1900-02-29 leap year bug), 44197 (2021-01-01)

**Equivalence tests:**
- For each stage, run same input through old stage and new stage, assert `old_output == new_output`
- This is the primary regression safety net

### Existing integration tests

The existing `tests/test_conversion.py` (requires Excel) remains unchanged and serves as end-to-end validation. Run manually after implementation to confirm no regressions with real Excel.

## Rollback

Three levels:

1. **Runtime:** `xw.USE_FAST_CONVERSION = False` or `XLWINGS_FAST=0`
2. **Branch:** `git checkout main` restores original code
3. **Package:** `uv pip install xlwings==0.11.4` restores the PyPI release

## Branch and versioning

- **Branch:** `perf/fast-conversion` (already created)
- **Version:** No source change needed — `__version__` is set at build time. For local testing: `uv pip install -e .` from the branch
- **Install/test cycle:**
  ```bash
  # Install from branch (editable)
  uv pip install -e /path/to/xlwings

  # Restore release version
  uv pip install xlwings==0.11.4
  ```
