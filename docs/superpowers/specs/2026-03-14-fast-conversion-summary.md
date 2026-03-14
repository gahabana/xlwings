# Fast Conversion Pipeline — Implementation Summary

**Branch:** `perf/fast-conversion` (9 commits ahead of main)
**Date:** 2026-03-14
**Tests:** 38/38 passing (0.19s, no Excel required)

---

## What was built

Four NumPy-accelerated replacements for the per-cell Python loops in the xlwings conversion pipeline, plus a toggle system and comprehensive tests.

### Components

| Component | File | Replaces | Expected speedup |
|-----------|------|----------|-----------------|
| `FastTransposeStage` | `conversion/fast.py` | Nested list comprehension transpose | 2-5x |
| `FastCleanDataForWriteStage` | `conversion/fast.py` | Per-cell `prepare_xl_data_element()` | 10-50x on large ranges |
| `fast_clean_value_data()` | `conversion/fast.py` | Per-cell `_clean_value_data_element()` | 10-50x on large ranges |
| `fast_xlserial_to_datetime_series()` | `conversion/fast.py` | `df.apply(xlserial_to_datetime)` | 50-100x |

### Files created

- `xlwings/conversion/fast.py` — all fast stage implementations
- `tests/test_fast_conversion.py` — 38 unit + equivalence + integration tests

### Files modified

- `xlwings/__init__.py` — `USE_FAST_CONVERSION` toggle
- `xlwings/conversion/standard.py` — `ValueAccessor.reader()` / `.writer()` swap in fast stages
- `xlwings/conversion/pandas_conv.py` — `_parse_dates()` uses vectorized date conversion
- `xlwings/_xlwindows.py` — `Engine.clean_value_data()` delegates to fast path (Windows)
- `xlwings/_xlmac.py` — `Engine.clean_value_data()` delegates to fast path (Mac)

---

## Toggle system

Two mechanisms (runtime takes precedence):

```python
# Environment variable (checked at import time)
# Set XLWINGS_FAST=0 to disable
export XLWINGS_FAST=0

# Python-level runtime toggle (overrides env var)
import xlwings as xw
xw.USE_FAST_CONVERSION = False
```

Default: **enabled** (`USE_FAST_CONVERSION = True`).

---

## Rollback

| Level | Command |
|-------|---------|
| Runtime | `xw.USE_FAST_CONVERSION = False` or `XLWINGS_FAST=0` |
| Branch | `git checkout main` |
| Package | `uv pip install xlwings==0.11.4` |

---

## Installation

```bash
# Install from branch (editable)
uv pip install -e /path/to/xlwings

# Restore release version
uv pip install xlwings==0.11.4
```

---

## Running tests

```bash
source .venv/bin/activate
python -m pytest tests/test_fast_conversion.py -v
```

No Excel instance required. Tests use synthetic data and mock engine objects.

---

## Commit history

```
f08e2795 test: add end-to-end equivalence tests for fast conversion
3e948443 feat: add fast_xlserial_to_datetime_series with vectorized date conversion
dc39a530 feat: add fast_clean_value_data with vectorized read cleaning
64208211 feat: add FastCleanDataForWriteStage with vectorized write cleaning
f58f8cc3 feat: wire FastTransposeStage into ValueAccessor with toggle
b3e464ba feat: add FastTransposeStage with NumPy-accelerated transpose
abe0b539 feat: add USE_FAST_CONVERSION toggle with XLWINGS_FAST env var
b633b294 Add implementation plan for fast conversion pipeline
da8a25c0 Add design spec for NumPy-accelerated conversion pipeline
```

---

## Platform handling

The fast path handles Windows and Mac differences:

- **Windows:** Error codes are integers in a dict. Empty sentinels: `("", None)`.
- **Mac:** Error codes are strings (pass through unchanged). Empty sentinels: `("", None, kw.missing_value)`.

Each engine's `clean_value_data()` passes platform-specific parameters to `fast_clean_value_data()`.
