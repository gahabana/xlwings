# Fast Conversion Pipeline — Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add NumPy-accelerated data conversion stages that replace the per-cell Python loops, yielding 10-50x speedup on large ranges.

**Architecture:** New module `xlwings/conversion/fast.py` with drop-in stage replacements. Toggle via `xw.USE_FAST_CONVERSION` (bool) and `XLWINGS_FAST` env var. Original stages untouched as fallback.

**Tech Stack:** NumPy (object arrays + vectorize masks), pandas (vectorized date conversion), pytest (unit tests)

**Spec:** `docs/superpowers/specs/2026-03-14-fast-conversion-design.md`

---

## Chunk 1: Toggle System + FastTransposeStage

### Task 1: Add USE_FAST_CONVERSION toggle to xlwings/__init__.py

**Files:**
- Modify: `xlwings/__init__.py:123` (after `on_server` line)

- [ ] **Step 1: Write the toggle**

Add after line 123 (`on_server = os.environ.get("XLWINGS_ON_SERVER") == "true"`):

```python
USE_FAST_CONVERSION = os.environ.get("XLWINGS_FAST", "1") != "0"
```

- [ ] **Step 2: Verify import works**

Run: `python -c "import xlwings as xw; print(xw.USE_FAST_CONVERSION)"`
Expected: `True`

- [ ] **Step 3: Verify env var override works**

Run: `XLWINGS_FAST=0 python -c "import xlwings as xw; print(xw.USE_FAST_CONVERSION)"`
Expected: `False`

- [ ] **Step 4: Commit**

```bash
git add xlwings/__init__.py
git commit -m "feat: add USE_FAST_CONVERSION toggle with XLWINGS_FAST env var"
```

---

### Task 2: Create fast.py with FastTransposeStage

**Files:**
- Create: `xlwings/conversion/fast.py`
- Test: `tests/test_fast_conversion.py`

This is the simplest stage — good foundation to validate the pattern before tackling the complex stages.

- [ ] **Step 1: Write the failing test for FastTransposeStage**

Create `tests/test_fast_conversion.py`:

```python
import pytest
import numpy as np

from xlwings.conversion.standard import TransposeStage


class FakeContext:
    """Minimal mock of ConversionContext for unit testing stages."""

    def __init__(self, value):
        self.value = value


class TestFastTransposeStage:
    def test_square_matrix(self):
        from xlwings.conversion.fast import FastTransposeStage

        stage = FastTransposeStage()
        ctx = FakeContext([[1, 2, 3], [4, 5, 6], [7, 8, 9]])
        stage(ctx)
        assert ctx.value == [[1, 4, 7], [2, 5, 8], [3, 6, 9]]

    def test_rectangular_wide(self):
        from xlwings.conversion.fast import FastTransposeStage

        stage = FastTransposeStage()
        ctx = FakeContext([[1, 2, 3], [4, 5, 6]])
        stage(ctx)
        assert ctx.value == [[1, 4], [2, 5], [3, 6]]

    def test_rectangular_tall(self):
        from xlwings.conversion.fast import FastTransposeStage

        stage = FastTransposeStage()
        ctx = FakeContext([[1], [2], [3]])
        stage(ctx)
        assert ctx.value == [[1, 2, 3]]

    def test_single_row(self):
        from xlwings.conversion.fast import FastTransposeStage

        stage = FastTransposeStage()
        ctx = FakeContext([[1, 2, 3]])
        stage(ctx)
        assert ctx.value == [[1], [2], [3]]

    def test_single_cell(self):
        from xlwings.conversion.fast import FastTransposeStage

        stage = FastTransposeStage()
        ctx = FakeContext([[42]])
        stage(ctx)
        assert ctx.value == [[42]]

    def test_mixed_types_preserved(self):
        """Transpose must not coerce types — strings, ints, floats, None all survive."""
        from xlwings.conversion.fast import FastTransposeStage

        stage = FastTransposeStage()
        ctx = FakeContext([["a", 1, 2.5], [None, True, ""]])
        stage(ctx)
        assert ctx.value == [["a", None], [1, True], [2.5, ""]]

    def test_empty_list(self):
        from xlwings.conversion.fast import FastTransposeStage

        stage = FastTransposeStage()
        ctx = FakeContext([])
        stage(ctx)
        assert ctx.value == []

    def test_equivalence_with_original(self):
        """FastTransposeStage must produce identical output to TransposeStage."""
        from xlwings.conversion.fast import FastTransposeStage

        data = [["hello", 1, 2.5, None], [True, "", 0, -1]]

        old_ctx = FakeContext([row[:] for row in data])
        new_ctx = FakeContext([row[:] for row in data])

        TransposeStage()(old_ctx)
        FastTransposeStage()(new_ctx)

        assert old_ctx.value == new_ctx.value
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_fast_conversion.py::TestFastTransposeStage -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'xlwings.conversion.fast'`

- [ ] **Step 3: Write FastTransposeStage implementation**

Create `xlwings/conversion/fast.py`:

```python
"""
NumPy-accelerated conversion stages for xlwings.

Drop-in replacements for the stages in standard.py. Activated when
xw.USE_FAST_CONVERSION is True (default). Original stages serve as
fallback when USE_FAST_CONVERSION is False.
"""

import numpy as np


class FastTransposeStage:
    def __call__(self, c):
        if not c.value:
            return
        c.value = np.array(c.value, dtype=object).T.tolist()
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_fast_conversion.py::TestFastTransposeStage -v`
Expected: all 7 tests PASS

- [ ] **Step 5: Commit**

```bash
git add xlwings/conversion/fast.py tests/test_fast_conversion.py
git commit -m "feat: add FastTransposeStage with NumPy-accelerated transpose"
```

---

### Task 3: Wire FastTransposeStage into ValueAccessor

**Files:**
- Modify: `xlwings/conversion/standard.py:248-269` (ValueAccessor class)

- [ ] **Step 1: Add conditional import at top of standard.py**

After the existing try/except blocks near line 19-21, add:

```python
try:
    import numpy as np
except ImportError:
    np = None
```

Note: `np` may already be imported — check first. If not present, add it.

- [ ] **Step 2: Modify ValueAccessor.reader() and .writer()**

In `ValueAccessor.reader()` (line 250), replace the TransposeStage usage with a conditional:

```python
    @staticmethod
    def reader(options):
        import xlwings as xw
        from .fast import FastTransposeStage

        use_fast = xw.USE_FAST_CONVERSION and np is not None
        transpose_stage = FastTransposeStage() if use_fast else TransposeStage()

        return (
            BaseAccessor.reader(options)
            .append_stage(ReadValueFromRangeStage(options))
            .append_stage(Ensure2DStage())
            .append_stage(CleanDataFromReadStage(options))
            .append_stage(transpose_stage, only_if=options.get("transpose", False))
            .append_stage(AdjustDimensionsStage(options))
        )

    @staticmethod
    def writer(options):
        import xlwings as xw
        from .fast import FastTransposeStage

        use_fast = xw.USE_FAST_CONVERSION and np is not None
        transpose_stage = FastTransposeStage() if use_fast else TransposeStage()

        return (
            Pipeline()
            .prepend_stage(FormatStage(options))
            .prepend_stage(WriteValueToRangeStage(options))
            .prepend_stage(CleanDataForWriteStage(options))
            .prepend_stage(transpose_stage, only_if=options.get("transpose", False))
            .prepend_stage(Ensure2DStage())
        )
```

- [ ] **Step 3: Add test that toggle controls stage selection**

Append to `tests/test_fast_conversion.py`:

```python
class TestToggle:
    def test_toggle_default_is_true(self):
        import xlwings as xw

        assert xw.USE_FAST_CONVERSION is True

    def test_toggle_can_be_disabled_at_runtime(self):
        import xlwings as xw

        original = xw.USE_FAST_CONVERSION
        try:
            xw.USE_FAST_CONVERSION = False
            assert xw.USE_FAST_CONVERSION is False
        finally:
            xw.USE_FAST_CONVERSION = original
```

- [ ] **Step 4: Run all tests**

Run: `python -m pytest tests/test_fast_conversion.py -v`
Expected: all PASS

- [ ] **Step 5: Commit**

```bash
git add xlwings/conversion/standard.py tests/test_fast_conversion.py
git commit -m "feat: wire FastTransposeStage into ValueAccessor with toggle"
```

---

## Chunk 2: FastCleanDataForWriteStage

### Task 4: Write tests for FastCleanDataForWriteStage

**Files:**
- Modify: `tests/test_fast_conversion.py`

The write stage calls `c.engine.impl.prepare_xl_data_element(x, options)` per element. Our fast version skips elements that are already plain strings or floats (the common case) and only calls the engine function for elements that need transformation.

- [ ] **Step 1: Add FakeEngine mock and write tests**

Append to `tests/test_fast_conversion.py`:

```python
import datetime as dt

try:
    import pandas as pd
except ImportError:
    pd = None


class FakeEngineImpl:
    """Mock engine impl that mimics the Windows prepare_xl_data_element."""

    @staticmethod
    def prepare_xl_data_element(x, options):
        if isinstance(x, (dt.datetime, dt.date)):
            return x  # real impl converts to COM time; we just pass through for testing
        elif pd and pd.isna(x):
            return ""
        elif isinstance(x, (np.floating, float)) and np.isnan(x):
            return ""
        elif isinstance(x, np.number):
            return float(x)
        elif x is None:
            return ""
        else:
            return x


class FakeEngine:
    def __init__(self):
        self.impl = FakeEngineImpl()


class FakeContextWithEngine:
    """Mock context with engine for write stage testing."""

    def __init__(self, value):
        self.value = value
        self.engine = FakeEngine()
        self.meta = {}


class TestFastCleanDataForWriteStage:
    def test_none_becomes_empty_string(self):
        from xlwings.conversion.fast import FastCleanDataForWriteStage

        stage = FastCleanDataForWriteStage({})
        ctx = FakeContextWithEngine([[None, "hello"], [None, None]])
        stage(ctx)
        assert ctx.value == [["", "hello"], ["", ""]]

    def test_nan_becomes_empty_string(self):
        from xlwings.conversion.fast import FastCleanDataForWriteStage

        stage = FastCleanDataForWriteStage({})
        ctx = FakeContextWithEngine([[float("nan"), 1.0], [np.nan, "text"]])
        stage(ctx)
        assert ctx.value == [["", 1.0], ["", "text"]]

    def test_numpy_numbers_become_float(self):
        from xlwings.conversion.fast import FastCleanDataForWriteStage

        stage = FastCleanDataForWriteStage({})
        ctx = FakeContextWithEngine([[np.float64(3.14), np.int64(42)]])
        stage(ctx)
        assert ctx.value == [[3.14, 42.0]]
        assert all(type(v) is float for v in ctx.value[0])

    def test_strings_and_floats_pass_through(self):
        from xlwings.conversion.fast import FastCleanDataForWriteStage

        stage = FastCleanDataForWriteStage({})
        ctx = FakeContextWithEngine([["hello", 1.5, 0, "world"]])
        stage(ctx)
        assert ctx.value == [["hello", 1.5, 0, "world"]]

    def test_bool_not_converted(self):
        from xlwings.conversion.fast import FastCleanDataForWriteStage

        stage = FastCleanDataForWriteStage({})
        ctx = FakeContextWithEngine([[True, False, 1]])
        stage(ctx)
        assert ctx.value[0][0] is True
        assert ctx.value[0][1] is False

    def test_mixed_types(self):
        from xlwings.conversion.fast import FastCleanDataForWriteStage

        stage = FastCleanDataForWriteStage({})
        ctx = FakeContextWithEngine([
            ["text", 1.0, None, np.float64(2.5)],
            [float("nan"), np.int64(10), True, ""],
        ])
        stage(ctx)
        assert ctx.value == [
            ["text", 1.0, "", 2.5],
            ["", 10.0, True, ""],
        ]

    def test_single_cell(self):
        from xlwings.conversion.fast import FastCleanDataForWriteStage

        stage = FastCleanDataForWriteStage({})
        ctx = FakeContextWithEngine([[None]])
        stage(ctx)
        assert ctx.value == [[""]]

    @pytest.mark.skipif(pd is None, reason="pandas not installed")
    def test_pd_nat_becomes_none(self):
        from xlwings.conversion.fast import FastCleanDataForWriteStage

        stage = FastCleanDataForWriteStage({})
        ctx = FakeContextWithEngine([[pd.NaT, "text"]])
        stage(ctx)
        assert ctx.value[0][0] is None or ctx.value[0][0] == ""
        assert ctx.value[0][1] == "text"

    @pytest.mark.skipif(pd is None, reason="pandas not installed")
    def test_pd_timestamp(self):
        from xlwings.conversion.fast import FastCleanDataForWriteStage

        ts = pd.Timestamp("2024-01-15 10:30:00")
        stage = FastCleanDataForWriteStage({})
        ctx = FakeContextWithEngine([[ts, "text"]])
        stage(ctx)
        # Engine converts Timestamp; exact result depends on engine impl
        assert ctx.value[0][1] == "text"

    def test_equivalence_with_original(self):
        """FastCleanDataForWriteStage must match CleanDataForWriteStage output."""
        from xlwings.conversion.fast import FastCleanDataForWriteStage
        from xlwings.conversion.standard import CleanDataForWriteStage

        data = [
            ["text", 1.0, None, np.float64(2.5)],
            [float("nan"), np.int64(10), True, ""],
        ]

        old_ctx = FakeContextWithEngine([row[:] for row in data])
        new_ctx = FakeContextWithEngine([row[:] for row in data])

        CleanDataForWriteStage({})(old_ctx)
        FastCleanDataForWriteStage({})(new_ctx)

        assert old_ctx.value == new_ctx.value
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_fast_conversion.py::TestFastCleanDataForWriteStage -v`
Expected: FAIL — `ImportError: cannot import name 'FastCleanDataForWriteStage'`

- [ ] **Step 3: Commit test file**

```bash
git add tests/test_fast_conversion.py
git commit -m "test: add failing tests for FastCleanDataForWriteStage"
```

---

### Task 5: Implement FastCleanDataForWriteStage

**Files:**
- Modify: `xlwings/conversion/fast.py`

- [ ] **Step 1: Implement the fast write stage**

Append to `xlwings/conversion/fast.py`:

```python
def _fast_prepare_xl_data(data, engine_prepare_fn, options):
    """
    Vectorized replacement for the nested list comprehension in CleanDataForWriteStage.

    Strategy: convert to NumPy object array, build a mask of elements that need
    transformation (anything that is NOT a plain str or plain float), then call
    the engine's prepare function only on those elements.
    """
    arr = np.array(data, dtype=object)
    rows, cols = arr.shape

    # Build "needs processing" mask — True for anything that isn't str/float/int passthrough
    # We use np.vectorize for the type check since elements are heterogeneous
    _needs_work = np.vectorize(
        lambda x: not isinstance(x, (str, float, bool))
        or (isinstance(x, float) and x != x)  # NaN check via self-inequality
        or x is None
    )
    mask = _needs_work(arr)

    # Fast path: if nothing needs work, skip entirely
    if not mask.any():
        return arr.tolist()

    # Apply engine's prepare function only to masked elements
    indices = np.argwhere(mask)
    for row_idx, col_idx in indices:
        arr[row_idx, col_idx] = engine_prepare_fn(arr[row_idx, col_idx], options)

    return arr.tolist()


class FastCleanDataForWriteStage:
    def __init__(self, options):
        self.options = options

    def __call__(self, c):
        c.value = _fast_prepare_xl_data(
            c.value, c.engine.impl.prepare_xl_data_element, self.options
        )
```

- [ ] **Step 2: Run tests**

Run: `python -m pytest tests/test_fast_conversion.py::TestFastCleanDataForWriteStage -v`
Expected: all 8 tests PASS

- [ ] **Step 3: Run all tests so far**

Run: `python -m pytest tests/test_fast_conversion.py -v`
Expected: all PASS

- [ ] **Step 4: Commit**

```bash
git add xlwings/conversion/fast.py
git commit -m "feat: add FastCleanDataForWriteStage with vectorized write cleaning"
```

---

### Task 6: Wire FastCleanDataForWriteStage into ValueAccessor

**Files:**
- Modify: `xlwings/conversion/standard.py` (ValueAccessor.writer)

- [ ] **Step 1: Update writer() to use fast stage**

In `ValueAccessor.writer()`, add the fast write stage import and conditional:

```python
    @staticmethod
    def writer(options):
        import xlwings as xw
        from .fast import FastCleanDataForWriteStage, FastTransposeStage

        use_fast = xw.USE_FAST_CONVERSION and np is not None
        transpose_stage = FastTransposeStage() if use_fast else TransposeStage()
        clean_write_stage = (
            FastCleanDataForWriteStage(options)
            if use_fast
            else CleanDataForWriteStage(options)
        )

        return (
            Pipeline()
            .prepend_stage(FormatStage(options))
            .prepend_stage(WriteValueToRangeStage(options))
            .prepend_stage(clean_write_stage)
            .prepend_stage(transpose_stage, only_if=options.get("transpose", False))
            .prepend_stage(Ensure2DStage())
        )
```

- [ ] **Step 2: Run all tests**

Run: `python -m pytest tests/test_fast_conversion.py -v`
Expected: all PASS

- [ ] **Step 3: Commit**

```bash
git add xlwings/conversion/standard.py
git commit -m "feat: wire FastCleanDataForWriteStage into ValueAccessor.writer"
```

---

## Chunk 3: FastCleanDataFromReadStage

### Task 7: Write tests for FastCleanDataFromReadStage

**Files:**
- Modify: `tests/test_fast_conversion.py`

The read stage is more complex — it delegates to `c.engine.impl.clean_value_data()` which differs between Windows and Mac. The fast path provides a `fast_clean_value_data()` function that both engine impls can call.

- [ ] **Step 1: Write tests**

Append to `tests/test_fast_conversion.py`:

```python
class TestFastCleanValueData:
    """Test the fast_clean_value_data function directly (platform-independent logic)."""

    def test_empty_string_replaced(self):
        from xlwings.conversion.fast import fast_clean_value_data

        result = fast_clean_value_data(
            [["", "hello"], ["", ""]],
            datetime_builder=dt.datetime,
            empty_as=None,
            number_builder=None,
            err_to_str=False,
            cell_errors={},
            empty_sentinels=("", None),
        )
        assert result == [[None, "hello"], [None, None]]

    def test_none_replaced(self):
        from xlwings.conversion.fast import fast_clean_value_data

        result = fast_clean_value_data(
            [[None, "text"]],
            datetime_builder=dt.datetime,
            empty_as="N/A",
            number_builder=None,
            err_to_str=False,
            cell_errors={},
            empty_sentinels=("", None),
        )
        assert result == [["N/A", "text"]]

    def test_number_builder_applied_to_floats(self):
        from xlwings.conversion.fast import fast_clean_value_data

        result = fast_clean_value_data(
            [[1.6, 2.4, "text"]],
            datetime_builder=dt.datetime,
            empty_as=None,
            number_builder=lambda x: int(round(x)),
            err_to_str=False,
            cell_errors={},
            empty_sentinels=("", None),
        )
        assert result == [[2, 2, "text"]]

    def test_windows_error_codes_to_none(self):
        from xlwings.conversion.fast import fast_clean_value_data

        win_errors = {
            -2146826281: "#DIV/0!",
            -2146826246: "#N/A",
        }
        result = fast_clean_value_data(
            [[-2146826281, "text", -2146826246]],
            datetime_builder=dt.datetime,
            empty_as=None,
            number_builder=None,
            err_to_str=False,
            cell_errors=win_errors,
            empty_sentinels=("", None),
        )
        assert result == [[None, "text", None]]

    def test_windows_error_codes_to_string(self):
        from xlwings.conversion.fast import fast_clean_value_data

        win_errors = {
            -2146826281: "#DIV/0!",
            -2146826246: "#N/A",
        }
        result = fast_clean_value_data(
            [[-2146826281, "text"]],
            datetime_builder=dt.datetime,
            empty_as=None,
            number_builder=None,
            err_to_str=True,
            cell_errors=win_errors,
            empty_sentinels=("", None),
        )
        assert result == [["#DIV/0!", "text"]]

    def test_mac_errors_are_strings_passthrough(self):
        """On Mac, cell_errors is empty dict — string errors pass through unchanged."""
        from xlwings.conversion.fast import fast_clean_value_data

        result = fast_clean_value_data(
            [["#DIV/0!", "text", "#N/A"]],
            datetime_builder=dt.datetime,
            empty_as=None,
            number_builder=None,
            err_to_str=False,
            cell_errors={},
            empty_sentinels=("", None),
        )
        assert result == [["#DIV/0!", "text", "#N/A"]]

    def test_datetime_passthrough_when_builder_is_datetime(self):
        from xlwings.conversion.fast import fast_clean_value_data

        now = dt.datetime(2024, 1, 15, 10, 30, 0)
        result = fast_clean_value_data(
            [[now, "text"]],
            datetime_builder=dt.datetime,
            empty_as=None,
            number_builder=None,
            err_to_str=False,
            cell_errors={},
            empty_sentinels=("", None),
        )
        assert result == [[now, "text"]]

    def test_datetime_converted_to_date(self):
        from xlwings.conversion.fast import fast_clean_value_data

        now = dt.datetime(2024, 1, 15, 10, 30, 0)
        date_builder = lambda year, month, day, **kwargs: dt.date(year, month, day)
        result = fast_clean_value_data(
            [[now, "text"]],
            datetime_builder=date_builder,
            empty_as=None,
            number_builder=None,
            err_to_str=False,
            cell_errors={},
            empty_sentinels=("", None),
        )
        assert result == [[dt.date(2024, 1, 15), "text"]]

    def test_mixed_types_all_handled(self):
        from xlwings.conversion.fast import fast_clean_value_data

        win_errors = {-2146826281: "#DIV/0!"}
        now = dt.datetime(2024, 6, 1, 12, 0, 0)
        result = fast_clean_value_data(
            [["text", 3.14, None, now], ["", -2146826281, 42, True]],
            datetime_builder=dt.datetime,
            empty_as=None,
            number_builder=None,
            err_to_str=False,
            cell_errors=win_errors,
            empty_sentinels=("", None),
        )
        assert result == [["text", 3.14, None, now], [None, None, 42, True]]

    def test_single_cell(self):
        from xlwings.conversion.fast import fast_clean_value_data

        result = fast_clean_value_data(
            [[""]],
            datetime_builder=dt.datetime,
            empty_as=None,
            number_builder=None,
            err_to_str=False,
            cell_errors={},
            empty_sentinels=("", None),
        )
        assert result == [[None]]

    def test_single_column(self):
        from xlwings.conversion.fast import fast_clean_value_data

        result = fast_clean_value_data(
            [["a"], [""], [None]],
            datetime_builder=dt.datetime,
            empty_as="EMPTY",
            number_builder=None,
            err_to_str=False,
            cell_errors={},
            empty_sentinels=("", None),
        )
        assert result == [["a"], ["EMPTY"], ["EMPTY"]]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_fast_conversion.py::TestFastCleanValueData -v`
Expected: FAIL — `ImportError: cannot import name 'fast_clean_value_data'`

- [ ] **Step 3: Commit test file**

```bash
git add tests/test_fast_conversion.py
git commit -m "test: add failing tests for fast_clean_value_data"
```

---

### Task 8: Implement fast_clean_value_data

**Files:**
- Modify: `xlwings/conversion/fast.py`

- [ ] **Step 1: Implement the function**

Append to `xlwings/conversion/fast.py`:

```python
import datetime as dt


def fast_clean_value_data(
    data,
    datetime_builder,
    empty_as,
    number_builder,
    err_to_str,
    cell_errors,
    empty_sentinels=("", None),
):
    """
    Vectorized replacement for per-cell _clean_value_data_element.

    Parameters
    ----------
    data : list[list]
        2D list of cell values from Excel.
    datetime_builder : callable
        Factory for datetime objects (e.g., datetime.datetime or datetime.date lambda).
    empty_as : any
        Value to substitute for empty cells.
    number_builder : callable or None
        Optional transform for float values (e.g., int rounding).
    err_to_str : bool
        If True, convert error codes to their string representation.
    cell_errors : dict
        Mapping of error codes to error strings. On Windows: {int: str}.
        On Mac: empty dict (errors are strings, pass through unchanged).
    empty_sentinels : tuple
        Values that count as "empty" — ("", None) on Windows,
        ("", None, kw.missing_value) on Mac.
    """
    arr = np.array(data, dtype=object)

    # --- Empty mask ---
    _is_empty = np.vectorize(lambda x: x is None or x == "" or any(x is s for s in empty_sentinels))
    empty_mask = _is_empty(arr)
    if empty_mask.any():
        arr[empty_mask] = empty_as

    # --- Datetime mask ---
    _is_datetime = np.vectorize(lambda x: isinstance(x, (dt.datetime, dt.date)))
    datetime_mask = _is_datetime(arr)
    if datetime_mask.any() and datetime_builder is not dt.datetime:
        for row_idx, col_idx in np.argwhere(datetime_mask):
            value = arr[row_idx, col_idx]
            arr[row_idx, col_idx] = datetime_builder(
                year=value.year,
                month=value.month,
                day=value.day,
                hour=getattr(value, "hour", 0),
                minute=getattr(value, "minute", 0),
                second=getattr(value, "second", 0),
                microsecond=getattr(value, "microsecond", 0),
                tzinfo=None,
            )

    # --- Number builder mask ---
    if number_builder is not None:
        _is_float = np.vectorize(lambda x: isinstance(x, float))
        float_mask = _is_float(arr)
        if float_mask.any():
            for row_idx, col_idx in np.argwhere(float_mask):
                arr[row_idx, col_idx] = number_builder(arr[row_idx, col_idx])

    # --- Error mask (Windows only — cell_errors is a dict with int keys) ---
    if cell_errors:
        _is_error = np.vectorize(lambda x: isinstance(x, int) and x in cell_errors)
        error_mask = _is_error(arr)
        if error_mask.any():
            if err_to_str:
                for row_idx, col_idx in np.argwhere(error_mask):
                    arr[row_idx, col_idx] = cell_errors[arr[row_idx, col_idx]]
            else:
                arr[error_mask] = None

    return arr.tolist()
```

- [ ] **Step 2: Move the `import datetime as dt` to the top of fast.py**

Make sure the import is at module level, not duplicated.

- [ ] **Step 3: Run tests**

Run: `python -m pytest tests/test_fast_conversion.py::TestFastCleanValueData -v`
Expected: all 12 tests PASS

- [ ] **Step 4: Run full test suite so far**

Run: `python -m pytest tests/test_fast_conversion.py -v`
Expected: all PASS

- [ ] **Step 5: Commit**

```bash
git add xlwings/conversion/fast.py
git commit -m "feat: add fast_clean_value_data with vectorized read cleaning"
```

---

### Task 9: Wire fast_clean_value_data into engine impls

**Files:**
- Modify: `xlwings/_xlwindows.py:467-477` (Engine.clean_value_data)
- Modify: `xlwings/_xlmac.py:86-96` (Engine.clean_value_data)
- Modify: `xlwings/conversion/standard.py` (ValueAccessor.reader)

The engine `clean_value_data` methods gain a fast path that delegates to `fast_clean_value_data` with platform-specific parameters.

- [ ] **Step 1: Modify Windows engine**

In `xlwings/_xlwindows.py`, modify `Engine.clean_value_data()`:

```python
    @staticmethod
    def clean_value_data(data, datetime_builder, empty_as, number_builder, err_to_str):
        import xlwings as xw

        if xw.USE_FAST_CONVERSION:
            try:
                from .conversion.fast import fast_clean_value_data

                return fast_clean_value_data(
                    data,
                    datetime_builder,
                    empty_as,
                    number_builder,
                    err_to_str,
                    cell_errors=cell_errors,
                    empty_sentinels=("", None),
                )
            except ImportError:
                pass

        return [
            [
                _clean_value_data_element(
                    c, datetime_builder, empty_as, number_builder, err_to_str
                )
                for c in row
            ]
            for row in data
        ]
```

- [ ] **Step 2: Modify Mac engine**

In `xlwings/_xlmac.py`, modify `Engine.clean_value_data()`:

```python
    @staticmethod
    def clean_value_data(data, datetime_builder, empty_as, number_builder, err_to_str):
        import xlwings as xw

        if xw.USE_FAST_CONVERSION:
            try:
                from .conversion.fast import fast_clean_value_data

                return fast_clean_value_data(
                    data,
                    datetime_builder,
                    empty_as,
                    number_builder,
                    err_to_str,
                    cell_errors={},  # Mac errors are strings, pass through unchanged
                    empty_sentinels=("", None, kw.missing_value),
                )
            except ImportError:
                pass

        return [
            [
                _clean_value_data_element(
                    c, datetime_builder, empty_as, number_builder, err_to_str
                )
                for c in row
            ]
            for row in data
        ]
```

- [ ] **Step 3: Run all tests**

Run: `python -m pytest tests/test_fast_conversion.py -v`
Expected: all PASS

- [ ] **Step 4: Commit**

```bash
git add xlwings/_xlwindows.py xlwings/_xlmac.py
git commit -m "feat: wire fast_clean_value_data into Windows and Mac engine impls"
```

---

## Chunk 4: Vectorized Pandas Date Conversion

### Task 10: Write tests for fast_xlserial_to_datetime_series

**Files:**
- Modify: `tests/test_fast_conversion.py`

- [ ] **Step 1: Write tests**

Append to `tests/test_fast_conversion.py`:

```python
@pytest.mark.skipif(pd is None, reason="pandas not installed")
class TestFastXlserialToDatetimeSeries:
    def test_known_date_44197(self):
        """44197.0 = 2021-01-01"""
        from xlwings.conversion.fast import fast_xlserial_to_datetime_series

        series = pd.Series([44197.0])
        result = fast_xlserial_to_datetime_series(series)
        assert result.iloc[0] == dt.datetime(2021, 1, 1)

    def test_known_date_serial_1(self):
        """Serial 1 = 1900-01-01 (in xlserial_to_datetime's convention)"""
        from xlwings.conversion.fast import fast_xlserial_to_datetime_series
        from xlwings.utils import xlserial_to_datetime

        series = pd.Series([1.0])
        result = fast_xlserial_to_datetime_series(series)
        expected = xlserial_to_datetime(1.0)
        assert result.iloc[0] == expected

    def test_known_date_serial_60(self):
        """Serial 60 — Excel's fake 1900-02-29 leap year bug."""
        from xlwings.conversion.fast import fast_xlserial_to_datetime_series
        from xlwings.utils import xlserial_to_datetime

        series = pd.Series([60.0])
        result = fast_xlserial_to_datetime_series(series)
        expected = xlserial_to_datetime(60.0)
        assert result.iloc[0] == expected

    def test_multiple_dates(self):
        from xlwings.conversion.fast import fast_xlserial_to_datetime_series
        from xlwings.utils import xlserial_to_datetime

        serials = [44197.0, 44562.0, 44927.0]
        series = pd.Series(serials)
        result = fast_xlserial_to_datetime_series(series)
        for i, serial in enumerate(serials):
            assert result.iloc[i] == xlserial_to_datetime(serial)

    def test_non_numeric_becomes_nat(self):
        from xlwings.conversion.fast import fast_xlserial_to_datetime_series

        series = pd.Series([44197.0, "not a date", None])
        result = fast_xlserial_to_datetime_series(series)
        assert result.iloc[0] == dt.datetime(2021, 1, 1)
        assert pd.isna(result.iloc[1])
        assert pd.isna(result.iloc[2])

    def test_equivalence_with_apply(self):
        """Vectorized version must match df.apply(xlserial_to_datetime) exactly."""
        from xlwings.conversion.fast import fast_xlserial_to_datetime_series
        from xlwings.utils import xlserial_to_datetime

        serials = [1.0, 60.0, 44197.0, 44562.0, 44927.5, 43831.75]
        series = pd.Series(serials)

        expected = series.apply(xlserial_to_datetime)
        result = fast_xlserial_to_datetime_series(series)

        for i in range(len(serials)):
            assert result.iloc[i] == expected.iloc[i], (
                f"Mismatch at index {i}: serial={serials[i]}, "
                f"fast={result.iloc[i]}, apply={expected.iloc[i]}"
            )
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_fast_conversion.py::TestFastXlserialToDatetimeSeries -v`
Expected: FAIL — `ImportError`

- [ ] **Step 3: Commit**

```bash
git add tests/test_fast_conversion.py
git commit -m "test: add failing tests for fast_xlserial_to_datetime_series"
```

---

### Task 11: Implement fast_xlserial_to_datetime_series

**Files:**
- Modify: `xlwings/conversion/fast.py`

- [ ] **Step 1: Add the function**

Append to `xlwings/conversion/fast.py`:

```python
try:
    import pandas as pd
except ImportError:
    pd = None


def fast_xlserial_to_datetime_series(series):
    """
    Vectorized replacement for df[col].apply(xlserial_to_datetime).

    Converts a pandas Series of Excel date serials to datetime objects.
    Matches the precision of utils.xlserial_to_datetime by rounding
    timestamps to 3 decimal places.

    Non-numeric values become NaT.
    """
    numeric = pd.to_numeric(series, errors="coerce")
    timestamps = ((numeric - 25569) * 86400).round(3)
    result = pd.to_datetime(timestamps, unit="s", utc=True)
    return result.dt.tz_convert(None)
```

Note: Move the `import pandas` try/except to the top of fast.py if not already there.

- [ ] **Step 2: Run tests**

Run: `python -m pytest tests/test_fast_conversion.py::TestFastXlserialToDatetimeSeries -v`
Expected: all 6 tests PASS

- [ ] **Step 3: Commit**

```bash
git add xlwings/conversion/fast.py
git commit -m "feat: add fast_xlserial_to_datetime_series with vectorized date conversion"
```

---

### Task 12: Wire fast_xlserial_to_datetime_series into pandas_conv.py

**Files:**
- Modify: `xlwings/conversion/pandas_conv.py:12-24` (_parse_dates function)

- [ ] **Step 1: Modify _parse_dates to use fast path**

Replace the `_parse_dates` function:

```python
    def _parse_dates(df, parse_dates):
        # Office.js UDFs don't have the info whether the cell is in date format
        if parse_dates is True:
            parse_dates = [0]
        elif not isinstance(parse_dates, list):
            parse_dates = [parse_dates]

        import xlwings as xw

        use_fast = xw.USE_FAST_CONVERSION
        if use_fast:
            try:
                from .fast import fast_xlserial_to_datetime_series
            except ImportError:
                use_fast = False

        for col in parse_dates:
            if isinstance(col, str):
                if use_fast:
                    df[col] = fast_xlserial_to_datetime_series(df[col])
                else:
                    df[col] = df[col].apply(xlserial_to_datetime)
            else:
                col_name = df.columns[col]
                if use_fast:
                    df[col_name] = fast_xlserial_to_datetime_series(df.iloc[:, col])
                else:
                    df[col_name] = df.iloc[:, col].apply(xlserial_to_datetime)
        return df
```

- [ ] **Step 2: Run all tests**

Run: `python -m pytest tests/test_fast_conversion.py -v`
Expected: all PASS

- [ ] **Step 3: Commit**

```bash
git add xlwings/conversion/pandas_conv.py
git commit -m "feat: wire fast_xlserial_to_datetime_series into _parse_dates"
```

---

## Chunk 5: Final Integration + Verification

### Task 13: Add integration-level equivalence tests

**Files:**
- Modify: `tests/test_fast_conversion.py`

These tests simulate the full read/write pipeline behavior by running the same data through old and new stages in sequence, verifying end-to-end equivalence.

- [ ] **Step 1: Write integration equivalence tests**

Append to `tests/test_fast_conversion.py`:

```python
class TestEndToEndEquivalence:
    """Run same data through old and new full pipelines, assert identical results."""

    def _run_read_pipeline(self, data, use_fast, datetime_builder=dt.datetime,
                           empty_as=None, number_builder=None, err_to_str=False):
        """Simulate CleanDataFromRead + Transpose stages."""
        from xlwings.conversion.fast import fast_clean_value_data, FastTransposeStage
        from xlwings.conversion.standard import TransposeStage

        win_errors = {
            -2146826281: "#DIV/0!",
            -2146826246: "#N/A",
            -2146826259: "#NAME?",
            -2146826288: "#NULL!",
            -2146826252: "#NUM!",
            -2146826265: "#REF!",
            -2146826273: "#VALUE!",
        }

        if use_fast:
            cleaned = fast_clean_value_data(
                data, datetime_builder, empty_as, number_builder,
                err_to_str, cell_errors=win_errors, empty_sentinels=("", None),
            )
        else:
            # Inline the Windows _clean_value_data_element logic to avoid
            # importing xlwings._xlwindows (requires pywintypes on Windows only)
            def _clean_element(value):
                if value in ("", None):
                    return empty_as
                elif isinstance(value, (dt.datetime, dt.date)):
                    return datetime_builder(
                        year=value.year, month=value.month, day=value.day,
                        hour=getattr(value, "hour", 0), minute=getattr(value, "minute", 0),
                        second=getattr(value, "second", 0),
                        microsecond=getattr(value, "microsecond", 0), tzinfo=None,
                    ) if datetime_builder is not dt.datetime else value
                elif number_builder is not None and isinstance(value, float):
                    return number_builder(value)
                elif isinstance(value, int) and value in win_errors:
                    return win_errors[value] if err_to_str else None
                return value

            cleaned = [[_clean_element(c) for c in row] for row in data]

        return cleaned

    def test_large_mixed_data_read(self):
        """1000-row mixed-type dataset, old vs new must match."""
        import random
        random.seed(42)

        win_errors = {
            -2146826281: "#DIV/0!",
            -2146826246: "#N/A",
        }

        rows = []
        for _ in range(1000):
            row = [
                random.choice(["text", "", None, 3.14, 42, -2146826281,
                               dt.datetime(2024, 1, 1), True]),
                random.random() * 100,
                random.choice(["hello", "", None]),
            ]
            rows.append(row)

        old_result = self._run_read_pipeline([r[:] for r in rows], use_fast=False)
        new_result = self._run_read_pipeline([r[:] for r in rows], use_fast=True)
        assert old_result == new_result

    def test_large_mixed_data_write(self):
        """1000-row write, old vs new must match."""
        from xlwings.conversion.fast import FastCleanDataForWriteStage
        from xlwings.conversion.standard import CleanDataForWriteStage
        import random
        random.seed(42)

        rows = []
        for _ in range(1000):
            row = [
                random.choice(["text", None, 3.14, np.float64(2.5), np.int64(10),
                               float("nan"), np.nan, True, False, ""]),
                random.random() * 100,
                random.choice(["hello", None, ""]),
            ]
            rows.append(row)

        old_ctx = FakeContextWithEngine([r[:] for r in rows])
        new_ctx = FakeContextWithEngine([r[:] for r in rows])

        CleanDataForWriteStage({})(old_ctx)
        FastCleanDataForWriteStage({})(new_ctx)

        assert old_ctx.value == new_ctx.value
```

- [ ] **Step 2: Run integration tests**

Run: `python -m pytest tests/test_fast_conversion.py::TestEndToEndEquivalence -v`
Expected: all PASS

- [ ] **Step 3: Commit**

```bash
git add tests/test_fast_conversion.py
git commit -m "test: add end-to-end equivalence tests for fast conversion"
```

---

### Task 14: Run full test suite and verify

- [ ] **Step 1: Run all fast conversion tests**

Run: `python -m pytest tests/test_fast_conversion.py -v`
Expected: all PASS

- [ ] **Step 2: Run with toggle disabled to verify fallback**

Run: `XLWINGS_FAST=0 python -m pytest tests/test_fast_conversion.py -v`
Expected: all PASS (toggle tests still pass; equivalence tests still compare correctly)

- [ ] **Step 3: Verify editable install works**

Run: `uv pip install -e /Users/zh/gd/git/xlwings`
Expected: successful install

- [ ] **Step 4: Quick smoke test**

```bash
python -c "
import xlwings as xw
print('USE_FAST_CONVERSION:', xw.USE_FAST_CONVERSION)
from xlwings.conversion.fast import (
    FastTransposeStage,
    FastCleanDataForWriteStage,
    fast_clean_value_data,
    fast_xlserial_to_datetime_series,
)
print('All fast components imported successfully')
"
```
Expected: prints True and success message

- [ ] **Step 5: Final commit**

```bash
git add -A
git commit -m "chore: final verification — all fast conversion components wired and tested"
```

---

## File Structure Summary

| File | Status | Responsibility |
|------|--------|----------------|
| `xlwings/__init__.py` | Modified | `USE_FAST_CONVERSION` toggle |
| `xlwings/conversion/fast.py` | **New** | `FastTransposeStage`, `FastCleanDataForWriteStage`, `fast_clean_value_data`, `fast_xlserial_to_datetime_series` |
| `xlwings/conversion/standard.py` | Modified | `ValueAccessor.reader()/.writer()` swap stages based on toggle |
| `xlwings/conversion/pandas_conv.py` | Modified | `_parse_dates()` uses vectorized path |
| `xlwings/_xlwindows.py` | Modified | `Engine.clean_value_data()` delegates to fast path |
| `xlwings/_xlmac.py` | Modified | `Engine.clean_value_data()` delegates to fast path with Mac sentinels |
| `tests/test_fast_conversion.py` | **New** | All unit + equivalence tests |

## Rollback Verification

After all tasks complete, verify all three rollback paths work:

1. `xw.USE_FAST_CONVERSION = False` — runtime disable
2. `XLWINGS_FAST=0` — env var disable
3. `uv pip install xlwings==0.11.4` — restore release
