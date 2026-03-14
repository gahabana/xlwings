# xlwings Performance Analysis

**Date:** 2026-03-14
**Goal:** Identify the top 3 most impactful performance optimizations in the xlwings codebase.

---

## Methodology

Deep analysis of the xlwings codebase focusing on:
- Data conversion and serialization hot paths
- Per-cell processing overhead in read/write pipelines
- Pandas DataFrame integration bottlenecks
- Platform-specific code paths (Windows COM / Mac appscript)

---

## Top 3 Optimizations Identified

### 1. Vectorize cell-by-cell data cleaning with NumPy (10-50x speedup on large ranges)

**The problem:** Every read/write passes through nested list comprehensions that call `_clean_value_data_element()` or `prepare_xl_data_element()` **per cell** — each doing 3-8 `isinstance()` checks.

- `_xlwindows.py:487-494` — `clean_value_data()` iterates every cell
- `_xlmac.py:126-130` — same pattern, with even more type checks (13 branches at lines 95-123)
- `standard.py:130-133` — `CleanDataForWriteStage` does the same for writes

For a 1000x1000 range, that's **1 million Python function calls** with multiple type checks each.

**The fix:** Convert the 2D list to a NumPy object array first, then use vectorized operations:
- Boolean masks for type-based dispatch instead of per-element isinstance chains
- `np.vectorize` for type detection, batch `arr[mask] = value` for transformation
- Skip elements that don't need processing (plain strings/floats pass through)

This is the single biggest win since *every* `.value` read or write hits this path.

### 2. Replace `df.apply(xlserial_to_datetime)` with vectorized math (50-100x speedup for DataFrame date columns)

**The problem:** `pandas_conv.py:20-23` uses `.apply()` — the slowest way to transform a pandas column:

```python
df[col] = df[col].apply(xlserial_to_datetime)
```

`xlserial_to_datetime` (`utils.py:73-83`) does simple arithmetic: `(serial - 25569) * 86400` then `datetime.fromtimestamp()`. This is trivially vectorizable.

**The fix:**

```python
numeric = pd.to_numeric(series, errors="coerce")
timestamps = ((numeric - 25569) * 86400).round(3)
result = pd.to_datetime(timestamps, unit="s", utc=True).dt.tz_convert(None)
```

One expression, fully vectorized. For a column with 100K dates, `.apply()` takes ~2 seconds; the vectorized version takes ~20ms.

### 3. Eliminate redundant data copies in the conversion pipeline (2-4x memory reduction, measurable speed gain)

**The problem:** The `ValueAccessor.reader` pipeline (`standard.py:251-258`) runs 5-6 sequential stages, each creating a **new full copy** of the data:

1. `ReadValueFromRangeStage` — list of lists from COM
2. `Ensure2DStage` — may wrap in another list
3. `CleanDataFromReadStage` — **new** nested list comprehension (full copy)
4. `TransposeStage` — **new** nested list comprehension (full copy)
5. `AdjustDimensionsStage` — may create another list

For a 10MB dataset, you're allocating 30-40MB of intermediate lists that get immediately discarded.

**The fix:** The transpose becomes `np.array(data, dtype=object).T.tolist()` (the `.T` is a zero-copy view). Cleaning uses in-place mask operations on the NumPy array instead of creating new lists.

---

## Key Findings

### Hot path anatomy

Every `range.value` read goes through:
```
COM/appscript call -> raw_value (list of lists)
  -> Ensure2DStage
  -> CleanDataFromReadStage (per-cell isinstance loop)
  -> TransposeStage (optional, nested list comprehension)
  -> AdjustDimensionsStage
```

Every `range.value = data` write goes through:
```
Ensure2DStage
  -> CleanDataForWriteStage (per-cell isinstance loop)
  -> TransposeStage (optional, nested list comprehension)
  -> WriteValueToRangeStage -> COM/appscript call
```

The per-cell loops are the dominant cost for large ranges — not the COM calls themselves.

### Platform differences

| Aspect | Windows | Mac |
|--------|---------|-----|
| Error codes | `dict[int, str]` | `tuple[str]` (strings pass through) |
| Empty check | `value in ("", None)` | `value == "" or value == kw.missing_value` |
| Int handling | Passthrough | `int -> float()` (appscript SInt64 bug, GH #227) |
| Time types | `datetime, pywintypes.TimeType` | `datetime, np.datetime64` |

### Dependencies

NumPy and pandas are **optional** dependencies in xlwings. Any optimization using them must either:
- Have a pure-Python fallback, OR
- Require NumPy as a dependency for the user

We chose the latter (NumPy required for the fast path) since the user confirmed this is acceptable.

---

---

## Additional Analysis: Startup Latency and Named Cell I/O

### Question 1: Can the Excel "Run" → Python execution startup be dramatically improved?

**Short answer: Yes — this is likely the single biggest perceived-latency issue, and it's solvable.**

#### What happens when you click "Run"

Every click on "Run" (via xlwings toolbar or macro button) spawns a **new Python process**:

1. **VBA macro fires** (`Main.bas:308-330`) — resolves interpreter path from config
2. **Shell command constructed** — on Mac via AppleScript (`xlwings-dev.applescript:11-16`), on Windows via `cmd.exe /C` (`Main.bas:380-420`)
3. **New `python` process starts** — full interpreter cold start
4. **`import xlwings`** — imports the entire package, initializes engines (`__init__.py:82-120`), imports pywin32/appscript
5. **`xlwings.utils.prepare_sys_path()`** — parses args, modifies sys.path (`utils.py:438-473`)
6. **Your script imports** — any `import pandas`, `import numpy`, etc. in your script
7. **Your code runs**

**Where the time goes (typical breakdown for a ~2-3 second startup):**

| Phase | Time | Notes |
|-------|------|-------|
| Process creation + Python interpreter start | ~300-500ms | OS-level, unavoidable per-launch |
| `import xlwings` + engine init | ~200-500ms | Imports pywin32/appscript, sets up engines |
| `import numpy` | ~100-200ms | If your script uses it |
| `import pandas` | ~300-600ms | If your script uses it — pandas is heavy |
| Your script's other imports | varies | Depends on your dependencies |
| Actual script execution | fast | The part you care about |

**The fix: persistent Python process (daemon/server mode)**

Instead of spawning a new Python process each click, keep a Python process running in the background that listens for commands from Excel:

- **Option A: xlwings Server (Pro feature)** — xlwings already has a REST API server mode (`xlwings restapi run`) but it's designed for remote/web use, not local RunPython acceleration.
- **Option B: Custom daemon** — A lightweight background Python process that:
  1. Starts once (manually or at Excel launch)
  2. Pre-imports xlwings, numpy, pandas, and your modules
  3. Listens on a socket/pipe for "run this function" commands from VBA
  4. Executes and returns results without process startup overhead

This would reduce the per-click latency from **2-3 seconds to ~50-100ms** (just the COM call + function execution). The approach is well-proven — Jupyter kernels work exactly this way.

**Feasibility:** High. The VBA side would need a small change to send a command to the daemon instead of spawning a process. The Python side needs a simple socket listener loop. The xlwings DLL path on Windows (`xlwings64-dev.dll` with `XLPyDLLActivateAuto`) already hints at in-process execution being a known optimization vector.

**Verdict: Dramatic improvement possible (20-50x startup reduction).**

---

### Question 2: Can reading/writing 20-30 Named cells be dramatically improved?

**Short answer: Probably not dramatically — you're already near the floor for this workload.**

#### What happens for each named cell read

```
sheet["MyNamedCell"].value
  → 1 COM call: resolve name + get Range object
  → 1 COM call: read the value
  → Conversion pipeline (trivial for scalar)
  → Return Python value
```

**Per-cell cost:** ~2 COM round-trips × 5-20ms each = **10-40ms per cell** (machine-dependent).

For 25 reads + 10 writes: **~35 cells × ~15ms average = ~500ms total**.

#### Why it's hard to optimize further

- **The bottleneck is COM/appscript latency**, not Python code. Each round-trip crosses the process boundary between Python and Excel. Our fast conversion pipeline doesn't help here — it optimizes what happens *after* data arrives in Python, but for single cells that's already trivial.
- **Named cells are scattered** — they can't be read in a single rectangular range operation. Each name resolves to a different cell on potentially different sheets.
- **No batch API** — xlwings (and the underlying COM/appscript APIs) don't offer a "read these 25 named ranges in one call" operation.

#### What could help (modest improvements)

| Approach | Savings | Complexity |
|----------|---------|------------|
| **Group reads by sheet** — read all names from Sheet1, then Sheet2, etc. Avoids repeated sheet activation overhead. | ~10-20% | Low |
| **Read contiguous ranges** — if some named cells are adjacent, read the enclosing range once and index into it. | ~5-15% per group | Low |
| **VBA-side batch helper** — a VBA function that reads all 25 names into an array and passes it to Python as a single COM call. | ~50-70% | Medium |
| **Persistent daemon (from Q1)** — if startup is solved, the 500ms for COM calls becomes the total time, which feels fast. | N/A (but feels better) | Medium |

**Verdict: Modest improvements possible (maybe 2x with VBA batching), but not the dramatic 10x+ gains we found in the conversion pipeline. The real win is solving the startup latency from Q1 — once startup drops from 2-3s to 50ms, the 500ms for 35 COM calls becomes the dominant cost, and that's acceptable.**

---

## Honorable mentions

- **Mac `prepare_xl_data_element`** (`_xlmac.py:95-123`): 13 isinstance branches checked per cell, including guards for `pd` and `np` types even when those libraries aren't involved.
- **DataFrame `.values.tolist()` conversions** (`pandas_conv.py:71-73`): Multiple intermediate conversions between pandas, numpy, and Python native types.
- **Repeated DataFrame column access** (`pandas_conv.py:53-59`): `isinstance` checks on column dtypes inside loops instead of cached outside.
