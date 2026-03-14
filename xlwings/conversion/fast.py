"""
NumPy-accelerated conversion stages for xlwings.

Drop-in replacements for the stages in standard.py. Activated when
xw.USE_FAST_CONVERSION is True (default). Original stages serve as
fallback when USE_FAST_CONVERSION is False.
"""

import datetime as dt

import numpy as np


class FastTransposeStage:
    def __call__(self, c):
        if not c.value:
            return
        c.value = np.array(c.value, dtype=object).T.tolist()


def _fast_prepare_xl_data(data, engine_prepare_fn, options):
    """
    Vectorized replacement for the nested list comprehension in CleanDataForWriteStage.

    Strategy: convert to NumPy object array, build a mask of elements that need
    transformation (anything that is NOT a plain str or plain float), then call
    the engine's prepare function only on those elements.
    """
    arr = np.array(data, dtype=object)

    # Build "needs processing" mask — True for anything that isn't str/float/bool passthrough
    # int is NOT excluded because Mac needs int -> float() conversion (GH #227)
    # NaN floats also need processing
    _needs_work = np.vectorize(
        lambda x: not isinstance(x, (str, float, bool))
        or isinstance(x, np.floating)  # np.floating is a float subclass but needs conversion
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
    _is_empty = np.vectorize(
        lambda x: x is None or x == "" or any(x is s for s in empty_sentinels)
    )
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
