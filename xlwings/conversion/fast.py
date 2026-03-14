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
