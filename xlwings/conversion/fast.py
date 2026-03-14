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
