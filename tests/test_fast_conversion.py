import datetime as dt

import pytest
import numpy as np

try:
    import pandas as pd
except ImportError:
    pd = None

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


class FakeEngineImpl:
    """Mock engine impl that mimics the Windows prepare_xl_data_element."""

    @staticmethod
    def prepare_xl_data_element(x, options):
        if x is None:
            return ""
        elif pd and isinstance(x, type(pd.NaT)) and pd.isna(x):
            return ""
        elif isinstance(x, (np.floating, float)) and np.isnan(x):
            return ""
        elif isinstance(x, np.number):
            return float(x)
        elif isinstance(x, (dt.datetime, dt.date)):
            return x  # real impl converts to COM time; we just pass through for testing
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
    def test_pd_nat_becomes_empty_string(self):
        from xlwings.conversion.fast import FastCleanDataForWriteStage

        stage = FastCleanDataForWriteStage({})
        ctx = FakeContextWithEngine([[pd.NaT, "text"]])
        stage(ctx)
        # FakeEngineImpl's pd.isna(pd.NaT) returns True -> ""
        assert ctx.value == [["", "text"]]

    @pytest.mark.skipif(pd is None, reason="pandas not installed")
    def test_pd_timestamp(self):
        from xlwings.conversion.fast import FastCleanDataForWriteStage

        ts = pd.Timestamp("2024-01-15 10:30:00")
        stage = FastCleanDataForWriteStage({})
        ctx = FakeContextWithEngine([[ts, "text"]])
        stage(ctx)
        # FakeEngineImpl treats Timestamps as passthrough (not datetime subclass in newer pandas)
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
