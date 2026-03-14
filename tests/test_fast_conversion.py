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


class TestEndToEndEquivalence:
    """Run same data through old and new full pipelines, assert identical results."""

    def _run_read_pipeline(self, data, use_fast, datetime_builder=dt.datetime,
                           empty_as=None, number_builder=None, err_to_str=False):
        """Simulate CleanDataFromRead stage."""
        from xlwings.conversion.fast import fast_clean_value_data

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
                    if datetime_builder is not dt.datetime:
                        return datetime_builder(
                            year=value.year, month=value.month, day=value.day,
                            hour=getattr(value, "hour", 0),
                            minute=getattr(value, "minute", 0),
                            second=getattr(value, "second", 0),
                            microsecond=getattr(value, "microsecond", 0),
                            tzinfo=None,
                        )
                    return value
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

    def test_large_mixed_data_read_with_number_builder(self):
        """1000-row dataset with number_builder, old vs new must match."""
        import random
        random.seed(123)

        rows = []
        for _ in range(1000):
            row = [random.random() * 100, random.choice(["text", "", None])]
            rows.append(row)

        old_result = self._run_read_pipeline(
            [r[:] for r in rows], use_fast=False,
            number_builder=lambda x: int(round(x)),
        )
        new_result = self._run_read_pipeline(
            [r[:] for r in rows], use_fast=True,
            number_builder=lambda x: int(round(x)),
        )
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
