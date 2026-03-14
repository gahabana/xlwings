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
