import datetime

from openpyxl.cell import Cell

import pytest

from openpyxl_evaluator import Evaluator


class TestEvaluator:
    def test_string(self, build_cell):
        cell = build_cell('Hello, World!')

        assert Evaluator(cell).value == 'Hello, World!'

    def test_number(self, build_cell):
        cell = build_cell(42)

        assert Evaluator(cell).value == 42

    def test_datevalue(self, build_cell):
        cell = build_cell('=DATEVALUE("2024-03-03")')

        assert Evaluator(cell).value == datetime.date(2024, 3, 3)

    @pytest.fixture
    def build_cell(self):
        def _build_cell(value):
            return Cell(None, None, None, value)

        return _build_cell
