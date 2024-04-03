import datetime

from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

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

    def test_reference(self, build_cell, worksheet):
        worksheet['A1'] = 42
        cell = build_cell('=A1')

        assert Evaluator(cell).value == 42

    def test_range(self, build_cell, worksheet):
        worksheet['A1'] = 42
        worksheet['A2'] = 43
        worksheet['B1'] = 44
        cell = build_cell('=A1:B2')

        assert Evaluator(cell).value == ((42, 44), (43, None))

    @pytest.fixture
    def build_cell(self, worksheet):
        def _build_cell(value):
            return Cell(worksheet, None, None, value)

        return _build_cell

    @pytest.fixture
    def worksheet(self):
        return Worksheet(None, 'Sheet1')
