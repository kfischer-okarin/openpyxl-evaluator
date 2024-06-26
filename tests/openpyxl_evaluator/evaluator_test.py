import datetime

import pytest
from openpyxl.cell import Cell
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from openpyxl_evaluator import Evaluator


class TestEvaluator:
    def test_string(self, build_cell):
        cell = build_cell("Hello, World!")

        assert Evaluator(cell).value == "Hello, World!"

    def test_number(self, build_cell):
        cell = build_cell(42)

        assert Evaluator(cell).value == 42

    def test_date(self, build_cell):
        cell = build_cell(datetime.date(2024, 3, 3))

        assert Evaluator(cell).value == datetime.date(2024, 3, 3)

    def test_datevalue(self, build_cell):
        cell = build_cell('=DATEVALUE("2024-03-03")')

        assert Evaluator(cell).value == datetime.date(2024, 3, 3)

    def test_reference(self, build_cell, worksheet):
        worksheet["A1"] = 42
        cell = build_cell("=A1")

        assert Evaluator(cell).value == 42

    def test_range(self, build_cell, worksheet):
        worksheet["A1"] = 42
        worksheet["A2"] = 43
        worksheet["B1"] = 44
        cell = build_cell("=A1:B2")

        assert Evaluator(cell).value == ((42, 44), (43, None))

    def test_sum(self, build_cell, worksheet):
        worksheet["A1"] = 10
        worksheet["A2"] = 11
        worksheet["B1"] = 12
        cell = build_cell("=SUM(A1:B2)")

        assert Evaluator(cell).value == 33

    def test_iferror(self, build_cell, worksheet):
        worksheet["A1"] = 42
        cell = build_cell('=IFERROR(42/A1, "Error")')

        assert Evaluator(cell).value == 1

        worksheet["A1"] = 0

        assert Evaluator(cell).value == "Error"

    @pytest.mark.parametrize(("formula", "result"), [("=1+1", 2), ("=1-1", 0), ("=1*4", 4), ("=5.2/2", 2.6)])
    def test_infix_operators(self, build_cell, formula, result):
        cell = build_cell(formula)

        assert Evaluator(cell).value == result

    @pytest.fixture
    def build_cell(self, worksheet):
        def _build_cell(value):
            return Cell(worksheet, None, None, value)

        return _build_cell

    @pytest.fixture
    def worksheet(self, workbook):
        return Worksheet(workbook, "Sheet1")

    @pytest.fixture
    def workbook(self):
        return Workbook()
