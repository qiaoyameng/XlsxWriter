###############################################################################
#
# End-to-end XML tests for typed cached formula results.
#
# These tests only use zipfile and xml.etree.ElementTree to inspect the
# generated workbook, checking each cell's c/@t attribute, <f> element and
# <v> element independently.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import os
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree

from xlsxwriter import FormulaResult, Workbook
from xlsxwriter.exceptions import (
    FormulaRangeConflict,
    FormulaResultNotSupported,
    InvalidFormulaResult,
)

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _read_sheet_cells(path, index=1):
    """Return a dict of cell reference -> <c> Element for a worksheet."""
    with zipfile.ZipFile(path) as archive:
        xml_data = archive.read(f"xl/worksheets/sheet{index}.xml")

    root = ElementTree.fromstring(xml_data)
    cells = {}
    for cell in root.iter(f"{{{NS}}}c"):
        cells[cell.attrib["r"]] = cell
    return cells


def _assert_cell(
    testcase, cell, expected_type, expected_formula, expected_value
):
    """Assert the c/@t, <f> and <v> contents of a formula cell."""
    formula = cell.find(f"{{{NS}}}f")
    value = cell.find(f"{{{NS}}}v")

    testcase.assertEqual(cell.attrib.get("t"), expected_type)

    if expected_formula is None:
        testcase.assertIsNone(formula)
    else:
        testcase.assertIsNotNone(formula)
        testcase.assertEqual(formula.text, expected_formula)

    if expected_value is _NO_V:
        testcase.assertIsNone(value)
    else:
        testcase.assertIsNotNone(value)
        # ElementTree reports an empty <v></v> element as text=None.
        actual_value = value.text if value.text is not None else ""
        testcase.assertEqual(actual_value, expected_value)


class _NoV:
    """Sentinel meaning the cell must not contain a <v> element."""


_NO_V = _NoV()


class TestTypedFormulaXml(unittest.TestCase):
    """Verify cached result XML for single formula cells."""

    def setUp(self):
        fd, self.filename = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)

    def tearDown(self):
        if os.path.exists(self.filename):
            os.unlink(self.filename)

    def _workbook(self, constant_memory=False):
        options = {"constant_memory": True} if constant_memory else {}
        return Workbook(self.filename, options)

    def _check_single_formula_cells(self):
        cells = _read_sheet_cells(self.filename)

        # number: no t attribute, numeric <v>.
        _assert_cell(self, cells["A1"], None, "1+2", "3")

        # string: t="str" and "001" is not turned into a number.
        _assert_cell(self, cells["A2"], "str", '"001"', "001")

        # boolean: t="b" and 1/0 <v>.
        _assert_cell(self, cells["A3"], "b", "TRUE()", "1")
        _assert_cell(self, cells["A4"], "b", "FALSE()", "0")

        # error: t="e" and the error code in <v>.
        _assert_cell(self, cells["A5"], "e", "NA()", "#N/A")

        # blank cached result: formula present, no <v> at all.
        _assert_cell(self, cells["A6"], None, "1+1", _NO_V)

        # empty string cached result: t="str" with an empty <v></v>.
        empty = cells["A7"]
        _assert_cell(self, empty, "str", '""', "")
        value = empty.find(f"{{{NS}}}v")
        self.assertIsNotNone(value)
        self.assertEqual(value.text if value.text is not None else "", "")

        # A blank and an empty string must serialize differently.
        self.assertIsNone(
            cells["A6"].find(f"{{{NS}}}v"),
            "blank formula result must omit <v>",
        )
        self.assertIsNotNone(
            cells["A7"].find(f"{{{NS}}}v"),
            "empty string formula result must keep <v>",
        )

    def test_single_formula_cells(self):
        workbook = self._workbook()
        worksheet = workbook.add_worksheet()

        worksheet.write_typed_formula(0, 0, "=1+2", FormulaResult.number(3))
        worksheet.write_typed_formula(1, 0, '="001"', FormulaResult.string("001"))
        worksheet.write_typed_formula(2, 0, "=TRUE()", FormulaResult.boolean(True))
        worksheet.write_typed_formula(3, 0, "=FALSE()", FormulaResult.boolean(False))
        worksheet.write_typed_formula(4, 0, "=NA()", FormulaResult.error("#N/A"))
        worksheet.write_typed_formula(5, 0, "=1+1", FormulaResult.blank())
        worksheet.write_typed_formula(6, 0, '=""', FormulaResult.string(""))

        workbook.close()
        self._check_single_formula_cells()

    def test_single_formula_cells_constant_memory(self):
        workbook = self._workbook(constant_memory=True)
        worksheet = workbook.add_worksheet()

        worksheet.write_typed_formula(0, 0, "=1+2", FormulaResult.number(3))
        worksheet.write_typed_formula(1, 0, '="001"', FormulaResult.string("001"))
        worksheet.write_typed_formula(2, 0, "=TRUE()", FormulaResult.boolean(True))
        worksheet.write_typed_formula(3, 0, "=FALSE()", FormulaResult.boolean(False))
        worksheet.write_typed_formula(4, 0, "=NA()", FormulaResult.error("#N/A"))
        worksheet.write_typed_formula(5, 0, "=1+1", FormulaResult.blank())
        worksheet.write_typed_formula(6, 0, '=""', FormulaResult.string(""))

        workbook.close()
        self._check_single_formula_cells()

    def _write_array_workbook(self, workbook):
        worksheet = workbook.add_worksheet()
        matrix = [
            [FormulaResult.number(1), FormulaResult.string("001")],
            [FormulaResult.boolean(True), FormulaResult.error("#DIV/0!")],
            [FormulaResult.blank(), FormulaResult.string("")],
        ]
        worksheet.write_typed_array_formula(
            0, 0, 2, 1, "{=A1:B3}", matrix
        )
        return worksheet

    def _check_array_cells(self):
        cells = _read_sheet_cells(self.filename)

        # Anchor keeps the array formula, its range and its own result.
        anchor = cells["A1"]
        formula = anchor.find(f"{{{NS}}}f")
        self.assertEqual(formula.attrib["t"], "array")
        self.assertEqual(formula.attrib["ref"], "A1:B3")
        self.assertEqual(formula.text, "A1:B3")
        _assert_cell(self, anchor, None, "A1:B3", "1")

        # Dependent cells have no formula, only their typed cached results.
        _assert_cell(self, cells["B1"], "str", None, "001")
        _assert_cell(self, cells["A2"], "b", None, "1")
        _assert_cell(self, cells["B2"], "e", None, "#DIV/0!")

        # Blank dependent cell is omitted entirely; empty string keeps <v>.
        self.assertNotIn("A3", cells)
        _assert_cell(self, cells["B3"], "str", None, "")

    def test_array_formula_matrix(self):
        workbook = self._workbook()
        self._write_array_workbook(workbook)
        workbook.close()
        self._check_array_cells()

    def test_array_formula_matrix_constant_memory(self):
        workbook = self._workbook(constant_memory=True)
        self._write_array_workbook(workbook)
        workbook.close()
        self._check_array_cells()

    def test_dynamic_array_formula(self):
        workbook = self._workbook()
        worksheet = workbook.add_worksheet()
        worksheet.write_typed_dynamic_array_formula(
            0,
            0,
            1,
            1,
            "=SORT(A1:B2)",
            [
                [FormulaResult.number(1), FormulaResult.string("a")],
                [FormulaResult.number(2), FormulaResult.blank()],
            ],
        )
        workbook.close()

        cells = _read_sheet_cells(self.filename)
        anchor = cells["A1"]
        self.assertEqual(anchor.attrib.get("cm"), "1")
        formula = anchor.find(f"{{{NS}}}f")
        self.assertEqual(formula.attrib["t"], "array")
        self.assertEqual(formula.attrib["ref"], "A1:B2")
        _assert_cell(self, anchor, None, "_xlfn._xlws.SORT(A1:B2)", "1")
        _assert_cell(self, cells["B1"], "str", None, "a")
        _assert_cell(self, cells["A2"], None, None, "2")
        self.assertNotIn("B2", cells)


class TestTypedFormulaValidation(unittest.TestCase):
    """Validation and conflict checks must happen before sheet mutation."""

    def setUp(self):
        fd, self.filename = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)

    def tearDown(self):
        if os.path.exists(self.filename):
            os.unlink(self.filename)

    def test_rejects_non_formula_result(self):
        workbook = Workbook(self.filename)
        worksheet = workbook.add_worksheet()
        with self.assertRaises(InvalidFormulaResult):
            worksheet.write_typed_formula(0, 0, "=1+1", 3)
        workbook.close()

        # The rejected call must not have written anything.
        self.assertEqual(_read_sheet_cells(self.filename), {})

    def test_matrix_shape_mismatch(self):
        workbook = Workbook(self.filename)
        worksheet = workbook.add_worksheet()

        with self.assertRaises(InvalidFormulaResult):
            worksheet.write_typed_array_formula(
                0, 0, 1, 1, "=X", [[FormulaResult.number(1)]]
            )

        with self.assertRaises(InvalidFormulaResult):
            worksheet.write_typed_array_formula(
                0,
                0,
                0,
                1,
                "=X",
                [
                    [
                        FormulaResult.number(1),
                        FormulaResult.number(2),
                        FormulaResult.number(3),
                    ]
                ],
            )

        with self.assertRaises(InvalidFormulaResult):
            worksheet.write_typed_array_formula(
                0, 0, 0, 0, "=X", FormulaResult.number(1)
            )

        workbook.close()
        self.assertEqual(_read_sheet_cells(self.filename), {})

    def test_already_written_cell_conflict(self):
        workbook = Workbook(self.filename)
        worksheet = workbook.add_worksheet()
        worksheet.write_number(0, 1, 9)

        with self.assertRaises(FormulaRangeConflict):
            worksheet.write_typed_formula(0, 1, "=1", FormulaResult.number(1))

        with self.assertRaises(FormulaRangeConflict):
            worksheet.write_typed_array_formula(
                0,
                0,
                1,
                1,
                "=X",
                [
                    [FormulaResult.number(1), FormulaResult.number(2)],
                    [FormulaResult.number(3), FormulaResult.number(4)],
                ],
            )

    def test_overlapping_array_ranges(self):
        workbook = Workbook(self.filename)
        worksheet = workbook.add_worksheet()
        matrix = [
            [FormulaResult.number(1), FormulaResult.number(2)],
            [FormulaResult.number(3), FormulaResult.number(4)],
        ]
        worksheet.write_typed_array_formula(0, 0, 1, 1, "=X", matrix)

        with self.assertRaises(FormulaRangeConflict):
            worksheet.write_typed_array_formula(1, 1, 2, 2, "=Y", matrix)

    def test_constant_memory_rejects_backward_row(self):
        workbook = Workbook(self.filename, {"constant_memory": True})
        worksheet = workbook.add_worksheet()
        worksheet.write_number(3, 0, 1)

        with self.assertRaises(FormulaResultNotSupported):
            worksheet.write_typed_formula(1, 0, "=1", FormulaResult.number(1))

        with self.assertRaises(FormulaResultNotSupported):
            worksheet.write_typed_array_formula(
                1,
                0,
                2,
                0,
                "=X",
                [[FormulaResult.number(1)], [FormulaResult.number(2)]],
            )

        workbook.close()

    def test_constant_memory_pending_cell_overwrite(self):
        workbook = Workbook(self.filename, {"constant_memory": True})
        worksheet = workbook.add_worksheet()
        worksheet.write_typed_array_formula(
            0,
            0,
            2,
            0,
            "=X",
            [
                [FormulaResult.number(1)],
                [FormulaResult.number(2)],
                [FormulaResult.number(3)],
            ],
        )

        # A later user write to a dependent cell wins, matching the
        # overwrite behavior of the normal (non constant_memory) mode.
        worksheet.write_number(1, 0, 9)

        workbook.close()
        cells = _read_sheet_cells(self.filename)
        self.assertEqual(cells["A2"].find(f"{{{NS}}}v").text, "9")

    def test_constant_memory_pending_user_cell_merge(self):
        workbook = Workbook(self.filename, {"constant_memory": True})
        worksheet = workbook.add_worksheet()
        # Array occupies column B (index 1); a user cell is written to the
        # same pending row in column A, so the two must merge into one row.
        worksheet.write_typed_array_formula(
            0,
            1,
            1,
            2,
            "=X",
            [
                [FormulaResult.number(1), FormulaResult.number(2)],
                [FormulaResult.number(3), FormulaResult.number(4)],
            ],
        )

        worksheet.write_number(1, 0, 9)

        workbook.close()
        cells = _read_sheet_cells(self.filename)
        # The user cell A2 and the array dependent B2 coexist in one row.
        self.assertEqual(cells["A2"].find(f"{{{NS}}}v").text, "9")
        self.assertEqual(cells["B2"].find(f"{{{NS}}}v").text, "3")
        self.assertEqual(cells["C2"].find(f"{{{NS}}}v").text, "4")


if __name__ == "__main__":
    unittest.main()
