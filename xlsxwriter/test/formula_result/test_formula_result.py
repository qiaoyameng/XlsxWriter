###############################################################################
#
# Tests for the FormulaResult typed cached formula result class.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import math
import unittest
from fractions import Fraction

from xlsxwriter.exceptions import XlsxInputError
from xlsxwriter.formula_result import (
    BLANK,
    BOOLEAN,
    ERROR,
    NUMBER,
    STRING,
    FormulaResult,
)


class TestFormulaResult(unittest.TestCase):
    """Test FormulaResult factory methods and validation."""

    def test_number_types(self):
        for value in (1, 2.5, Fraction(1, 2)):
            result = FormulaResult.number(value)
            self.assertEqual(result.type, NUMBER)
            self.assertEqual(result.value, value)

    def test_number_rejects_bool(self):
        with self.assertRaises(XlsxInputError):
            FormulaResult.number(True)

    def test_number_rejects_string(self):
        # A numeric looking string must not be silently converted.
        with self.assertRaises(XlsxInputError):
            FormulaResult.number("3")

    def test_number_rejects_nan_inf(self):
        with self.assertRaises(XlsxInputError):
            FormulaResult.number(math.nan)
        with self.assertRaises(XlsxInputError):
            FormulaResult.number(math.inf)

    def test_string_keeps_type(self):
        # "001" stays a string even though float("001") works.
        result = FormulaResult.string("001")
        self.assertEqual(result.type, STRING)
        self.assertEqual(result.value, "001")
        self.assertTrue(float("001") == 1.0)

    def test_empty_string_is_a_string(self):
        result = FormulaResult.string("")
        self.assertEqual(result.type, STRING)
        self.assertEqual(result.value, "")

    def test_string_rejects_non_str(self):
        with self.assertRaises(XlsxInputError):
            FormulaResult.string(1)

    def test_boolean(self):
        self.assertEqual(FormulaResult.boolean(True).type, BOOLEAN)
        self.assertIs(FormulaResult.boolean(False).value, False)

    def test_boolean_rejects_int(self):
        with self.assertRaises(XlsxInputError):
            FormulaResult.boolean(1)

    def test_error_codes(self):
        for code in (
            "#NULL!",
            "#DIV/0!",
            "#VALUE!",
            "#REF!",
            "#NAME?",
            "#NUM!",
            "#N/A",
        ):
            result = FormulaResult.error(code)
            self.assertEqual(result.type, ERROR)
            self.assertEqual(result.value, code)

    def test_error_rejects_unknown_code(self):
        with self.assertRaises(XlsxInputError):
            FormulaResult.error("#BOGUS!")

    def test_blank(self):
        result = FormulaResult.blank()
        self.assertEqual(result.type, BLANK)
        self.assertIsNone(result.value)

    def test_equality(self):
        self.assertEqual(
            FormulaResult.string("001"), FormulaResult.string("001")
        )
        self.assertNotEqual(
            FormulaResult.string("1"), FormulaResult.number(1)
        )


if __name__ == "__main__":
    unittest.main()
