###############################################################################
#
# Tests for XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import os
import re
import tempfile
import unittest
import warnings
import zipfile

from xlsxwriter.exceptions import OverlappingRange
from xlsxwriter.workbook import Workbook


def _temp_xlsx():
    fd, filename = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    return filename


def _zip_parts(filename):
    with zipfile.ZipFile(filename) as xlsx_file:
        return {name: xlsx_file.read(name) for name in xlsx_file.namelist()}


def _streaming_workbook(filename):
    return Workbook(filename, {"constant_memory": True, "in_memory": False})


class TestStreamTable(unittest.TestCase):
    """
    Test pre-declared tables in constant_memory mode.

    """

    def test_table_xml_matches_non_streaming(self):
        """Table XML and relationships match the normal mode output."""

        options = {
            "name": "SalesData",
            "style": "Table Style Medium 10",
            "banded_rows": True,
            "banded_columns": True,
            "first_column": True,
            "last_column": True,
            "total_row": True,
            "columns": [
                {"header": "Item", "total_string": "Totals"},
                {"header": "Qty", "total_function": "sum"},
                {"header": "Price", "total_function": "average",
                 "total_value": 5.5},
                {"header": "Note"},
            ],
        }
        data = [
            ("Apple", 3, 1.5, "a"),
            ("Pear", 5, 2.5, "b"),
            ("Plum", 7, 3.5, "c"),
        ]

        normal_path = _temp_xlsx()
        workbook = Workbook(normal_path)
        worksheet = workbook.add_worksheet()
        normal_options = dict(options)
        normal_options["data"] = [list(row) for row in data]
        worksheet.add_table("A1:D5", normal_options)
        workbook.close()
        normal = _zip_parts(normal_path)

        stream_path = _temp_xlsx()
        workbook = _streaming_workbook(stream_path)
        worksheet = workbook.add_worksheet()
        self.assertEqual(worksheet.add_table_stream("A1:D5", options), 0)
        for row_num, row_data in enumerate(data, start=1):
            worksheet.write_row(row_num, 0, row_data)
        workbook.close()
        stream = _zip_parts(stream_path)

        self.assertEqual(
            normal["xl/tables/table1.xml"], stream["xl/tables/table1.xml"]
        )
        self.assertEqual(
            normal["xl/worksheets/_rels/sheet1.xml.rels"],
            stream["xl/worksheets/_rels/sheet1.xml.rels"],
        )
        self.assertIn(
            b"spreadsheetml.table+xml", stream["[Content_Types].xml"]
        )
        # The totals row is emitted even though the caller never writes it.
        self.assertIn(b'r="5"', stream["xl/worksheets/sheet1.xml"])
        self.assertIn(
            b"SUBTOTAL(109,[Qty])", stream["xl/worksheets/sheet1.xml"]
        )

        for path in (normal_path, stream_path):
            os.remove(path)

    def test_multiple_tables(self):
        """Several streaming tables produce several table parts."""

        path = _temp_xlsx()
        workbook = _streaming_workbook(path)
        worksheet = workbook.add_worksheet()
        worksheet.add_table_stream(
            "A1:B4",
            {"name": "TableAlpha", "total_row": True,
             "columns": [{"header": "K", "total_string": "S"},
                         {"header": "V", "total_function": "sum"}]},
        )
        worksheet.add_table_stream(
            "D1:E4",
            {"name": "TableBeta", "total_row": True,
             "columns": [{"header": "P"},
                         {"header": "Q", "total_function": "max"}]},
        )
        for row_num in range(1, 4):
            worksheet.write_number(row_num, 0, row_num)
            worksheet.write_number(row_num, 1, row_num * 2)
            worksheet.write_number(row_num, 3, row_num)
            worksheet.write_number(row_num, 4, row_num * 3)
        workbook.close()

        parts = _zip_parts(path)
        self.assertIn("xl/tables/table1.xml", parts)
        self.assertIn("xl/tables/table2.xml", parts)
        self.assertEqual(
            parts["xl/worksheets/_rels/sheet1.xml.rels"].count(
                b"tables/table"
            ),
            2,
        )
        self.assertEqual(
            parts["[Content_Types].xml"].count(b"spreadsheetml.table+xml"), 2
        )
        os.remove(path)

    def test_forward_jump_still_writes_totals(self):
        """A totals row skipped by a forward jump is emitted exactly once."""

        path = _temp_xlsx()
        workbook = _streaming_workbook(path)
        worksheet = workbook.add_worksheet()
        worksheet.add_table_stream(
            "A1:C6",
            {"name": "JumpTable", "total_row": True,
             "columns": [{"header": "A", "total_string": "T"},
                         {"header": "B", "total_function": "sum"},
                         {"header": "C"}]},
        )
        worksheet.write_row(1, 0, (1, 2, 3))
        # Forward jump over rows 2-4 straight to a data row below them.
        worksheet.write_row(5, 0, (9, 9, 9))
        workbook.close()

        sheet = _zip_parts(path)["xl/worksheets/sheet1.xml"].decode()
        self.assertEqual(sheet.count('r="6"'), 1)
        self.assertIn("SUBTOTAL(109", sheet)
        os.remove(path)

    def test_forward_jump_preserves_order_and_incoming_row(self):
        """A jump over a totals row keeps ascending order and pending cells."""

        path = _temp_xlsx()
        workbook = _streaming_workbook(path)
        worksheet = workbook.add_worksheet()
        worksheet.add_table_stream(
            "A1:C6",
            {"name": "OrderTable", "total_row": True,
             "columns": [{"header": "A", "total_string": "T"},
                         {"header": "B", "total_function": "sum"},
                         {"header": "C"}]},
        )
        worksheet.write_row(1, 0, (1, 2, 3))
        # Assemble cells on the incoming row, then trigger a forward jump
        # that flushes the intermediate totals row. The pending cells must
        # survive the intermediate flushes.
        worksheet.write(7, 0, "keep-a")
        worksheet.write(7, 1, "keep-b")
        worksheet.write(7, 2, "keep-c")
        workbook.close()

        sheet = _zip_parts(path)["xl/worksheets/sheet1.xml"].decode()
        row_order = [
            int(num)
            for num in re.findall(r'<row r="(\d+)"', sheet)
        ]
        self.assertEqual(row_order, sorted(row_order))
        self.assertIn("keep-a", sheet)
        self.assertIn("keep-b", sheet)
        self.assertIn("keep-c", sheet)
        # The totals label and function are on the totals row (row 6).
        self.assertIn('<c r="A6" t="inlineStr"><is><t>T</t></is></c>', sheet)
        self.assertIn("SUBTOTAL(109,[B])", sheet)
        os.remove(path)

    def test_declaration_too_late_is_rejected(self):
        """A declaration starting on a flushed row returns -2."""

        path = _temp_xlsx()
        workbook = _streaming_workbook(path)
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, "first")
        worksheet.write(1, 0, "second")  # flushes row 0

        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            result = worksheet.add_table_stream("A1:C5")

        self.assertEqual(result, -2)
        self.assertTrue(
            any("declared before" in str(w.message) for w in caught)
        )
        self.assertEqual(worksheet.tables, [])
        self.assertEqual(worksheet.stream_table_ranges, [])
        workbook.fileclosed = True
        os.remove(path)

    def test_overlapping_declaration_is_atomic(self):
        """A failed overlap check leaves no partial table state."""

        path = _temp_xlsx()
        workbook = _streaming_workbook(path)
        worksheet = workbook.add_worksheet()
        self.assertEqual(
            worksheet.add_table_stream("A1:C5", {"name": "FirstTable"}), 0
        )

        with self.assertRaises(OverlappingRange):
            worksheet.add_table_stream("B2:D6", {"name": "SecondTable"})

        self.assertEqual(len(worksheet.stream_table_ranges), 1)
        self.assertEqual(len(worksheet.tables), 1)
        # A later, non-overlapping declaration still succeeds.
        self.assertEqual(
            worksheet.add_table_stream("E1:G5", {"name": "ThirdTable"}), 0
        )
        workbook.fileclosed = True
        os.remove(path)

    def test_merge_overlap_in_both_orders(self):
        """Merge ranges and streaming tables reject overlaps both ways."""

        path = _temp_xlsx()
        workbook = _streaming_workbook(path)
        worksheet = workbook.add_worksheet()
        worksheet.add_table_stream("A1:C5", {"name": "FirstMergeTable"})
        with self.assertRaises(OverlappingRange):
            worksheet.merge_range("B2:B3", "x")
        workbook.fileclosed = True
        os.remove(path)

        path = _temp_xlsx()
        workbook = _streaming_workbook(path)
        worksheet = workbook.add_worksheet()
        worksheet.merge_range("B2:B3", "x")
        with self.assertRaises(OverlappingRange):
            worksheet.add_table_stream("A1:C5", {"name": "SecondMergeTable"})
        workbook.fileclosed = True
        os.remove(path)

    def test_unsupported_options_leave_no_dangling_parts(self):
        """The data and column formula options are rejected cleanly."""

        path = _temp_xlsx()
        workbook = _streaming_workbook(path)
        worksheet = workbook.add_worksheet()
        with warnings.catch_warnings(record=True):
            warnings.simplefilter("always")
            self.assertEqual(
                worksheet.add_table_stream("A1:B3", {"data": [[1, 2]]}), -2
            )
            self.assertEqual(
                worksheet.add_table_stream(
                    "D1:E3",
                    {"columns": [
                        {"header": "D"},
                        {"header": "E", "formula": "=D1*2"},
                    ]},
                ),
                -2,
            )
        for row_num in range(3):
            worksheet.write(row_num, 0, row_num)
        workbook.close()

        parts = _zip_parts(path)
        self.assertFalse(any("tables/" in name for name in parts))
        self.assertNotIn(b"tableParts", parts["xl/worksheets/sheet1.xml"])
        os.remove(path)

    def test_only_supported_in_constant_memory(self):
        """add_table_stream() rejects normal (non optimized) workbooks."""

        from io import BytesIO

        workbook = Workbook(BytesIO(), {"in_memory": True})
        worksheet = workbook.add_worksheet()
        with warnings.catch_warnings(record=True):
            warnings.simplefilter("always")
            self.assertEqual(worksheet.add_table_stream("A1:B3"), -3)
        workbook.fileclosed = True

    def test_add_table_routes_to_streaming(self):
        """add_table() works in constant_memory via the streaming path."""

        path = _temp_xlsx()
        workbook = _streaming_workbook(path)
        worksheet = workbook.add_worksheet()
        self.assertEqual(
            worksheet.add_table("A1:C4", {"name": "RoutedTable"}), 0
        )
        for row_num in range(1, 4):
            worksheet.write_row(row_num, 0, (row_num, row_num, row_num))
        workbook.close()

        self.assertIn("xl/tables/table1.xml", _zip_parts(path))
        os.remove(path)
