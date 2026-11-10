###############################################################################
#
# Tests for the XlsxWriter 'format_snapshot' Workbook option.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import re
import os
import tempfile
import unittest
import zipfile
from io import BytesIO

from xlsxwriter.workbook import Workbook


def _make_workbook(constant_memory=False):
    # Build a workbook where a single Format object is modified between
    # writes, with 'format_snapshot' mode enabled.
    if constant_memory:
        handle, output = tempfile.mkstemp(suffix=".xlsx")
        os.close(handle)
        options = {"constant_memory": True, "format_snapshot": True}
    else:
        output = BytesIO()
        options = {"in_memory": True, "format_snapshot": True}

    workbook = Workbook(output, options)
    worksheet = workbook.add_worksheet()

    cell_format = workbook.add_format()

    # First batch of cells/rows with a red font.
    cell_format.set_font_color("red")
    worksheet.write_string(0, 0, "early1", cell_format)
    worksheet.write_string(0, 1, "early2", cell_format)
    worksheet.set_row(1, None, cell_format)
    worksheet.write_number(1, 0, 123)

    # Second batch after changing the same Format object to blue.
    cell_format.set_font_color("blue")
    worksheet.write_string(2, 0, "late1", cell_format)
    worksheet.write_string(2, 1, "late2", cell_format)
    worksheet.set_column(3, 3, None, cell_format)

    # The caller's Format object must remain editable after snapshotting.
    cell_format.set_bold()
    worksheet.write_string(4, 0, "bold-blue", cell_format)

    # Two distinct Format objects with identical properties.
    format1 = workbook.add_format({"font_color": "green", "italic": True})
    format2 = workbook.add_format({"font_color": "green", "italic": True})
    worksheet.write_string(6, 0, "g1", format1)
    worksheet.write_string(6, 1, "g2", format2)

    # A conditional format (DXF) and a cell format (XF) with the same
    # properties must not be mixed.
    cond_format = workbook.add_format({"bg_color": "yellow"})
    worksheet.conditional_format("A8:B8", {"type": "blanks", "format": cond_format})
    worksheet.write_string(8, 0, "xf-yellow", cond_format)

    workbook.close()

    if constant_memory:
        with open(output, "rb") as fh:
            data = fh.read()
        os.remove(output)
        return data

    return output.getvalue()


def _cell_style(sheet_xml, ref):
    match = re.search(r'<c r="%s"[^>]*s="(\d+)"' % ref, sheet_xml)
    return int(match.group(1)) if match else None


class TestFormatSnapshot(unittest.TestCase):
    """
    Test the 'format_snapshot' Workbook option by unzipping the generated
    workbook and checking the cell style references and styles.xml.

    """

    def _check_snapshot_workbook(self, data):
        zip_file = zipfile.ZipFile(BytesIO(data))
        sheet = zip_file.read("xl/worksheets/sheet1.xml").decode()
        styles = zip_file.read("xl/styles.xml").decode()

        fonts = re.findall(r"<font>.*?</font>|<font/>", styles, re.DOTALL)
        fills = re.findall(r"<fill>.*?</fill>", styles, re.DOTALL)
        cell_xfs = re.findall(r"<xf\b[^>]*>", styles.split("<cellXfs")[1])
        dxfs = re.findall(r"<dxf[ >]", styles)

        def cell_font(index):
            font_id = int(re.search(r'fontId="(\d+)"', cell_xfs[index]).group(1))
            return fonts[font_id]

        def cell_fill(index):
            fill_id = int(re.search(r'fillId="(\d+)"', cell_xfs[index]).group(1))
            return fills[fill_id]

        style_a1 = _cell_style(sheet, "A1")
        style_b1 = _cell_style(sheet, "B1")
        style_a3 = _cell_style(sheet, "A3")
        style_b3 = _cell_style(sheet, "B3")
        style_a5 = _cell_style(sheet, "A5")
        style_a7 = _cell_style(sheet, "A7")
        style_b7 = _cell_style(sheet, "B7")
        style_a9 = _cell_style(sheet, "A9")

        # Cells written at the same time share a style; cells written
        # before and after the format change have different styles.
        self.assertEqual(style_a1, style_b1)
        self.assertEqual(style_a3, style_b3)
        self.assertNotEqual(style_a1, style_a3)

        # The early batch keeps the red font, the late batch is blue.
        self.assertIn('rgb="FFFF0000"', cell_font(style_a1))
        self.assertIn('rgb="FF0000FF"', cell_font(style_a3))

        # Editing the caller's Format after snapshotting creates a new
        # style for new cells but doesn't affect previously written cells.
        self.assertNotIn(style_a5, (style_a1, style_a3))
        self.assertIn("<b/>", cell_font(style_a5))
        self.assertNotIn("<b/>", cell_font(style_a3))

        # Two distinct Format objects with identical properties produce a
        # single equivalent style record.
        self.assertEqual(style_a7, style_b7)
        self.assertIn('rgb="FF008000"', cell_font(style_a7))

        # The conditional format produces a DXF record, separate from the
        # cell XF record with the same properties.
        self.assertEqual(len(dxfs), 1)
        self.assertIn('rgb="FFFFFF00"', cell_fill(style_a9))

        # The row format was snapshotted while red.
        match = re.search(r'<row r="2"[^>]*s="(\d+)"', sheet)
        self.assertIsNotNone(match)
        self.assertEqual(int(match.group(1)), style_a1)

        # A cell without an explicit format falls back to the row format.
        self.assertEqual(_cell_style(sheet, "A2"), style_a1)

        # The column format was snapshotted while blue.
        match = re.search(r"<col\b[^>]*/>", sheet)
        self.assertIsNotNone(match)
        style_col = int(re.search(r'style="(\d+)"', match.group(0)).group(1))
        self.assertEqual(style_col, style_a3)

    def test_format_snapshot(self):
        """Test format snapshots in the default mode."""
        self._check_snapshot_workbook(_make_workbook())

    def test_format_snapshot_constant_memory(self):
        """Test format snapshots in constant_memory mode."""
        self._check_snapshot_workbook(_make_workbook(constant_memory=True))

    def test_default_behavior_unchanged(self):
        """Test that the default behavior isn't changed without the option."""
        output = BytesIO()
        workbook = Workbook(output, {"in_memory": True})
        worksheet = workbook.add_worksheet()

        cell_format = workbook.add_format()
        cell_format.set_font_color("red")
        worksheet.write_string(0, 0, "early", cell_format)

        cell_format.set_font_color("blue")
        worksheet.write_string(1, 0, "late", cell_format)

        workbook.close()

        zip_file = zipfile.ZipFile(output)
        sheet = zip_file.read("xl/worksheets/sheet1.xml").decode()

        # Without 'format_snapshot' both cells reference the same style.
        self.assertEqual(_cell_style(sheet, "A1"), _cell_style(sheet, "A2"))
