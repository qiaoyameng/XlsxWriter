###############################################################################
#
# Tests for the 'reproducible' workbook option.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2013-2026, John McNamara, jmcnamara@cpan.org
#

import hashlib
import io
import os
import tempfile
import unittest
import zipfile
from datetime import datetime
from xml.etree import ElementTree

from xlsxwriter.exceptions import XlsxInputError
from xlsxwriter.workbook import Workbook

CREATED = datetime(2024, 5, 17, 12, 0, 0)

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

NS = {"m": MAIN_NS, "r": REL_NS, "pr": PKG_REL_NS}


def _build_workbook(order, filename=None, change=None):
    # Build a logically identical workbook using two different creation
    # orders for the Format objects, the cell writes and the tables.
    if filename is not None:
        workbook = Workbook(filename, {"reproducible": True, "created": CREATED})
    else:
        filename = io.BytesIO()
        workbook = Workbook(
            filename, {"in_memory": True, "reproducible": True, "created": CREATED}
        )

    if order == 1:
        bold = workbook.add_format({"bold": True, "font_color": "red"})
        date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
        money = workbook.add_format({"num_format": "#,##0.00"})
    else:
        money = workbook.add_format({"num_format": "#,##0.00"})
        date_format = workbook.add_format({"num_format": "yyyy-mm-dd"})
        bold = workbook.add_format({"bold": True, "font_color": "red"})

    worksheet = workbook.add_worksheet("Data")
    other = workbook.add_worksheet("Second")

    cells = [
        (worksheet, 0, 0, "Name", bold),
        (worksheet, 0, 1, "Value", bold),
        (worksheet, 1, 0, "apple", None),
        (worksheet, 1, 1, 42, money),
        (worksheet, 2, 0, "banana", None),
        (worksheet, 2, 1, 3.5, money),
        (worksheet, 3, 0, "apple", None),
        (worksheet, 3, 1, datetime(2024, 1, 15), date_format),
        (other, 0, 0, "Name", bold),
        (other, 1, 0, "cherry", None),
    ]

    if change == "string":
        cells[2] = (worksheet, 1, 0, "apricot", None)
    elif change == "number":
        cells[3] = (worksheet, 1, 1, 43, money)

    if order == 2:
        cells = list(reversed(cells))

    for sheet, row, col, value, cell_format in cells:
        sheet.write(row, col, value, cell_format)

    worksheet.write_url(4, 0, "https://example.com", bold, "Example")
    worksheet.write_formula(5, 1, "=SUM(B2:B3)", money, 45.5)

    table_1 = {
        "columns": [{"header": "A"}, {"header": "B", "format": money}],
        "data": [["x", 1], ["y", 2]],
    }
    table_2 = {
        "columns": [{"header": "P"}, {"header": "Q", "format": bold}],
        "data": [["u", 3], ["v", 4]],
    }

    if order == 1:
        worksheet.add_table(7, 0, 9, 1, table_1)
        worksheet.add_table(7, 4, 9, 5, table_2)
    else:
        worksheet.add_table(7, 4, 9, 5, table_2)
        worksheet.add_table(7, 0, 9, 1, table_1)

    worksheet.conditional_format(
        1,
        1,
        2,
        1,
        {"type": "cell", "criteria": ">", "value": 10, "format": bold},
    )

    workbook.close()

    if isinstance(filename, io.BytesIO):
        return filename.getvalue()

    with open(filename, "rb") as file:
        return file.read()


def _sha256(data):
    return hashlib.sha256(data).hexdigest()


def _zip(data):
    return zipfile.ZipFile(io.BytesIO(data))


def _xml(zip_file, name):
    return ElementTree.fromstring(zip_file.read(name))


def _shared_strings(zip_file):
    root = _xml(zip_file, "xl/sharedStrings.xml")
    strings = []
    for si in root.findall("m:si", NS):
        strings.append("".join(t.text or "" for t in si.iter(f"{{{MAIN_NS}}}t")))
    return strings


def _rels_map(zip_file, name):
    root = _xml(zip_file, name)
    return {rel.get("Id"): rel for rel in root.findall("pr:Relationship", NS)}


class TestReproducibleOutput(unittest.TestCase):
    """Test that logically identical workbooks are byte identical."""

    def test_identical_sha256_for_different_creation_orders(self):
        data_1 = _build_workbook(order=1)
        data_2 = _build_workbook(order=2)

        self.assertEqual(_sha256(data_1), _sha256(data_2))

    def test_identical_sha256_in_file_mode(self):
        tmpdir = tempfile.mkdtemp()
        path_1 = os.path.join(tmpdir, "repro_1.xlsx")
        path_2 = os.path.join(tmpdir, "repro_2.xlsx")

        data_1 = _build_workbook(order=1, filename=path_1)
        data_2 = _build_workbook(order=2, filename=path_2)

        self.assertEqual(_sha256(data_1), _sha256(data_2))

    def test_zip_entries_have_fixed_timestamps_and_order(self):
        data_1 = _build_workbook(order=1)
        data_2 = _build_workbook(order=2)

        zip_1 = _zip(data_1)
        zip_2 = _zip(data_2)

        # The zip entries are in the same order.
        self.assertEqual(zip_1.namelist(), zip_2.namelist())

        # The zip entry timestamps are fixed.
        for info in zip_1.infolist():
            self.assertEqual(info.date_time, (1980, 1, 1, 0, 0, 0))

    def test_core_xml_uses_the_user_supplied_date(self):
        data = _build_workbook(order=1)
        core = _zip(data).read("docProps/core.xml").decode("utf-8")

        self.assertIn("2024-05-17T12:00:00Z", core)

    def test_data_change_changes_the_output(self):
        data_1 = _build_workbook(order=1)
        changed_string = _build_workbook(order=1, change="string")
        changed_number = _build_workbook(order=1, change="number")

        self.assertNotEqual(_sha256(data_1), _sha256(changed_string))
        self.assertNotEqual(_sha256(data_1), _sha256(changed_number))


class TestReproducibleReferences(unittest.TestCase):
    """Test that style, string and relationship references are valid."""

    def setUp(self):
        self.data = _build_workbook(order=1)
        self.zip_file = _zip(self.data)

    def tearDown(self):
        self.zip_file.close()

    def test_shared_string_references(self):
        strings = _shared_strings(self.zip_file)

        # The shared strings are stored in a sorted, content based order.
        self.assertEqual(strings, sorted(strings))

        expected = {
            "A1": "Name",
            "B1": "Value",
            "A2": "apple",
            "A3": "banana",
            "A4": "apple",
            "A5": "Example",
            "A8": "A",
            "B8": "B",
            "E8": "P",
            "F8": "Q",
        }

        root = _xml(self.zip_file, "xl/worksheets/sheet1.xml")

        for cell in root.iter(f"{{{MAIN_NS}}}c"):
            if cell.get("t") != "s":
                continue

            index = int(cell.find("m:v", NS).text)

            # The index refers to a valid shared string.
            self.assertLess(index, len(strings))

            # And to the expected string for this cell.
            if cell.get("r") in expected:
                self.assertEqual(strings[index], expected[cell.get("r")])

        # Duplicate strings share the same index.
        indices = {}
        for cell in root.iter(f"{{{MAIN_NS}}}c"):
            if cell.get("t") == "s" and cell.get("r") in ("A2", "A4"):
                indices[cell.get("r")] = int(cell.find("m:v", NS).text)

        self.assertEqual(indices["A2"], indices["A4"])

    def test_style_references(self):
        styles = _xml(self.zip_file, "xl/styles.xml")

        cell_xfs = styles.find("m:cellXfs", NS)
        xfs = cell_xfs.findall("m:xf", NS)
        fonts = styles.find("m:fonts", NS).findall("m:font", NS)
        num_fmts = {
            num_fmt.get("numFmtId"): num_fmt.get("formatCode")
            for num_fmt in styles.findall("m:numFmts/m:numFmt", NS)
        }

        sheet = _xml(self.zip_file, "xl/worksheets/sheet1.xml")
        cells = {c.get("r"): c for c in sheet.iter(f"{{{MAIN_NS}}}c")}

        # Every cell style index refers to a valid xf entry.
        for cell in cells.values():
            if cell.get("s") is not None:
                self.assertLess(int(cell.get("s")), len(xfs))

        # The bold header cell refers to an xf with a bold font.
        xf = xfs[int(cells["A1"].get("s"))]
        font = fonts[int(xf.get("fontId"))]
        self.assertIsNotNone(font.find("m:b", NS))

        # The date cell refers to an xf with the yyyy-mm-dd num format.
        xf = xfs[int(cells["B4"].get("s"))]
        self.assertEqual(num_fmts[xf.get("numFmtId")], "yyyy-mm-dd")

        # Cells with the same format share the same style index.
        self.assertEqual(cells["A1"].get("s"), cells["B1"].get("s"))

    def test_dxf_references(self):
        styles = _xml(self.zip_file, "xl/styles.xml")
        dxfs = styles.find("m:dxfs", NS).findall("m:dxf", NS)

        def dxf_is_bold(dxf):
            font = dxf.find("m:font", NS)
            return font is not None and font.find("m:b", NS) is not None

        def dxf_num_format(dxf):
            num_fmt = dxf.find("m:numFmt", NS)
            return num_fmt.get("formatCode") if num_fmt is not None else None

        # The conditional format refers to the bold dxf.
        sheet = _xml(self.zip_file, "xl/worksheets/sheet1.xml")
        cf_rule = sheet.find("m:conditionalFormatting/m:cfRule", NS)
        self.assertTrue(dxf_is_bold(dxfs[int(cf_rule.get("dxfId"))]))

        # The table column formats refer to the correct dxf entries.
        table_1 = _xml(self.zip_file, "xl/tables/table1.xml")
        columns = table_1.find("m:tableColumns", NS).findall("m:tableColumn", NS)
        self.assertEqual(
            dxf_num_format(dxfs[int(columns[1].get("dataDxfId"))]), "#,##0.00"
        )

        table_2 = _xml(self.zip_file, "xl/tables/table2.xml")
        columns = table_2.find("m:tableColumns", NS).findall("m:tableColumn", NS)
        self.assertTrue(dxf_is_bold(dxfs[int(columns[1].get("dataDxfId"))]))

    def test_relationship_references(self):
        names = set(self.zip_file.namelist())

        # The workbook relationships refer to existing files.
        workbook = _xml(self.zip_file, "xl/workbook.xml")
        workbook_rels = _rels_map(self.zip_file, "xl/_rels/workbook.xml.rels")

        for sheet in workbook.findall("m:sheets/m:sheet", NS):
            rel_id = sheet.get(f"{{{REL_NS}}}id")
            target = workbook_rels[rel_id].get("Target")
            self.assertIn("xl/" + target, names)

        # The hyperlink refers to the correct external target.
        sheet = _xml(self.zip_file, "xl/worksheets/sheet1.xml")
        sheet_rels = _rels_map(self.zip_file, "xl/worksheets/_rels/sheet1.xml.rels")

        hyperlink = sheet.find("m:hyperlinks/m:hyperlink", NS)
        self.assertEqual(hyperlink.get("ref"), "A5")
        rel_id = hyperlink.get(f"{{{REL_NS}}}id")
        rel = sheet_rels[rel_id]
        self.assertEqual(rel.get("Target"), "https://example.com")
        self.assertEqual(rel.get("TargetMode"), "External")

        # The table parts refer to existing table files in range order.
        table_parts = sheet.findall("m:tableParts/m:tablePart", NS)
        self.assertEqual(len(table_parts), 2)

        tables = []
        for table_part in table_parts:
            rel_id = table_part.get(f"{{{REL_NS}}}id")
            target = sheet_rels[rel_id].get("Target")
            table_name = "xl/tables/" + target.split("/")[-1]
            self.assertIn(table_name, names)
            tables.append(_xml(self.zip_file, table_name))

        # Tables are numbered in range order, not creation order.
        self.assertEqual(tables[0].get("ref"), "A8:B10")
        self.assertEqual(tables[1].get("ref"), "E8:F10")

        # The default table names match the deterministic table ids.
        self.assertEqual(tables[0].get("name"), "Table1")
        self.assertEqual(tables[1].get("name"), "Table2")


class TestReproducibleUnsupported(unittest.TestCase):
    """Test that unsupported features are reported before close."""

    def _workbook(self, **options):
        options.setdefault("in_memory", True)
        options["reproducible"] = True
        options["created"] = CREATED
        return Workbook(io.BytesIO(), options)

    def test_missing_created_date(self):
        workbook = Workbook(io.BytesIO(), {"in_memory": True, "reproducible": True})
        workbook.add_worksheet()

        with self.assertRaises(XlsxInputError) as context:
            workbook.close()

        self.assertIn("creation date", str(context.exception))

    def test_created_date_via_set_properties(self):
        workbook = self._workbook()
        workbook.set_properties({"created": CREATED, "author": "Test"})
        workbook.add_worksheet().write(0, 0, "Hello")
        workbook.close()

    def test_charts_are_rejected(self):
        workbook = self._workbook()
        worksheet = workbook.add_worksheet()
        worksheet.write_column(0, 0, [1, 2, 3])
        chart = workbook.add_chart({"type": "column"})
        chart.add_series({"values": "=Sheet1!$A$1:$A$3"})
        worksheet.insert_chart(0, 2, chart)

        with self.assertRaises(XlsxInputError) as context:
            workbook.close()

        self.assertIn("charts", str(context.exception))

    def test_comments_are_rejected(self):
        workbook = self._workbook()
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, "Hello")
        worksheet.write_comment(0, 0, "A comment")

        with self.assertRaises(XlsxInputError) as context:
            workbook.close()

        self.assertIn("comments", str(context.exception))

    def test_constant_memory_is_rejected(self):
        workbook = self._workbook(in_memory=False, constant_memory=True)
        workbook.add_worksheet().write(0, 0, "Hello")

        with self.assertRaises(XlsxInputError) as context:
            workbook.close()

        self.assertIn("constant_memory", str(context.exception))


class TestCreatedOption(unittest.TestCase):
    """Test the 'created' option outside of 'reproducible' mode."""

    def test_created_option_without_reproducible(self):
        buffer = io.BytesIO()
        workbook = Workbook(buffer, {"in_memory": True, "created": CREATED})
        workbook.add_worksheet().write(0, 0, "Hello")
        workbook.close()

        core = zipfile.ZipFile(buffer).read("docProps/core.xml").decode("utf-8")
        self.assertIn("2024-05-17T12:00:00Z", core)
