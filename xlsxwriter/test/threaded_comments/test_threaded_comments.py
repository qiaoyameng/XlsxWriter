###############################################################################
#
# Tests for XlsxWriter threaded comments.
#
# These tests unzip the generated XLSX and parse the XML directly, without
# relying on a local copy of Excel.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2026, John McNamara, jmcnamara@cpan.org
#

import unittest
import zipfile
from datetime import datetime, timedelta, timezone
from io import BytesIO
from xml.etree import ElementTree as ET

from xlsxwriter.exceptions import (
    DuplicateThreadedCommentId,
    ThreadedCommentConflict,
    ThreadedCommentParentError,
)
from xlsxwriter.workbook import Workbook


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2018/11/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _close_workbook(build):
    # Build a workbook into an in-memory XLSX and return its zip entries.
    buffer = BytesIO()
    workbook = Workbook(buffer, {"in_memory": True})
    build(workbook)
    workbook.close()

    return zipfile.ZipFile(BytesIO(buffer.getvalue()))


def _parse(xlsx_file, name):
    return ET.fromstring(xlsx_file.read(name))


def _comments(root):
    return root.findall(f".//{{{MAIN_NS}}}comment")


def _comment_ids(root):
    return [comment.get("id") for comment in _comments(root)]


class TestThreadedCommentParts(unittest.TestCase):
    """
    Test the threaded comment, person list, relationship and content type
    parts written to the XLSX container.

    """

    def test_person_is_shared_across_worksheets(self):
        """The same person produces one person record across two sheets."""

        def build(workbook):
            person = workbook.add_person("Alice")
            other = workbook.add_person("Bob")

            sheet1 = workbook.add_worksheet()
            sheet2 = workbook.add_worksheet()

            sheet1.add_threaded_comment(0, 0, "s1-root", person, "first")
            sheet2.add_threaded_comment(0, 0, "s2-root", person, "again")
            sheet2.add_threaded_comment(1, 0, "s2-bob", other, "bob")

        xlsx_file = _close_workbook(build)

        person_root = _parse(xlsx_file, "xl/persons/person.xml")
        persons = person_root.findall(f"{{{MAIN_NS}}}person")

        self.assertEqual(
            [person.get("displayName") for person in persons],
            ["Alice", "Bob"],
        )
        self.assertEqual([person.get("id") for person in persons], ["0", "1"])

        # Both sheets reference the same personId for Alice.
        sheet1 = _parse(
            xlsx_file, "xl/threadedComments/threadedComment1.xml"
        )
        sheet2 = _parse(
            xlsx_file, "xl/threadedComments/threadedComment2.xml"
        )
        alice_ids = {
            comment.get("personId")
            for comment in _comments(sheet1) + _comments(sheet2)
            if comment.get("id") in ("s1-root", "s2-root")
        }
        self.assertEqual(alice_ids, {"0"})

    def test_out_of_order_reply_is_grouped_under_root(self):
        """A reply submitted before its root is placed after it in the part."""

        def build(workbook):
            person = workbook.add_person("Alice")
            sheet = workbook.add_worksheet()

            # Reply first, root comment afterwards.
            sheet.add_threaded_reply(
                2, 2, "reply-1", person, "early reply", parent_id="root-1"
            )
            sheet.add_threaded_comment(2, 2, "root-1", person, "root text")

        xlsx_file = _close_workbook(build)

        root = _parse(
            xlsx_file, "xl/threadedComments/threadedComment1.xml"
        )
        cells = root.findall(f"{{{MAIN_NS}}}threadedComment")
        self.assertEqual(len(cells), 1)
        self.assertEqual(cells[0].get("ref"), "C3")

        comments = _comments(cells[0])
        self.assertEqual(_comment_ids(cells[0]), ["root-1", "reply-1"])
        self.assertEqual(comments[1].get("parentId"), "root-1")
        self.assertEqual(
            comments[1].find(f"{{{MAIN_NS}}}text").text, "early reply"
        )

    def test_content_types_and_relationships(self):
        """Content_Types and rels reference every threaded comment part."""

        def build(workbook):
            person = workbook.add_person("Alice")
            workbook.add_worksheet().add_threaded_comment(
                0, 0, "c1", person, "text"
            )

        xlsx_file = _close_workbook(build)

        content_types = _parse(xlsx_file, "[Content_Types].xml")
        overrides = {
            node.get("PartName"): node.get("ContentType")
            for node in content_types.findall(f"{{{CT_NS}}}Override")
        }
        self.assertEqual(
            overrides["/xl/threadedComments/threadedComment1.xml"],
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.threadedComment+xml",
        )
        self.assertEqual(
            overrides["/xl/persons/person.xml"],
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.person+xml",
        )

        sheet_rels = _parse(
            xlsx_file, "xl/worksheets/_rels/sheet1.xml.rels"
        )
        targets = {
            node.get("Type").rsplit("/", 1)[-1]: node.get("Target")
            for node in sheet_rels.findall(f"{{{REL_NS}}}Relationship")
        }
        self.assertEqual(
            targets["threadedComment"],
            "../threadedComments/threadedComment1.xml",
        )

        workbook_rels = _parse(xlsx_file, "xl/_rels/workbook.xml.rels")
        person_targets = [
            node.get("Target")
            for node in workbook_rels.findall(f"{{{REL_NS}}}Relationship")
            if node.get("Type").endswith("/person")
        ]
        self.assertEqual(person_targets, ["persons/person.xml"])

    def test_newline_unicode_and_utc_round_trip(self):
        """Newlines, non-ASCII text and UTC timestamps survive the XML."""

        def build(workbook):
            person = workbook.add_person("小茗")
            sheet = workbook.add_worksheet()
            sheet.add_threaded_comment(
                0,
                0,
                "c1",
                person,
                "第一行\n第二行 🎉",
                datetime(2026, 11, 10, 8, 30, 0, tzinfo=timezone.utc),
            )
            sheet.add_threaded_comment(
                1,
                0,
                "c2",
                person,
                "naive",
                datetime(2026, 11, 10, 16, 30, 0,
                         tzinfo=timezone(timedelta(hours=8))),
            )

        xlsx_file = _close_workbook(build)
        root = _parse(
            xlsx_file, "xl/threadedComments/threadedComment1.xml"
        )
        comments = {
            comment.get("id"): comment for comment in _comments(root)
        }

        self.assertEqual(
            comments["c1"].find(f"{{{MAIN_NS}}}text").text,
            "第一行\n第二行 🎉",
        )
        self.assertEqual(
            comments["c1"].get("dateTime"), "2026-11-10T08:30:00Z"
        )
        # +08:00 is converted to UTC on serialization.
        self.assertEqual(
            comments["c2"].get("dateTime"), "2026-11-10T08:30:00Z"
        )


class TestThreadedCommentValidation(unittest.TestCase):
    """
    Test that invalid threaded comment relationships fail deterministically.

    """

    def test_duplicate_id_is_rejected(self):
        """Comment ids must be unique within a worksheet."""

        def build(workbook):
            person = workbook.add_person("Alice")
            sheet = workbook.add_worksheet()
            sheet.add_threaded_comment(0, 0, "dup", person, "first")
            sheet.add_threaded_comment(1, 0, "dup", person, "second")

        with self.assertRaises(DuplicateThreadedCommentId):
            _close_workbook(build)

    def test_missing_parent_is_rejected(self):
        """A reply whose parent id does not exist fails at close."""

        def build(workbook):
            person = workbook.add_person("Alice")
            sheet = workbook.add_worksheet()
            sheet.add_threaded_reply(
                0, 0, "reply", person, "text", parent_id="missing"
            )

        with self.assertRaises(ThreadedCommentParentError):
            _close_workbook(build)

    def test_parent_cycle_is_rejected(self):
        """A parent chain that loops back on itself fails at close."""

        def build(workbook):
            person = workbook.add_person("Alice")
            sheet = workbook.add_worksheet()
            sheet.add_threaded_reply(
                0, 0, "a", person, "one", parent_id="b"
            )
            sheet.add_threaded_reply(
                0, 0, "b", person, "two", parent_id="a"
            )

        with self.assertRaises(ThreadedCommentParentError):
            _close_workbook(build)

    def test_parent_in_another_cell_is_rejected(self):
        """Replies must stay inside the same cell conversation."""

        def build(workbook):
            person = workbook.add_person("Alice")
            sheet = workbook.add_worksheet()
            sheet.add_threaded_comment(0, 0, "root", person, "root")
            sheet.add_threaded_reply(
                1, 1, "reply", person, "reply", parent_id="root"
            )

        with self.assertRaises(ThreadedCommentParentError):
            _close_workbook(build)

    def test_legacy_and_threaded_comment_conflict(self):
        """The same cell cannot hold both comment representations."""

        def build_legacy_first(workbook):
            person = workbook.add_person("Alice")
            sheet = workbook.add_worksheet()
            sheet.write_comment(0, 0, "legacy")
            sheet.add_threaded_comment(0, 0, "t1", person, "threaded")

        def build_threaded_first(workbook):
            person = workbook.add_person("Alice")
            sheet = workbook.add_worksheet()
            sheet.add_threaded_comment(0, 0, "t1", person, "threaded")
            sheet.write_comment(0, 0, "legacy")

        with self.assertRaises(ThreadedCommentConflict):
            _close_workbook(build_legacy_first)

        with self.assertRaises(ThreadedCommentConflict):
            _close_workbook(build_threaded_first)


if __name__ == "__main__":
    unittest.main()
