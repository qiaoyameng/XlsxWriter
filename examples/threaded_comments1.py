###############################################################################
#
# An example of writing threaded comments to a worksheet using Python and
# XlsxWriter.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2026, John McNamara, jmcnamara@cpan.org
#
from datetime import datetime, timezone

import xlsxwriter

workbook = xlsxwriter.Workbook("threaded_comments1.xlsx")
worksheet = workbook.add_worksheet()

# Persons are declared on the workbook and shared across worksheets.
alice = workbook.add_person("Alice")
bob = workbook.add_person("Bob")

worksheet.write("A1", "Hello")
worksheet.add_threaded_comment(
    "A1",
    "comment-1",
    alice,
    "First line of the discussion\nSecond line",
    datetime(2026, 11, 10, 8, 30, 0, tzinfo=timezone.utc),
)

# Replies may be submitted before their root comment.
worksheet.add_threaded_reply(
    "A1",
    "comment-2",
    bob,
    "A reply from Bob",
    parent_id="comment-1",
)

workbook.close()
