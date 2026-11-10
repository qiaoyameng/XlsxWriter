###############################################################################
#
# ThreadedComments - Classes for writing Excel XLSX threaded comment files
# and the workbook person list.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c), 2026, John McNamara, jmcnamara@cpan.org
#

# Standard packages.
from datetime import datetime, timezone
from typing import Dict, List, Optional

# Package imports.
from xlsxwriter import xmlwriter
from xlsxwriter.utility import _preserve_whitespace, xl_rowcol_to_cell


# Namespaces used by the threaded comment and person list parts.
XMLNS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2018/11/main"
XMLNS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XMLNS_XSD = "http://www.w3.org/2001/XMLSchema"
XMLNS_XSI = "http://www.w3.org/2001/XMLSchema-instance"


###########################################################################
#
# A person type class.
#
###########################################################################
class PersonType:
    """
    A class to represent a person who can author threaded comments.

    Persons are declared on the workbook and shared by every worksheet, so
    each distinct person is written once to the workbook person list.

    """

    def __init__(self, name: str, provider_id: str = "None", user_id: str = "") -> None:
        """
        Initialize a Person instance.

        Args:
            name:        The display name of the person.
            provider_id: The identity provider, such as "None" or "AD".
            user_id:     The optional user id within the provider.
        """
        self.name: str = name
        self.provider_id: str = provider_id
        self.user_id: str = user_id


###########################################################################
#
# A threaded comment type class.
#
###########################################################################
class ThreadedCommentType:
    """
    A class to represent a single comment in a cell threaded comment
    conversation.

    """

    def __init__(
        self,
        row: int,
        col: int,
        comment_id: str,
        text: str,
        person: PersonType,
        date_time: Optional[datetime] = None,
        parent_id: Optional[str] = None,
    ) -> None:
        """
        Initialize a ThreadedComment instance.

        Args:
            row:        The zero indexed row of the commented cell.
            col:        The zero indexed column of the commented cell.
            comment_id: The stable, application defined, unique comment id.
            text:       The comment text.
            person:     The PersonType author of the comment.
            date_time:  An optional timezone aware or naive UTC datetime.
            parent_id:  The optional id of the parent comment. Replies may
                        be submitted before their root comment.
        """
        self.row: int = row
        self.col: int = col
        self.id: str = comment_id
        self.text: str = text
        self.person: PersonType = person
        self.date_time: Optional[datetime] = date_time
        self.parent_id: Optional[str] = parent_id

        # The person id in the workbook person list, set during close.
        self.person_id: int = 0

        # The threaded comment part index used by the worksheet rels.
        self.file_index: int = 0

    @property
    def ref(self) -> str:
        # Return the A1 cell reference of the commented cell.
        return xl_rowcol_to_cell(self.row, self.col)

    @property
    def iso_date(self) -> Optional[str]:
        # Return the timestamp as a UTC xsd:dateTime string, or None.
        if self.date_time is None:
            return None

        date_time = self.date_time
        if date_time.tzinfo is not None:
            date_time = date_time.astimezone(timezone.utc).replace(tzinfo=None)

        return date_time.strftime("%Y-%m-%dT%H:%M:%SZ")


###########################################################################
#
# The file writer class for the personList part.
#
###########################################################################
class PersonList(xmlwriter.XMLwriter):
    """
    A class for writing the Excel XLSX personList.xml file.

    """

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(self, persons: Optional[List[PersonType]] = None) -> None:
        # Assemble and write the XML file.
        if persons is None:
            persons = []

        # Write the XML declaration.
        self._xml_declaration()

        # Write the personList element.
        self._write_person_list()

        for person_id, person in enumerate(persons):
            self._write_person(person, person_id)

        self._xml_end_tag("personList")

        # Close the file.
        self._xml_close()

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_person_list(self) -> None:
        # Write the <personList> element.
        attributes = [
            ("xmlns", XMLNS_MAIN),
            ("xmlns:r", XMLNS_R),
            ("xmlns:xsd", XMLNS_XSD),
            ("xmlns:xsi", XMLNS_XSI),
        ]

        self._xml_start_tag("personList", attributes)

    def _write_person(self, person: PersonType, person_id: int) -> None:
        # Write a <person> element.
        attributes = [
            ("displayName", person.name),
            ("id", str(person_id)),
            ("providerId", person.provider_id),
        ]

        if person.user_id:
            attributes.append(("userId", person.user_id))

        self._xml_empty_tag("person", attributes)


###########################################################################
#
# The file writer class for the threadedComments part.
#
###########################################################################
class ThreadedComments(xmlwriter.XMLwriter):
    """
    A class for writing an Excel XLSX threadedComments.xml file.

    """

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _assemble_xml_file(
        self, comments: Optional[List[ThreadedCommentType]] = None
    ) -> None:
        # Assemble and write the XML file.
        if comments is None:
            comments = []

        # Write the XML declaration.
        self._xml_declaration()

        # Write the threadedComments element.
        self._write_threaded_comments()

        # Group the comments by cell so every conversation stays together.
        cells: Dict[str, List[ThreadedCommentType]] = {}
        cell_order: List[str] = []
        for comment in comments:
            if comment.ref not in cells:
                cells[comment.ref] = []
                cell_order.append(comment.ref)
            cells[comment.ref].append(comment)

        for ref in cell_order:
            self._write_threaded_comment_cell(ref, cells[ref])

        self._xml_end_tag("threadedComments")

        # Close the file.
        self._xml_close()

    ###########################################################################
    #
    # XML methods.
    #
    ###########################################################################

    def _write_threaded_comments(self) -> None:
        # Write the <threadedComments> element.
        attributes = [
            ("xmlns", XMLNS_MAIN),
            ("xmlns:r", XMLNS_R),
            ("xmlns:xsd", XMLNS_XSD),
            ("xmlns:xsi", XMLNS_XSI),
        ]

        self._xml_start_tag("threadedComments", attributes)

    def _write_threaded_comment_cell(
        self, ref: str, comments: List[ThreadedCommentType]
    ) -> None:
        # Write a <threadedComment> block for one commented cell.
        self._xml_start_tag("threadedComment", [("ref", ref)])

        # Roots before replies so out of order submissions read as a thread.
        comments = sorted(
            comments,
            key=lambda item: (item.parent_id is not None, item.id),
        )

        for comment in comments:
            self._write_comment(comment)

        self._xml_end_tag("threadedComment")

    def _write_comment(self, comment: ThreadedCommentType) -> None:
        # Write an <comment> element.
        attributes = [
            ("ref", comment.ref),
            ("personId", str(comment.person_id)),
            ("id", comment.id),
        ]

        if comment.iso_date is not None:
            attributes.append(("dateTime", comment.iso_date))

        if comment.parent_id is not None:
            attributes.append(("parentId", comment.parent_id))

        self._xml_start_tag("comment", attributes)

        # Write the text element.
        self._xml_data_element("text", comment.text)

        self._xml_end_tag("comment")

    def _xml_data_element(self, tag: str, data: str, attributes: List = []) -> None:
        # Write a data element, preserving whitespace and newlines.
        if _preserve_whitespace(data) or "\n" in data:
            attributes = attributes + [("xml:space", "preserve")]

        super()._xml_data_element(tag, data, attributes)
