###############################################################################
#
# FormulaResult - Explicitly typed cached results for formula cells.
#
# SPDX-License-Identifier: BSD-2-Clause
#
# Copyright (c) 2013-2026, John McNamara, jmcnamara@cpan.org
#

from __future__ import annotations

import math
from fractions import Fraction
from typing import Any, Union

from xlsxwriter.exceptions import XlsxInputError

# The set of Excel error codes that are valid cached formula results.
ERROR_CODES = frozenset(
    (
        "#NULL!",
        "#DIV/0!",
        "#VALUE!",
        "#REF!",
        "#NAME?",
        "#NUM!",
        "#N/A",
    )
)

# Supported cached result type tags.
NUMBER = "number"
STRING = "string"
BOOLEAN = "boolean"
ERROR = "error"
BLANK = "blank"


class FormulaResult:
    """
    An explicitly typed cached result for a formula cell.

    XlsxWriter doesn't calculate formula values. Some spreadsheet
    applications read the cached result stored in the file instead of
    recalculating the formula, so this class allows the caller to supply a
    result whose type is preserved exactly when the worksheet XML is written.

    The type is always taken from the factory method used to create the
    object, never inferred from the Python value. In particular the string
    ``"001"`` is written as a string, even though it could be converted to a
    number.

    Use the factory class methods rather than instantiating the class
    directly: :func:`number`, :func:`string`, :func:`boolean`,
    :func:`error` and :func:`blank`.

    """

    __slots__ = ("type", "value")

    def __init__(self, result_type: str, value: Any = None) -> None:
        self.type = result_type
        self.value = value

    def __repr__(self) -> str:
        return f"FormulaResult({self.type!r}, {self.value!r})"

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, FormulaResult):
            return NotImplemented
        return self.type == other.type and self.value == other.value

    @classmethod
    def number(cls, value: Union[int, float, Fraction]) -> "FormulaResult":
        """
        Create a numeric cached result.

        Args:
            value: An ``int``, ``float`` or ``Fraction``. Booleans are
                rejected because :func:`boolean` must be used instead.

        Returns:
            A ``FormulaResult`` of type ``number``.

        """
        if isinstance(value, bool) or not isinstance(
            value, (int, float, Fraction)
        ):
            raise XlsxInputError(
                "FormulaResult.number() requires an int, float or Fraction "
                f"value, got {type(value).__name__}"
            )

        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            raise XlsxInputError(
                "FormulaResult.number() doesn't support NaN or Infinity values"
            )

        return cls(NUMBER, value)

    @classmethod
    def string(cls, value: str) -> "FormulaResult":
        """
        Create a string cached result.

        The string is written with the formula string type (``t="str"``),
        even if it contains only digits, such as ``"001"``. An empty string
        is written as an empty ``<v></v>`` element, which is distinct from a
        :func:`blank` result that has no ``<v>`` element at all.

        Args:
            value: A ``str`` value.

        Returns:
            A ``FormulaResult`` of type ``string``.

        """
        if not isinstance(value, str):
            raise XlsxInputError(
                "FormulaResult.string() requires a str value, got "
                f"{type(value).__name__}"
            )

        return cls(STRING, value)

    @classmethod
    def boolean(cls, value: bool) -> "FormulaResult":
        """
        Create a boolean cached result.

        Args:
            value: A ``bool`` value.

        Returns:
            A ``FormulaResult`` of type ``boolean``.

        """
        if not isinstance(value, bool):
            raise XlsxInputError(
                "FormulaResult.boolean() requires a bool value, got "
                f"{type(value).__name__}"
            )

        return cls(BOOLEAN, value)

    @classmethod
    def error(cls, value: str) -> "FormulaResult":
        """
        Create an error cached result.

        Args:
            value: One of the allowed Excel error codes: ``#NULL!``,
                ``#DIV/0!``, ``#VALUE!``, ``#REF!``, ``#NAME?``,
                ``#NUM!`` or ``#N/A``.

        Returns:
            A ``FormulaResult`` of type ``error``.

        """
        if value not in ERROR_CODES:
            allowed = ", ".join(sorted(ERROR_CODES))
            raise XlsxInputError(
                f"FormulaResult.error() got invalid error code {value!r}. "
                f"Allowed error codes are: {allowed}"
            )

        return cls(ERROR, value)

    @classmethod
    def blank(cls) -> "FormulaResult":
        """
        Create a blank cached result.

        A blank result is written as a formula cell without a ``<v>``
        element, which is distinct from a :func:`string` result containing
        an empty string.

        Returns:
            A ``FormulaResult`` of type ``blank``.

        """
        return cls(BLANK, None)
