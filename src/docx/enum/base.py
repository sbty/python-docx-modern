"""Base classes and other objects used by enumerations."""

from __future__ import annotations

import enum
from typing import TYPE_CHECKING, Any, TypeVar, cast

if TYPE_CHECKING:
    from typing_extensions import Self

_T = TypeVar("_T", bound="BaseXmlEnum")


class BaseEnum(int, enum.Enum):
    """Base class for Enums that do not map XML attr values.

    The enum's value will be an integer, corresponding to the integer assigned the
    corresponding member in the MS API enum of the same name.
    """

    def __new__(cls, ms_api_value: int, docstr: str) -> Self:
        self = int.__new__(cls, ms_api_value)
        self._value_ = ms_api_value
        self.__doc__ = docstr.strip()
        return self

    def __str__(self) -> str:
        """The symbolic name and string value of this member, e.g. 'MIDDLE (3)'."""
        return f"{self.name} ({self.value})"


class BaseXmlEnum(int, enum.Enum):
    """Base class for Enums that also map XML attr values.

    The enum's value will be an integer, corresponding to the integer assigned the
    corresponding member in the MS API enum of the same name.
    """

    xml_value: str | None

    def __new__(cls, ms_api_value: int, xml_value: str | None, docstr: str) -> Self:
        self = int.__new__(cls, ms_api_value)
        self._value_ = ms_api_value
        self.xml_value = xml_value
        self.__doc__ = docstr.strip()
        return self

    def __str__(self) -> str:
        """The symbolic name and string value of this member, e.g. 'MIDDLE (3)'."""
        return f"{self.name} ({self.value})"

    @classmethod
    def from_xml(cls, xml_value: str | None) -> Self:
        """Enumeration member corresponding to XML attribute value `xml_value`.

        Example::

            >>> WD_PARAGRAPH_ALIGNMENT.from_xml("center")
            WD_PARAGRAPH_ALIGNMENT.CENTER

        """
        member = next((member for member in cls if member.xml_value == xml_value), None)
        if member is None:
            raise ValueError(f"{cls.__name__} has no XML mapping for '{xml_value}'")
        return member

    @classmethod
    def to_xml(cls: type[_T], value: int | _T | None) -> str | None:
        """XML value of this enum member, generally an XML attribute value."""
        # -- presence of multi-arg `__new__()` method fools type-checker, but getting a
        # -- member by its value using EnumCls(val) works as usual.
        member: _T = cast(Any, cls)(value)
        # -- a local named `xml_value` makes mypy treat the enum's `xml_value` attribute as
        # -- final and reject its assignment in `__new__()`, so this local is named differently --
        member_xml_value = member.xml_value
        if not member_xml_value:
            raise ValueError(f"{cls.__name__}.{member.name} has no XML representation")
        return member_xml_value
