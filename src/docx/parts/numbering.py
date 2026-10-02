"""|NumberingPart| and closely related objects."""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, cast

from ..opc.constants import CONTENT_TYPE as CT
from ..opc.packuri import PackURI
from ..opc.part import XmlPart
from ..oxml.parser import parse_xml
from ..shared import lazyproperty

if TYPE_CHECKING:
    from ..oxml.numbering import CT_Numbering
    from ..package import Package


class NumberingPart(XmlPart):
    """Proxy for the numbering.xml part containing numbering definitions for a document
    or glossary."""

    @classmethod
    def new(cls, package: Package) -> NumberingPart:
        """Newly created numbering part, containing only the root ``<w:numbering>`` element."""
        partname = PackURI("/word/numbering.xml")
        content_type = CT.WML_NUMBERING
        element = parse_xml(cls._default_numbering_xml())
        return cls(partname, content_type, element, package)

    @lazyproperty
    def numbering_definitions(self) -> _NumberingDefinitions:
        """The |_NumberingDefinitions| instance containing the numbering definitions
        (<w:num> element proxies) for this numbering part."""
        return _NumberingDefinitions(cast("CT_Numbering", self._element))

    @classmethod
    def _default_numbering_xml(cls) -> bytes:
        """A byte-string containing XML for a default (empty) numbering part."""
        path = os.path.join(os.path.split(__file__)[0], "..", "templates", "default-numbering.xml")
        with open(path, "rb") as f:
            return f.read()


class _NumberingDefinitions:
    """Collection of |_NumberingDefinition| instances corresponding to the ``<w:num>``
    elements in a numbering part."""

    def __init__(self, numbering_elm: CT_Numbering) -> None:
        super().__init__()
        self._numbering = numbering_elm

    def __len__(self) -> int:
        return len(self._numbering.num_lst)
