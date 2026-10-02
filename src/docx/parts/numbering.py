"""|NumberingPart| and closely related objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from ..opc.part import XmlPart
from ..shared import lazyproperty

if TYPE_CHECKING:
    from ..oxml.numbering import CT_Numbering


class NumberingPart(XmlPart):
    """Proxy for the numbering.xml part containing numbering definitions for a document
    or glossary."""

    @classmethod
    def new(cls) -> NumberingPart:
        """Newly created numbering part, containing only the root ``<w:numbering>`` element."""
        raise NotImplementedError

    @lazyproperty
    def numbering_definitions(self) -> _NumberingDefinitions:
        """The |_NumberingDefinitions| instance containing the numbering definitions
        (<w:num> element proxies) for this numbering part."""
        return _NumberingDefinitions(cast("CT_Numbering", self._element))


class _NumberingDefinitions:
    """Collection of |_NumberingDefinition| instances corresponding to the ``<w:num>``
    elements in a numbering part."""

    def __init__(self, numbering_elm: CT_Numbering) -> None:
        super().__init__()
        self._numbering = numbering_elm

    def __len__(self) -> int:
        return len(self._numbering.num_lst)
