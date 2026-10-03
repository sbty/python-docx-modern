"""Custom element classes related to hyperlinks (CT_Hyperlink)."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.oxml.simpletypes import ST_OnOff, ST_String, XsdString
from docx.oxml.text.run import CT_R
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    ZeroOrMore,
)

if TYPE_CHECKING:
    from docx.oxml.text.pagebreak import CT_LastRenderedPageBreak


class CT_Hyperlink(BaseOxmlElement):
    """`<w:hyperlink>` element, containing the text and address for a hyperlink."""

    r_lst: list[CT_R]

    rId: str | None = OptionalAttribute("r:id", XsdString)  # type: ignore[assignment]
    anchor: str | None = OptionalAttribute(  # type: ignore[assignment]
        "w:anchor", ST_String
    )
    history: bool = OptionalAttribute(  # type: ignore[assignment]
        "w:history", ST_OnOff, default=True
    )

    r = ZeroOrMore("w:r")

    @property
    def lastRenderedPageBreaks(self) -> list[CT_LastRenderedPageBreak]:
        """All `w:lastRenderedPageBreak` descendants of this hyperlink."""
        return cast("list[CT_LastRenderedPageBreak]", self.xpath("./w:r/w:lastRenderedPageBreak"))

    @property
    def text(self) -> str:
        """The textual content of this hyperlink.

        `CT_Hyperlink` stores the hyperlink-text as one or more `w:r` children.
        """
        return "".join(r.text for r in self.xpath("w:r"))

    @text.setter
    def text(self, value: object) -> None:
        # -- read-only; the setter exists only because lxml's `_Element.text` is writable.
        # -- An override must stay writable and accept whatever the base accepts, so raise
        # -- exactly what a property without a setter raises --
        raise AttributeError(f"property 'text' of '{type(self).__name__}' object has no setter")
