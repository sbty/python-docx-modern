"""Objects related to shapes.

A shape is a visual object that appears on the drawing layer of a document.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, cast

from docx.enum.shape import WD_INLINE_SHAPE
from docx.oxml.ns import nsmap
from docx.shared import Parented

if TYPE_CHECKING:
    from docx.oxml.document import CT_Body
    from docx.oxml.shape import CT_Inline, CT_Picture
    from docx.parts.story import StoryPart
    from docx.shared import Length


class InlineShapes(Parented):
    """Sequence of |InlineShape| instances, supporting len(), iteration, and indexed access."""

    def __init__(self, body_elm: CT_Body, parent: StoryPart) -> None:
        super().__init__(parent)
        self._body = body_elm

    def __getitem__(self, idx: int) -> InlineShape:
        """Provide indexed access, e.g. 'inline_shapes[idx]'."""
        try:
            inline = self._inline_lst[idx]
        except IndexError:
            msg = f"inline shape index [{idx:d}] out of range"
            raise IndexError(msg) from None

        return InlineShape(inline)

    def __iter__(self) -> Iterator[InlineShape]:
        return (InlineShape(inline) for inline in self._inline_lst)

    def __len__(self) -> int:
        return len(self._inline_lst)

    @property
    def _inline_lst(self) -> list[CT_Inline]:
        body = self._body
        xpath = "//w:p/w:r/w:drawing/wp:inline"
        inlines: list[CT_Inline] = body.xpath(xpath)
        return inlines


class InlineShape:
    """Proxy for an ``<wp:inline>`` element, representing the container for an inline
    graphical object."""

    def __init__(self, inline: CT_Inline) -> None:
        super().__init__()
        self._inline = inline

    @property
    def height(self) -> Length:
        """Read/write.

        The display height of this inline shape as an |Emu| instance.
        """
        return self._inline.extent.cy

    @height.setter
    def height(self, cy: Length) -> None:
        self._inline.extent.cy = cy
        pic = self._inline.graphic.graphicData.pic
        # -- only a picture has a `pic:pic` element, whose shape size mirrors the extent --
        if pic is not None:
            pic.spPr.cy = cy

    @property
    def type(self) -> WD_INLINE_SHAPE:
        """The type of this inline shape as a member of
        ``docx.enum.shape.WD_INLINE_SHAPE``, e.g. ``LINKED_PICTURE``.

        Read-only.
        """
        graphicData = self._inline.graphic.graphicData
        uri = graphicData.uri
        if uri == nsmap["pic"]:
            # -- a picture's `a:graphicData` always contains a `pic:pic` element --
            blip = cast("CT_Picture", graphicData.pic).blipFill.blip
            # -- a picture without an `a:blip` has no image data, so it is not linked --
            if blip is not None and blip.link is not None:
                return WD_INLINE_SHAPE.LINKED_PICTURE
            return WD_INLINE_SHAPE.PICTURE
        if uri == nsmap["c"]:
            return WD_INLINE_SHAPE.CHART
        if uri == nsmap["dgm"]:
            return WD_INLINE_SHAPE.SMART_ART
        return WD_INLINE_SHAPE.NOT_IMPLEMENTED

    @property
    def width(self) -> Length:
        """Read/write.

        The display width of this inline shape as an |Emu| instance.
        """
        return self._inline.extent.cx

    @width.setter
    def width(self, cx: Length) -> None:
        self._inline.extent.cx = cx
        pic = self._inline.graphic.graphicData.pic
        # -- only a picture has a `pic:pic` element, whose shape size mirrors the extent --
        if pic is not None:
            pic.spPr.cx = cx
