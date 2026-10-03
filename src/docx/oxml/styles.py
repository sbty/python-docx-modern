"""Custom element classes related to the styles part."""

from __future__ import annotations

from collections.abc import Callable, Iterator
from typing import TYPE_CHECKING, cast

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import nsdecls
from docx.oxml.parser import parse_xml
from docx.oxml.simpletypes import ST_DecimalNumber, ST_OnOff, ST_String
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)


def styleId_from_name(name: str) -> str:
    """Return the style id corresponding to `name`, taking into account special-case
    names such as 'Heading 1'."""
    return {
        "caption": "Caption",
        "heading 1": "Heading1",
        "heading 2": "Heading2",
        "heading 3": "Heading3",
        "heading 4": "Heading4",
        "heading 5": "Heading5",
        "heading 6": "Heading6",
        "heading 7": "Heading7",
        "heading 8": "Heading8",
        "heading 9": "Heading9",
    }.get(name, name.replace(" ", ""))


class CT_LatentStyles(BaseOxmlElement):
    """`w:latentStyles` element, defining behavior defaults for latent styles and
    containing `w:lsdException` child elements that each override those defaults for a
    named latent style."""

    add_lsdException: Callable[[], CT_LsdException]
    lsdException_lst: list[CT_LsdException]

    lsdException = ZeroOrMore("w:lsdException", successors=())

    count: int | None = OptionalAttribute("w:count", ST_DecimalNumber)  # type: ignore[assignment]
    defLockedState: bool | None = OptionalAttribute("w:defLockedState", ST_OnOff)  # type: ignore[assignment]
    defQFormat: bool | None = OptionalAttribute("w:defQFormat", ST_OnOff)  # type: ignore[assignment]
    defSemiHidden: bool | None = OptionalAttribute("w:defSemiHidden", ST_OnOff)  # type: ignore[assignment]
    defUIPriority: int | None = OptionalAttribute("w:defUIPriority", ST_DecimalNumber)  # type: ignore[assignment]
    defUnhideWhenUsed: bool | None = OptionalAttribute("w:defUnhideWhenUsed", ST_OnOff)  # type: ignore[assignment]

    def bool_prop(self, attr_name: str) -> bool:
        """Return the boolean value of the attribute having `attr_name`, or |False| if
        not present."""
        value: bool | None = getattr(self, attr_name)
        if value is None:
            return False
        return value

    def get_by_name(self, name: str) -> CT_LsdException | None:
        """Return the `w:lsdException` child having `name`, or |None| if not found."""
        found: list[CT_LsdException] = self.xpath(f'w:lsdException[@w:name="{name}"]')
        if not found:
            return None
        return found[0]

    def set_bool_prop(self, attr_name: str, value: bool) -> None:
        """Set the on/off attribute having `attr_name` to `value`."""
        setattr(self, attr_name, bool(value))


class CT_LsdException(BaseOxmlElement):
    """``<w:lsdException>`` element, defining override visibility behaviors for a named
    latent style."""

    locked: bool | None = OptionalAttribute("w:locked", ST_OnOff)  # type: ignore[assignment]
    name: str = RequiredAttribute("w:name", ST_String)  # type: ignore[assignment]
    qFormat: bool | None = OptionalAttribute("w:qFormat", ST_OnOff)  # type: ignore[assignment]
    semiHidden: bool | None = OptionalAttribute("w:semiHidden", ST_OnOff)  # type: ignore[assignment]
    uiPriority: int | None = OptionalAttribute("w:uiPriority", ST_DecimalNumber)  # type: ignore[assignment]
    unhideWhenUsed: bool | None = OptionalAttribute("w:unhideWhenUsed", ST_OnOff)  # type: ignore[assignment]

    def delete(self) -> None:
        """Remove this `w:lsdException` element from the XML document."""
        cast("_Element", self.getparent()).remove(self)

    def on_off_prop(self, attr_name: str) -> bool | None:
        """Return the boolean value of the attribute having `attr_name`, or |None| if
        not present."""
        value: bool | None = getattr(self, attr_name)
        return value

    def set_on_off_prop(self, attr_name: str, value: bool | None) -> None:
        """Set the on/off attribute having `attr_name` to `value`."""
        setattr(self, attr_name, value)


if TYPE_CHECKING:
    from lxml.etree import _Element  # pyright: ignore[reportPrivateUsage]

    from docx.oxml.shared import CT_DecimalNumber, CT_OnOff, CT_String
    from docx.oxml.text.font import CT_RPr
    from docx.oxml.text.parfmt import CT_PPr


class CT_Style(BaseOxmlElement):
    """A ``<w:style>`` element, representing a style definition."""

    get_or_add_basedOn: Callable[[], CT_String]
    get_or_add_next: Callable[[], CT_String]
    get_or_add_pPr: Callable[[], CT_PPr]
    get_or_add_rPr: Callable[[], CT_RPr]
    _add_locked: Callable[[], CT_OnOff]
    _add_name: Callable[[], CT_String]
    _add_qFormat: Callable[[], CT_OnOff]
    _add_semiHidden: Callable[[], CT_OnOff]
    _add_uiPriority: Callable[[], CT_DecimalNumber]
    _add_unhideWhenUsed: Callable[[], CT_OnOff]
    _remove_basedOn: Callable[[], None]
    _remove_locked: Callable[[], None]
    _remove_name: Callable[[], None]
    _remove_next: Callable[[], None]
    _remove_qFormat: Callable[[], None]
    _remove_semiHidden: Callable[[], None]
    _remove_uiPriority: Callable[[], None]
    _remove_unhideWhenUsed: Callable[[], None]

    _tag_seq = (
        "w:name",
        "w:aliases",
        "w:basedOn",
        "w:next",
        "w:link",
        "w:autoRedefine",
        "w:hidden",
        "w:uiPriority",
        "w:semiHidden",
        "w:unhideWhenUsed",
        "w:qFormat",
        "w:locked",
        "w:personal",
        "w:personalCompose",
        "w:personalReply",
        "w:rsid",
        "w:pPr",
        "w:rPr",
        "w:tblPr",
        "w:trPr",
        "w:tcPr",
        "w:tblStylePr",
    )
    name: CT_String | None = ZeroOrOne(  # type: ignore[assignment]
        "w:name", successors=_tag_seq[1:]
    )
    basedOn: CT_String | None = ZeroOrOne(  # type: ignore[assignment]
        "w:basedOn", successors=_tag_seq[3:]
    )
    next: CT_String | None = ZeroOrOne(  # type: ignore[assignment]
        "w:next", successors=_tag_seq[4:]
    )
    uiPriority: CT_DecimalNumber | None = ZeroOrOne(  # type: ignore[assignment]
        "w:uiPriority", successors=_tag_seq[8:]
    )
    semiHidden: CT_OnOff | None = ZeroOrOne(  # type: ignore[assignment]
        "w:semiHidden", successors=_tag_seq[9:]
    )
    unhideWhenUsed: CT_OnOff | None = ZeroOrOne(  # type: ignore[assignment]
        "w:unhideWhenUsed", successors=_tag_seq[10:]
    )
    qFormat: CT_OnOff | None = ZeroOrOne(  # type: ignore[assignment]
        "w:qFormat", successors=_tag_seq[11:]
    )
    locked: CT_OnOff | None = ZeroOrOne(  # type: ignore[assignment]
        "w:locked", successors=_tag_seq[12:]
    )
    pPr: CT_PPr | None = ZeroOrOne(  # type: ignore[assignment]
        "w:pPr", successors=_tag_seq[17:]
    )
    rPr: CT_RPr | None = ZeroOrOne(  # type: ignore[assignment]
        "w:rPr", successors=_tag_seq[18:]
    )
    del _tag_seq

    type: WD_STYLE_TYPE | None = OptionalAttribute(  # type: ignore[assignment]
        "w:type", WD_STYLE_TYPE
    )
    styleId: str | None = OptionalAttribute(  # type: ignore[assignment]
        "w:styleId", ST_String
    )
    default: bool | None = OptionalAttribute("w:default", ST_OnOff)  # type: ignore[assignment]
    customStyle: bool | None = OptionalAttribute("w:customStyle", ST_OnOff)  # type: ignore[assignment]

    @property
    def basedOn_val(self) -> str | None:
        """Value of `w:basedOn/@w:val` or |None| if not present."""
        basedOn = self.basedOn
        if basedOn is None:
            return None
        return basedOn.val

    @basedOn_val.setter
    def basedOn_val(self, value: str | None) -> None:
        if value is None:
            self._remove_basedOn()
        else:
            self.get_or_add_basedOn().val = value

    @property
    def base_style(self) -> CT_Style | None:
        """Sibling CT_Style element this style is based on or |None| if no base style or
        base style not found."""
        basedOn = self.basedOn
        if basedOn is None:
            return None
        styles = cast("CT_Styles", self.getparent())
        base_style = styles.get_by_id(basedOn.val)
        if base_style is None:
            return None
        return base_style

    @property
    def effective_type(self) -> WD_STYLE_TYPE:
        """Style type of this `w:style`, `WD_STYLE_TYPE.PARAGRAPH` when `w:type` is absent.

        `w:type` is optional and the schema default is paragraph. `.type` cannot carry
        that default itself: assigning a descriptor's default removes the attribute, so new
        paragraph styles would be written without `w:type`.
        """
        type_ = self.type
        return WD_STYLE_TYPE.PARAGRAPH if type_ is None else type_

    def delete(self) -> None:
        """Remove this `w:style` element from its parent `w:styles` element."""
        cast("_Element", self.getparent()).remove(self)

    @property
    def locked_val(self) -> bool:
        """Value of `w:locked/@w:val` or |False| if not present."""
        locked = self.locked
        if locked is None:
            return False
        return locked.val

    @locked_val.setter
    def locked_val(self, value: bool) -> None:
        self._remove_locked()
        if bool(value) is True:
            locked = self._add_locked()
            locked.val = value

    @property
    def name_val(self) -> str | None:
        """Value of ``<w:name>`` child or |None| if not present."""
        name = self.name
        if name is None:
            return None
        return name.val

    @name_val.setter
    def name_val(self, value: str | None) -> None:
        self._remove_name()
        if value is not None:
            name = self._add_name()
            name.val = value

    @property
    def next_style(self) -> CT_Style | None:
        """Sibling CT_Style element identified by the value of `w:name/@w:val` or |None|
        if no value is present or no style with that style id is found."""
        next = self.next
        if next is None:
            return None
        styles = cast("CT_Styles", self.getparent())
        return styles.get_by_id(next.val)  # None if not found

    @property
    def qFormat_val(self) -> bool:
        """Value of `w:qFormat/@w:val` or |False| if not present."""
        qFormat = self.qFormat
        if qFormat is None:
            return False
        return qFormat.val

    @qFormat_val.setter
    def qFormat_val(self, value: bool) -> None:
        self._remove_qFormat()
        if bool(value):
            self._add_qFormat()

    @property
    def semiHidden_val(self) -> bool:
        """Value of ``<w:semiHidden>`` child or |False| if not present."""
        semiHidden = self.semiHidden
        if semiHidden is None:
            return False
        return semiHidden.val

    @semiHidden_val.setter
    def semiHidden_val(self, value: bool) -> None:
        self._remove_semiHidden()
        if bool(value) is True:
            semiHidden = self._add_semiHidden()
            semiHidden.val = value

    @property
    def uiPriority_val(self) -> int | None:
        """Value of ``<w:uiPriority>`` child or |None| if not present."""
        uiPriority = self.uiPriority
        if uiPriority is None:
            return None
        return uiPriority.val

    @uiPriority_val.setter
    def uiPriority_val(self, value: int | None) -> None:
        self._remove_uiPriority()
        if value is not None:
            uiPriority = self._add_uiPriority()
            uiPriority.val = value

    @property
    def unhideWhenUsed_val(self) -> bool:
        """Value of `w:unhideWhenUsed/@w:val` or |False| if not present."""
        unhideWhenUsed = self.unhideWhenUsed
        if unhideWhenUsed is None:
            return False
        return unhideWhenUsed.val

    @unhideWhenUsed_val.setter
    def unhideWhenUsed_val(self, value: bool) -> None:
        self._remove_unhideWhenUsed()
        if bool(value) is True:
            unhideWhenUsed = self._add_unhideWhenUsed()
            unhideWhenUsed.val = value


# -- built-in styles that comment markup refers to, as Word defines them; `w:basedOn` is
# -- added from the document's own default style of the same type --
_COMMENT_STYLES = (
    (
        "CommentReference",
        "annotation reference",
        WD_STYLE_TYPE.CHARACTER,
        f'<w:style {nsdecls("w")} w:type="character" w:styleId="CommentReference">'
        '<w:name w:val="annotation reference"/>'
        '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>'
        '<w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
        "</w:style>",
    ),
    (
        "CommentText",
        "annotation text",
        WD_STYLE_TYPE.PARAGRAPH,
        f'<w:style {nsdecls("w")} w:type="paragraph" w:styleId="CommentText">'
        '<w:name w:val="annotation text"/>'
        '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>'
        '<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        "</w:style>",
    ),
)


class CT_Styles(BaseOxmlElement):
    """``<w:styles>`` element, the root element of a styles part, i.e. styles.xml."""

    _tag_seq = ("w:docDefaults", "w:latentStyles", "w:style")
    add_style: Callable[[], CT_Style]
    get_or_add_latentStyles: Callable[[], CT_LatentStyles]
    style_lst: list[CT_Style]

    latentStyles: CT_LatentStyles | None = ZeroOrOne(  # type: ignore[assignment]
        "w:latentStyles", successors=_tag_seq[2:]
    )
    style = ZeroOrMore("w:style", successors=())
    del _tag_seq

    def add_style_of_type(self, name: str, style_type: WD_STYLE_TYPE, builtin: bool) -> CT_Style:
        """Return a newly added `w:style` element having `name` and `style_type`.

        `w:style/@customStyle` is set based on the value of `builtin`.
        """
        style = self.add_style()
        style.type = style_type
        style.customStyle = None if builtin else True
        style.styleId = styleId_from_name(name)
        style.name_val = name
        return style

    def ensure_comment_styles(self) -> None:
        """Add the built-in styles that comment markup refers to, when they are missing.

        Comment reference runs use the "CommentReference" character style and comment
        paragraphs use the "CommentText" paragraph style. The default template defines
        neither, so without this the references dangle and Word formats them with default
        properties.

        A style is not added when the document already defines it, either under its usual
        id or, as in localized templates, under another id with the same built-in name.
        Each added style is based on the document's default style of its type, whatever
        that style's id is.
        """
        for style_id, name, style_type, style_xml in _COMMENT_STYLES:
            if self.get_by_id(style_id) is not None or self.get_by_name(name) is not None:
                continue
            style = cast(CT_Style, parse_xml(style_xml))
            default_style = self.default_for(style_type)
            if default_style is not None:
                style.basedOn_val = default_style.styleId
            self.append(style)

    def default_for(self, style_type: WD_STYLE_TYPE) -> CT_Style | None:
        """Return `w:style[@w:type="*{style_type}*][-1]` or |None| if not found."""
        default_styles_for_type = [
            s for s in self._iter_styles() if s.effective_type == style_type and s.default
        ]
        if not default_styles_for_type:
            return None
        # spec calls for last default in document order
        return default_styles_for_type[-1]

    def get_by_id(self, styleId: str) -> CT_Style | None:
        """`w:style` child where @styleId = `styleId`.

        |None| if not found.
        """
        xpath = f'w:style[@w:styleId="{styleId}"]'
        return next(iter(self.xpath(xpath)), None)

    def get_by_name(self, name: str) -> CT_Style | None:
        """`w:style` child with `w:name` grandchild having value `name`.

        |None| if not found.
        """
        xpath = f'w:style[w:name/@w:val="{name}"]'
        return next(iter(self.xpath(xpath)), None)

    def _iter_styles(self) -> Iterator[CT_Style]:
        """Generate each of the `w:style` child elements in document order."""
        return (style for style in self.xpath("w:style"))
