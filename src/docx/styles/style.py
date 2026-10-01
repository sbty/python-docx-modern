"""Style object hierarchy."""

from __future__ import annotations

from typing import cast

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.styles import CT_Style
from docx.shared import ElementProxy
from docx.styles import BabelFish
from docx.text.font import Font
from docx.text.parfmt import ParagraphFormat


def StyleFactory(style_elm: CT_Style) -> BaseStyle:
    """Return `Style` object of appropriate |BaseStyle| subclass for `style_elm`."""
    style_cls: type[BaseStyle] = {
        WD_STYLE_TYPE.PARAGRAPH: ParagraphStyle,
        WD_STYLE_TYPE.CHARACTER: CharacterStyle,
        WD_STYLE_TYPE.TABLE: _TableStyle,
        WD_STYLE_TYPE.LIST: _NumberingStyle,
        # -- a style without `w:type` raises KeyError here, as it always has --
    }[cast("WD_STYLE_TYPE", style_elm.type)]

    return style_cls(style_elm)


class BaseStyle(ElementProxy):
    """Base class for the various types of style object, paragraph, character, table,
    and numbering.

    These properties and methods are inherited by all style objects.
    """

    def __init__(self, style_elm: CT_Style) -> None:
        super().__init__(style_elm)
        self._style_elm = style_elm

    @property
    def _style_element(self) -> CT_Style:
        """The `w:style` element, typed (`._element` is None once deleted)."""
        return cast("CT_Style", self._element)

    @property
    def builtin(self) -> bool:
        """Read-only.

        |True| if this style is a built-in style. |False| indicates it is a custom
        (user-defined) style. Note this value is based on the presence of a
        `customStyle` attribute in the XML, not on specific knowledge of which styles
        are built into Word.
        """
        return not self._style_element.customStyle

    def delete(self) -> None:
        """Remove this style definition from the document.

        Note that calling this method does not remove or change the style applied to any
        document content. Content items having the deleted style will be rendered using
        the default style, as is any content with a style not defined in the document.
        """
        self._style_element.delete()
        # -- later access to most properties raises AttributeError, as before --
        # -- setattr rather than plain assignment: mypy rejects assigning None to the
        # -- inherited `_element` type, while pyright treats a `type: ignore` there as
        # -- unnecessary --
        setattr(self, "_element", None)  # noqa: B010

    @property
    def hidden(self) -> bool:
        """|True| if display of this style in the style gallery and list of recommended
        styles is suppressed.

        |False| otherwise. In order to be shown in the style gallery, this value must be
        |False| and :attr:`.quick_style` must be |True|.
        """
        return self._style_element.semiHidden_val

    @hidden.setter
    def hidden(self, value: bool) -> None:
        self._style_element.semiHidden_val = value

    @property
    def locked(self) -> bool:
        """Read/write Boolean.

        |True| if this style is locked. A locked style does not appear in the styles
        panel or the style gallery and cannot be applied to document content. This
        behavior is only active when formatting protection is turned on for the document
        (via the Developer menu).
        """
        return self._style_element.locked_val

    @locked.setter
    def locked(self, value: bool) -> None:
        self._style_element.locked_val = value

    @property
    def name(self) -> str | None:
        """The UI name of this style."""
        name = self._style_element.name_val
        if name is None:
            return None
        return BabelFish.internal2ui(name)

    @name.setter
    def name(self, value: str | None) -> None:
        self._style_element.name_val = value

    @property
    def priority(self) -> int | None:
        """The integer sort key governing display sequence of this style in the Word UI.

        |None| indicates no setting is defined, causing Word to use the default value of
        0. Style name is used as a secondary sort key to resolve ordering of styles
        having the same priority value.
        """
        return self._style_element.uiPriority_val

    @priority.setter
    def priority(self, value: int | None) -> None:
        self._style_element.uiPriority_val = value

    @property
    def quick_style(self) -> bool:
        """|True| if this style should be displayed in the style gallery when
        :attr:`.hidden` is |False|.

        Read/write Boolean.
        """
        return self._style_element.qFormat_val

    @quick_style.setter
    def quick_style(self, value: bool) -> None:
        self._style_element.qFormat_val = value

    @property
    def style_id(self) -> str:
        """The unique key name (string) for this style.

        This value is subject to rewriting by Word and should generally not be changed
        unless you are familiar with the internals involved.
        """
        # -- `w:styleId` is optional in the schema but present on every style in practice;
        # -- keep the public `str` type rather than widening it for that edge case --
        return cast(str, self._style_elm.styleId)

    @style_id.setter
    def style_id(self, value: str | None) -> None:
        self._style_element.styleId = value

    @property
    def type(self) -> WD_STYLE_TYPE:
        """Member of :ref:`WdStyleType` corresponding to the type of this style, e.g.
        ``WD_STYLE_TYPE.PARAGRAPH``."""
        type = self._style_elm.type
        if type is None:
            return WD_STYLE_TYPE.PARAGRAPH
        return type

    @property
    def unhide_when_used(self) -> bool:
        """|True| if an application should make this style visible the next time it is
        applied to content.

        False otherwise. Note that |docx| does not automatically unhide a style having
        |True| for this attribute when it is applied to content.
        """
        return self._style_element.unhideWhenUsed_val

    @unhide_when_used.setter
    def unhide_when_used(self, value: bool) -> None:
        self._style_element.unhideWhenUsed_val = value


class CharacterStyle(BaseStyle):
    """A character style.

    A character style is applied to a |Run| object and primarily provides character-
    level formatting via the |Font| object in its :attr:`.font` property.
    """

    @property
    def base_style(self) -> BaseStyle | None:
        """Style object this style inherits from or |None| if this style is not based on
        another style."""
        base_style = self._style_element.base_style
        if base_style is None:
            return None
        return StyleFactory(base_style)

    @base_style.setter
    def base_style(self, style: BaseStyle | None) -> None:
        style_id = style.style_id if style is not None else None
        self._style_element.basedOn_val = style_id

    @property
    def font(self) -> Font:
        """The |Font| object providing access to the character formatting properties for
        this style, such as font name and size."""
        return Font(self._style_element)


# -- just in case someone uses the old name in an extension function --
_CharacterStyle = CharacterStyle


class ParagraphStyle(CharacterStyle):
    """A paragraph style.

    A paragraph style provides both character formatting and paragraph formatting such
    as indentation and line-spacing.
    """

    def __repr__(self) -> str:
        return f"_ParagraphStyle('{self.name}') id: {id(self)}"

    @property
    def next_paragraph_style(self) -> ParagraphStyle:
        """|_ParagraphStyle| object representing the style to be applied automatically
        to a new paragraph inserted after a paragraph of this style.

        Returns self if no next paragraph style is defined. Assigning |None| or `self`
        removes the setting such that new paragraphs are created using this same style.
        """
        next_style_elm = self._style_element.next_style
        if next_style_elm is None:
            return self
        if next_style_elm.type != WD_STYLE_TYPE.PARAGRAPH:
            return self
        return cast(ParagraphStyle, StyleFactory(next_style_elm))

    @next_paragraph_style.setter
    def next_paragraph_style(self, style: ParagraphStyle | None) -> None:
        if style is None or style.style_id == self.style_id:
            self._style_element._remove_next()  # pyright: ignore[reportPrivateUsage]
        else:
            self._style_element.get_or_add_next().val = style.style_id

    @property
    def paragraph_format(self) -> ParagraphFormat:
        """The |ParagraphFormat| object providing access to the paragraph formatting
        properties for this style such as indentation."""
        return ParagraphFormat(self._style_element)


# -- just in case someone uses the old name in an extension function --
_ParagraphStyle = ParagraphStyle


class _TableStyle(ParagraphStyle):
    """A table style.

    A table style provides character and paragraph formatting for its contents as well
    as special table formatting properties.
    """

    def __repr__(self) -> str:
        return f"_TableStyle('{self.name}') id: {id(self)}"


class _NumberingStyle(BaseStyle):
    """A numbering style.

    Not yet implemented.
    """
