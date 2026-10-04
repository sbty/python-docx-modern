"""DrawingML objects related to color, ColorFormat being the most prominent."""

from __future__ import annotations

from typing import TYPE_CHECKING, TypeAlias, cast

from docx.enum.dml import MSO_COLOR_TYPE
from docx.oxml.simpletypes import ST_HexColorAuto
from docx.shared import ElementProxy, RGBColor

if TYPE_CHECKING:
    from docx.enum.dml import MSO_THEME_COLOR
    from docx.oxml.styles import CT_Style
    from docx.oxml.text.font import CT_Color
    from docx.oxml.text.run import CT_R

# -- the `w:rPr` parents a color can be read from: a run or a style definition --
RPrParent: TypeAlias = "CT_R | CT_Style"


class ColorFormat(ElementProxy):
    """Provides access to color settings like RGB color, theme color, and luminance adjustments."""

    def __init__(self, rPr_parent: RPrParent) -> None:
        super().__init__(rPr_parent)
        self._element = rPr_parent
        # -- the `w:rPr` parent, typed more narrowly than the inherited `._element` --
        self._rPr_parent = rPr_parent

    @property
    def rgb(self) -> RGBColor | None:
        """An |RGBColor| value or |None| if no RGB color is specified.

        When :attr:`type` is `MSO_COLOR_TYPE.RGB`, the value of this property will always be an
        |RGBColor| value. It may also be an |RGBColor| value if :attr:`type` is
        `MSO_COLOR_TYPE.THEME`, as Word writes the current value of a theme color when one is
        assigned. In that case, the RGB value should be interpreted as no more than a good guess
        however, as the theme color takes precedence at rendering time. Its value is |None|
        whenever :attr:`type` is either |None| or `MSO_COLOR_TYPE.AUTO`.

        Assigning an |RGBColor| value causes :attr:`type` to become `MSO_COLOR_TYPE.RGB` and any
        theme color is removed. Assigning |None| causes any color to be removed such that the
        effective color is inherited from the style hierarchy.
        """
        color = self._color
        if color is None:
            return None
        if color.val == ST_HexColorAuto.AUTO:
            return None
        return cast(RGBColor, color.val)

    @rgb.setter
    def rgb(self, value: RGBColor | None) -> None:
        if value is None and self._color is None:
            return
        rPr = self._rPr_parent.get_or_add_rPr()
        rPr._remove_color()  # pyright: ignore[reportPrivateUsage]
        if value is not None:
            rPr.get_or_add_color().val = value

    @property
    def theme_color(self) -> MSO_THEME_COLOR | None:
        """Member of :ref:`MsoThemeColorIndex` or |None| if no theme color is specified.

        When :attr:`type` is `MSO_COLOR_TYPE.THEME`, the value of this property will always be a
        member of :ref:`MsoThemeColorIndex`. When :attr:`type` has any other value, the value of
        this property is |None|.

        Assigning a member of :ref:`MsoThemeColorIndex` causes :attr:`type` to become
        `MSO_COLOR_TYPE.THEME`. Any existing RGB value is retained but ignored by Word. Assigning
        |None| causes any color specification to be removed such that the effective color is
        inherited from the style hierarchy.
        """
        color = self._color
        if color is None:
            return None
        return color.themeColor

    @theme_color.setter
    def theme_color(self, value: MSO_THEME_COLOR | None) -> None:
        if value is None:
            rPr = self._rPr_parent.rPr
            if self._color is not None and rPr is not None:
                rPr._remove_color()  # pyright: ignore[reportPrivateUsage]
            return
        self._rPr_parent.get_or_add_rPr().get_or_add_color().themeColor = value

    @property
    def type(self) -> MSO_COLOR_TYPE | None:
        """Read-only.

        A member of :ref:`MsoColorType`, one of RGB, THEME, or AUTO, corresponding to the way this
        color is defined. Its value is |None| if no color is applied at this level, which causes
        the effective color to be inherited from the style hierarchy.
        """
        color = self._color
        if color is None:
            return None
        if color.themeColor is not None:
            return MSO_COLOR_TYPE.THEME
        if color.val == ST_HexColorAuto.AUTO:
            return MSO_COLOR_TYPE.AUTO
        return MSO_COLOR_TYPE.RGB

    @property
    def _color(self) -> CT_Color | None:
        """Return `w:rPr/w:color` or |None| if not present.

        Helper to factor out repetitive element access.
        """
        rPr = self._rPr_parent.rPr
        if rPr is None:
            return None
        return rPr.color
