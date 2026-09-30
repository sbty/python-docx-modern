"""Unit-test suite for the `docx.oxml.simpletypes` module."""

from __future__ import annotations

import pytest

from docx.oxml.simpletypes import (
    BaseSimpleType,
    ST_HpsMeasure,
    ST_SignedTwipsMeasure,
    ST_TblWidthTwips,
    ST_TwipsMeasure,
    ST_UniversalMeasure,
)
from docx.shared import Inches, Pt, Twips


class DescribeMeasureParsing:
    """Invalid measure values raise `ValueError`, not `OverflowError` or `KeyError`."""

    @pytest.mark.parametrize(
        "simple_type",
        [ST_HpsMeasure, ST_SignedTwipsMeasure, ST_TblWidthTwips, ST_TwipsMeasure],
    )
    @pytest.mark.parametrize("str_value", ["INF", "-INF", "1E999"])
    def it_raises_ValueError_on_a_non_finite_number(
        self, simple_type: type[BaseSimpleType], str_value: str
    ):
        with pytest.raises(ValueError, match="finite"):
            simple_type.convert_from_xml(str_value)

    def it_raises_ValueError_on_a_non_finite_universal_measure(self):
        # -- the message shows the whole attribute value, unit included --
        with pytest.raises(ValueError, match="finite number, got '1e999in'"):
            ST_UniversalMeasure.convert_from_xml("1e999in")

    def it_raises_ValueError_on_an_unknown_universal_measure_unit(self):
        with pytest.raises(ValueError, match="unit"):
            ST_UniversalMeasure.convert_from_xml("12xx")


class DescribeST_HpsMeasure:
    """Unit-test suite for `docx.oxml.simpletypes.ST_HpsMeasure`."""

    @pytest.mark.parametrize(
        ("str_value", "expected_value"),
        [
            ("28", Pt(14)),
            # -- fractional half-points are kept, not rounded (#1475) --
            ("36.5625", Pt(18.28125)),
            ("12pt", Pt(12)),
        ],
    )
    def it_converts_an_XML_value_to_a_Length(self, str_value: str, expected_value: int):
        assert ST_HpsMeasure.convert_from_xml(str_value) == expected_value

    @pytest.mark.parametrize(
        ("value", "expected_str"),
        [
            (Pt(12), "24"),
            (Pt(10.5), "21"),
            # -- round to the nearest half-point rather than truncating --
            (Pt(10.75), "22"),
            # -- an exact quarter-point tie rounds to the even half-point count --
            (Pt(10.25), "20"),
            (Pt(18.28125), "37"),
        ],
    )
    def it_converts_a_Length_to_an_XML_value(self, value: int, expected_str: str):
        assert ST_HpsMeasure.convert_to_xml(value) == expected_str

    @pytest.mark.parametrize("str_value", ["abc", "NaN"])
    def it_raises_on_a_value_that_is_not_a_measure(self, str_value: str):
        # -- any ValueError is the contract; the message comes from int()/float() --
        with pytest.raises(ValueError):  # noqa: PT011
            ST_HpsMeasure.convert_from_xml(str_value)


class DescribeST_TwipsMeasure:
    """Unit-test suite for `docx.oxml.simpletypes.ST_TwipsMeasure`."""

    @pytest.mark.parametrize(
        ("str_value", "expected_value"),
        [
            ("1440", Twips(1440)),
            ("120.6", Twips(121)),
            # -- halves round to even, as in ST_SignedTwipsMeasure --
            ("0.5", Twips(0)),
            ("1.5", Twips(2)),
            ("1in", Inches(1)),
        ],
    )
    def it_converts_an_XML_value_to_a_Length(self, str_value: str, expected_value: int):
        assert ST_TwipsMeasure.convert_from_xml(str_value) == expected_value

    @pytest.mark.parametrize("str_value", ["abc", "NaN", "50%"])
    def it_raises_on_a_value_that_is_not_a_measure(self, str_value: str):
        # -- any ValueError is the contract; the message comes from int()/float() --
        with pytest.raises(ValueError):  # noqa: PT011
            ST_TwipsMeasure.convert_from_xml(str_value)

    @pytest.mark.parametrize("str_value", ["0.5", "1.5", "120.6"])
    def it_rounds_fractional_twips_like_ST_SignedTwipsMeasure(self, str_value: str):
        assert ST_TwipsMeasure.convert_from_xml(
            str_value
        ) == ST_SignedTwipsMeasure.convert_from_xml(str_value)
