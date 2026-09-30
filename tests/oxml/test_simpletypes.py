"""Unit-test suite for the `docx.oxml.simpletypes` module."""

from __future__ import annotations

import pytest

from docx.oxml.simpletypes import ST_SignedTwipsMeasure, ST_TwipsMeasure
from docx.shared import Inches, Twips


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
