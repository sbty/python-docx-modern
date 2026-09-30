"""Test suite for the docx.oxml.styles module."""

from typing import cast

import pytest

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.styles import CT_Styles

from ..unitutil.cxml import element, xml


class DescribeCT_Styles:
    def it_can_add_the_comment_styles_when_they_are_missing(self):
        styles = cast(CT_Styles, element("w:styles"))

        styles.ensure_comment_styles()

        comment_reference = styles.get_by_id("CommentReference")
        assert comment_reference is not None
        assert comment_reference.type == WD_STYLE_TYPE.CHARACTER
        assert comment_reference.name_val == "annotation reference"
        comment_text = styles.get_by_id("CommentText")
        assert comment_text is not None
        assert comment_text.type == WD_STYLE_TYPE.PARAGRAPH
        assert comment_text.name_val == "annotation text"
        # -- Word marks both styles semi-hidden --
        assert comment_reference.semiHidden_val is True
        assert comment_text.semiHidden_val is True

    def it_bases_the_comment_styles_on_the_document_default_styles(self):
        # -- localized templates use other ids for "Normal" and "Default Paragraph Font" --
        styles = cast(
            CT_Styles,
            element(
                "w:styles/(w:style{w:type=paragraph,w:default=1,w:styleId=a}"
                "/w:name{w:val=Normal},w:style{w:type=character,w:default=1,w:styleId=a0}"
                "/w:name{w:val=Default Paragraph Font})"
            ),
        )

        styles.ensure_comment_styles()

        comment_reference = styles.get_by_id("CommentReference")
        assert comment_reference is not None
        assert comment_reference.basedOn_val == "a0"
        comment_text = styles.get_by_id("CommentText")
        assert comment_text is not None
        assert comment_text.basedOn_val == "a"

    def but_it_does_not_duplicate_a_comment_style_defined_under_another_id(self):
        styles = cast(
            CT_Styles,
            element(
                "w:styles/(w:style{w:type=character,w:styleId=a3}"
                "/w:name{w:val=annotation reference},w:style{w:type=paragraph,w:styleId=a4}"
                "/w:name{w:val=annotation text})"
            ),
        )
        expected_xml = styles.xml

        styles.ensure_comment_styles()

        assert styles.xml == expected_xml

    def but_it_leaves_existing_comment_styles_alone(self):
        styles = cast(
            CT_Styles,
            element(
                "w:styles/(w:style{w:type=character,w:styleId=CommentReference}"
                "/w:name{w:val=Custom},w:style{w:type=paragraph,w:styleId=CommentText})"
            ),
        )
        expected_xml = styles.xml

        styles.ensure_comment_styles()
        styles.ensure_comment_styles()

        assert styles.xml == expected_xml

    def it_can_add_a_style_of_type(self, add_fixture):
        styles, name, style_type, builtin, expected_xml = add_fixture
        style = styles.add_style_of_type(name, style_type, builtin)
        assert styles.xml == expected_xml
        assert style is styles[-1]

    # fixtures -------------------------------------------------------

    @pytest.fixture(
        params=[
            (
                "w:styles",
                "Foo Bar",
                WD_STYLE_TYPE.LIST,
                False,
                "w:styles/w:style{w:type=numbering,w:customStyle=1,w:styleId=FooBar"
                "}/w:name{w:val=Foo Bar}",
            ),
            (
                "w:styles",
                "heading 1",
                WD_STYLE_TYPE.PARAGRAPH,
                True,
                "w:styles/w:style{w:type=paragraph,w:styleId=Heading1}/w:name{w:val=heading 1}",
            ),
        ]
    )
    def add_fixture(self, request):
        styles_cxml, name, style_type, builtin, expected_cxml = request.param
        styles = element(styles_cxml)
        expected_xml = xml(expected_cxml)
        return styles, name, style_type, builtin, expected_xml
