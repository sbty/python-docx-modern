"""Test suite for the docx.parts.numbering module."""

from typing import cast

import pytest

import docx
from docx.opc.constants import CONTENT_TYPE as CT
from docx.oxml.ns import qn
from docx.oxml.numbering import CT_Numbering
from docx.package import Package
from docx.parts.numbering import NumberingPart, _NumberingDefinitions

from ..oxml.unitdata.numbering import a_num, a_numbering
from ..unitutil.file import test_file
from ..unitutil.mock import class_mock, instance_mock


class DescribeNumberingPart:
    def it_constructs_a_new_empty_numbering_part(self):
        package = Package()

        numbering_part = NumberingPart.new(package)

        assert isinstance(numbering_part, NumberingPart)
        assert numbering_part.partname == "/word/numbering.xml"
        assert numbering_part.content_type == CT.WML_NUMBERING
        assert numbering_part.package is package
        assert numbering_part.element.tag == qn("w:numbering")
        assert len(numbering_part.element) == 0

    # -- a document with no numbering part used to raise NotImplementedError (#1541) --
    def it_is_added_to_a_document_that_has_none(self):
        document = docx.Document(test_file("having-images.docx"))

        numbering_part = document.part.numbering_part

        assert len(numbering_part.numbering_definitions) == 0
        assert document.part.numbering_part is numbering_part

    def it_provides_access_to_the_numbering_definitions(self, num_defs_fixture):
        (
            numbering_part,
            _NumberingDefinitions_,
            numbering_elm_,
            numbering_definitions_,
        ) = num_defs_fixture
        numbering_definitions = numbering_part.numbering_definitions
        _NumberingDefinitions_.assert_called_once_with(numbering_elm_)
        assert numbering_definitions is numbering_definitions_

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def num_defs_fixture(self, _NumberingDefinitions_, numbering_elm_, numbering_definitions_):
        numbering_part = NumberingPart(None, None, numbering_elm_, None)
        return (
            numbering_part,
            _NumberingDefinitions_,
            numbering_elm_,
            numbering_definitions_,
        )

    # fixture components ---------------------------------------------

    @pytest.fixture
    def _NumberingDefinitions_(self, request, numbering_definitions_):
        return class_mock(
            request,
            "docx.parts.numbering._NumberingDefinitions",
            return_value=numbering_definitions_,
        )

    @pytest.fixture
    def numbering_definitions_(self, request):
        return instance_mock(request, _NumberingDefinitions)

    @pytest.fixture
    def numbering_elm_(self, request):
        return instance_mock(request, CT_Numbering)


class Describe_NumberingDefinitions:
    def it_knows_how_many_numbering_definitions_it_contains(self, len_fixture):
        numbering_definitions, numbering_definition_count = len_fixture
        assert len(numbering_definitions) == numbering_definition_count

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[0, 1, 2, 3])
    def len_fixture(self, request):
        numbering_definition_count = request.param
        numbering_bldr = a_numbering().with_nsdecls()
        for idx in range(numbering_definition_count):
            numbering_bldr.with_child(a_num())
        numbering_elm = numbering_bldr.element
        numbering_definitions = _NumberingDefinitions(cast(CT_Numbering, numbering_elm))
        return numbering_definitions, numbering_definition_count
