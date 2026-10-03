# pyright: reportPrivateUsage=false

"""Temporary stand-in for main oxml module.

This module came across with the PackageReader transplant. Probably much will get
replaced with objects from the pptx.oxml.core and then this module will either get
deleted or only hold the package related custom element classes.
"""

from __future__ import annotations

from typing import cast

from lxml import etree

from docx.opc.constants import NAMESPACE as NS
from docx.opc.constants import RELATIONSHIP_TARGET_MODE as RTM

# configure XML parser
element_class_lookup = etree.ElementNamespaceClassLookup()
oxml_parser = etree.XMLParser(remove_blank_text=True, resolve_entities=False)
oxml_parser.set_element_class_lookup(element_class_lookup)

nsmap = {
    "ct": NS.OPC_CONTENT_TYPES,
    "pr": NS.OPC_RELATIONSHIPS,
    "r": NS.OFC_RELATIONSHIPS,
}


# ===========================================================================
# functions
# ===========================================================================


def parse_xml(text: str | bytes) -> etree._Element:
    """`etree.fromstring()` replacement that uses oxml parser."""
    return etree.fromstring(text, oxml_parser)


def qn(tag: str) -> str:
    """Stands for "qualified name", a utility function to turn a namespace prefixed tag
    name into a Clark-notation qualified tag name for lxml.

    For
    example, ``qn('p:cSld')`` returns ``'{http://schemas.../main}cSld'``.
    """
    prefix, tagroot = tag.split(":")
    uri = nsmap[prefix]
    return f"{{{uri}}}{tagroot}"


def serialize_part_xml(part_elm: etree._Element) -> bytes:
    """Serialize `part_elm` etree element to XML suitable for storage as an XML part.

    That is to say, no insignificant whitespace added for readability, and an
    appropriate XML declaration added with UTF-8 encoding specified.
    """
    return etree.tostring(part_elm, encoding="UTF-8", standalone=True)


def serialize_for_reading(element: etree._Element) -> str:
    """Serialize `element` to human-readable XML suitable for tests.

    No XML declaration.
    """
    return etree.tostring(element, encoding="unicode", pretty_print=True)


# ===========================================================================
# Custom element classes
# ===========================================================================


class BaseOxmlElement(etree.ElementBase):
    """Base class for all custom element classes, to add standardized behavior to all
    classes in one place."""

    @property
    def xml(self) -> str:
        """Return XML string for this element, suitable for testing purposes.

        Pretty printed for readability and without an XML declaration at the top.
        """
        return serialize_for_reading(self)


class CT_Default(BaseOxmlElement):
    """`<Default>` element that appears in `[Content_Types].xml` part.

    Used to specify a default content type to be applied to any part with the specified extension.
    """

    @property
    def content_type(self) -> str:
        """String held in the ``ContentType`` attribute of this ``<Default>``
        element."""
        # -- a required attribute in the schema --
        return cast(str, self.get("ContentType"))

    @property
    def extension(self) -> str:
        """String held in the ``Extension`` attribute of this ``<Default>`` element."""
        # -- a required attribute in the schema --
        return cast(str, self.get("Extension"))

    @staticmethod
    def new(ext: str, content_type: str) -> CT_Default:
        """Return a new ``<Default>`` element with attributes set to parameter values."""
        xml = f'<Default xmlns="{nsmap["ct"]}"/>'
        default = cast(CT_Default, parse_xml(xml))
        default.set("Extension", ext)
        default.set("ContentType", content_type)
        return default


class CT_Override(BaseOxmlElement):
    """``<Override>`` element, specifying the content type to be applied for a part with
    the specified partname."""

    @property
    def content_type(self) -> str:
        """String held in the ``ContentType`` attribute of this ``<Override>``
        element."""
        # -- a required attribute in the schema --
        return cast(str, self.get("ContentType"))

    @staticmethod
    def new(partname: str, content_type: str) -> CT_Override:
        """Return a new ``<Override>`` element with attributes set to parameter values."""
        xml = f'<Override xmlns="{nsmap["ct"]}"/>'
        override = cast(CT_Override, parse_xml(xml))
        override.set("PartName", partname)
        override.set("ContentType", content_type)
        return override

    @property
    def partname(self) -> str:
        """String held in the ``PartName`` attribute of this ``<Override>`` element."""
        # -- a required attribute in the schema --
        return cast(str, self.get("PartName"))


class CT_Relationship(BaseOxmlElement):
    """`<Relationship>` element, representing a single relationship from source to target part."""

    @staticmethod
    def new(
        rId: str, reltype: str, target: str, target_mode: str = RTM.INTERNAL
    ) -> CT_Relationship:
        """Return a new ``<Relationship>`` element."""
        xml = f'<Relationship xmlns="{nsmap["pr"]}"/>'
        relationship = cast(CT_Relationship, parse_xml(xml))
        relationship.set("Id", rId)
        relationship.set("Type", reltype)
        relationship.set("Target", target)
        if target_mode == RTM.EXTERNAL:
            relationship.set("TargetMode", RTM.EXTERNAL)
        return relationship

    @property
    def rId(self) -> str:
        """String held in the ``Id`` attribute of this ``<Relationship>`` element."""
        # -- a required attribute in the schema --
        return cast(str, self.get("Id"))

    @property
    def reltype(self) -> str:
        """String held in the ``Type`` attribute of this ``<Relationship>`` element."""
        # -- a required attribute in the schema --
        return cast(str, self.get("Type"))

    @property
    def target_ref(self) -> str:
        """String held in the ``Target`` attribute of this ``<Relationship>``
        element."""
        # -- a required attribute in the schema --
        return cast(str, self.get("Target"))

    @property
    def target_mode(self) -> str:
        """String held in the ``TargetMode`` attribute of this ``<Relationship>``
        element, either ``Internal`` or ``External``.

        Defaults to ``Internal``.
        """
        return self.get("TargetMode", RTM.INTERNAL)


class CT_Relationships(BaseOxmlElement):
    """``<Relationships>`` element, the root element in a .rels file."""

    def add_rel(self, rId: str, reltype: str, target: str, is_external: bool = False) -> None:
        """Add a child ``<Relationship>`` element with attributes set according to
        parameter values."""
        target_mode = RTM.EXTERNAL if is_external else RTM.INTERNAL
        relationship = CT_Relationship.new(rId, reltype, target, target_mode)
        self.append(relationship)

    @staticmethod
    def new() -> CT_Relationships:
        """Return a new ``<Relationships>`` element."""
        xml = f'<Relationships xmlns="{nsmap["pr"]}"/>'
        return cast(CT_Relationships, parse_xml(xml))

    @property
    def Relationship_lst(self) -> list[CT_Relationship]:
        """Return a list containing all the ``<Relationship>`` child elements."""
        return cast("list[CT_Relationship]", self.findall(qn("pr:Relationship")))

    @property
    def xml(self) -> bytes:  # type: ignore[override]
        """Return XML string for this element, suitable for saving in a .rels stream,
        not pretty printed and with an XML declaration at the top."""
        return serialize_part_xml(self)


class CT_Types(BaseOxmlElement):
    """``<Types>`` element, the container element for Default and Override elements in
    [Content_Types].xml."""

    def add_default(self, ext: str, content_type: str) -> None:
        """Add a child ``<Default>`` element with attributes set to parameter values."""
        default = CT_Default.new(ext, content_type)
        self.append(default)

    def add_override(self, partname: str, content_type: str) -> None:
        """Add a child ``<Override>`` element with attributes set to parameter
        values."""
        override = CT_Override.new(partname, content_type)
        self.append(override)

    @property
    def defaults(self) -> list[CT_Default]:
        return cast("list[CT_Default]", self.findall(qn("ct:Default")))

    @staticmethod
    def new() -> CT_Types:
        """Return a new ``<Types>`` element."""
        xml = f'<Types xmlns="{nsmap["ct"]}"/>'
        types = cast(CT_Types, parse_xml(xml))
        return types

    @property
    def overrides(self) -> list[CT_Override]:
        return cast("list[CT_Override]", self.findall(qn("ct:Override")))


ct_namespace = element_class_lookup.get_namespace(nsmap["ct"])
ct_namespace["Default"] = CT_Default
ct_namespace["Override"] = CT_Override
ct_namespace["Types"] = CT_Types

pr_namespace = element_class_lookup.get_namespace(nsmap["pr"])
pr_namespace["Relationship"] = CT_Relationship
pr_namespace["Relationships"] = CT_Relationships
