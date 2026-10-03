"""Provides CoreProperties, Dublin-Core attributes of the document.

These are broadly-standardized attributes like author, last-modified, etc.
"""

from __future__ import annotations

import datetime as dt
from typing import TYPE_CHECKING

from docx.oxml.coreprops import CT_CoreProperties

if TYPE_CHECKING:
    from docx.oxml.coreprops import CT_CoreProperties


class CoreProperties:
    """Corresponds to part named ``/docProps/core.xml``, containing the core document
    properties for this document package."""

    def __init__(self, element: CT_CoreProperties) -> None:
        self._element = element

    @property
    def author(self) -> str:
        return self._element.author_text

    @author.setter
    def author(self, value: str) -> None:
        self._element.author_text = value

    @property
    def category(self) -> str:
        return self._element.category_text

    @category.setter
    def category(self, value: str) -> None:
        self._element.category_text = value

    @property
    def comments(self) -> str:
        return self._element.comments_text

    @comments.setter
    def comments(self, value: str) -> None:
        self._element.comments_text = value

    @property
    def content_status(self) -> str:
        return self._element.contentStatus_text

    @content_status.setter
    def content_status(self, value: str) -> None:
        self._element.contentStatus_text = value

    @property
    def created(self) -> dt.datetime | None:
        return self._element.created_datetime

    @created.setter
    def created(self, value: dt.datetime) -> None:
        self._element.created_datetime = value

    @property
    def identifier(self) -> str:
        return self._element.identifier_text

    @identifier.setter
    def identifier(self, value: str) -> None:
        self._element.identifier_text = value

    @property
    def keywords(self) -> str:
        return self._element.keywords_text

    @keywords.setter
    def keywords(self, value: str) -> None:
        self._element.keywords_text = value

    @property
    def language(self) -> str:
        return self._element.language_text

    @language.setter
    def language(self, value: str) -> None:
        self._element.language_text = value

    @property
    def last_modified_by(self) -> str:
        return self._element.lastModifiedBy_text

    @last_modified_by.setter
    def last_modified_by(self, value: str) -> None:
        self._element.lastModifiedBy_text = value

    @property
    def last_printed(self) -> dt.datetime | None:
        return self._element.lastPrinted_datetime

    @last_printed.setter
    def last_printed(self, value: dt.datetime) -> None:
        self._element.lastPrinted_datetime = value

    @property
    def modified(self) -> dt.datetime | None:
        return self._element.modified_datetime

    @modified.setter
    def modified(self, value: dt.datetime) -> None:
        self._element.modified_datetime = value

    @property
    def revision(self) -> int:
        return self._element.revision_number

    @revision.setter
    def revision(self, value: int) -> None:
        self._element.revision_number = value

    @property
    def subject(self) -> str:
        return self._element.subject_text

    @subject.setter
    def subject(self, value: str) -> None:
        self._element.subject_text = value

    @property
    def title(self) -> str:
        return self._element.title_text

    @title.setter
    def title(self, value: str) -> None:
        self._element.title_text = value

    @property
    def version(self) -> str:
        return self._element.version_text

    @version.setter
    def version(self, value: str) -> None:
        self._element.version_text = value
