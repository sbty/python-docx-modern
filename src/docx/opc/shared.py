"""Objects shared by opc modules."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any, TypeVar, cast

_T = TypeVar("_T")


class CaseInsensitiveDict(dict[str, str]):
    """Mapping type that behaves like dict except that it matches without respect to the
    case of the key.

    E.g. cid['A'] == cid['a']. Note this is not general-purpose, just complete enough to
    satisfy opc package needs. It assumes str keys, and that it is created empty; keys
    passed in constructor are not accounted for
    """

    def __contains__(self, key: object) -> bool:
        return super().__contains__(cast(str, key).lower())

    def __getitem__(self, key: str) -> str:
        return super().__getitem__(key.lower())

    def __setitem__(self, key: str, value: str) -> None:
        return super().__setitem__(key.lower(), value)


def cls_method_fn(cls: type, method_name: str) -> Callable[..., Any]:
    """Return method of `cls` having `method_name`."""
    return cast("Callable[..., Any]", getattr(cls, method_name))
