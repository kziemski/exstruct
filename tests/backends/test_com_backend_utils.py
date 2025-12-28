from collections.abc import Callable
from typing import TypeVar, cast

import pytest
from typing_extensions import ParamSpec

from exstruct.core.backends.com_backend import _split_csv_respecting_quotes

P = ParamSpec("P")
R = TypeVar("R")


def _parametrize(
    *args: object, **kwargs: object
) -> Callable[[Callable[P, R]], Callable[P, R]]:
    return cast(
        Callable[[Callable[P, R]], Callable[P, R]],
        pytest.mark.parametrize(*args, **kwargs),
    )


@_parametrize(
    "raw,expected",
    [
        ("A,B", ["A", "B"]),
        ("  A  ,  B ", ["A", "B"]),
        ("'Sheet,1'!A1:B2,'Sheet2'!C3:D4", ["'Sheet,1'!A1:B2", "'Sheet2'!C3:D4"]),
        ("'O''Brien'!A1,'X'!B2", ["'O''Brien'!A1", "'X'!B2"]),
        ("OnlyOne", ["OnlyOne"]),
    ],
)
def test_split_csv_respecting_quotes(raw: str, expected: list[str]) -> None:
    """シングルクォート内のカンマを保持して分割する。"""
    assert _split_csv_respecting_quotes(raw) == expected
