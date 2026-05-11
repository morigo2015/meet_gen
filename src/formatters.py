from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from numbers import Integral, Real
from typing import Any


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\u00a0", " ").split())


def normalize_header(value: Any) -> str:
    return normalize_text(value)


def is_empty_value(value: Any) -> bool:
    return normalize_text(value) == ""


def format_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, Integral):
        return str(value)
    if isinstance(value, Decimal):
        return _format_decimal(value)
    if isinstance(value, Real):
        return _format_decimal(Decimal(str(value)))
    return normalize_text(value)


def format_row(row: dict[str, Any]) -> dict[str, str]:
    return {key: format_value(value) for key, value in row.items() if key != "_row_number"}


def _format_decimal(value: Decimal) -> str:
    try:
        normalized = value.normalize()
    except InvalidOperation:
        return str(value)
    if normalized == normalized.to_integral_value():
        return f"{normalized.to_integral_value():f}"
    return f"{value:.2f}"
