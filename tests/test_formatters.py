from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal

from src.formatters import format_value, normalize_header, normalize_text


def test_normalizes_text_and_headers() -> None:
    assert normalize_text("  ТОВ\u00a0 Альфа  ") == "ТОВ Альфа"
    assert normalize_header("  номер_договору  ") == "номер_договору"


def test_formats_dates_and_numbers() -> None:
    assert format_value(date(2026, 5, 1)) == "01.05.2026"
    assert format_value(datetime(2026, 5, 2, 13, 30)) == "02.05.2026"
    assert format_value(10) == "10"
    assert format_value(10.0) == "10"
    assert format_value(Decimal("10.50")) == "10.50"
    assert format_value(Decimal("10.5")) == "10.50"
