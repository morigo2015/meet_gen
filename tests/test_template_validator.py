from __future__ import annotations

import pytest

from src.errors import TemplateValidationError
from src.template_validator import (
    extract_placeholders,
    validate_required_fields,
    validate_template_columns,
)


def test_extracts_simple_placeholders(make_template) -> None:
    path = make_template(["Договір {{ номер_договору }}", "Клієнт {{ назва_клієнта }}"])

    assert extract_placeholders(path) == ["номер_договору", "назва_клієнта"]


def test_rejects_unsupported_placeholder_syntax(make_template) -> None:
    path = make_template(["Сума {{ сума | currency }}"])

    with pytest.raises(TemplateValidationError, match="Invalid placeholder"):
        extract_placeholders(path)


def test_reports_missing_columns_and_required_fields() -> None:
    column_issues = validate_template_columns(["назва_клієнта"], ["номер_договору"])
    assert column_issues
    assert "назва_клієнта" in column_issues[0].format()

    row_issues = validate_required_fields(
        [{"_row_number": 5, "номер_договору": "", "назва_клієнта": "ТОВ Альфа"}],
        ["номер_договору", "назва_клієнта"],
    )
    assert row_issues[0].format() == 'Row 5: missing required field "номер_договору"'
