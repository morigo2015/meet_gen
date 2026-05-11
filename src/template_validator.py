from __future__ import annotations

import html
import re
import zipfile
from pathlib import Path
from typing import Iterable

from .errors import TemplateValidationError, ValidationIssue
from .formatters import is_empty_value, normalize_header

PLACEHOLDER_PATTERN = re.compile(r"{{\s*(.*?)\s*}}", re.DOTALL)
SIMPLE_FIELD_PATTERN = re.compile(r"^[\w ]+$", re.UNICODE)
WORD_XML_PREFIXES = (
    "word/document.xml",
    "word/header",
    "word/footer",
    "word/footnotes.xml",
    "word/endnotes.xml",
)


def extract_placeholders(template_path: str | Path) -> list[str]:
    text = _read_word_xml(template_path)
    issues = _find_syntax_issues(text)
    if issues:
        raise TemplateValidationError("; ".join(issues))

    placeholders: list[str] = []
    seen: set[str] = set()
    for match in PLACEHOLDER_PATTERN.finditer(text):
        name = normalize_header(match.group(1))
        if name not in seen:
            placeholders.append(name)
            seen.add(name)
    return placeholders


def validate_template_columns(placeholders: Iterable[str], columns: Iterable[str]) -> list[ValidationIssue]:
    normalized_columns = {normalize_header(column) for column in columns}
    issues: list[ValidationIssue] = []
    for placeholder in placeholders:
        normalized_placeholder = normalize_header(placeholder)
        if normalized_placeholder not in normalized_columns:
            issues.append(
                ValidationIssue(
                    message=(
                        f'Placeholder "{{{{ {normalized_placeholder} }}}}" has no matching '
                        f'Excel column "{normalized_placeholder}".'
                    ),
                    field=normalized_placeholder,
                )
            )
    return issues


def validate_required_fields(
    rows: Iterable[dict[str, object]],
    placeholders: Iterable[str],
) -> list[ValidationIssue]:
    required = [normalize_header(placeholder) for placeholder in placeholders]
    issues: list[ValidationIssue] = []
    for row in rows:
        row_number = _row_number(row)
        for field in required:
            if is_empty_value(row.get(field)):
                issues.append(
                    ValidationIssue(
                        message=f'missing required field "{field}"',
                        row_number=row_number,
                        field=field,
                    )
                )
    return issues


def validate_row_required_fields(
    row: dict[str, object],
    placeholders: Iterable[str],
) -> list[ValidationIssue]:
    return validate_required_fields([row], placeholders)


def _read_word_xml(template_path: str | Path) -> str:
    source = Path(template_path)
    if not source.exists():
        raise TemplateValidationError(f"Word template not found: {source}")
    if source.suffix.lower() != ".docx":
        raise TemplateValidationError(f"Unsupported Word template format: {source.suffix or '<none>'}")
    try:
        with zipfile.ZipFile(source) as docx:
            parts = []
            for name in docx.namelist():
                if name.startswith(WORD_XML_PREFIXES) and name.endswith(".xml"):
                    parts.append(docx.read(name).decode("utf-8"))
    except zipfile.BadZipFile as exc:
        raise TemplateValidationError(f"Invalid Word template: {source}") from exc
    if not parts:
        raise TemplateValidationError(f"Word template has no readable document XML: {source}")
    return html.unescape("\n".join(parts))


def _find_syntax_issues(text: str) -> list[str]:
    issues: list[str] = []
    if "{%" in text or "{#" in text:
        issues.append("Unsupported Jinja tags found; only simple {{ column_name }} placeholders are supported")

    matches = list(PLACEHOLDER_PATTERN.finditer(text))
    if text.count("{{") != len(matches) or text.count("}}") != len(matches):
        issues.append("Malformed placeholder braces found")

    for match in matches:
        expression = normalize_header(match.group(1))
        if not expression:
            issues.append("Empty placeholder found")
        elif not SIMPLE_FIELD_PATTERN.fullmatch(expression):
            issues.append(
                f'Invalid placeholder "{{{{ {expression} }}}}"; only simple column names are supported'
            )
    return issues


def _row_number(row: dict[str, object]) -> int | None:
    value = row.get("_row_number")
    return value if isinstance(value, int) else None
