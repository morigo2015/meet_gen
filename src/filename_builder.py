from __future__ import annotations

import re
import string
from pathlib import Path
from typing import Any

from .formatters import format_value, is_empty_value

DEFAULT_FILENAME_TEMPLATE = "{номер_договору}_{назва_клієнта}"
FORBIDDEN_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
WHITESPACE = re.compile(r"\s+")
RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


class FilenameBuilder:
    def __init__(
        self,
        output_dir: str | Path | None = None,
        max_length: int = 180,
    ) -> None:
        self.max_length = max_length
        self.used_names: set[str] = set()
        if output_dir is not None:
            directory = Path(output_dir)
            if directory.exists():
                self.used_names.update(path.name for path in directory.glob("*.docx"))

    def build(
        self,
        row: dict[str, Any],
        template: str = DEFAULT_FILENAME_TEMPLATE,
        sequence: int = 1,
    ) -> str:
        stem = self._render_stem(row, template, sequence)
        stem = _sanitize_stem(stem)
        if not stem:
            stem = f"document_{sequence:04d}"
        stem = _avoid_reserved_name(stem)
        filename = _with_extension(stem, self.max_length)
        filename = self._deduplicate(filename)
        self.used_names.add(filename)
        return filename

    def _render_stem(self, row: dict[str, Any], template: str, sequence: int) -> str:
        field_names = _field_names(template)
        if not field_names:
            return template
        if any(is_empty_value(row.get(field)) for field in field_names):
            return f"document_{sequence:04d}"
        values = {field: format_value(row.get(field)) for field in field_names}
        try:
            return template.format(**values)
        except (KeyError, ValueError):
            return f"document_{sequence:04d}"

    def _deduplicate(self, filename: str) -> str:
        if filename not in self.used_names:
            return filename
        stem = filename[:-5]
        for counter in range(2, 10000):
            suffix = f"_{counter}"
            available_stem_length = self.max_length - len(".docx") - len(suffix)
            candidate_stem = stem[:available_stem_length].rstrip(" ._") + suffix
            candidate = f"{candidate_stem}.docx"
            if candidate not in self.used_names:
                return candidate
        raise RuntimeError("Could not build a unique filename")


def build_filename(
    row: dict[str, Any],
    template: str = DEFAULT_FILENAME_TEMPLATE,
    sequence: int = 1,
    output_dir: str | Path | None = None,
    max_length: int = 180,
) -> str:
    return FilenameBuilder(output_dir=output_dir, max_length=max_length).build(row, template, sequence)


def _field_names(template: str) -> list[str]:
    formatter = string.Formatter()
    names: list[str] = []
    for _, field_name, _, _ in formatter.parse(template):
        if field_name and field_name not in names:
            names.append(field_name)
    return names


def _sanitize_stem(value: str) -> str:
    value = FORBIDDEN_CHARS.sub("_", value)
    value = WHITESPACE.sub("_", value)
    value = re.sub(r"_+", "_", value)
    return value.strip(" ._")


def _avoid_reserved_name(stem: str) -> str:
    if stem.upper() in RESERVED_NAMES:
        return f"{stem}_"
    return stem


def _with_extension(stem: str, max_length: int) -> str:
    extension = ".docx"
    available_stem_length = max(1, max_length - len(extension))
    return f"{stem[:available_stem_length].rstrip(' ._')}{extension}"
