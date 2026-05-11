from __future__ import annotations

from dataclasses import dataclass


class DocumentGeneratorError(Exception):
    """Base class for user-facing generator errors."""


class ExcelReadError(DocumentGeneratorError):
    """Raised when the Excel source cannot be read or validated."""


class TemplateValidationError(DocumentGeneratorError):
    """Raised when the Word template cannot be parsed or validated."""


class FilenameBuildError(DocumentGeneratorError):
    """Raised when a filesystem-safe filename cannot be built."""


class RenderingError(DocumentGeneratorError):
    """Raised when a Word document cannot be rendered or saved."""


@dataclass(frozen=True)
class ValidationIssue:
    message: str
    row_number: int | None = None
    field: str | None = None

    def format(self) -> str:
        prefix = f"Row {self.row_number}: " if self.row_number is not None else ""
        return f"{prefix}{self.message}"
