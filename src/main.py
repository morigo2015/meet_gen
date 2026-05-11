from __future__ import annotations

from collections import defaultdict
from pathlib import Path
import sys
from typing import Annotated

import typer
from pydantic import BaseModel, Field

from .errors import DocumentGeneratorError, TemplateValidationError, ValidationIssue
from .excel_reader import read_excel
from .filename_builder import DEFAULT_FILENAME_TEMPLATE, FilenameBuilder
from .formatters import format_row
from .template_validator import (
    extract_placeholders,
    validate_required_fields,
    validate_template_columns,
)
from .word_renderer import render_docx

app = typer.Typer(no_args_is_help=True)


class SkippedRow(BaseModel):
    row_number: int
    errors: list[str] = Field(default_factory=list)


class GeneratedDocument(BaseModel):
    row_number: int
    filename: str


class GenerationReport(BaseModel):
    processed_rows: int
    generated_documents: int
    skipped_rows: int
    output_directory: str
    dry_run: bool = False
    errors: list[str] = Field(default_factory=list)
    skipped: list[SkippedRow] = Field(default_factory=list)
    generated: list[GeneratedDocument] = Field(default_factory=list)

    @property
    def has_errors(self) -> bool:
        return bool(self.errors or self.skipped)


@app.callback()
def main() -> None:
    """Generate Word documents from Excel rows."""
    _configure_utf8_stdio()


def generate_documents(
    template: str | Path,
    data: str | Path,
    output: str | Path,
    sheet: str | None = None,
    filename_template: str = DEFAULT_FILENAME_TEMPLATE,
    strict: bool = False,
    dry_run: bool = False,
) -> GenerationReport:
    template_path = Path(template)
    data_path = Path(data)
    output_dir = Path(output)
    _ensure_output_dir(output_dir)

    excel_data = read_excel(data_path, sheet_name=sheet)
    placeholders = extract_placeholders(template_path)
    global_issues = validate_template_columns(placeholders, excel_data.headers)
    if global_issues:
        report = _build_report(
            processed_rows=len(excel_data.rows),
            output_dir=output_dir,
            dry_run=dry_run,
            errors=[issue.format() for issue in global_issues],
        )
        _write_log(output_dir, report)
        return report

    formatted_rows = []
    for row in excel_data.rows:
        formatted = format_row(row)
        formatted["_row_number"] = row["_row_number"]
        formatted_rows.append(formatted)

    row_issues = validate_required_fields(formatted_rows, placeholders)
    issues_by_row = _group_issues_by_row(row_issues)
    if strict and row_issues:
        skipped = [
            SkippedRow(row_number=row_number, errors=[issue.message for issue in issues])
            for row_number, issues in issues_by_row.items()
        ]
        report = _build_report(
            processed_rows=len(formatted_rows),
            output_dir=output_dir,
            dry_run=dry_run,
            skipped=skipped,
        )
        _write_log(output_dir, report)
        return report

    valid_rows = [row for row in formatted_rows if int(row["_row_number"]) not in issues_by_row]
    skipped_rows = [
        SkippedRow(row_number=row_number, errors=[issue.message for issue in issues])
        for row_number, issues in issues_by_row.items()
    ]

    filename_builder = FilenameBuilder(output_dir=output_dir)
    generated: list[GeneratedDocument] = []
    for index, row in enumerate(valid_rows, start=1):
        row_number = int(row["_row_number"])
        context = {key: value for key, value in row.items() if key != "_row_number"}
        filename = filename_builder.build(context, filename_template, sequence=index)
        generated.append(GeneratedDocument(row_number=row_number, filename=filename))
        if not dry_run:
            render_docx(template_path, output_dir / filename, context)

    report = _build_report(
        processed_rows=len(formatted_rows),
        output_dir=output_dir,
        dry_run=dry_run,
        skipped=skipped_rows,
        generated=generated,
    )
    _write_log(output_dir, report)
    return report


@app.command()
def generate(
    template: Annotated[Path, typer.Option("--template", help="Path to .docx template")],
    data: Annotated[Path, typer.Option("--data", help="Path to .xlsx data file")],
    output: Annotated[Path, typer.Option("--output", help="Directory for generated files")],
    sheet: Annotated[str | None, typer.Option("--sheet", help="Worksheet name")] = None,
    filename_template: Annotated[
        str,
        typer.Option("--filename-template", help="Filename template with {column_name} fields"),
    ] = DEFAULT_FILENAME_TEMPLATE,
    strict: Annotated[bool, typer.Option("--strict", help="Do not generate if validation errors exist")] = False,
    dry_run: Annotated[bool, typer.Option("--dry-run", help="Validate and report without writing documents")] = False,
) -> None:
    try:
        report = generate_documents(
            template=template,
            data=data,
            output=output,
            sheet=sheet,
            filename_template=filename_template,
            strict=strict,
            dry_run=dry_run,
        )
    except (DocumentGeneratorError, OSError) as exc:
        typer.echo(f"Error: {exc}", err=True)
        raise typer.Exit(code=1) from exc

    _echo_report(report)
    if report.errors or (strict and report.skipped) or (report.skipped and report.generated_documents == 0):
        raise typer.Exit(code=1)


def _ensure_output_dir(output_dir: Path) -> None:
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise TemplateValidationError(f"Could not create output directory: {output_dir}") from exc


def _configure_utf8_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, OSError, ValueError):
            continue


def _group_issues_by_row(issues: list[ValidationIssue]) -> dict[int, list[ValidationIssue]]:
    grouped: dict[int, list[ValidationIssue]] = defaultdict(list)
    for issue in issues:
        if issue.row_number is not None:
            grouped[issue.row_number].append(issue)
    return dict(sorted(grouped.items()))


def _build_report(
    processed_rows: int,
    output_dir: Path,
    dry_run: bool,
    errors: list[str] | None = None,
    skipped: list[SkippedRow] | None = None,
    generated: list[GeneratedDocument] | None = None,
) -> GenerationReport:
    generated = generated or []
    skipped = skipped or []
    return GenerationReport(
        processed_rows=processed_rows,
        generated_documents=len(generated),
        skipped_rows=len(skipped),
        output_directory=str(output_dir),
        dry_run=dry_run,
        errors=errors or [],
        skipped=skipped,
        generated=generated,
    )


def _echo_report(report: GenerationReport) -> None:
    if report.dry_run:
        typer.echo("Dry run: yes")
    typer.echo(f"Processed rows: {report.processed_rows}")
    typer.echo(f"Generated documents: {report.generated_documents}")
    typer.echo(f"Skipped rows: {report.skipped_rows}")
    typer.echo(f"Output directory: {report.output_directory}")
    for error in report.errors:
        typer.echo(error)
    for skipped in report.skipped:
        for error in skipped.errors:
            typer.echo(f"Row {skipped.row_number}: {error}")


def _write_log(output_dir: Path, report: GenerationReport) -> None:
    lines = []
    if report.dry_run:
        lines.append("Dry run: yes")
    lines.extend(
        [
            f"Processed rows: {report.processed_rows}",
            f"Generated documents: {report.generated_documents}",
            f"Skipped rows: {report.skipped_rows}",
            f"Output directory: {report.output_directory}",
        ]
    )
    if report.errors:
        lines.append("")
        lines.append("Errors:")
        lines.extend(report.errors)
    if report.skipped:
        lines.append("")
        lines.append("Skipped rows:")
        for skipped in report.skipped:
            for error in skipped.errors:
                lines.append(f"Row {skipped.row_number}: {error}")
    if report.generated:
        lines.append("")
        lines.append("Documents:")
        for item in report.generated:
            lines.append(f"Row {item.row_number}: {item.filename}")
    (output_dir / "generation.log").write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    app()
