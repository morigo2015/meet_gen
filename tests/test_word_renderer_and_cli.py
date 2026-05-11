from __future__ import annotations

from pathlib import Path

from docx import Document
from typer.testing import CliRunner

from src.main import app, generate_documents
from src.word_renderer import render_docx


def _document_text(path: Path) -> str:
    return "\n".join(paragraph.text for paragraph in Document(path).paragraphs)


def test_renders_one_docx_and_keeps_template_unchanged(make_template, tmp_path) -> None:
    template = make_template(["Договір {{ номер_договору }}", "Клієнт {{ назва_клієнта }}"])
    before = template.read_bytes()
    output = tmp_path / "result.docx"

    render_docx(template, output, {"номер_договору": "001", "назва_клієнта": "ТОВ Альфа"})

    assert "001" in _document_text(output)
    assert "ТОВ Альфа" in _document_text(output)
    assert template.read_bytes() == before


def test_generate_documents_mult_row_non_strict_and_dry_run(make_template, make_workbook, tmp_path) -> None:
    template = make_template(["{{ номер_договору }} {{ назва_клієнта }}"])
    workbook = make_workbook(
        ["номер_договору", "назва_клієнта"],
        [["001", "ТОВ Альфа"], ["", "ТОВ Бета"], ["003", "ТОВ Гама"]],
    )
    output = tmp_path / "generated"

    report = generate_documents(template, workbook, output, dry_run=True)

    assert report.processed_rows == 3
    assert report.generated_documents == 2
    assert report.skipped_rows == 1
    assert not list(output.glob("*.docx"))
    assert (output / "generation.log").exists()

    report = generate_documents(template, workbook, output)
    assert report.generated_documents == 2
    assert len(list(output.glob("*.docx"))) == 2


def test_strict_mode_writes_no_documents_when_any_row_is_invalid(make_template, make_workbook, tmp_path) -> None:
    template = make_template(["{{ номер_договору }} {{ назва_клієнта }}"])
    workbook = make_workbook(
        ["номер_договору", "назва_клієнта"],
        [["001", "ТОВ Альфа"], ["", "ТОВ Бета"]],
    )
    output = tmp_path / "generated"

    report = generate_documents(template, workbook, output, strict=True)

    assert report.generated_documents == 0
    assert report.skipped_rows == 1
    assert not list(output.glob("*.docx"))


def test_cli_generate_command(make_template, make_workbook, tmp_path) -> None:
    template = make_template(["{{ номер_договору }} {{ назва_клієнта }}"])
    workbook = make_workbook(["номер_договору", "назва_клієнта"], [["001", "ТОВ Альфа"]])
    output = tmp_path / "generated"
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "generate",
            "--template",
            str(template),
            "--data",
            str(workbook),
            "--output",
            str(output),
        ],
    )

    assert result.exit_code == 0, result.output
    assert "Processed rows: 1" in result.output
    assert len(list(output.glob("*.docx"))) == 1
