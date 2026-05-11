from __future__ import annotations

from src.filename_builder import FilenameBuilder, build_filename


def test_preserves_cyrillic_and_replaces_forbidden_characters() -> None:
    row = {"номер_договору": "001/А", "назва_клієнта": 'ТОВ "Альфа"'}

    filename = build_filename(row)

    assert filename == "001_А_ТОВ_Альфа.docx"


def test_fallback_truncation_and_duplicate_suffixes(tmp_path) -> None:
    builder = FilenameBuilder(max_length=24)

    assert builder.build({"номер_договору": "", "назва_клієнта": "ТОВ"}, sequence=1) == "document_0001.docx"
    assert builder.build({"номер_договору": "", "назва_клієнта": "ТОВ"}, sequence=2) == "document_0002.docx"

    first = builder.build({"номер_договору": "001", "назва_клієнта": "ТОВ Альфа"}, sequence=3)
    second = builder.build({"номер_договору": "001", "назва_клієнта": "ТОВ Альфа"}, sequence=4)
    assert first == "001_ТОВ_Альфа.docx"
    assert second == "001_ТОВ_Альфа_2.docx"

    long_name = FilenameBuilder(max_length=20).build(
        {"номер_договору": "001", "назва_клієнта": "ДужеДовгаНазваКлієнта"},
        sequence=1,
    )
    assert long_name.endswith(".docx")
    assert len(long_name) <= 20
