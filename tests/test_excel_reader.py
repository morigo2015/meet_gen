from __future__ import annotations

import pytest

from src.errors import ExcelReadError
from src.excel_reader import read_excel


def test_reads_cyrillic_headers_values_and_preserves_row_numbers(make_workbook) -> None:
    path = make_workbook(
        [" номер_договору ", "назва_клієнта"],
        [
            ["001", "ТОВ Альфа"],
            [None, None],
            ["002", "ТОВ Бета"],
        ],
    )

    data = read_excel(path)

    assert data.headers == ["номер_договору", "назва_клієнта"]
    assert len(data.rows) == 2
    assert data.rows[0]["_row_number"] == 2
    assert data.rows[0]["назва_клієнта"] == "ТОВ Альфа"
    assert data.rows[1]["_row_number"] == 4


def test_detects_duplicate_normalized_headers(make_workbook) -> None:
    path = make_workbook(["номер_договору", " номер_договору "], [["001", "002"]])

    with pytest.raises(ExcelReadError, match="Duplicate column headers"):
        read_excel(path)


def test_missing_file_and_sheet_errors(make_workbook, tmp_path) -> None:
    with pytest.raises(ExcelReadError, match="not found"):
        read_excel(tmp_path / "missing.xlsx")

    path = make_workbook(["name"], [["value"]], sheet_name="Data")
    with pytest.raises(ExcelReadError, match="Worksheet not found"):
        read_excel(path, sheet_name="Other")
