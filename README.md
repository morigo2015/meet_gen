# Document Generator MVP

Generates one Word `.docx` document for each valid row in an Excel `.xlsx` file.

## Setup

Use Python 3.11+.

```bash
python -m pip install -e ".[dev]"
```

## Run

```bash
python -m src.main generate --template ./templates/template.docx --data ./input/data.xlsx --output ./output/generated
```

Optional flags:

```bash
--sheet "Sheet1"
--filename-template "{номер_договору}_{назва_клієнта}"
--strict
--dry-run
```

`--dry-run` validates data and writes `generation.log`, but does not create `.docx` files. `--strict` stops generation when validation errors are found and writes no documents.

## Input Format

The first Excel row must contain column names. Every later non-empty row is one document.

Example columns:

```text
номер_договору | назва_клієнта | дата_договору | сума
001            | ТОВ Альфа     | 01.05.2026    | 10000
002            | ТОВ Бета      | 02.05.2026    | 25000
```

The Word template uses simple placeholders only:

```text
{{ номер_договору }}
{{ назва_клієнта }}
{{ дата_договору }}
{{ сума }}
```

Jinja expressions, filters, loops, and conditions are intentionally out of scope for the MVP.

## Test

```bash
pytest
```
