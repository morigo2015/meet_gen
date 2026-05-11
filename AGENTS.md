# Repository Guidelines

## Project Structure & Module Organization
This repository is defined by a single MVP spec for a Python document generator. The expected layout is:

```text
src/              # application code
tests/            # pytest test files
templates/        # .docx Word templates
input/            # sample .xlsx source data
output/generated/ # generated .docx files
```

Keep core responsibilities split by module: `excel_reader.py`, `template_validator.py`, `word_renderer.py`, `filename_builder.py`, `formatters.py`, and `errors.py`. Use `README.md` for user-facing setup notes and examples.

## Build, Test, and Development Commands
The project is intended to run with Python 3.11+.

```bash
python -m src.main generate --template ./templates/template.docx --data ./input/data.xlsx --output ./output/generated
pytest
```

Use `--sheet`, `--filename-template`, `--strict`, and `--dry-run` when validating generation behavior. `pytest` is the required test command.

## Coding Style & Naming Conventions
Use standard Python style: 4-space indentation, `snake_case` for functions and modules, and `PascalCase` for classes. Keep file and module names descriptive and aligned with their responsibility. Preserve UTF-8 and Cyrillic text handling throughout CLI messages, filenames, and validation errors. Prefer explicit validation and small, testable functions over monolithic logic.

## Testing Guidelines
Tests should live under `tests/` and be named `test_*.py`. Cover the behaviors called out in the spec: Excel parsing, duplicate column detection, placeholder extraction, missing-field validation, filename generation, formatting, and `dry-run`/`strict` modes. Include cases for Cyrillic column names and values. Run the full suite with `pytest` before opening a PR.

## Commit & Pull Request Guidelines
No Git history is available in this workspace, so use short, imperative commit messages with a narrow scope, for example: `Add Excel row validation`. Pull requests should describe the change, note any CLI or template impact, and include sample output when generated `.docx` files or filenames change. Link related issues when available.

## Agent-Specific Instructions
Do not overwrite the Word template or generated output artifacts unless the task explicitly requires it. When changing formatting or validation rules, update the relevant tests in the same change.
