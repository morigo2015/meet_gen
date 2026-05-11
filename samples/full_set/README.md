# Full Sample Set

Input files:
- `templates/contract_template.docx`
- `input/contracts.xlsx`

Run:

```bash
python -m src.main generate --template ./samples/full_set/templates/contract_template.docx --data ./samples/full_set/input/contracts.xlsx --output ./samples/full_set/output/generated
```

Expected non-strict result: 5 processed non-empty rows, 4 generated documents, 1 skipped row.
