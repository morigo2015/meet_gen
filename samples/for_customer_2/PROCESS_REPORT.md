# Process and Results Report

Generated on 2026-05-12 for `samples/for_customer_2`.

## Inputs

- Workbook: `input/developer_work_2026.xlsx`.
- Sheet `contracts`: 3 rows, one framework contract per developer.
- Sheet `monthly_work`: 6 rows, two months per developer.
- Templates: `framework_service_contract.docx`, `monthly_work_request.docx`, `monthly_acceptance_act.docx`.

## Generation Results

- Contracts: 3 generated, 0 skipped.
- Work requests: 6 generated, 0 skipped.
- Acceptance acts: 6 generated, 0 skipped.
- Total generated documents: 15.

## Verification

- Output filenames were derived from workbook values and matched expected names.
- All generated documents were opened with `python-docx`.
- No unresolved `{{ ... }}` placeholders were found.
- Contracts contain contract number, developer name, tax ID, period, total gross fee, monthly request workflow, non-employment wording, and IP-transfer wording.
- Each monthly request contains contract number, request number, month, work description, expected result, acceptance criteria, deadline, and gross fee.
- Each act references the same contract and request, repeats accepted result/criteria, includes gross amount, PIT, military levy, net amount, and SSC accrual values.
- Tax/SSC arithmetic in generated acts was recalculated from workbook gross amounts and matched generated values.
- Generation logs exist in all output folders and show 0 skipped rows.

## Legal Alignment Check

See `LEGAL_CHECK.md`. Template-level check passed for the civil-law service workflow sample. Real production use still requires review by a Ukrainian lawyer and accountant for actual party status, tax treatment, and signing process.

## Files

### contracts
- `FW-DEV-2026-001_Іваненко_Олег_Петрович_договір.docx`
- `FW-DEV-2026-002_Петренко_Марія_Андріївна_договір.docx`
- `FW-DEV-2026-003_Шевчук_Дмитро_Ігорович_договір.docx`

### work_requests
- `REQ-2026-001-06_Іваненко_Олег_Петрович_заявка.docx`
- `REQ-2026-001-07_Іваненко_Олег_Петрович_заявка.docx`
- `REQ-2026-002-06_Петренко_Марія_Андріївна_заявка.docx`
- `REQ-2026-002-07_Петренко_Марія_Андріївна_заявка.docx`
- `REQ-2026-003-06_Шевчук_Дмитро_Ігорович_заявка.docx`
- `REQ-2026-003-07_Шевчук_Дмитро_Ігорович_заявка.docx`

### acts
- `ACT-2026-001-06_Іваненко_Олег_Петрович_акт.docx`
- `ACT-2026-001-07_Іваненко_Олег_Петрович_акт.docx`
- `ACT-2026-002-06_Петренко_Марія_Андріївна_акт.docx`
- `ACT-2026-002-07_Петренко_Марія_Андріївна_акт.docx`
- `ACT-2026-003-06_Шевчук_Дмитро_Ігорович_акт.docx`
- `ACT-2026-003-07_Шевчук_Дмитро_Ігорович_акт.docx`

## Failures

None.
