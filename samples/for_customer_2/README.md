# Customer Sample 2: Framework Contract + Monthly Requests + Acts

This sample models a more realistic Ukrainian software-development service workflow:

- 3 individual developers.
- Each developer has one framework service contract for June-July 2026.
- Each developer has one monthly work request for June and one for July.
- Each monthly act is based on the framework contract and the relevant work request.

Input files:

- `input/developer_work_2026.xlsx` with sheets `contracts` and `monthly_work`.
- `templates/framework_service_contract.docx`.
- `templates/monthly_work_request.docx`.
- `templates/monthly_acceptance_act.docx`.

Generation commands:

```bash
python -m src.main generate --template ./samples/for_customer_2/templates/framework_service_contract.docx --data ./samples/for_customer_2/input/developer_work_2026.xlsx --sheet contracts --output ./samples/for_customer_2/output/generated/contracts --filename-template "{договір_номер}_{виконавець_піб}_договір"

python -m src.main generate --template ./samples/for_customer_2/templates/monthly_work_request.docx --data ./samples/for_customer_2/input/developer_work_2026.xlsx --sheet monthly_work --output ./samples/for_customer_2/output/generated/work_requests --filename-template "{заявка_номер}_{виконавець_піб}_заявка"

python -m src.main generate --template ./samples/for_customer_2/templates/monthly_acceptance_act.docx --data ./samples/for_customer_2/input/developer_work_2026.xlsx --sheet monthly_work --output ./samples/for_customer_2/output/generated/acts --filename-template "{акт_номер}_{виконавець_піб}_акт"
```

Expected output: 3 contracts, 6 work requests, and 6 acceptance acts.

These files are sample templates for document-generation testing and should be reviewed by a Ukrainian lawyer and accountant before real use.
