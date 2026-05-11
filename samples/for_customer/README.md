# Customer Sample: Software Development Services

This sample set models a Ukrainian software-development service package where a limited liability company orders development services from an individual contractor.

Input files:
- `input/developer_contractors.xlsx` - contractor and company requisites.
- `templates/service_contract.docx` - service contract template.
- `templates/act_acceptance.docx` - acceptance act template.
- `templates/invoice.docx` - invoice template.

Recommended generation commands:

```bash
python -m src.main generate --template ./samples/for_customer/templates/service_contract.docx --data ./samples/for_customer/input/developer_contractors.xlsx --output ./samples/for_customer/output/generated/contracts --filename-template "{номер_договору}_{назва_клієнта}_договір"

python -m src.main generate --template ./samples/for_customer/templates/act_acceptance.docx --data ./samples/for_customer/input/developer_contractors.xlsx --output ./samples/for_customer/output/generated/acts --filename-template "{акт_номер}_{назва_клієнта}_акт"

python -m src.main generate --template ./samples/for_customer/templates/invoice.docx --data ./samples/for_customer/input/developer_contractors.xlsx --output ./samples/for_customer/output/generated/invoices --filename-template "{рахунок_номер}_{назва_клієнта}_рахунок"
```

The templates are realistic samples for testing document generation and are not legal advice.
