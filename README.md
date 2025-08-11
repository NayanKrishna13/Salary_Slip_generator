## Payslip PDF Generator

Generate a formatted PDF payslip for a single employee from an Excel salary sheet.

### Prerequisites
- Python 3.9+
- Install dependencies:

```bash
pip install -r requirements.txt
```

### Usage
Assuming your Excel file is at `/workspace/salary.xlsx` and contains a `Code` column:

```bash
python /workspace/generate_payslip.py \
  --excel-path /workspace/salary.xlsx \
  # or set env var once: EXCEL_PATH=/workspace/salary.xlsx
  --sheet-name 0 \
  --employee-code 61 \
  --company-name "Your Org Name" \
  --month "Jul 2025" \
  --output-dir /workspace/output
```

The generated file will be saved as `/workspace/output/payslip_<code>.pdf`.

### Notes
- The sheet must contain a `Code` column; the script matches employees using that column.
- Most numeric fields are formatted as currency with the `₹` symbol by default. Change with `--currency-symbol`.
- If your sheet uses a different layout, pass `--sheet-name` as the sheet title or index.