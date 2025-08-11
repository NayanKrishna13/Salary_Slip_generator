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

### Usage (interactive only)
```bash
python /workspace/generate_payslip.py
```
The program will prompt for:
- Excel path (defaults to `/workspace/salary.xlsx` or env `EXCEL_PATH`)
- Sheet name/index (default: 0)
- Employee code (required)
- Company name (optional)
- Month label (defaults to current month)
- Output directory (default: `./output`)
- Currency symbol (default: `₹`)

Optional: you can set a default Excel path via env var:
```bash
export EXCEL_PATH=/workspace/salary.xlsx
```
The generated file will be saved as `/workspace/output/payslip_<code>.pdf`.

### Notes
- The sheet must contain a `Code` column; the script matches employees using that column.
- Most numeric fields are formatted as currency with the `₹` symbol by default. Change with `--currency-symbol`.
- If your sheet uses a different layout, pass `--sheet-name` as the sheet title or index.