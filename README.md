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

### Usage (employee code prompt only)
The Excel path and defaults are set in the script. You will only be asked for the employee code (and optionally company name/month if you want to override defaults).

```bash
python /workspace/generate_payslip.py
```

- Hardcoded defaults in `generate_payslip.py`:
  - Excel file: `/workspace/salary.xlsx`
  - Sheet: `0`
  - Output directory: `./output`
  - Currency symbol: `₹`
  - Company name: empty by default
  - Month label: current month

Output PDF name will include both employee name and code, e.g. `payslip_Chandan_Sharma_61.pdf`.

The generated file will be saved as `/workspace/output/payslip_<code>.pdf`.

### Notes
- The sheet must contain a `Code` column; the script matches employees using that column.
- Most numeric fields are formatted as currency with the `₹` symbol by default. Change with `--currency-symbol`.
- If your sheet uses a different layout, pass `--sheet-name` as the sheet title or index.