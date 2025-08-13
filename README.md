# Payslip Image Generator

Generate payslip images that closely match the provided format (header with logo, two-column details box, earnings/deductions tables, totals, net pay in words).

## Quick start

1) Install dependencies

```bash
python3 -m venv /workspace/.venv && source /workspace/.venv/bin/activate
pip install -r /workspace/requirements.txt
```

2) Prepare your data Excel/CSV and the logo image (PNG preferred)

- Place your input file anywhere and reference it by absolute path.
- The script expects the following columns (case sensitive):
  - Identity box (left): `Name`, `Joining Date`, `Designation`, `Department`, `Location`, `Effective Work Days`, `LOP`
  - Identity box (right): `Employee No`, `Bank Name`, `Bank Account No`, `PAN Number`, `PF No`, `PF UAN`
  - Earnings (Master/Actual pairs): `BASIC_master`, `BASIC_actual`, `DA_master`, `DA_actual`, `HRA_master`, `HRA_actual`, `TRANSPORT_ALLOWANCE_master`, `TRANSPORT_ALLOWANCE_actual`, `DA_TPT_master`, `DA_TPT_actual`
  - Deductions (Actual): `PF_actual`, `GLIS_actual`, `REFUND_SPF_actual`, `REFUND_SWF_actual`
  - Optional: `Month` (e.g., `July 2025`). You can also override via `--month`.

3) Run

```bash
python /workspace/scripts/generate_payslips.py \
  --excel /absolute/path/to/your_data.xlsx \
  --logo /absolute/path/to/csds_logo.png \
  --output-dir /workspace/output \
  --month "July 2025"
```

Notes:
- Use the same logo seen in your sample image. Pass its absolute path with `--logo`.
- The script also supports CSV input (via `--excel` when the file has a `.csv` extension).
- Fonts: the script auto-picks common system fonts; to force specific TTFs, pass `--font-regular` and `--font-bold`.
- Output: individual PNG files named `<EmployeeNo>_<Name>_<Month>.png` are written to the given `--output-dir`.

## Sample input

A small sample CSV is provided at `examples/sample_input.csv`. You can use it to validate the rendering:

```bash
python /workspace/scripts/generate_payslips.py \
  --excel /workspace/examples/sample_input.csv \
  --logo /absolute/path/to/csds_logo.png \
  --output-dir /workspace/output
```

## Customization

- Image size: `--width` and `--height` (default 1600x1120)
- Month label: `--month` overrides the value in the data file
- If you need additional earnings/deductions rows, add more columns with the same naming convention and extend the script if necessary.