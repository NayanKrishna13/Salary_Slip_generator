#!/usr/bin/env python3
import argparse
import os
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any

from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Generate a PDF payslip from an Excel salary sheet by employee code."
        )
    )
    parser.add_argument(
        "--excel-path",
        required=True,
        help="Absolute or relative path to the Excel file (e.g., /path/to/salary.xlsx)",
    )
    parser.add_argument(
        "--sheet-name",
        default=0,
        help=(
            "Excel sheet name or index (default 0). Use the sheet that contains the salary table."
        ),
    )
    parser.add_argument(
        "--employee-code",
        required=True,
        help="Employee code to lookup (as shown in the 'Code' column)",
    )
    parser.add_argument(
        "--output-dir",
        default="./output",
        help="Directory where the generated PDF will be saved (default: ./output)",
    )
    parser.add_argument(
        "--month",
        default=None,
        help=(
            "Optional month string to show on the payslip header (e.g., 'Jul 2025')."
        ),
    )
    parser.add_argument(
        "--company-name",
        default="",
        help="Optional company/department name to show on the payslip header.",
    )
    parser.add_argument(
        "--currency-symbol",
        default="₹",
        help="Currency symbol to prefix amounts (default: ₹).",
    )
    return parser.parse_args()


def normalize_column(name: str) -> str:
    return "".join(ch for ch in str(name).strip().lower() if ch.isalnum())


def find_code_column(columns: List[str]) -> Optional[str]:
    normalized_to_original: Dict[str, str] = {normalize_column(c): c for c in columns}
    candidates = [
        "code",
        "employeecode",
        "empcode",
        "employeeno",
        "employeenumber",
        "ecode",
    ]
    for cand in candidates:
        if cand in normalized_to_original:
            return normalized_to_original[cand]
    # Fallback: pick any column that contains 'code'
    for n, original in normalized_to_original.items():
        if "code" in n:
            return original
    return None


def format_amount(value: Any, currency_symbol: str) -> str:
    if value is None or value == "":
        return "-"
    try:
        number = float(value)
    except Exception:
        return str(value)
    if number.is_integer():
        return f"{currency_symbol} {int(number):,}"
    return f"{currency_symbol} {number:,.2f}"


def read_sheet_headers_and_rows(excel_path: str, sheet_name_or_index: Any) -> Tuple[List[str], List[List[Any]]]:
    wb = load_workbook(excel_path, data_only=True)
    if isinstance(sheet_name_or_index, int) or (isinstance(sheet_name_or_index, str) and sheet_name_or_index.isdigit()):
        idx = int(sheet_name_or_index)
        ws = wb[wb.sheetnames[idx]]
    else:
        ws = wb[str(sheet_name_or_index)]

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []

    # Find first non-empty row to serve as header
    header_row: Optional[List[Any]] = None
    header_index: int = 0
    for i, r in enumerate(rows):
        if any(cell is not None and str(cell).strip() != "" for cell in r):
            header_row = [str(c) if c is not None else "" for c in r]
            header_index = i
            break
    if header_row is None:
        return [], []

    # Remaining rows are data
    data_rows = rows[header_index + 1 :]
    # Trim trailing None values for each row to align with header length
    header_len = len(header_row)
    normalized_rows: List[List[Any]] = []
    for r in data_rows:
        row_values = list(r) + [None] * (header_len - len(r))
        normalized_rows.append(row_values[:header_len])

    return header_row, normalized_rows


def find_matching_row(headers: List[str], rows: List[List[Any]], code_col: str, target_code: str) -> Optional[int]:
    if code_col not in headers:
        return None
    col_idx = headers.index(code_col)
    for i, r in enumerate(rows):
        value = r[col_idx]
        if value is None:
            continue
        if str(value).strip().lower() == target_code:
            return i
    return None


def row_to_key_values(headers: List[str], row: List[Any], currency_symbol: str) -> List[Tuple[str, str]]:
    kv: List[Tuple[str, str]] = []
    for col, value in zip(headers, row):
        if isinstance(value, (int, float)) and normalize_column(col) not in {"slno", "sl", "dp", "glis"}:
            display = format_amount(value, currency_symbol)
        else:
            display = "-" if value in (None, "") else str(value)
        kv.append((str(col), display))
    return kv


def make_table_data(kv_pairs: List[Tuple[str, str]]) -> List[List[str]]:
    rows: List[List[str]] = [["Field", "Value"]]
    for key, value in kv_pairs:
        rows.append([key, value])
    return rows


def build_pdf(
    output_pdf_path: str,
    header_title: str,
    company_name: Optional[str],
    table_data: List[List[str]],
) -> None:
    os.makedirs(os.path.dirname(output_pdf_path), exist_ok=True)

    doc = SimpleDocTemplate(
        output_pdf_path,
        pagesize=A4,
        rightMargin=20 * mm,
        leftMargin=20 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title=header_title,
        author=company_name or "",
    )

    styles = getSampleStyleSheet()
    story: List[Any] = []

    if company_name:
        story.append(Paragraph(f"<b>{company_name}</b>", styles["Title"]))
        story.append(Spacer(1, 4 * mm))

    story.append(Paragraph(header_title, styles["Heading2"]))
    story.append(Spacer(1, 4 * mm))

    table = Table(table_data, colWidths=[70 * mm, None])

    style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
    ])

    for row_idx in range(1, len(table_data)):
        if row_idx % 2 == 0:
            style.add("BACKGROUND", (0, row_idx), (-1, row_idx), colors.whitesmoke)

    table.setStyle(style)
    story.append(table)

    doc.build(story)


def main() -> None:
    args = parse_args()

    month_str = args.month or datetime.now().strftime("%b %Y")

    headers, rows = read_sheet_headers_and_rows(args.excel_path, args.sheet_name)
    if not headers or not rows:
        raise SystemExit("The selected sheet appears to be empty or has no headers.")

    code_col = "Code"
    if code_col not in headers:
        raise SystemExit("The sheet must contain a 'Code' column.")

    target_code = str(args.employee_code).strip().lower()
    match_idx = find_matching_row(headers, rows, code_col, target_code)
    if match_idx is None:
        sample_values = ", ".join(sorted({str(r[headers.index(code_col)]) for r in rows if r[headers.index(code_col)] is not None})[:20])
        raise SystemExit(
            f"No employee found with code '{args.employee_code}'. Available samples: {sample_values} ..."
        )

    row = rows[match_idx]

    # Reorder to show Name and Code first if present
    ordered_headers = list(headers)

    def move_front(col_name: str) -> None:
        if col_name in ordered_headers:
            ordered_headers.insert(0, ordered_headers.pop(ordered_headers.index(col_name)))

    for possible_name_col in ["Name", "Employee Name", "Employee", "Emp Name"]:
        if possible_name_col in ordered_headers:
            move_front(possible_name_col)
            break

    move_front(code_col)

    # Rebuild row according to new order
    ordered_row = [row[headers.index(c)] for c in ordered_headers]

    kv_pairs = row_to_key_values(ordered_headers, ordered_row, currency_symbol=args.currency_symbol)

    table_data = make_table_data(kv_pairs)

    emp_identifier = str(row[headers.index(code_col)])
    emp_name = None
    for c in headers:
        if normalize_column(c) in {"name", "employeename", "empname"}:
            emp_name = str(row[headers.index(c)])
            break

    header_title = (
        f"Salary Statement for {month_str}"
        + (f" — {emp_name}" if emp_name else "")
        + f" (Code: {emp_identifier})"
    )

    output_filename = f"payslip_{emp_identifier}.pdf"
    output_path = os.path.join(args.output_dir, output_filename)

    build_pdf(
        output_pdf_path=output_path,
        header_title=header_title,
        company_name=args.company_name,
        table_data=table_data,
    )

    print(f"Generated: {os.path.abspath(output_path)}")


if __name__ == "__main__":
    main()