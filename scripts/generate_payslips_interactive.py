#!/usr/bin/env python3
import os
from typing import Optional
from scripts.generate_payslips import read_dataframe, ensure_output_dir, render_payslip

# You can change these defaults directly in code if you prefer not to type them
DEFAULT_EXCEL_ABS_PATH = "/workspace/examples/sample_input.csv"
DEFAULT_OUTPUT_DIR_ABS_PATH = "/workspace/output"
DEFAULT_OUTPUT_FILENAME = ""  # If set to an absolute path ending with .pdf, it will be used directly
DEFAULT_LOGO_ABS_PATH = ""  # e.g. "/workspace/assets/csds_logo.png"
DEFAULT_MONTH_TEXT: Optional[str] = None  # e.g. "July 2025" to override
DEFAULT_IMAGE_WIDTH = 1600
DEFAULT_IMAGE_HEIGHT = 1120
DEFAULT_FONT_REGULAR = None  # e.g. "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
DEFAULT_FONT_BOLD = None     # e.g. "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"


def prompt_abs_path(prompt_text: str, default_value: str = "", allow_blank: bool = False) -> str:
    while True:
        raw = input(f"{prompt_text} [{default_value}]: ").strip()
        value = raw or default_value
        if value == "" and allow_blank:
            return ""
        if not os.path.isabs(value):
            print("Please provide an absolute path (starting with '/').")
            continue
        return value


def prompt_int(prompt_text: str, default_value: int) -> int:
    while True:
        raw = input(f"{prompt_text} [{default_value}]: ").strip()
        if raw == "":
            return default_value
        try:
            return int(raw)
        except ValueError:
            print("Please enter a valid integer.")


def main():
    print("Payslip PDF Generator (Interactive, by Employee Code)")
    print("Provide absolute paths. Press Enter to accept defaults shown in brackets.")

    excel_path = prompt_abs_path("Absolute path to Excel/CSV data file", DEFAULT_EXCEL_ABS_PATH)
    output_dir = prompt_abs_path("Absolute path to output directory", DEFAULT_OUTPUT_DIR_ABS_PATH)
    # Optional: exact PDF filename
    output_filename = input(f"Absolute output PDF filename (leave blank to auto-name) [{DEFAULT_OUTPUT_FILENAME or 'auto'}]: ").strip() or DEFAULT_OUTPUT_FILENAME
    if output_filename and (not os.path.isabs(output_filename) or not output_filename.lower().endswith(".pdf")):
        print("If provided, output filename must be an absolute path ending with .pdf. Ignoring.")
        output_filename = ""

    logo_path = prompt_abs_path("Absolute path to logo image (or leave blank)", DEFAULT_LOGO_ABS_PATH, allow_blank=True)

    employee_code = input("Enter Employee Code (as in 'Employee No' column): ").strip()
    if not employee_code:
        print("Employee Code is required.")
        return

    month_text = input(f"Month label to display (e.g. 'July 2025') [auto or data]: ").strip() or DEFAULT_MONTH_TEXT or None

    width = prompt_int("Image width (px)", DEFAULT_IMAGE_WIDTH)
    height = prompt_int("Image height (px)", DEFAULT_IMAGE_HEIGHT)

    font_regular = input(f"Absolute path to regular TTF font [{DEFAULT_FONT_REGULAR or 'auto'}]: ").strip() or (DEFAULT_FONT_REGULAR or None)
    if font_regular and not os.path.isabs(font_regular):
        print("Ignoring regular font path since it's not absolute.")
        font_regular = None

    font_bold = input(f"Absolute path to bold TTF font [{DEFAULT_FONT_BOLD or 'auto'}]: ").strip() or (DEFAULT_FONT_BOLD or None)
    if font_bold and not os.path.isabs(font_bold):
        print("Ignoring bold font path since it's not absolute.")
        font_bold = None

    # Validate and prepare
    if not os.path.exists(excel_path):
        print(f"Data file not found: {excel_path}")
        return

    ensure_output_dir(output_dir)

    # Load data
    try:
        df = read_dataframe(excel_path)
    except Exception as exc:
        print(f"Failed to read data: {exc}")
        return

    # Find employee by code (string compare on 'Employee No')
    if 'Employee No' not in df.columns:
        print("Column 'Employee No' not found in data.")
        return

    # Normalize to string for matching
    df['__emp_no_str__'] = df['Employee No'].astype(str).str.strip()
    match_rows = df[df['__emp_no_str__'] == str(employee_code).strip()]
    if match_rows.empty:
        print(f"No record found for Employee Code: {employee_code}")
        return

    row = match_rows.iloc[0]

    # Render a single PDF
    print(f"Generating payslip PDF for Employee Code {employee_code} ...")
    try:
        render_payslip(
            row=row,
            output_path=output_dir,
            logo_path=(logo_path or None),
            regular_font_path=font_regular,
            bold_font_path=font_bold,
            month_text=month_text,
            image_width=width,
            image_height=height,
            output_format="pdf",
            output_filename=(output_filename or None),
        )
        final_name = output_filename if output_filename else "auto-named file in output directory"
        print(f"Done. Wrote PDF to {final_name}.")
    except Exception as exc:
        print(f"Failed to render PDF: {exc}")


if __name__ == "__main__":
    main()