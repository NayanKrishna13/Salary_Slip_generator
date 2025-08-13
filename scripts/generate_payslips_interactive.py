#!/usr/bin/env python3
import os
from typing import Optional, List, Set
from scripts.generate_payslips import read_dataframe, ensure_output_dir, render_payslip

# You can change these defaults directly in code if you prefer not to type them
DEFAULT_EXCEL_ABS_PATH = "/workspace/examples/sample_input.csv"
DEFAULT_OUTPUT_DIR_ABS_PATH = "/workspace/output"
DEFAULT_OUTPUT_FILENAME = ""  # If set to an absolute path ending with .pdf, it will be used directly (only applied when exactly one employee is generated)
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


def prompt_mode() -> int:
    while True:
        print("Select mode:")
        print("  1) Generate for specific employee Code(s)")
        print("  2) Generate for ALL employees")
        choice = input("Enter 1 or 2 [1]: ").strip() or "1"
        if choice in {"1", "2"}:
            return int(choice)
        print("Invalid choice. Please enter 1 or 2.")


def parse_codes(raw_codes: str) -> List[str]:
    # Accept comma/space/newline separated codes
    tokens = []
    for part in raw_codes.replace("\n", ",").replace(" ", ",").split(","):
        t = part.strip()
        if t:
            tokens.append(t)
    # Preserve order but remove duplicates
    seen: Set[str] = set()
    result: List[str] = []
    for t in tokens:
        if t not in seen:
            seen.add(t)
            result.append(t)
    return result


def main():
    print("Payslip PDF Generator (Interactive)")
    print("Provide absolute paths. Press Enter to accept defaults shown in brackets.")

    mode = prompt_mode()

    excel_path = prompt_abs_path("Absolute path to Excel/CSV data file", DEFAULT_EXCEL_ABS_PATH)
    output_dir = prompt_abs_path("Absolute path to output directory", DEFAULT_OUTPUT_DIR_ABS_PATH)

    # Optional: exact PDF filename (only when generating exactly one employee)
    output_filename = input(f"Absolute output PDF filename (leave blank to auto-name) [{DEFAULT_OUTPUT_FILENAME or 'auto'}]: ").strip() or DEFAULT_OUTPUT_FILENAME
    if output_filename and (not os.path.isabs(output_filename) or not output_filename.lower().endswith(".pdf")):
        print("If provided, output filename must be an absolute path ending with .pdf. Ignoring.")
        output_filename = ""

    logo_path = prompt_abs_path("Absolute path to logo image (or leave blank)", DEFAULT_LOGO_ABS_PATH, allow_blank=True)

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

    # Load data (header is the first non-empty row)
    try:
        df = read_dataframe(excel_path, use_first_row_as_header=True)
    except Exception as exc:
        print(f"Failed to read data: {exc}")
        return

    # Determine which column carries the code
    code_col = None
    if 'Code' in df.columns:
        code_col = 'Code'
    elif 'Employee No' in df.columns:
        code_col = 'Employee No'
    else:
        print("No 'Code' or 'Employee No' column found in data.")
        return

    df['__code_norm__'] = df[code_col].astype(str).str.strip()

    # Mode handling
    if mode == 1:
        raw_codes = input("Enter one or more Codes (comma/space separated): ").strip()
        codes = parse_codes(raw_codes)
        if not codes:
            print("No Codes provided.")
            return

        # If exactly one and user provided explicit filename, we will use it; otherwise auto-name
        for idx_code, code in enumerate(codes):
            matches = df[df['__code_norm__'] == code]
            if matches.empty:
                print(f"! No record found for Code: {code}")
                continue
            row = matches.iloc[0]
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
                    output_filename=(output_filename if (len(codes) == 1 and output_filename) else None),
                )
                dest = output_filename if (len(codes) == 1 and output_filename) else "auto-named file in output directory"
                print(f" - Generated PDF for Code {code} -> {dest}")
            except Exception as exc:
                print(f" ! Failed to render for Code {code}: {exc}")
        print("Done.")
        return

    # Mode 2: All employees
    print("Generating PDFs for ALL employees ...")
    count = 0
    for _, row in df.iterrows():
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
                output_filename=None,
            )
            count += 1
        except Exception as exc:
            code_val = str(row.get(code_col, ""))
            print(f" ! Failed to render for Code {code_val}: {exc}")
    print(f"Done. Generated {count} PDF(s).")


if __name__ == "__main__":
    main()