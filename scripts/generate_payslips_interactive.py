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

    mode = prompt_mode()

    # Use constants set in code (no prompts for absolute paths)
    excel_path = DEFAULT_EXCEL_ABS_PATH
    output_dir = DEFAULT_OUTPUT_DIR_ABS_PATH
    output_filename = DEFAULT_OUTPUT_FILENAME  # used only for single employee case
    logo_path = DEFAULT_LOGO_ABS_PATH

    month_text = DEFAULT_MONTH_TEXT
    width = DEFAULT_IMAGE_WIDTH
    height = DEFAULT_IMAGE_HEIGHT
    font_regular = DEFAULT_FONT_REGULAR
    font_bold = DEFAULT_FONT_BOLD

    # Validate and prepare
    if not os.path.isabs(excel_path) or not os.path.exists(excel_path):
        print(f"Data file not found or not absolute: {excel_path}")
        return
    if not os.path.isabs(output_dir):
        print(f"Output directory path must be absolute: {output_dir}")
        return
    if logo_path and not os.path.isabs(logo_path):
        print(f"Logo path must be absolute when provided: {logo_path}")
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