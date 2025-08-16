#!/usr/bin/env python3
import math
import os
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Union

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from num2words import num2words


DEFAULT_FONT_REGULAR_CANDIDATES = [
	"/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
	"/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
]
DEFAULT_FONT_BOLD_CANDIDATES = [
	"/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
	"/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
]


def load_font(preferred_path: Optional[str], fallback_candidates: List[str], size: int) -> ImageFont.FreeTypeFont:

	if preferred_path and os.path.isfile(preferred_path):
		try:
			return ImageFont.truetype(preferred_path, size=size)
		except Exception:
			pass
	for path in fallback_candidates:
		if os.path.isfile(path):
			try:
				return ImageFont.truetype(path, size=size)
			except Exception:
				continue
	# Final fallback
	return ImageFont.load_default()


def format_currency(value: float) -> str:
	if value is None or (isinstance(value, float) and math.isnan(value)):
		value = 0.0
	return f"{float(value):,.2f}"


def amount_in_words(value: float) -> str:
	try:
		integer = int(round(float(value)))
		words = num2words(integer, lang="en_IN").replace(" and ", " ")
		words = words.replace(",", "")
		words = words.title()
		return f"Rupees {words} Only"
	except Exception:
		return "Rupees Zero Only"


def safe_get(row: pd.Series, key: str, default: str = "") -> str:
	if key in row and pd.notna(row[key]):
		return str(row[key])
	return default


def safe_get_any(row: pd.Series, keys: List[str], default: str = "") -> str:
	for key in keys:
		if key in row and pd.notna(row[key]):
			return str(row[key])
	return default


def safe_get_num(row: pd.Series, key: str, default: float = 0.0) -> float:
	if key in row and pd.notna(row[key]):
		try:
			return float(row[key])
		except Exception:
			return default
	return default


def get_num_any(row: pd.Series, keys: List[str], default: float = 0.0) -> float:
	for key in keys:
		if key in row and pd.notna(row[key]):
			try:
				return float(row[key])
			except Exception:
				continue
	return default


def get_earning_pair(row: pd.Series, base_keys: List[str]) -> Tuple[float, float]:
	# Try explicit master/actual suffixed columns first for any base variant
	for base in base_keys:
		master_key = f"{base}_master"
		actual_key = f"{base}_actual"
		has_master = master_key in row and pd.notna(row[master_key])
		has_actual = actual_key in row and pd.notna(row[actual_key])
		if has_master or has_actual:
			try:
				master_val = float(row[master_key]) if has_master else float(row[actual_key]) if has_actual else 0.0
			except Exception:
				master_val = 0.0
			try:
				actual_val = float(row[actual_key]) if has_actual else master_val
			except Exception:
				actual_val = master_val
			return master_val, actual_val
	# Fallback: single column provided (no suffix). Use same value for both
	for base in base_keys:
		if base in row and pd.notna(row[base]):
			try:
				val = float(row[base])
			except Exception:
				val = 0.0
			return val, val
	return 0.0, 0.0

# ---- Header and join normalization helpers ----

def _normalize_header_name(name: str) -> str:
	return "".join(ch for ch in str(name).strip().lower() if ch.isalnum())


def _build_normalized_column_map(df: pd.DataFrame) -> Dict[str, str]:
	mapping: Dict[str, str] = {}
	for col in df.columns:
		norm = _normalize_header_name(col)
		if norm and norm not in mapping:
			mapping[norm] = col
	return mapping


def draw_centered_text(draw: ImageDraw.ImageDraw, text: str, center_x: int, y: int, font: ImageFont.FreeTypeFont, fill: Tuple[int, int, int] = (0, 0, 0)) -> int:

	bbox = draw.textbbox((0, 0), text, font=font)
	width = bbox[2] - bbox[0]
	draw.text((center_x - width // 2, y), text, font=font, fill=fill)
	return bbox[3] - bbox[1]


def draw_key_value(draw: ImageDraw.ImageDraw, x_key: int, x_val: int, y: int, key_text: str, val_text: str, key_font: ImageFont.FreeTypeFont, val_font: ImageFont.FreeTypeFont, key_fill=(0, 0, 0), val_fill=(0, 0, 0)) -> int:

	draw.text((x_key, y), key_text, font=key_font, fill=key_fill)
	draw.text((x_val, y), val_text, font=val_font, fill=val_fill)
	bbox = draw.textbbox((x_val, y), val_text, font=val_font)
	return bbox[3] - bbox[1]


def read_dataframe(input_path: str, use_first_row_as_header: bool = True, sheet_name: Optional[Union[str, int]] = None, header_row_index: Optional[int] = None) -> pd.DataFrame:

	ext = os.path.splitext(input_path.lower())[1]
	if ext in [".xlsx", ".xlsm", ".xls"]:
		if sheet_name is not None:
			df = pd.read_excel(input_path, header=None, sheet_name=sheet_name)
		else:
			df = pd.read_excel(input_path, header=None)
	elif ext in [".csv", ".tsv"]:
		sep = "," if ext == ".csv" else "\t"
		df = pd.read_csv(input_path, sep=sep, header=None)
	else:
		raise ValueError(f"Unsupported input file type: {ext}")

	if use_first_row_as_header:
		# If explicit header row is provided, use it
		if header_row_index is not None:
			if header_row_index < 0 or header_row_index >= len(df):
				return pd.DataFrame()
			header_series = df.iloc[header_row_index].astype(str).str.strip()
			df = df.iloc[header_row_index + 1:].copy()
			df.columns = header_series.values
			df.reset_index(drop=True, inplace=True)
			return df
		# Otherwise, pick the first row with at least 2 non-empty cells to avoid title rows
		candidate_idx: Optional[int] = None
		for r in range(len(df)):
			row = df.iloc[r]
			if row.notna().sum() >= 2:
				candidate_idx = r
				break
		if candidate_idx is None:
			return pd.DataFrame()
		header_series = df.iloc[candidate_idx].astype(str).str.strip()
		df = df.iloc[candidate_idx + 1:].copy()
		df.columns = header_series.values
		df.reset_index(drop=True, inplace=True)
	return df


def ensure_output_dir(path: str) -> None:
	os.makedirs(path, exist_ok=True)


def render_payslip(
	row: pd.Series,
	output_path: str,
	logo_path: Optional[str],
	regular_font_path: Optional[str],
	bold_font_path: Optional[str],
	month_text: Optional[str] = None,
	image_width: int = 1600,
	image_height: int = 1120,
	background_color: Tuple[int, int, int] = (255, 255, 255),
	output_format: str = "png",
	output_filename: Optional[str] = None,
) -> None:
	image = Image.new("RGB", (image_width, image_height), color=background_color)

	draw = ImageDraw.Draw(image)

	# Fonts
	font_title = load_font(bold_font_path, DEFAULT_FONT_BOLD_CANDIDATES, 44)
	font_subtitle = load_font(regular_font_path, DEFAULT_FONT_REGULAR_CANDIDATES, 24)

	font_section = load_font(bold_font_path, DEFAULT_FONT_BOLD_CANDIDATES, 26)
	font_label = load_font(bold_font_path, DEFAULT_FONT_BOLD_CANDIDATES, 22)
	font_value = load_font(regular_font_path, DEFAULT_FONT_REGULAR_CANDIDATES, 22)

	font_table_header = load_font(bold_font_path, DEFAULT_FONT_BOLD_CANDIDATES, 22)

	font_table_cell = load_font(regular_font_path, DEFAULT_FONT_REGULAR_CANDIDATES, 22)

	font_footer = load_font(regular_font_path, DEFAULT_FONT_REGULAR_CANDIDATES, 18)


	margin = 40
	center_x = image_width // 2
	y = margin

	# Logo
	if logo_path and os.path.isfile(logo_path):
		try:
			logo = Image.open(logo_path).convert("RGBA")
			max_logo_h = 120
			ratio = max_logo_h / logo.height
			new_w = int(logo.width * ratio)
			logo = logo.resize((new_w, max_logo_h), Image.LANCZOS)
			image.paste(logo, (margin, y), mask=logo)
		except Exception:
			pass

	# Header
	draw_centered_text(draw, "Centre for the Study of Developing Societies", center_x, y + 10, font_title)

	y += 70
	draw_centered_text(draw, "29 Rajpur Road civil lines delhi 110054.", center_x, y, font_subtitle)

	y += 35

	# Month
	if not month_text:
		month_text = safe_get(row, "Month", "")
		if not month_text:
			# Try to infer from data if available, otherwise use current month
			now = datetime.now()
			month_text = now.strftime("%B %Y")
	y += draw_centered_text(draw, f"Payslip for the month of {month_text}", center_x, y + 10, font_section)

	y += 20

	# Details box
	box_top = y + 10
	box_left = margin
	box_right = image_width - margin
	box_bottom = box_top + 210
	draw.rectangle([box_left, box_top, box_right, box_bottom], outline=(0, 0, 0), width=2)


	# Vertical divider
	mid_x = (box_left + box_right) // 2
	draw.line([(mid_x, box_top), (mid_x, box_bottom)], fill=(0, 0, 0), width=2)

	# Left column fields
	left_fields = [
		("Name:", safe_get(row, "Name")),
		("Joining Date:", safe_get(row, "Joining Date")),
		("Designation:", safe_get(row, "Designation")),
		("Department:", safe_get(row, "Department")),
		("Location:", safe_get(row, "Location")),
		("Effective Work Days:", safe_get(row, "Effective Work Days")),
		("LOP:", safe_get(row, "LOP")),
	]

	employee_no_val = safe_get_any(row, ["Employee No", "Employee Number", "Code"]) or ""


	right_fields = [
		("Employee No:", employee_no_val),
		("Bank Name:", safe_get(row, "Bank Name")),
		("Bank Account No:", safe_get(row, "Bank Account No")),
		("PAN Number:", safe_get(row, "PAN Number")),
		("PF No:", ""),
		("PF UAN:", ""),
	]

	row_height = 28
	pad_x = 14
	key_x_left = box_left + pad_x
	val_x_left = box_left + 220
	key_x_right = mid_x + pad_x
	val_x_right = mid_x + 220

	cursor_y = box_top + 12
	for key, val in left_fields:
		draw_key_value(draw, key_x_left, val_x_left, cursor_y, key, val, font_label, font_value)

		cursor_y += row_height

	cursor_y = box_top + 12
	for key, val in right_fields:
		draw_key_value(draw, key_x_right, val_x_right, cursor_y, key, val, font_label, font_value)

		cursor_y += row_height

	y = box_bottom + 20

	# Earnings and Deductions tables
	table_left = margin
	table_right = image_width - margin
	table_top = y

	# Column geometry
	earn_right = table_left + int(0.58 * (table_right - table_left))
	# Earnings sub-table columns: Title | Master | Actual
	earn_col1 = table_left + 10
	# Treat these as RIGHT edges for numeric alignment
	earn_col2 = earn_right - 200
	earn_col3 = earn_right - 60

	ded_col1 = earn_right + 10
	# Treat as RIGHT edge for numeric alignment
	ded_col2 = table_right - 100

	header_height = 34
	line_height = 28
	content_start_y = table_top + header_height + 8

	# Build rows first so we can size the box dynamically
	earnings_items = [
		("BASIC",) + get_earning_pair(row, ["BASIC", "BASIC "]),
		("DA",) + get_earning_pair(row, ["DA", "DA "]),
		("HRA",) + get_earning_pair(row, ["HRA", "HRA "]),
		("TRANSPORT ALLOWANCE",) + get_earning_pair(row, ["TRANSPORT_ALLOWANCE", "TRANSPORT ALLOWANCE"]),
		("DA TPT",) + get_earning_pair(row, ["DA_TPT", "DA TPT"]),
	]

	deductions_items = [
		("PF", get_num_any(row, ["PF_actual", "PF"])),
		("GLIS", get_num_any(row, ["GLIS_actual", "GLIS"])),
		("REFUND SPF", get_num_any(row, ["REFUND_SPF_actual", "REFUND_SPF", "REFUND SPF"])),
		("REFUND SWF", get_num_any(row, ["REFUND_SWF_actual", "REFUND_SWF", "REFUND SWF"])),
	]

	# Filter out zero rows (if entirely zero and empty)
	earnings_items = [item for item in earnings_items if (item[1] != 0 or item[2] != 0)]
	deductions_items = [item for item in deductions_items if item[1] != 0]

	max_rows = max(len(earnings_items), len(deductions_items), 1)
	content_height = max_rows * line_height
	totals_y = content_start_y + content_height + 10
	table_bottom = totals_y + 40

	# Draw outer box and divider
	draw.rectangle([table_left, table_top, table_right, table_bottom], outline=(0, 0, 0), width=2)
	draw.line([(earn_right, table_top), (earn_right, table_bottom)], fill=(0, 0, 0), width=2)

	# Headers
	draw.text((earn_col1, table_top + 6), "Earnings", font=font_table_header, fill=(0, 0, 0))
	draw.text((earn_col2 - 30, table_top + 6), "Master", font=font_table_header, fill=(0, 0, 0))
	draw.text((earn_col3 - 30, table_top + 6), "Actual", font=font_table_header, fill=(0, 0, 0))

	draw.text((ded_col1, table_top + 6), "Deductions", font=font_table_header, fill=(0, 0, 0))
	draw.text((ded_col2 - 20, table_top + 6), "Actual", font=font_table_header, fill=(0, 0, 0))

	# Header divider
	draw.line([(table_left, table_top + header_height), (table_right, table_top + header_height)], fill=(0, 0, 0), width=2)

	# Rows
	row_y = content_start_y
	total_earn_master = 0.0
	total_earn_actual = 0.0
	for label, master, actual in earnings_items:
		# Label left-aligned
		draw.text((earn_col1, row_y), label, font=font_table_cell, fill=(0, 0, 0))
		# Right-aligned numbers
		master_text = format_currency(master)
		master_w = draw.textbbox((0, 0), master_text, font=font_table_cell)[2]
		draw.text((earn_col2 - master_w, row_y), master_text, font=font_table_cell, fill=(0, 0, 0))
		actual_text = format_currency(actual)
		actual_w = draw.textbbox((0, 0), actual_text, font=font_table_cell)[2]
		draw.text((earn_col3 - actual_w, row_y), actual_text, font=font_table_cell, fill=(0, 0, 0))
		total_earn_master += master
		total_earn_actual += actual
		row_y += line_height

	total_deduct = 0.0
	row_y_ded = content_start_y
	for label, actual in deductions_items:
		draw.text((ded_col1, row_y_ded), label, font=font_table_cell, fill=(0, 0, 0))
		actual_text = format_currency(actual)
		actual_w = draw.textbbox((0, 0), actual_text, font=font_table_cell)[2]
		draw.text((ded_col2 - actual_w, row_y_ded), actual_text, font=font_table_cell, fill=(0, 0, 0))
		total_deduct += actual
		row_y_ded += line_height

	# Totals row
	draw.line([(table_left, totals_y), (table_right, totals_y)], fill=(0, 0, 0), width=2)
	# Totals labels and right-aligned numbers
	draw.text((earn_col1, totals_y + 8), "Total Earnings:INR.", font=font_table_header, fill=(0, 0, 0))
	te_m_text = format_currency(total_earn_master)
	te_m_w = draw.textbbox((0, 0), te_m_text, font=font_table_header)[2]
	draw.text((earn_col2 - te_m_w, totals_y + 8), te_m_text, font=font_table_header, fill=(0, 0, 0))
	te_a_text = format_currency(total_earn_actual)
	te_a_w = draw.textbbox((0, 0), te_a_text, font=font_table_header)[2]
	draw.text((earn_col3 - te_a_w, totals_y + 8), te_a_text, font=font_table_header, fill=(0, 0, 0))

	draw.text((ded_col1, totals_y + 8), "Total Deductions:INR.", font=font_table_header, fill=(0, 0, 0))
	td_text = format_currency(total_deduct)
	td_w = draw.textbbox((0, 0), td_text, font=font_table_header)[2]
	draw.text((ded_col2 - td_w, totals_y + 8), td_text, font=font_table_header, fill=(0, 0, 0))


	# Net Pay
	y = table_bottom + 20
	net_pay = total_earn_actual - total_deduct
	label_text = "Net Pay for the month:"
	draw.text((margin, y), label_text, font=font_label, fill=(0, 0, 0))
	# Bold value
	value_text = format_currency(net_pay)
	bbox_label = draw.textbbox((margin, y), label_text, font=font_label)
	draw.text((bbox_label[2] + 20, y), value_text, font=font_section, fill=(0, 0, 0))

	y += 40

	# Amount in words
	draw.text((margin, y), f"({amount_in_words(net_pay)})", font=font_value, fill=(0, 0, 0))


	# Footer
	footer_text = "This is a system generated payslip and does not require a signature"

	draw.text((center_x - draw.textlength(footer_text, font=font_footer) // 2, image_height - 50), footer_text, font=font_footer, fill=(100, 100, 100))


	# Output filename: Name and Code/Employee No
	code_val = employee_no_val
	name_val = safe_get(row, "Name", "Employee").strip()
	base = f"{name_val}_{code_val}" if code_val else name_val
	safe_base = "".join(c for c in base if c.isalnum() or c in ("-", "_"))
	if not safe_base:
		safe_base = "payslip"
	suffix = f"_{month_text.replace(' ', '_')}" if month_text else ""
	auto_filename = f"{safe_base}{suffix}"

	# Decide final path and save
	if output_filename:
		final_path = output_filename
	else:
		ext = ".pdf" if output_format.lower() == "pdf" else ".png"
		final_path = os.path.join(output_path, f"{auto_filename}{ext}")

	if output_format.lower() == "pdf":
		image.convert("RGB").save(final_path, "PDF", resolution=300.0)
	else:
		image.save(final_path)


def _detect_join_key(emp_df: pd.DataFrame, sal_df: pd.DataFrame) -> Tuple[str, str]:
	# Build normalized header maps
	emp_map = _build_normalized_column_map(emp_df)
	sal_map = _build_normalized_column_map(sal_df)
	emp_keys = set(emp_map.keys())
	sal_keys = set(sal_map.keys())
	# Candidate normalized keys, in priority order (code-like first, then name)
	candidates = [
		"code", "employeeno", "employeenumber", "empno", "empnumber", "employeeid", "empid", "employeecode",
		"name",
	]
	for cand in candidates:
		if cand in emp_keys and cand in sal_keys:
			return emp_map[cand], sal_map[cand]
	# No common normalized headers found
	raise ValueError("Could not find a common join key between employee and salary sheets. Expected one of: Code, Employee No, Employee Number, or Name")


def _normalize_join_column(df: pd.DataFrame, join_key: str, new_col: str = "__join_key__") -> pd.DataFrame:
	df = df.copy()
	df[new_col] = df[join_key].astype(str).str.strip().str.lower()
	return df


def prepare_two_file_merged_dataframe(employee_excel_path: str, salary_excel_path: str, employee_sheet_name: str = "Employee Information", salary_sheet_name: str = "sheet0") -> pd.DataFrame:
	emp_df = read_dataframe(employee_excel_path, use_first_row_as_header=True, sheet_name=employee_sheet_name)
	sal_df = read_dataframe(salary_excel_path, use_first_row_as_header=True, sheet_name=salary_sheet_name, header_row_index=1)

	if emp_df.empty:
		raise ValueError("Employee Information sheet resulted in empty DataFrame")
	if sal_df.empty:
		raise ValueError("Salary sheet resulted in empty DataFrame")

	left_key, right_key = _detect_join_key(emp_df, sal_df)

	emp_df = _normalize_join_column(emp_df, left_key)
	sal_df = _normalize_join_column(sal_df, right_key)

	# Avoid duplicate non-join columns from salary overwriting identity fields
	merged = emp_df.merge(sal_df, on="__join_key__", how="inner", suffixes=("", "_sal"))
	return merged


if __name__ == "__main__":
	raise SystemExit("CLI mode is disabled. Please run scripts/generate_payslips_interactive.py")