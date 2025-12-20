
from datetime import datetime
from io import BytesIO
import os
import logging
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from terms_template import get_terms_section

logging.basicConfig(filename='ibm_v2_debug.log', level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

USD_TO_AED = 3.6725

def create_styled_excel_v2(
	data: list,
	header_info: dict,
	logo_path: str,
	output: BytesIO,
	compliance_text: str,
	ibm_terms_text: str
):
	"""
	Template 1 Excel generation (for v2):
	data rows: [SKU, Product Description, Quantity, Start Date, End Date, Unit Price AED, Total Price AED]
	Table data is provided from Excel, not PDF.
	"""
	wb = Workbook()
	ws = wb.active
	ws.title = "Quotation"
	ws.sheet_view.showGridLines = False

	# --- Debug: Log Raw Input Data to File ---
	with open('debug_excel_input.log', 'w', encoding='utf-8') as dbg:
		dbg.write('Row\tSKU\tDescription\tQuantity\tStart Date\tEnd Date\tCost\n')
		for idx, row in enumerate(data, start=1):
			sku = row[0] if len(row) > 0 else ""
			desc = row[1] if len(row) > 1 else ""
			qty = row[2] if len(row) > 2 else ""
			start_date = row[3] if len(row) > 3 else ""
			end_date = row[4] if len(row) > 4 else ""
			cost = row[5] if len(row) > 5 else ""
			dbg.write(f'{idx}\t{sku}\t{desc}\t{qty}\t{start_date}\t{end_date}\t{cost}\n')

	# --- Header / Branding ---
	ws.merge_cells("B1:C2")  # Move logo to row 1-2
	if logo_path and os.path.exists(logo_path):
		img = Image(logo_path)
		img.width = 1.87 * 96
		img.height = 0.56 * 96
		ws.add_image(img, "B1")
		ws.row_dimensions[1].height = 25
		ws.row_dimensions[2].height = 25
	ws.merge_cells("D3:G3")
	ws["D3"] = "Quotation"
	ws["D3"].font = Font(size=20, color="1F497D")
	ws["D3"].alignment = Alignment(horizontal="center", vertical="center")

	ws.column_dimensions[get_column_letter(2)].width  = 8
	ws.column_dimensions[get_column_letter(3)].width  = 15
	ws.column_dimensions[get_column_letter(4)].width  = 50
	ws.column_dimensions[get_column_letter(5)].width  = 10
	ws.column_dimensions[get_column_letter(6)].width  = 14
	ws.column_dimensions[get_column_letter(7)].width  = 14
	ws.column_dimensions[get_column_letter(8)].width  = 15
	ws.column_dimensions[get_column_letter(9)].width  = 15
	ws.column_dimensions[get_column_letter(10)].width = 18
	ws.column_dimensions[get_column_letter(11)].width = 15
	ws.column_dimensions[get_column_letter(12)].width = 18

	left_labels = ["Date:", "From:", "Email:", "Contact:", "", "Company:", "Attn:", "Email:"]
	left_values = [
		datetime.today().strftime('%d/%m/%Y'),
		"Sneha Lokhandwala",
		"s.lokhandwala@mindware.net",
		"+971 55 456 6650",
		"",
		header_info.get('Reseller Name', 'empty'),
		"empty",
		"empty"
	]
	row_positions = [5, 6, 7, 8, 9, 10, 11, 12]
	for row, label, value in zip(row_positions, left_labels, left_values):
		if label:
			ws[f"C{row}"] = label
			ws[f"C{row}"].font = Font(bold=True, color="1F497D")
		if value:
			ws[f"D{row}"] = value
			ws[f"D{row}"].font = Font(color="1F497D")

	right_labels = [
		"End User:", "Bid Number:", "Agreement Number:", "PA Site Number:", "",
		"Select Territory:", "Government Entity (GOE):", "Payment Terms:"
	]
	right_values = [
		header_info.get('Customer Name', ''),
		header_info.get('Bid Number', ''),
		header_info.get('PA Agreement Number', ''),
		header_info.get('PA Site Number', ''),
		"",
		header_info.get('Select Territory', ''),
		header_info.get('Government Entity (GOE)', ''),
		"As aligned with Mindware"
	]
	for row, label, value in zip(row_positions, right_labels, right_values):
		ws.merge_cells(f"H{row}:L{row}")
		ws[f"H{row}"] = f"{label} {value}"
		ws[f"H{row}"].font = Font(bold=True, color="1F497D")
		ws[f"H{row}"].alignment = Alignment(horizontal="left", vertical="center")

	# --- Table Extraction from Excel (Second Sheet) ---
	# Input 'data' must be a list of lists, each row:
	# [SKU (A), Description (B), Quantity (G), Start Date (H), End Date (I), Cost (S)]
	# All values start from Excel row 10 (pandas index 9).

	# --- Table Headers ---
	headers = [
		"Sl", "SKU", "Product Description", "Quantity", "Start Date", "End Date",
		"Unit Price in AED", "Cost (USD)", "Total Price in AED", "Partner Discount", "Partner Price in AED"
	]
	header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
	for col, header in enumerate(headers, start=2):
		ws.merge_cells(start_row=16, start_column=col, end_row=17, end_column=col)
		cell = ws.cell(row=16, column=col, value=header)
		cell.font = Font(bold=True, size=13, color="1F497D")
		cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
		cell.fill = header_fill

	row_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
	start_row = 18

	for idx, row in enumerate(data, start=1):
		excel_row = start_row + idx - 1
		# row: [SKU, Description, Quantity, Start Date, End Date, Cost]
		sku = row[0] if len(row) > 0 else ""
		desc = row[1] if len(row) > 1 else ""
		qty = row[2] if len(row) > 2 else 0
		start_date = row[3] if len(row) > 3 else ""
		end_date = row[4] if len(row) > 4 else ""

		# Ensure cost is a float (handle string or None)
		raw_cost = row[5] if len(row) > 5 else 0
		try:
			cost = float(raw_cost) if raw_cost not in (None, "", "-") else 0
		except Exception:
			cost = 0

		# Log extracted values for debugging
		logging.info(f"Row {idx}: SKU={sku}, Desc={desc}, Qty={qty}, Start={start_date}, End={end_date}, Cost={cost}")

		# Serial number
		ws.cell(row=excel_row, column=2, value=idx).font = Font(size=11, color="1F497D")
		ws.cell(row=excel_row, column=2).alignment = Alignment(horizontal="center", vertical="center")

		# Write data columns
		ws.cell(row=excel_row, column=3, value=sku).font = Font(size=11, color="1F497D")
		ws.cell(row=excel_row, column=4, value=desc).font = Font(size=11, color="1F497D")
		ws.cell(row=excel_row, column=5, value=qty).font = Font(size=11, color="1F497D")
		ws.cell(row=excel_row, column=6, value=start_date).font = Font(size=11, color="1F497D")
		ws.cell(row=excel_row, column=7, value=end_date).font = Font(size=11, color="1F497D")
		# Unit Price in AED (leave blank, will be formula)
		# Cost (USD): from Excel (col H), convert from AED to USD
		cost_usd = cost / USD_TO_AED if cost else 0
		ws.cell(row=excel_row, column=9, value=cost_usd).number_format = '"USD"#,##0.00'
		ws.cell(row=excel_row, column=9).font = Font(size=11, color="1F497D")
		ws.cell(row=excel_row, column=9).alignment = Alignment(horizontal="center", vertical="center")

		# Formulas for calculated columns
		# J: Total Price in AED = Cost (I) * USD_TO_AED
		total_formula = f"=I{excel_row}*{USD_TO_AED}"
		ws.cell(row=excel_row, column=10, value=total_formula)
		ws.cell(row=excel_row, column=10).font = Font(size=11, color="1F497D")
		ws.cell(row=excel_row, column=10).alignment = Alignment(horizontal="center", vertical="center")

		# K: Partner Discount = Unit Price (H) * 0.99 (1% discount)
		discount_formula = f"=ROUNDUP(H{excel_row}*0.99,2)"
		ws.cell(row=excel_row, column=11, value=discount_formula)
		ws.cell(row=excel_row, column=11).font = Font(size=11, color="1F497D")
		ws.cell(row=excel_row, column=11).alignment = Alignment(horizontal="center", vertical="center")

		# L: Partner Price in AED = Partner Discount (K) * Quantity (E)
		partner_price_formula = f"=K{excel_row}*E{excel_row}"
		ws.cell(row=excel_row, column=12, value=partner_price_formula)
		ws.cell(row=excel_row, column=12).font = Font(size=11, color="1F497D")
		ws.cell(row=excel_row, column=12).alignment = Alignment(horizontal="center", vertical="center")

		# Currency formatting
		for price_col in [8, 10, 11, 12]:
			ws.cell(row=excel_row, column=price_col).number_format = '"AED"#,##0.00'
		ws.cell(row=excel_row, column=9).number_format = '"USD"#,##0.00'
		for col in range(2, 2 + len(headers)):
			ws.cell(row=excel_row, column=col).fill = row_fill
		ws.cell(row=excel_row, column=4).alignment = Alignment(wrap_text=True, horizontal="left", vertical="center")

	# --- Summary rows, terms, etc. remain identical to original (copy from previous logic if needed) ---

	# Ensure BytesIO is at the start for reading
	wb.save(output)
	output.seek(0)
