import os
import re
import io
import datetime
import pymupdf  # PyMuPDF
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

# The exact 14 extracted fields from the SRF + Source File (15 total columns in Excel)
# Exactly matching extract_to_excel.py and Site_Request_Summary.xlsx
CORE_FIELDS = [
    "Smart Site ID",
    "Smart Site Name",
    "Site Proposal Type",
    "Site Proposal Date",
    "Tower Co. Name",
    "Site Classification (Uptime SLA)",
    "Site Type",
    "Tower Type",
    "Tower Height (m)",
    "Building Height (m)",
    "Latitude",
    "Longitude",
    "Site Address",
    "Ref. No."
]

ALL_COLUMNS = CORE_FIELDS + ["Source File"]

KNOWN_LABELS = [
    "Site Proposal Date:",
    "Site Proposal Type:",
    "Site Classification (Uptime SLA):",
    "Tower Co. Name:",
    "Smart Site ID:",
    "Smart Site Name:",
    "Site Address:",
    "Latitude:",
    "Longitude:",
    "Site Type (Ground-based/ Rooftop):",
    "Tower Type:",
    "Tower Height (m):",
    "Building Height (m):",
    "Proposed Loadings:",
    "EQUIPMENT FOR DISMANTLING",
    "No. of Elements:",
    "Weight of Antenna/Equipment:",
    "Total EPA",
    "Total DC Power",
    "ADDITIONAL EQUIPMENT",
    "SITE REQUEST FORM",
    "Ref. No.:",
    "DOCUMENT REVISION:"
]

def clean_extracted_value(val):
    if not val:
        return ""
    val = str(val).strip()
    if val.endswith(":"):
        return ""
    for label in KNOWN_LABELS:
        if val.lower() == label.lower() or val.lower() == label.rstrip(":").lower():
            return ""
    return val

def extract_text_from_pdf(pdf_source):
    if isinstance(pdf_source, (bytes, bytearray)):
        doc = pymupdf.open(stream=pdf_source, filetype="pdf")
    else:
        doc = pymupdf.open(pdf_source)

    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"
    return doc, full_text

def extract_pdf_data(pdf_source, filename="Unknown.pdf"):
    doc, full_text = extract_text_from_pdf(pdf_source)

    data = {
        "Smart Site ID": "",
        "Smart Site Name": "",
        "Site Proposal Type": "",
        "Site Proposal Date": "",
        "Tower Co. Name": "",
        "Site Classification (Uptime SLA)": "",
        "Site Type": "",
        "Tower Type": "",
        "Tower Height (m)": "",
        "Building Height (m)": "",
        "Latitude": "",
        "Longitude": "",
        "Site Address": "",
        "Ref. No.": "",
        "Source File": os.path.basename(filename)
    }

    # Extract Ref. No.
    ref_match = re.search(r"Ref\.\s*No\.:\s*(.*?)(?=\s*DOCUMENT\s+REVISION|\n\s*\n)", full_text, re.DOTALL | re.IGNORECASE)
    if ref_match:
        lines = [l.strip() for l in ref_match.group(1).splitlines() if l.strip()]
        ref_no = ""
        for l in lines:
            if not ref_no:
                ref_no = l
            elif ref_no.endswith("-"):
                ref_no += l
            else:
                ref_no += " " + l
        data["Ref. No."] = ref_no

    patterns = {
        "Smart Site ID": r"Smart Site ID:\s*\n\s*([^\r\n]+)",
        "Smart Site Name": r"Smart Site Name:\s*\n\s*([^\r\n]+)",
        "Site Proposal Type": r"Site Proposal Type:?\s*\n?\s*([^\r\n]+)",
        "Site Proposal Date": r"Site Proposal Date:?\s*\n?\s*([^\r\n]+)",
        "Tower Co. Name": r"Tower Co\. Name:\s*\n\s*([^\r\n]+)",
        "Site Classification (Uptime SLA)": r"Site Classification \(Uptime SLA\):\s*\n\s*([^\r\n]+)",
        "Site Address": r"Site Address:\s*\n\s*([^\r\n]+)",
        "Latitude": r"Latitude:\s*\n\s*([^\r\n]+)",
        "Longitude": r"Longitude:\s*\n\s*([^\r\n]+)",
        "Site Type": r"Site Type \(Ground-based/\s*Rooftop\):\s*\n\s*([^\r\n]+)",
        "Tower Height (m)": r"Tower Height \(m\):\s*\n\s*([^\r\n]+)",
        "Building Height (m)": r"Building Height \(m\):\s*\n\s*([^\r\n]+)",
    }

    for key, pat in patterns.items():
        m = re.search(pat, full_text, re.IGNORECASE)
        if m:
            val = clean_extracted_value(m.group(1))
            data[key] = val

    # Tower Type handling
    tower_type_match = re.search(r"Tower Type:.*?\n(?:Bipod, etc\.\)\s*\n)?\s*([^\r\n]+)", full_text, re.IGNORECASE)
    if tower_type_match:
        data["Tower Type"] = clean_extracted_value(tower_type_match.group(1))

    return data

def create_styled_workbook(records, output_destination):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Site Requests"
    ws.views.sheetView[0].showGridLines = True

    columns_spec = [
        ("Smart Site ID", 18),
        ("Smart Site Name", 35),
        ("Site Proposal Type", 42),
        ("Site Proposal Date", 20),
        ("Tower Co. Name", 20),
        ("Site Classification (Uptime SLA)", 32),
        ("Site Type", 18),
        ("Tower Type", 18),
        ("Tower Height (m)", 18),
        ("Building Height (m)", 20),
        ("Latitude", 16),
        ("Longitude", 16),
        ("Site Address", 55),
        ("Ref. No.", 58),
        ("Source File", 58),
    ]

    # Styles - Official FTAP Palette
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    primary_header_fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")    # Frontier Crimson
    secondary_header_fill = PatternFill(start_color="1F2428", end_color="1F2428", fill_type="solid")  # Frontier Corporate Charcoal

    data_font = Font(name="Calibri", size=11, color="1F2029")
    data_font_bold = Font(name="Calibri", size=11, bold=True, color="C00000")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    # Header Row
    ws.row_dimensions[1].height = 32
    for col_idx, (col_name, width) in enumerate(columns_spec, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        if col_name in ["Smart Site ID", "Smart Site Name", "Site Proposal Type"]:
            cell.fill = primary_header_fill
        else:
            cell.fill = secondary_header_fill
        cell.alignment = center_align
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Data Rows
    for row_idx, record in enumerate(records, start=2):
        ws.row_dimensions[row_idx].height = 25
        row_fill = PatternFill(
            start_color="F8FAFC" if row_idx % 2 == 0 else "FFFFFF",
            end_color="F8FAFC" if row_idx % 2 == 0 else "FFFFFF",
            fill_type="solid"
        )
        for col_idx, (col_name, _) in enumerate(columns_spec, start=1):
            val = record.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)

            if col_name in ["Smart Site ID", "Site Proposal Type"]:
                cell.font = data_font_bold
            else:
                cell.font = data_font

            cell.fill = row_fill
            cell.border = thin_border

            if col_name in [
                "Smart Site ID", "Site Proposal Type", "Site Proposal Date",
                "Tower Co. Name", "Site Classification (Uptime SLA)", "Site Type",
                "Tower Type", "Tower Height (m)", "Building Height (m)",
                "Latitude", "Longitude"
            ]:
                cell.alignment = center_align
            else:
                cell.alignment = left_align

    ws.freeze_panes = "A2"

    if isinstance(output_destination, str):
        wb.save(output_destination)
    else:
        wb.save(output_destination)

def generate_excel_bytes(records):
    buf = io.BytesIO()
    create_styled_workbook(records, buf)
    buf.seek(0)
    return buf.getvalue()
