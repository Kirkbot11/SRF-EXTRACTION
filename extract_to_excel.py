import os
import re
import pymupdf  # PyMuPDF
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
    val = val.strip()
    # Check if the captured value is actually the next field label
    if val.endswith(":"):
        return ""
    for label in KNOWN_LABELS:
        if val.lower() == label.lower() or val.lower() == label.rstrip(":").lower():
            return ""
    return val

def extract_pdf_data(pdf_path):
    doc = pymupdf.open(pdf_path)
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"
    
    data = {
        "File Name": os.path.basename(pdf_path),
        "Smart Site ID": "",
        "Smart Site Name": "",
        "Site Proposal Type": "",
        "Site Proposal Date": "",
        "Tower Co. Name": "",
        "Site Classification (Uptime SLA)": "",
        "Site Address": "",
        "Latitude": "",
        "Longitude": "",
        "Site Type": "",
        "Tower Type": "",
        "Tower Height (m)": "",
        "Building Height (m)": "",
        "Ref. No.": "",
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
        "Site Proposal Type": r"Site Proposal Type:\s*\n\s*([^\r\n]+)",
        "Site Proposal Date": r"Site Proposal Date:\s*\n\s*([^\r\n]+)",
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

def create_excel(records, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracted Site Data"
    ws.views.sheetView[0].showGridLines = True

    columns = [
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

    field_map = {
        "Smart Site ID": "Smart Site ID",
        "Smart Site Name": "Smart Site Name",
        "Site Proposal Type": "Site Proposal Type",
        "Site Proposal Date": "Site Proposal Date",
        "Tower Co. Name": "Tower Co. Name",
        "Site Classification (Uptime SLA)": "Site Classification (Uptime SLA)",
        "Site Type": "Site Type",
        "Tower Type": "Tower Type",
        "Tower Height (m)": "Tower Height (m)",
        "Building Height (m)": "Building Height (m)",
        "Latitude": "Latitude",
        "Longitude": "Longitude",
        "Site Address": "Site Address",
        "Ref. No.": "Ref. No.",
        "Source File": "File Name",
    }

    # Styles
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    primary_header_fill = PatternFill(start_color="104E8B", end_color="104E8B", fill_type="solid")
    secondary_header_fill = PatternFill(start_color="2F4F4F", end_color="2F4F4F", fill_type="solid")
    
    data_font = Font(name="Calibri", size=11, color="1A1A1A")
    data_font_bold = Font(name="Calibri", size=11, bold=True, color="104E8B")

    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    # Header Row
    ws.row_dimensions[1].height = 30
    for col_idx, (col_name, _) in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        if col_name in ["Smart Site ID", "Smart Site Name", "Site Proposal Type"]:
            cell.fill = primary_header_fill
        else:
            cell.fill = secondary_header_fill
        cell.alignment = center_align
        cell.border = thin_border

    # Data Rows
    for row_idx, record in enumerate(records, start=2):
        ws.row_dimensions[row_idx].height = 25
        row_fill = PatternFill(
            start_color="F4F8FB" if row_idx % 2 == 0 else "FFFFFF",
            end_color="F4F8FB" if row_idx % 2 == 0 else "FFFFFF",
            fill_type="solid"
        )
        for col_idx, (col_name, _) in enumerate(columns, start=1):
            key = field_map[col_name]
            val = record.get(key, "")
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

    # Column Widths
    for col_idx, (_, width) in enumerate(columns, start=1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "A2"
    wb.save(output_path)
    print(f"Excel file successfully saved: {output_path}")

def main():
    folder = r"c:\Users\kirk.encallado\Downloads\TEST FOLDER"
    pdf_files = sorted([f for f in os.listdir(folder) if f.lower().endswith(".pdf")])
    
    records = []
    for pdf in pdf_files:
        pdf_path = os.path.join(folder, pdf)
        data = extract_pdf_data(pdf_path)
        records.append(data)
        print("Successfully extracted data for:", pdf)
        for k, v in data.items():
            print(f"  {k}: {v}")

    output_file = os.path.join(folder, "Site_Request_Summary.xlsx")
    create_excel(records, output_file)

if __name__ == "__main__":
    main()
