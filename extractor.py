import os
import re
import io
import datetime
import pymupdf  # PyMuPDF
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

ALL_COLUMNS = [
    "Filename", "Document Type", "Reference Number", "Extraction Date", "Change Request Initiator",
    "Site Proposal Date", "Site Proposal Type", "Site Classification", "Tower Company Name", "Site ID",
    "Site Name", "Site Address", "Latitude", "Longitude", "Site Type",
    "Tower Type", "Tower Height M", "Building Height M", "Proposed Loadings", "Existing Loadings Description",
    "Dismantling Equipment List", "Dismantling Number Of Elements", "Dismantling Total Weight KG", "Dismantling Total EPA SQM", "Dismantling Total DC Power Watts",
    "Dismantling Total Ground Space", "Smart Equipment Housing Type", "RF Antenna Dimension MM", "RF Antenna Number Of Antennas", "RF Antenna Azimuth",
    "RF Antenna Center Line Height AGL M", "RF Antenna Weight", "BTS BBU Number And U Space", "RRH RRU Number", "Access Antenna Mounting Bracket",
    "Transmission Solution", "MW Radio Antenna Diameter M", "MW Radio Antenna Height AGL M", "MW Radio Antenna Azimuth", "MW Radio Antenna Weight KG",
    "Transmission Antenna Mounting Bracket", "Transmission Rack", "Floor Space", "Cabinet Dimensions MM", "Cabinet Number Of Cabinets",
    "Cabinet Weight KG", "OLT Dimension MM", "OLT Number", "Additional Total EPA SQM", "Additional Total DC Power Watts",
    "Additional Total Ground Space SQM", "Declared Total EPA SQM", "Declared Total DC Power Watts", "Declared Total Ground Space SQM", "Power Declared Design Capacity",
    "EPA Declared Design Capacity", "Ground Space Declared Design Capacity", "Sf Power Remarks", "Sf Infra Remarks", "AC Power",
    "DC Power", "Cooling", "Space", "Other Requirement Notes", "Implementation Target Date",
    "Approval 1 Department", "Approval 1 Level", "Approval 1 Approver Name", "Approval 1 Approval Date",
    "Approval 2 Department", "Approval 2 Level", "Approval 2 Approver Name", "Approval 2 Approval Date",
    "Approval 3 Department", "Approval 3 Level", "Approval 3 Approver Name", "Approval 3 Approval Date",
    "Approval 4 Department", "Approval 4 Level", "Approval 4 Approver Name", "Approval 4 Approval Date",
    "Approval 5 Department", "Approval 5 Level", "Approval 5 Approver Name", "Approval 5 Approval Date",
    "Approval 6 Department", "Approval 6 Level", "Approval 6 Approver Name", "Approval 6 Approval Date",
    "Verification 1 Role", "Verification 1 Name", "Verification 1 Date",
    "Verification 2 Role", "Verification 2 Name", "Verification 2 Date",
    "Verification 3 Role", "Verification 3 Name", "Verification 3 Date",
    "Verification 4 Role", "Verification 4 Name", "Verification 4 Date"
]

LABEL_SUBSTRINGS = [
    "total ground space", "total epa", "total dc power", "existing loadings",
    "proposed loadings", "no. of", "smart equipment", "additional equipment",
    "smart declared", "declared design capacity", "smart site", "tower co.", "latitude:",
    "longitude:", "site type", "tower type", "tower height", "building height",
    "site classification", "transmission solution", "access antenna", "transmission antenna", "transmission rack",
    "floor space", "equipment/cabinet", "olt dimension", "power requirement",
    "cooling requirement", "space requirement", "other requirement", "project implementation",
    "document revision", "site request form", "infrastructure details", "infrastructure",
    "general site information", "fixed / wireless approvals", "reviewed by", "approver 1",
    "approver signature", "approver name", "approver date", "equipment for dismantling",
    "site proposal date", "site proposal type"
]

SECTION_REGEXES = [
    r"^[A-Z]\.\s+.*",
    r"^\b(?:EQUIPMENT\s+FOR\s+DISMANTLING|ADDITIONAL\s+EQUIPMENT|SMART\s+DECLARED|FIXED\s*/\s*WIRELESS\s+APPROVALS|REVIEWED\s+BY)\b",
    r".*\bINFRASTRUCTURE\s+DETAILS\b.*"
]

def clean_val(v, field_name=None):
    if v is None:
        return ""
    v = str(v).strip()
    if not v or v == "" or v.endswith(":"):
        return ""
    v_low = v.lower()

    if field_name == "Site Proposal Type":
        if v_low in ["site proposal type", "site proposal date", "site classification", "tower co. name"]:
            return ""

    for ls in LABEL_SUBSTRINGS:
        if ls in v_low:
            return ""
    for s_reg in SECTION_REGEXES:
        if re.search(s_reg, v, re.IGNORECASE):
            return ""
    return v

def extract_text_from_pdf(pdf_source):
    if isinstance(pdf_source, (bytes, bytearray)):
        doc = pymupdf.open(stream=pdf_source, filetype="pdf")
    else:
        doc = pymupdf.open(pdf_source)

    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"

    # OCR Fallback if scanned / no digital text
    if not full_text.strip():
        try:
            import winocr
            import asyncio
            from PIL import Image
            import io
            
            ocr_text_parts = []
            for page in doc:
                pix = page.get_pixmap(dpi=200)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                res = asyncio.run(winocr.recognize_pil(img, 'en'))
                ocr_text_parts.append(res.get('text', ''))
            full_text = "\n".join(ocr_text_parts)
        except Exception:
            pass

    return doc, full_text

def extract_pdf_data(pdf_source, filename="Unknown.pdf"):
    doc, full_text = extract_text_from_pdf(pdf_source)
    pages = [p for p in doc]

    # Initialize all 101 columns with blank string ""
    row = {c: "" for c in ALL_COLUMNS}
    row["Filename"] = os.path.basename(filename)
    row["Document Type"] = "Site Request Form (SRF)"
    row["Extraction Date"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Reference Number
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
        row["Reference Number"] = ref_no

    # Proposed Loadings Checkbox
    m_prop = re.search(r"Proposed\s+Loadings:\s*(.*?)(?=\n[A-Z]|\n\n|\Z)", full_text, re.DOTALL | re.IGNORECASE)
    if m_prop:
        area = m_prop.group(1)
        for opt in ["Both", "Dismantling", "Additional", "N/A"]:
            if f"◼{opt}" in area or f"■{opt}" in area or re.search(r"[◼■▪☑\[xX]]\s*" + re.escape(opt), area):
                row["Proposed Loadings"] = opt
                break

    # Spatial alignment extraction across pages (handles side-by-side table cells and stacked cells)
    def get_field_aligned_value(label_pat, y_tol=20):
        for page in pages:
            blocks = page.get_text("blocks")
            for b in blocks:
                text = b[4].strip()
                # 1. Label and value in the same block (e.g. "RF Antenna Weight:\n64kg")
                m = re.search(label_pat + r":?\s*\n\s*([^\r\n]+)", text, re.IGNORECASE)
                if m:
                    val = clean_val(m.group(1))
                    if val:
                        return val

                # 2. Label in block b, value in an adjacent table cell to the right on the same row
                if re.search(label_pat, text, re.IGNORECASE):
                    y0, y1, x1 = b[1], b[3], b[2]
                    cand = [ob for ob in blocks if ob != b and ob[0] >= x1 - 10 and abs(ob[1] - y0) <= y_tol]
                    if cand:
                        cand.sort(key=lambda c: c[0])
                        val = clean_val(cand[0][4].strip().replace("\n", " "))
                        if val:
                            return val

        # 3. Fallback to regex across full_text with clean_val validation
        m_reg = re.search(label_pat + r":?\s*\n?\s*([^\r\n]+)", full_text, re.IGNORECASE)
        if m_reg:
            val = clean_val(m_reg.group(1))
            if val:
                return val
        return ""

    row["Existing Loadings Description"] = get_field_aligned_value(r"Existing\s+Loadings")
    row["RF Antenna Dimension MM"] = get_field_aligned_value(r"RF\s+Antenna\s+Dimension")
    row["RF Antenna Number Of Antennas"] = get_field_aligned_value(r"No\.\s+of\s+RF\s+Antennas")
    row["RF Antenna Azimuth"] = get_field_aligned_value(r"RF\s+Antenna\s+Azimuth")
    row["RF Antenna Center Line Height AGL M"] = get_field_aligned_value(r"RF\s+Antenna\s+Center\s+Line\s+Height")
    row["RF Antenna Weight"] = get_field_aligned_value(r"RF\s+Antenna\s+Weight")

    # Standard Regex Patterns for other fields
    patterns = {
        "Site Proposal Date": r"Site Proposal Date:?\s*\n?\s*([^\r\n]+)",
        "Site Proposal Type": r"Site Proposal Type:?\s*\n?\s*([^\r\n]+)",
        "Site Classification": r"Site Classification \(Uptime SLA\):\s*\n?\s*([^\r\n]+)",
        "Tower Company Name": r"Tower Co\. Name:\s*\n?\s*([^\r\n]+)",
        "Site ID": r"Smart Site ID:\s*\n?\s*([^\r\n]+)",
        "Site Name": r"Smart Site Name:\s*\n?\s*([^\r\n]+)",
        "Site Address": r"Site Address:\s*\n?\s*([^\r\n]+)",
        "Latitude": r"Latitude:\s*\n?\s*([^\r\n]+)",
        "Longitude": r"Longitude:\s*\n?\s*([^\r\n]+)",
        "Site Type": r"Site Type \(Ground-based/\s*Rooftop\):\s*\n?\s*([^\r\n]+)",
        "Tower Height M": r"Tower Height \(m\):\s*\n?\s*([^\r\n]+)",
        "Building Height M": r"Building Height \(m\):\s*\n?\s*([^\r\n]+)",
        "Dismantling Number Of Elements": r"No\. of Elements:\s*\n?\s*([^\r\n]+)",
        "Dismantling Total Weight KG": r"Weight of Antenna/Equipment:\s*\n?\s*([^\r\n]+)",
        "Smart Equipment Housing Type": r"SMART Equipment Housing Type:\s*\n?\s*([^\r\n]+)",
        "BTS BBU Number And U Space": r"No\. of BTS/BBU & U Space:\s*\n?\s*([^\r\n]+)",
        "RRH RRU Number": r"No\. of RRH/RRU:\s*\n?\s*([^\r\n]+)",
        "Transmission Solution": r"Transmission Solution \(FOC or MW Radio\):\s*\n?\s*([^\r\n]+)",
        "MW Radio Antenna Diameter M": r"MW Radio Antenna Diameter \(m\):\s*\n?\s*([^\r\n]+)",
        "MW Radio Antenna Height AGL M": r"MW Radio Antenna Height, AGL \(m\):\s*\n?\s*([^\r\n]+)",
        "MW Radio Antenna Azimuth": r"MW Radio Antenna Azimuth:\s*\n?\s*([^\r\n]+)",
        "MW Radio Antenna Weight KG": r"MW Radio Antenna Weight \(kg\):\s*\n?\s*([^\r\n]+)",
        "Floor Space": r"Floor\s+Space\s+for\s+Outdoor/Indoor[^\n:]*(?:\n[^\n:]*)*:\s*\n?\s*([^\r\n]+)",
        "Cabinet Dimensions MM": r"Equipment/Cabinet\s+Dimensions[^\n:]*(?:\n[^\n:]*)*:\s*\n?\s*([^\r\n]+)",
        "Cabinet Number Of Cabinets": r"No\. of Cabinets:\s*\n?\s*([^\r\n]+)",
        "Cabinet Weight KG": r"Equipment/Cabinet\s+Weight\s*\(kg\):\s*\n?\s*([^\r\n]+)",
        "OLT Dimension MM": r"OLT Dimension \(mm\):\s*\n?\s*([^\r\n]+)",
        "OLT Number": r"No\. of OLT:\s*\n?\s*([^\r\n]+)",
        "Sf Power Remarks": r"SF - POWER Remarks:\s*\n?\s*([^\r\n]+)",
        "Sf Infra Remarks": r"SF - INFRA Remarks:\s*\n?\s*([^\r\n]+)",
        "AC Power": r"AC Power Requirement:\s*\n?\s*([^\r\n]+)",
        "DC Power": r"DC Power Requirement:\s*\n?\s*([^\r\n]+)",
        "Cooling": r"Cooling Requirement:\s*\n?\s*([^\r\n]+)",
        "Space": r"Space Requirement:\s*\n?\s*([^\r\n]+)",
    }

    for k, pat in patterns.items():
        m = re.search(pat, full_text, re.IGNORECASE)
        if m:
            row[k] = clean_val(m.group(1), field_name=k)

    # Clean RF Antenna Weight if it accidentally caught Existing Loading text
    if row["RF Antenna Weight"] == row["Existing Loadings Description"]:
        row["RF Antenna Weight"] = ""

    # Tower Type handling
    tt_m = re.search(r"Tower Type:.*?\n(?:Bipod, etc\.\)\s*\n)?\s*([^\r\n]+)", full_text, re.IGNORECASE)
    if tt_m:
        row["Tower Type"] = clean_val(tt_m.group(1))

    # Section-specific Metrics (Dismantling, Additional, Declared)
    dism_match = re.search(r"EQUIPMENT\s+FOR\s+DISMANTLING(.*?)(?=ADDITIONAL\s+EQUIPMENT|\Z)", full_text, re.DOTALL | re.IGNORECASE)
    add_match = re.search(r"ADDITIONAL\s+EQUIPMENT(.*?)(?=SMART\s+DECLARED\s+PLANNED|\Z)", full_text, re.DOTALL | re.IGNORECASE)
    decl_match = re.search(r"SMART\s+DECLARED\s+PLANNED\s+UTILIZED\s+LOADINGS(.*?)(?=C\.\s*POWER\s+REQUIREMENT|\Z)", full_text, re.DOTALL | re.IGNORECASE)

    def get_sec_metric(sec_text, pat):
        if not sec_text:
            return ""
        m = re.search(pat, sec_text, re.IGNORECASE)
        return clean_val(m.group(1)) if m else ""

    row["Dismantling Total EPA SQM"] = get_sec_metric(dism_match.group(1) if dism_match else "", r"Total\s+EPA\s*\(sq\.?\s*m\)[^:\n]*:\s*\n?\s*([^\r\n]+)")
    row["Dismantling Total DC Power Watts"] = get_sec_metric(dism_match.group(1) if dism_match else "", r"Total\s+DC\s+Power\s*\(watts\)[^:\n]*:\s*\n?\s*([^\r\n]+)")
    row["Dismantling Total Ground Space"] = get_sec_metric(dism_match.group(1) if dism_match else "", r"Total\s+Ground\s+Space[^:\n]*:\s*\n?\s*([^\r\n]+)")

    row["Additional Total EPA SQM"] = get_sec_metric(add_match.group(1) if add_match else "", r"Total\s+EPA\s*\(sq\.?\s*m\)[^:\n]*:\s*\n?\s*([^\r\n]+)")
    row["Additional Total DC Power Watts"] = get_sec_metric(add_match.group(1) if add_match else "", r"Total\s+DC\s+Power\s*\(watts\)[^:\n]*:\s*\n?\s*([^\r\n]+)")
    row["Additional Total Ground Space SQM"] = get_sec_metric(add_match.group(1) if add_match else "", r"Total\s+Ground\s+Space[^:\n]*:\s*\n?\s*([^\r\n]+)")

    row["Declared Total EPA SQM"] = get_sec_metric(decl_match.group(1) if decl_match else "", r"Total\s+EPA\s*\(sq\.?\s*m\)[^:\n]*:\s*\n?\s*([^\r\n]+)")
    row["Declared Total DC Power Watts"] = get_sec_metric(decl_match.group(1) if decl_match else "", r"Total\s+DC\s+Power\s*\(watts\)[^:\n]*:\s*\n?\s*([^\r\n]+)")
    row["Declared Total Ground Space SQM"] = get_sec_metric(decl_match.group(1) if decl_match else "", r"Total\s+Ground\s+Space[^:\n]*:\s*\n?\s*([^\r\n]+)")

    # Checkboxes on page 2
    def get_cb(label, options):
        m = re.search(re.escape(label) + r"[:\s]*(.*?)(?=\n[A-Z]|\n\n|\Z)", full_text, re.DOTALL | re.IGNORECASE)
        if m:
            area = m.group(1)
            for opt in options:
                if f"◼{opt}" in area or f"■{opt}" in area or re.search(r"[◼■▪☑\[xX]]\s*" + re.escape(opt), area):
                    return opt
        return ""

    row["Access Antenna Mounting Bracket"] = get_cb("Access Antenna Mounting Bracket", ["New", "Existing", "N/A"])
    row["Transmission Antenna Mounting Bracket"] = get_cb("Transmission Antenna Mounting Bracket", ["New", "Existing", "N/A"])
    row["Transmission Rack"] = get_cb("Transmission Rack", ["New", "Existing", "N/A"])
    row["Power Declared Design Capacity"] = get_cb("POWER - Declared Design Capacity", ["Within", "Exceeded"])
    row["EPA Declared Design Capacity"] = get_cb("EPA - Declared Design Capacity", ["Within", "Exceeded"])
    row["Ground Space Declared Design Capacity"] = get_cb("Ground Space - Declared Design Capacity", ["Within", "Exceeded"])

    # Other Requirement Notes
    m_other = re.search(r"D\.\s*OTHER\s+REQUIREMENT[S]?:?\s*(.*?)(?=\n\s*(?:[A-E]\.|\bFIXED\b|\Z))", full_text, re.DOTALL | re.IGNORECASE)
    val_o = m_other.group(1).strip() if m_other else ""
    if not val_o or val_o.startswith("E.") or "PROJ" in val_o:
        row["Other Requirement Notes"] = ""
    else:
        row["Other Requirement Notes"] = val_o.replace("\n", " ")

    # Implementation Target Date
    m_date = re.search(r"Project\s+Implementation\s+Target\s+Date\s*([0-9]{1,2}-[A-Za-z]{3}-[0-9]{4})?", full_text, re.IGNORECASE)
    if m_date and m_date.group(1):
        row["Implementation Target Date"] = m_date.group(1).strip()

    # Approvals parsing on last page
    if pages:
        p_last = pages[-1]
        blocks_last = p_last.get_text("blocks")
        base_y = 0
        for b in blocks_last:
            if "FIXED / WIRELESS APPROVALS" in b[4] or "REVIEWED BY" in b[4]:
                base_y = max(base_y, b[3])

        words_last = p_last.get_text("words")
        appr_words = [w for w in words_last if w[1] >= base_y]
        left = [w for w in appr_words if w[0] < 290]
        right = [w for w in appr_words if w[0] >= 290]
        date_regex = re.compile(r"^\d{1,2}-[A-Za-z]{3}-\d{4}$")

        def extract_appr_col(col_words, dept_name):
            name_words = []
            date_val = ""
            ignore = {"Approver", "1", "Signature", "(Engineer)", "Name", "Date", "Access", "Engineering", "Transmission", "FIXED", "/", "WIRELESS", "APPROVALS", "REVIEWED", "BY"}
            for w in col_words:
                text = w[4].strip()
                if date_regex.match(text):
                    date_val = text
                elif text not in ignore:
                    name_words.append(text)
            name_val = " ".join(name_words).strip()
            return {"Department": dept_name, "Level": "Approver 1", "Name": name_val, "Date": date_val}

        acc = extract_appr_col(left, "Access Engineering")
        tx = extract_appr_col(right, "Transmission Engineering")
        approvals = [a for a in [acc, tx] if a["Name"] != "" or a["Date"] != ""]

        for idx, appr in enumerate(approvals, start=1):
            if idx <= 6:
                row[f"Approval {idx} Department"] = appr["Department"]
                row[f"Approval {idx} Level"] = appr["Level"]
                row[f"Approval {idx} Approver Name"] = appr["Name"]
                row[f"Approval {idx} Approval Date"] = appr["Date"]

    # Final guarantee: Every single key must be a valid string (empty string for blanks)
    for c in ALL_COLUMNS:
        if c not in row or row[c] is None:
            row[c] = ""
        else:
            row[c] = str(row[c]).strip()

    return row

def get_daily_excel_path(base_folder=None, date_obj=None):
    if base_folder is None:
        base_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Daily_Reports")
    os.makedirs(base_folder, exist_ok=True)
    
    if date_obj is None:
        date_obj = datetime.date.today()
    date_str = date_obj.strftime("%Y-%m-%d")
    return os.path.join(base_folder, f"Site_Requests_{date_str}.xlsx")

def _apply_row_style(ws, row_idx, record):
    data_font = Font(name="Calibri", size=10, color="1A1A1A")
    data_font_bold = Font(name="Calibri", size=10, bold=True, color="104E8B")

    thin_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    ws.row_dimensions[row_idx].height = 22
    row_fill = PatternFill(
        start_color="F8FAFC" if row_idx % 2 == 0 else "FFFFFF",
        end_color="F8FAFC" if row_idx % 2 == 0 else "FFFFFF",
        fill_type="solid"
    )

    for col_idx, col_name in enumerate(ALL_COLUMNS, start=1):
        val = record.get(col_name, "")
        cell = ws.cell(row=row_idx, column=col_idx, value=val)

        if col_name in ["Site ID", "Site Proposal Type", "Reference Number"]:
            cell.font = data_font_bold
        else:
            cell.font = data_font

        cell.fill = row_fill
        cell.border = thin_border

        if col_name in [
            "Site ID", "Site Proposal Date", "Tower Company Name", "Site Type", 
            "Tower Type", "Tower Height M", "Building Height M", "Latitude", "Longitude",
            "Extraction Date", "Proposed Loadings", "Access Antenna Mounting Bracket",
            "Transmission Antenna Mounting Bracket", "Transmission Rack", "Implementation Target Date"
        ] or "Date" in col_name or "Level" in col_name:
            cell.alignment = center_align
        else:
            cell.alignment = left_align

def create_styled_workbook(records, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Site Requests"
    ws.views.sheetView[0].showGridLines = True

    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    primary_header_fill = PatternFill(start_color="104E8B", end_color="104E8B", fill_type="solid")
    secondary_header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    accent_header_fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
    approval_header_fill = PatternFill(start_color="4338CA", end_color="4338CA", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Header Row
    ws.row_dimensions[1].height = 36
    for col_idx, col_name in enumerate(ALL_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        if "Approval" in col_name or "Verification" in col_name:
            cell.fill = approval_header_fill
        elif col_name in ["Site ID", "Site Name", "Reference Number", "Site Proposal Type"]:
            cell.fill = primary_header_fill
        elif "Dismantling" in col_name or "Additional" in col_name or "Declared" in col_name:
            cell.fill = accent_header_fill
        else:
            cell.fill = secondary_header_fill
        cell.alignment = center_align
        cell.border = thin_border

        ws.column_dimensions[get_column_letter(col_idx)].width = max(len(col_name) + 3, 15)

    # Data Rows
    for row_idx, record in enumerate(records, start=2):
        _apply_row_style(ws, row_idx, record)

    ws.freeze_panes = "D2"
    wb.save(output_path)

def generate_excel_bytes(records):
    buf = io.BytesIO()
    create_styled_workbook(records, buf)
    buf.seek(0)
    return buf.getvalue()

def append_to_daily_excel(records, excel_path=None):
    if excel_path is None:
        excel_path = get_daily_excel_path()

    if not os.path.exists(excel_path):
        create_styled_workbook(records, excel_path)
        return {
            "status": "created",
            "added": len(records),
            "skipped": 0,
            "total": len(records),
            "excel_path": excel_path
        }

    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    ref_col_idx = headers.index("Reference Number") + 1 if "Reference Number" in headers else None
    file_col_idx = headers.index("Filename") + 1 if "Filename" in headers else None

    existing_identifiers = set()
    for row in range(2, ws.max_row + 1):
        ref_val = ws.cell(row=row, column=ref_col_idx).value if ref_col_idx else None
        file_val = ws.cell(row=row, column=file_col_idx).value if file_col_idx else None
        if ref_val and str(ref_val).strip():
            existing_identifiers.add(str(ref_val).strip().lower())
        elif file_val and str(file_val).strip():
            existing_identifiers.add(str(file_val).strip().lower())

    added_count = 0
    skipped_count = 0
    current_row = ws.max_row + 1

    for record in records:
        rec_ref = record.get("Reference Number", "").strip().lower()
        rec_file = record.get("Filename", "").strip().lower()
        
        if (rec_ref and rec_ref in existing_identifiers) or \
           (rec_file and rec_file in existing_identifiers):
            skipped_count += 1
            continue

        _apply_row_style(ws, current_row, record)
        if rec_ref:
            existing_identifiers.add(rec_ref)
        if rec_file:
            existing_identifiers.add(rec_file)
        
        current_row += 1
        added_count += 1

    wb.save(excel_path)
    total_in_file = ws.max_row - 1

    return {
        "status": "appended",
        "added": added_count,
        "skipped": skipped_count,
        "total": total_in_file,
        "excel_path": excel_path
    }

def load_daily_dataframe(excel_path=None):
    if excel_path is None:
        excel_path = get_daily_excel_path()
    if not os.path.exists(excel_path):
        return pd.DataFrame()
    df = pd.read_excel(excel_path, keep_default_na=False)
    # Ensure NaN or None are empty strings
    return df.fillna("")
