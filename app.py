import streamlit as st
import os
import io
import base64
import datetime
import pandas as pd
from extractor import (
    extract_pdf_data, 
    generate_excel_bytes,
    CORE_FIELDS,
    ALL_COLUMNS
)

# Base Paths & Asset Helpers
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")

def get_base64_image(filename):
    p = os.path.join(ASSETS_DIR, filename)
    if os.path.exists(p):
        with open(p, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    return ""

header_logo_b64 = get_base64_image("ftap_header_logo.png")
icon_logo_b64 = get_base64_image("ftap_logo_only.png")
icon_file_path = os.path.join(ASSETS_DIR, "ftap_logo_only.png")

# Page Configuration with FTAP Favicon
st.set_page_config(
    page_title="Frontier Towers Philippines - SRF Extractor",
    page_icon=icon_file_path if os.path.exists(icon_file_path) else "🗼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Frontier Towers Philippines (FTAP) Official Brand Styling
st.markdown("""
<style>
    /* Clean, soft executive canvas - eliminates white glare while maintaining corporate precision */
    .stApp {
        background-color: #F4F6F9 !important;
        color: #1F2029 !important;
        font-family: -apple-system, BlinkMacSystemFont, "Gibson", "Roboto", "Segoe UI", Helvetica, Arial, sans-serif;
    }

    /* Headings - Frontier Executive Slate / Charcoal */
    h1, h2, h3, h4, h5, h6 {
        color: #1F2029 !important;
        font-weight: 700 !important;
    }

    /* General prose styling - targeting text elements outside buttons */
    .stApp p, .stApp label {
        color: #334155;
    }

    /* FTAP Top Header Card */
    .ftap-header-box {
        background-color: #FFFFFF;
        border-radius: 16px;
        padding: 22px 28px;
        margin-bottom: 22px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 10px 25px -4px rgba(15, 23, 42, 0.08), 0 4px 10px -2px rgba(15, 23, 42, 0.04);
        position: relative;
        overflow: hidden;
    }
    .ftap-header-box::before {
        content: "";
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 5px;
        background: linear-gradient(90deg, #C00000 0%, #D81E2E 60%, #F15A2A 100%);
    }
    .ftap-header-content {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 20px;
        flex-wrap: wrap;
    }
    .ftap-brand-left {
        display: flex;
        align-items: center;
        gap: 20px;
    }
    .ftap-logo-img {
        height: 56px;
        width: auto;
        max-width: 230px;
        object-fit: contain;
    }
    .ftap-title-block h1 {
        color: #1F2029 !important;
        font-size: 23px !important;
        font-weight: 800 !important;
        margin: 0 !important;
        padding-bottom: 4px !important;
        letter-spacing: -0.3px;
    }
    .ftap-title-block p {
        color: #64748B !important;
        font-size: 14px !important;
        font-weight: 500 !important;
        margin: 0 !important;
    }
    .ftap-badge-enterprise {
        background: #FFF1F1;
        color: #C00000;
        border: 1px solid #FED7D7;
        padding: 6px 14px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 700;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        display: inline-flex;
        align-items: center;
        gap: 6px;
    }

    /* Metric Columns Layout */
    div[data-testid="stHorizontalBlock"] {
        align-items: stretch !important;
    }
    div[data-testid="column"] {
        display: flex !important;
        flex-direction: column !important;
        justify-content: stretch !important;
    }

    /* FTAP Metric Cards */
    .metric-card {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-top: 3px solid #C00000 !important;
        border-radius: 14px !important;
        padding: 12px 10px !important;
        text-align: center !important;
        box-shadow: 0 6px 18px -3px rgba(15, 23, 42, 0.08), 0 2px 6px -2px rgba(15, 23, 42, 0.04) !important;
        height: 120px !important;
        min-height: 120px !important;
        max-height: 120px !important;
        width: 100% !important;
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
        align-items: center !important;
        box-sizing: border-box !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 22px -4px rgba(192, 0, 0, 0.12), 0 4px 8px -2px rgba(15, 23, 42, 0.06) !important;
    }
    .metric-value {
        font-size: 23px !important;
        font-weight: 800 !important;
        color: #1F2029 !important;
        line-height: 1.15 !important;
        height: 48px !important;
        min-height: 48px !important;
        max-height: 48px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        text-align: center !important;
        margin: 0 !important;
        padding: 0 !important;
    }
    .metric-label {
        font-size: 11px !important;
        color: #64748B !important;
        text-transform: uppercase !important;
        font-weight: 700 !important;
        letter-spacing: 0.4px !important;
        height: 36px !important;
        min-height: 36px !important;
        max-height: 36px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        text-align: center !important;
        line-height: 1.25 !important;
        margin: 0 !important;
        padding: 0 4px !important;
    }

    /* Step 1 Card */
    .step1-card {
        background-color: #FFFFFF !important;
        border-radius: 16px !important;
        padding: 24px !important;
        margin-bottom: 22px !important;
        border: 1px solid #E2E8F0 !important;
        box-shadow: 0 8px 22px -4px rgba(15, 23, 42, 0.08) !important;
    }

    /* Verification Card */
    .verify-card {
        background-color: #FFFFFF !important;
        border-radius: 16px !important;
        padding: 22px 24px !important;
        margin-top: 18px !important;
        margin-bottom: 22px !important;
        border: 2px solid #C00000 !important;
        box-shadow: 0 10px 24px -4px rgba(192, 0, 0, 0.12) !important;
    }

    /* Step 2 Floating Container */
    .step2-floating-container {
        background-color: #FFFFFF !important;
        border-radius: 18px !important;
        padding: 26px !important;
        margin-top: 24px !important;
        margin-bottom: 28px !important;
        border: 2px solid #E2E8F0 !important;
        border-left: 6px solid #C00000 !important;
        box-shadow: 0 16px 36px -6px rgba(192, 0, 0, 0.1), 0 8px 16px -4px rgba(15, 23, 42, 0.05) !important;
    }

    /* FTAP Step Badges */
    .step-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #C00000;
        color: #FFFFFF !important;
        font-weight: 800;
        font-size: 13px;
        padding: 5px 14px;
        border-radius: 20px;
        box-shadow: 0 4px 10px rgba(192, 0, 0, 0.28);
        letter-spacing: 0.5px;
    }
    .step-badge-ready {
        background: #16A34A;
        color: #FFFFFF !important;
        font-weight: 700;
        font-size: 13px;
        padding: 5px 14px;
        border-radius: 20px;
        box-shadow: 0 3px 8px rgba(22, 163, 74, 0.25);
    }

    /* File Uploader Dropzone Enhancement (FTAP Crimson Accents) */
    [data-testid="stFileUploader"] {
        background-color: #FFFFFF !important;
        border-radius: 14px !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        border: 2px dashed #C00000 !important;
        background-color: #FFF8F8 !important;
        border-radius: 14px !important;
        padding: 24px !important;
        transition: background-color 0.2s ease, border-color 0.2s ease;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        background-color: #FFF0F0 !important;
        border-color: #9E0000 !important;
    }
    [data-testid="stFileUploaderDropzone"] span {
        color: #C00000 !important;
        font-weight: 700 !important;
    }
    [data-testid="stFileUploaderDropzone"] small {
        color: #64748B !important;
        font-weight: 500 !important;
    }

    /* ========================================================
       PROMINENT & HIGH-VISIBILITY "ADD MORE FILES" (+) BUTTON
       ======================================================== */
    [data-testid="stFileUploader"] button[aria-label="Add files"],
    [data-testid="stFileUploaderDropzone"] button[aria-label="Add files"],
    button[aria-label="Add files"],
    [data-testid="stFileUploader"] [data-testid="stBaseButton-borderless"] {
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        background: linear-gradient(135deg, #C00000 0%, #A00000 100%) !important;
        color: #FFFFFF !important;
        border: 2px solid #8B0000 !important;
        border-radius: 8px !important;
        padding: 6px 14px !important;
        height: 38px !important;
        min-width: 44px !important;
        margin-left: 8px !important;
        margin-top: 4px !important;
        margin-bottom: 4px !important;
        cursor: pointer !important;
        box-shadow: 0 4px 12px rgba(192, 0, 0, 0.35) !important;
        transition: all 0.2s ease-in-out !important;
        opacity: 1 !important;
        visibility: visible !important;
    }
    [data-testid="stFileUploader"] button[aria-label="Add files"]:hover,
    button[aria-label="Add files"]:hover {
        background: linear-gradient(135deg, #D61B2B 0%, #B30000 100%) !important;
        transform: translateY(-2px) scale(1.05) !important;
        box-shadow: 0 6px 16px rgba(192, 0, 0, 0.5) !important;
    }
    [data-testid="stFileUploader"] button[aria-label="Add files"] *,
    [data-testid="stFileUploader"] button[aria-label="Add files"] svg,
    [data-testid="stFileUploader"] button[aria-label="Add files"] span {
        color: #FFFFFF !important;
        fill: #FFFFFF !important;
        stroke: #FFFFFF !important;
        font-size: 20px !important;
        font-weight: 900 !important;
        -webkit-text-fill-color: #FFFFFF !important;
    }
    [data-testid="stFileUploader"] button[aria-label="Add files"]::after {
        content: " Add More Files" !important;
        color: #FFFFFF !important;
        font-size: 13px !important;
        font-weight: 800 !important;
        margin-left: 6px !important;
        white-space: nowrap !important;
        -webkit-text-fill-color: #FFFFFF !important;
    }

    /* ========================================================
       PRIMARY & SECONDARY BUTTON OVERRIDES (CRISP WHITE TEXT)
       ======================================================== */
    button[kind="primary"],
    div.stButton > button[kind="primary"],
    div[data-testid="stDownloadButton"] > button[kind="primary"] {
        background: linear-gradient(135deg, #C00000 0%, #A00000 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 10px 24px !important;
        font-weight: 800 !important;
        font-size: 15px !important;
        box-shadow: 0 4px 14px rgba(192, 0, 0, 0.35) !important;
        transition: all 0.2s ease-in-out !important;
    }
    button[kind="primary"]:hover,
    div.stButton > button[kind="primary"]:hover,
    div[data-testid="stDownloadButton"] > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #D41525 0%, #B30000 100%) !important;
        box-shadow: 0 6px 18px rgba(192, 0, 0, 0.45) !important;
        transform: translateY(-1px);
    }

    /* Force all nested text, divs, spans and paragraphs inside primary buttons to be pure white */
    button[kind="primary"] *,
    button[kind="primary"] p,
    button[kind="primary"] span,
    button[kind="primary"] div,
    button[kind="primary"] svg,
    div.stButton > button[kind="primary"] *,
    div.stButton > button[kind="primary"] p,
    div.stButton > button[kind="primary"] span,
    div.stButton > button[kind="primary"] div,
    div[data-testid="stDownloadButton"] > button[kind="primary"] *,
    div[data-testid="stDownloadButton"] > button[kind="primary"] p,
    div[data-testid="stDownloadButton"] > button[kind="primary"] span,
    div[data-testid="stDownloadButton"] > button[kind="primary"] div {
        color: #FFFFFF !important;
        fill: #FFFFFF !important;
        stroke: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-weight: 800 !important;
        font-size: 15px !important;
    }

    /* Secondary Buttons (Clear Data) */
    button[kind="secondary"],
    div.stButton > button[kind="secondary"],
    div[data-testid="stDownloadButton"] > button[kind="secondary"] {
        border: 1px solid #CBD5E1 !important;
        background-color: #FFFFFF !important;
        color: #334155 !important;
        border-radius: 10px !important;
        padding: 10px 20px !important;
        font-weight: 600 !important;
        transition: all 0.2s ease !important;
    }
    button[kind="secondary"] *,
    button[kind="secondary"] p,
    button[kind="secondary"] span,
    div.stButton > button[kind="secondary"] *,
    div[data-testid="stDownloadButton"] > button[kind="secondary"] * {
        color: #334155 !important;
        -webkit-text-fill-color: #334155 !important;
    }
    button[kind="secondary"]:hover,
    div.stButton > button[kind="secondary"]:hover,
    div[data-testid="stDownloadButton"] > button[kind="secondary"]:hover {
        border-color: #C00000 !important;
        background-color: #FFF5F5 !important;
    }
    button[kind="secondary"]:hover *,
    div.stButton > button[kind="secondary"]:hover *,
    div[data-testid="stDownloadButton"] > button[kind="secondary"]:hover * {
        color: #C00000 !important;
        -webkit-text-fill-color: #C00000 !important;
    }

    /* ========================================================
       SIDEBAR REFINED & NOTICEABLE ALIGNMENT
       ======================================================== */
    [data-testid="stSidebar"] {
        background-color: #F8FAFC !important;
        border-right: 1px solid #E2E8F0 !important;
    }
    [data-testid="stSidebarContent"] {
        padding: 1.25rem 1rem !important;
    }

    /* Branded Header Card in Sidebar */
    .sb-brand-card {
        background: #FFFFFF;
        border-radius: 14px;
        padding: 18px 14px;
        border: 1px solid #E2E8F0;
        border-top: 4px solid #C00000;
        box-shadow: 0 4px 16px rgba(15, 23, 42, 0.06);
        text-align: center;
        margin-bottom: 16px;
    }
    .sb-brand-logo {
        max-width: 175px;
        height: auto;
        margin-bottom: 10px;
        display: block;
        margin-left: auto;
        margin-right: auto;
    }
    .sb-brand-title {
        font-size: 13px;
        font-weight: 800;
        color: #1F2029;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        margin-bottom: 4px;
    }
    .sb-brand-badge {
        display: inline-block;
        background: #FFF1F1;
        color: #C00000;
        border: 1px solid #FED7D7;
        font-size: 11px;
        font-weight: 700;
        padding: 3px 10px;
        border-radius: 12px;
        letter-spacing: 0.4px;
    }

    /* Noticeable Section Card in Sidebar */
    .sb-section-card {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 14px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
        margin-bottom: 14px;
    }
    .sb-card-title {
        display: flex;
        align-items: center;
        justify-content: space-between;
        font-size: 13px;
        font-weight: 800;
        color: #1F2029;
        margin-bottom: 8px;
        padding-bottom: 6px;
        border-bottom: 1px solid #F1F5F9;
    }
    .sb-pill-count {
        background: #C00000;
        color: #FFFFFF;
        font-size: 11px;
        font-weight: 800;
        padding: 2px 8px;
        border-radius: 10px;
    }

    /* Structured Step-by-Step Rows */
    .sb-step-item {
        display: flex;
        align-items: flex-start;
        gap: 10px;
        padding: 8px 0;
        border-bottom: 1px dashed #F1F5F9;
    }
    .sb-step-item:last-child {
        border-bottom: none;
        padding-bottom: 0;
    }
    .sb-step-badge {
        display: flex;
        align-items: center;
        justify-content: center;
        width: 22px;
        height: 22px;
        min-width: 22px;
        background: #C00000;
        color: #FFFFFF;
        font-weight: 800;
        font-size: 11px;
        border-radius: 50%;
        margin-top: 2px;
    }
    .sb-step-text b {
        color: #1F2029;
        font-size: 13px;
        display: block;
        margin-bottom: 2px;
    }
    .sb-step-text span {
        color: #64748B;
        font-size: 12px;
        line-height: 1.3;
        display: block;
    }

    /* Field Chip List in Expander */
    .field-tag {
        display: inline-block;
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        color: #334155;
        font-size: 11px;
        font-weight: 600;
        padding: 3px 8px;
        border-radius: 6px;
        margin: 2px;
    }
</style>
""", unsafe_allow_html=True)

# Date calculations
today_obj = datetime.date.today()
today_str = today_obj.strftime("%Y-%m-%d")
today_readable = today_obj.strftime("%A, %B %d, %Y")
default_download_name = f"FTAP_Site_Requests_{today_str}.xlsx"

# Initialize In-Memory Staging Session State (Strictly in RAM)
if "staged_records" not in st.session_state:
    st.session_state.staged_records = {}  # filename -> record dict

# ========================================================
# SIDEBAR - HIGH-VISIBILITY, ALIGNED, ENTERPRISE BRANDING
# ========================================================
with st.sidebar:
    # 1. FTAP Brand Card
    if header_logo_b64:
        st.markdown(f"""
        <div class="sb-brand-card">
            <img src="data:image/png;base64,{header_logo_b64}" class="sb-brand-logo" alt="Frontier Towers Philippines" />
            <div class="sb-brand-title">Frontier Towers Philippines</div>
            <div class="sb-brand-badge">SRF Extraction Engine</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="sb-brand-card">
            <div style="font-size: 28px; margin-bottom: 4px;">🗼</div>
            <div class="sb-brand-title">Frontier Towers Philippines</div>
            <div class="sb-brand-badge">SRF Extraction Engine</div>
        </div>
        """, unsafe_allow_html=True)

    # 2. Extracted Fields Reference (1-to-1 Match with Python Code)
    st.markdown(f"""
    <div class="sb-section-card">
        <div class="sb-card-title">
            <span>📋 Target Extraction</span>
            <span class="sb-pill-count">{len(CORE_FIELDS)} Fields (1:1)</span>
        </div>
        <p style="font-size: 12px; color: #64748B; margin: 0 0 6px 0; line-height: 1.35;">
            Extracts the exact <b>{len(CORE_FIELDS)} core fields</b> 1-to-1 as specified in the Python extraction code.
        </p>
    </div>
    """, unsafe_allow_html=True)

    with st.expander(f"🔍 View all {len(CORE_FIELDS)} Fields (1:1 List)", expanded=False):
        for idx, fld in enumerate(CORE_FIELDS, 1):
            st.markdown(f"<span class='field-tag'><b>{idx}.</b> {fld}</span>", unsafe_allow_html=True)
        st.caption("+ Source File (Excel filename tracking)")

    # 3. Action Area (when files are ready)
    staged_total = len(st.session_state.staged_records)
    if staged_total > 0:
        st.markdown(f"""
        <div class="sb-section-card" style="border: 2px solid #C00000; background: #FFFBFB;">
            <div class="sb-card-title" style="color: #C00000;">
                <span>📥 Extracted Files</span>
                <span class="sb-pill-count">{staged_total} READY</span>
            </div>
            <p style="font-size: 12px; color: #475569; margin: 0 0 10px 0;">
                <b>{staged_total}</b> document{'s' if staged_total > 1 else ''} extracted and ready for instant export.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        records_list = list(st.session_state.staged_records.values())
        excel_bytes = generate_excel_bytes(records_list)
        st.download_button(
            label=f"⬇️ Download Excel ({staged_total} Files)",
            data=excel_bytes,
            file_name=default_download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        if st.button("🗑️ Clear Extracted Data", use_container_width=True):
            st.session_state.staged_records.clear()
            st.rerun()

    # 4. Status & Security Info Card
    st.markdown(f"""
    <div class="sb-section-card">
        <div class="sb-card-title">
            <span>📅 System Status</span>
            <span style="font-size: 11px; font-weight: 700; color: #16A34A;">● ACTIVE</span>
        </div>
        <div style="font-size: 12px; color: #334155; margin-bottom: 6px;">
            <b>Date:</b> {today_readable}
        </div>
        <div style="background: #F1F5F9; border-radius: 8px; padding: 6px 10px; font-size: 11px; color: #475569; display: flex; align-items: center; gap: 6px;">
            <span>🔒</span> <b>100% In-Memory RAM:</b> Zero disk storage
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 5. Step-by-Step Instructions Card
    st.markdown("""
    <div class="sb-section-card">
        <div class="sb-card-title">
            <span>💡 Quick Instructions</span>
            <span style="font-size: 11px; color: #64748B;">3 Steps</span>
        </div>
        <div class="sb-step-item">
            <span class="sb-step-badge">1</span>
            <div class="sb-step-text">
                <b>Upload SRFs</b>
                <span>Drop PDF Site Request Form(s) in Step 1.</span>
            </div>
        </div>
        <div class="sb-step-item">
            <span class="sb-step-badge">2</span>
            <div class="sb-step-text">
                <b>Verify & Extract</b>
                <span>Confirm file list & extract all 14 fields.</span>
            </div>
        </div>
        <div class="sb-step-item">
            <span class="sb-step-badge">3</span>
            <div class="sb-step-text">
                <b>Download Excel</b>
                <span>Export directly into formatted spreadsheet.</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ========================================================
# MAIN CONTENT HEADER BANNER
# ========================================================
logo_html = ""
if header_logo_b64:
    logo_html = f'<img src="data:image/png;base64,{header_logo_b64}" class="ftap-logo-img" alt="Frontier Towers Philippines Logo" />'
elif icon_logo_b64:
    logo_html = f'<img src="data:image/png;base64,{icon_logo_b64}" style="height: 50px; width: auto;" alt="FTAP Logo" />'

st.markdown(f"""
<div class="ftap-header-box">
    <div class="ftap-header-content">
        <div class="ftap-brand-left">
            {logo_html}
            <div class="ftap-title-block">
                <h1>Automated Site Request Form (SRF) Extractor</h1>
                <p>Frontier Tower Associates Philippines • Operations & Engineering Infrastructure</p>
            </div>
        </div>
        <div class="ftap-badge-enterprise">
            <span>🗼</span> {len(CORE_FIELDS)} Fields (1:1 Match)
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Metrics Summary Row (4 equal height cards)
staged_count = len(st.session_state.staged_records)
status_text = "READY" if staged_count > 0 else "WAITING"
status_color = "#16A34A" if staged_count > 0 else "#64748B"

col_m1, col_m2, col_m3, col_m4 = st.columns(4)
with col_m1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value">{today_obj.strftime('%b %d, %Y')}</div>
        <div class="metric-label">📅 Current Date</div>
    </div>
    """, unsafe_allow_html=True)

with col_m2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value" style="color: {'#C00000' if staged_count > 0 else '#64748B'};">{staged_count}</div>
        <div class="metric-label">📄 Extracted PDFs</div>
    </div>
    """, unsafe_allow_html=True)

with col_m3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value" style="color: #C00000;">{len(CORE_FIELDS)}</div>
        <div class="metric-label">📋 Extracted Fields</div>
    </div>
    """, unsafe_allow_html=True)

with col_m4:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value" style="color: {status_color} !important;">{status_text}</div>
        <div class="metric-label">⚡ Download Status</div>
    </div>
    """, unsafe_allow_html=True)

st.write("")

# STEP 1 CONTAINER (Floating Card)
st.markdown("""
<div class="step1-card">
    <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 8px;">
        <span class="step-badge">STEP 1</span>
        <span style="font-size: 19px; font-weight: 800; color: #1F2029;">Upload Incoming Site Request Forms</span>
    </div>
    <p style="color: #64748B; font-size: 14px; margin-top: 0; margin-bottom: 14px;">
        Drag & drop one or multiple SRF PDF documents below. You can verify document names and file sizes before initiating the 1-to-1 field extraction.
    </p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Drop your PDF Site Request Form(s) here:",
    type=["pdf"],
    accept_multiple_files=True,
    help="Drag and drop PDF files into this area."
)

# Appending & Add More Files Helper Banner
if uploaded_files or st.session_state.staged_records:
    current_count = len(st.session_state.staged_records)
    st.markdown(f"""
    <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; background: #FFF8F8; border: 1px solid #FECACA; border-left: 4px solid #C00000; border-radius: 10px; padding: 10px 16px; margin-top: 8px; margin-bottom: 14px;">
        <div style="display: flex; align-items: center; gap: 10px;">
            <span style="font-size: 20px;">➕</span>
            <span style="font-size: 13px; color: #7F1D1D; font-weight: 600;">
                <b>Appending More Files:</b> Click the prominent red <b>[+ Add More Files]</b> button inside the box above to add more incoming SRFs.
            </span>
        </div>
        <span style="background: #FEE2E2; color: #991B1B; font-size: 11px; font-weight: 800; padding: 4px 12px; border-radius: 12px; letter-spacing: 0.3px;">
            BATCH APPEND ACTIVE {'(' + str(current_count) + ' Staged)' if current_count > 0 else ''}
        </span>
    </div>
    """, unsafe_allow_html=True)

# Verification Card & Proceed Button Before Extraction
if uploaded_files:
    unextracted_files = [uf for uf in uploaded_files if uf.name not in st.session_state.staged_records]
    
    if unextracted_files:
        st.markdown(f"""
        <div class="verify-card">
            <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; margin-bottom: 12px;">
                <div style="display: flex; align-items: center; gap: 8px;">
                    <span style="background: #C00000; color: white; padding: 4px 12px; border-radius: 14px; font-weight: 800; font-size: 12px;">VERIFICATION</span>
                    <span style="font-weight: 800; font-size: 18px; color: #1F2029;">Verify Uploaded Files Before Extraction</span>
                </div>
                <span style="background: #FFF1F1; color: #C00000; font-weight: 700; font-size: 13px; padding: 4px 12px; border-radius: 12px; border: 1px solid #FED7D7;">
                    {len(unextracted_files)} File{'s' if len(unextracted_files) > 1 else ''} Ready
                </span>
            </div>
            <p style="color: #64748B; font-size: 14px; margin: 0 0 14px 0; font-weight: 500;">
                Review the uploaded files below. When ready, click <b>Proceed with Extraction</b> to extract all 14 core fields into memory.
            </p>
        </div>
        """, unsafe_allow_html=True)

        # Show file list for verification
        file_summary = []
        for idx, uf in enumerate(unextracted_files, start=1):
            size_kb = round(len(uf.getvalue()) / 1024, 1)
            file_summary.append({
                "#": idx,
                "File Name": uf.name,
                "File Size": f"{size_kb} KB",
                "Verification Status": "⏳ Ready to Extract"
            })
        
        st.dataframe(pd.DataFrame(file_summary), use_container_width=True, hide_index=True)
        
        col_p1, col_p2 = st.columns([3, 1])
        with col_p1:
            proceed_clicked = st.button(
                f"🚀 Proceed with Extraction ({len(unextracted_files)} File{'s' if len(unextracted_files) > 1 else ''})",
                type="primary",
                use_container_width=True
            )
        with col_p2:
            if st.button("🗑️ Clear Selected Files", use_container_width=True):
                st.session_state.staged_records.clear()
                st.rerun()

        if proceed_clicked:
            with st.spinner(f"Extracting all 14 core fields from {len(unextracted_files)} PDF(s)..."):
                for uf in unextracted_files:
                    bytes_data = uf.getvalue()
                    data = extract_pdf_data(bytes_data, filename=uf.name)
                    st.session_state.staged_records[uf.name] = data
            st.toast(f"Extracted {len(unextracted_files)} PDF(s) successfully!", icon="✅")
            st.rerun()
    else:
        st.caption("✓ All selected files have been extracted below. Drop more files above if you have additional incoming SRFs.")

# STEP 2 FLOATING CARD (Direct Download from Extracted Data)
if st.session_state.staged_records:
    rec_count = len(st.session_state.staged_records)
    records_list = list(st.session_state.staged_records.values())
    excel_bytes = generate_excel_bytes(records_list)
    
    st.markdown(f"""
    <div class="step2-floating-container">
        <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; margin-bottom: 14px;">
            <div style="display: flex; align-items: center; gap: 10px;">
                <span class="step-badge" style="background: #C00000; font-size: 14px; padding: 6px 16px;">STEP 2</span>
                <span style="font-size: 21px; font-weight: 800; color: #1F2029;">Extracted Data Preview</span>
            </div>
            <span class="step-badge-ready">✓ {rec_count} Record(s) Extracted & Ready</span>
        </div>
        <p style="color: #475569; font-size: 14px; margin-top: 0; margin-bottom: 16px; font-weight: 500;">
            Review the extracted fields in the table below. When ready, click 
            <b style="color: #C00000;">Download Extracted Excel File</b> to save your formatted spreadsheet.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Top Action Bar inside Step 2
    col_d1, col_d2 = st.columns([3, 1])
    with col_d1:
        st.download_button(
            label=f"⬇️ Download Extracted Excel File ({rec_count} Records)",
            data=excel_bytes,
            file_name=default_download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="dl_top",
            use_container_width=True
        )
    with col_d2:
        if st.button("🗑️ Clear Extracted Data", key="clear_top", use_container_width=True):
            st.session_state.staged_records.clear()
            st.rerun()

    st.write("")
    
    # Data Preview Table
    preview_df = pd.DataFrame(records_list)[ALL_COLUMNS]
    st.dataframe(
        preview_df,
        use_container_width=True,
        hide_index=True
    )

    st.write("")
    
    # Bottom Download Button
    st.download_button(
        label=f"⬇️ Download Extracted Excel File ({rec_count} Records)",
        data=excel_bytes,
        file_name=default_download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key="dl_bottom",
        use_container_width=True
    )
    
else:
    # Subtle placeholder when no files are waiting
    st.markdown("""
    <div style="background-color: #FFFFFF; border: 2px dashed #CBD5E1; border-radius: 14px; padding: 24px; text-align: center; margin-top: 20px;">
        <div style="display: inline-flex; align-items: center; gap: 8px; margin-bottom: 6px;">
            <span style="background: #94A3B8; color: white; padding: 4px 12px; border-radius: 14px; font-weight: 700; font-size: 12px;">STEP 2</span>
            <span style="color: #475569; font-weight: 700; font-size: 16px;">Extracted Data & Excel Export</span>
        </div>
        <p style="color: #64748B; font-size: 13px; margin: 0;">
            Drop your PDF Site Request Form(s) in Step 1 above. The extracted table and download button will appear here automatically.
        </p>
    </div>
    """, unsafe_allow_html=True)
