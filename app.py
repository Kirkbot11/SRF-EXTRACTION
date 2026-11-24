import streamlit as st
import os
import io
import datetime
import pandas as pd
from extractor import (
    extract_pdf_data, 
    generate_excel_bytes,
    ALL_COLUMNS
)

# Page Configuration
st.set_page_config(
    page_title="PDF Site Extractor",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Refined styling: Soft darker background + 3D floating cards with rich shadows
st.markdown("""
<style>
    /* Darker, softer canvas background - eliminates harsh white glare */
    .stApp {
        background-color: #E6EAF0 !important;
        color: #0F172A !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }

    /* Headings - Crisp Dark Navy */
    h1, h2, h3, h4, h5, h6 {
        color: #0F172A !important;
        font-weight: 700 !important;
    }

    /* Paragraphs and general text */
    p, span, label, div {
        color: #1E293B;
    }

    /* Header Banner */
    .header-box {
        background: linear-gradient(135deg, #1E3A8A 0%, #1D4ED8 100%);
        padding: 24px 28px;
        border-radius: 16px;
        color: #FFFFFF !important;
        margin-bottom: 24px;
        box-shadow: 0 10px 25px -4px rgba(30, 58, 138, 0.25);
    }
    .header-box h1 {
        color: #FFFFFF !important;
        font-size: 26px !important;
        font-weight: 800 !important;
        margin: 0 !important;
        padding-bottom: 6px !important;
    }
    .header-box p {
        color: #E2E8F0 !important;
        font-size: 15px !important;
        font-weight: 500 !important;
        margin: 0 !important;
    }

    /* Ensure all 4 metric columns stretch and align perfectly */
    div[data-testid="stHorizontalBlock"] {
        align-items: stretch !important;
    }
    div[data-testid="column"] {
        display: flex !important;
        flex-direction: column !important;
        justify-content: stretch !important;
    }

    /* Metric Cards - 3D Floating Style with Uniform Equal Height & Baseline */
    .metric-card {
        background-color: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 14px !important;
        padding: 12px 10px !important;
        text-align: center !important;
        box-shadow: 0 8px 20px -3px rgba(15, 23, 42, 0.12), 0 4px 8px -2px rgba(15, 23, 42, 0.06) !important;
        height: 120px !important;
        min-height: 120px !important;
        max-height: 120px !important;
        width: 100% !important;
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
        align-items: center !important;
        box-sizing: border-box !important;
    }
    .metric-value {
        font-size: 23px !important;
        font-weight: 800 !important;
        color: #1E3A8A !important;
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
        color: #475569 !important;
        text-transform: uppercase !important;
        font-weight: 700 !important;
        letter-spacing: 0.3px !important;
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

    /* Step 1 Card - Floating Clean White */
    .step1-card {
        background-color: #FFFFFF !important;
        border-radius: 16px !important;
        padding: 24px !important;
        margin-bottom: 24px !important;
        border: 1px solid #CBD5E1 !important;
        box-shadow: 0 10px 25px -4px rgba(15, 23, 42, 0.12), 0 4px 10px -2px rgba(15, 23, 42, 0.06) !important;
    }

    /* Verification Card Styling */
    .verify-card {
        background-color: #FFFFFF !important;
        border-radius: 16px !important;
        padding: 22px 24px !important;
        margin-top: 18px !important;
        margin-bottom: 22px !important;
        border: 2px solid #60A5FA !important;
        box-shadow: 0 12px 28px -4px rgba(37, 99, 235, 0.15) !important;
    }

    /* STEP 2 - Highlighted 3D Floating Card with Rich Shadow */
    .step2-floating-container {
        background-color: #FFFFFF !important;
        border-radius: 18px !important;
        padding: 26px !important;
        margin-top: 24px !important;
        margin-bottom: 28px !important;
        border: 2px solid #93C5FD !important;
        box-shadow: 0 20px 40px -8px rgba(30, 58, 138, 0.2), 0 10px 20px -4px rgba(15, 23, 42, 0.1) !important;
    }

    /* Floating Step Badges */
    .step-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #2563EB;
        color: #FFFFFF !important;
        font-weight: 800;
        font-size: 13px;
        padding: 4px 14px;
        border-radius: 20px;
        box-shadow: 0 4px 10px rgba(37, 99, 235, 0.35);
        letter-spacing: 0.5px;
    }
    .step-badge-ready {
        background: #16A34A;
        color: #FFFFFF !important;
        font-weight: 700;
        font-size: 13px;
        padding: 4px 12px;
        border-radius: 20px;
        box-shadow: 0 2px 6px rgba(22, 163, 74, 0.3);
    }

    /* File Uploader Dropzone Enhancement */
    [data-testid="stFileUploader"] {
        background-color: #FFFFFF !important;
        border-radius: 12px !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        border: 2px dashed #2563EB !important;
        background-color: #F0F7FF !important;
        border-radius: 12px !important;
        padding: 24px !important;
    }
    [data-testid="stFileUploaderDropzone"] span {
        color: #1E3A8A !important;
        font-weight: 700 !important;
    }
    [data-testid="stFileUploaderDropzone"] small {
        color: #475569 !important;
        font-weight: 500 !important;
    }

    /* Sidebar Contrast & Styling */
    [data-testid="stSidebar"] {
        background-color: #DCE3ED !important;
        border-right: 1px solid #CBD5E1 !important;
    }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        color: #0F172A !important;
    }
    [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] label {
        color: #1E293B !important;
        font-weight: 500 !important;
    }
</style>
""", unsafe_allow_html=True)

# Date calculations
today_obj = datetime.date.today()
today_str = today_obj.strftime("%Y-%m-%d")
today_readable = today_obj.strftime("%A, %B %d, %Y")
default_download_name = f"Site_Requests_{today_str}.xlsx"

# Initialize In-Memory Staging Session State (Strictly in RAM)
if "staged_records" not in st.session_state:
    st.session_state.staged_records = {}  # filename -> record dict

# Sidebar - Minimal, Clean & Helpful
with st.sidebar:
    st.markdown("### 📋 PDF Extractor")
    st.markdown(f"**Extracts all {len(ALL_COLUMNS)} standard fields directly to Excel.**")
    st.caption("100% in-browser processing. No files saved to your disk.")
    
    st.markdown("---")
    st.markdown(f"**📅 Date:** {today_readable}")
    
    staged_total = len(st.session_state.staged_records)
    if staged_total > 0:
        st.markdown(f"**📥 Extracted Files:** `{staged_total}` ready")
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
        if st.button("🗑️ Clear All Extracted Data", use_container_width=True):
            st.session_state.staged_records.clear()
            st.rerun()

    st.markdown("---")
    st.markdown("### 💡 Quick Instructions")
    st.markdown("""
    1. **Upload**: Drag & drop your PDF Site Request Forms into Step 1.
    2. **Review**: Check the extracted fields in the Step 2 table below.
    3. **Download**: Click **Download Extracted Excel** to get your file.
    """)

# Main Content Header Banner
st.markdown(f"""
<div class="header-box">
    <h1>📋 Automated PDF Site Request Extractor</h1>
    <p>Upload incoming PDFs, review the extracted fields in memory, and download your completed Excel spreadsheet directly.</p>
</div>
""", unsafe_allow_html=True)

# Metrics Summary Row (All 4 cards 100% equal size and aligned)
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
        <div class="metric-value" style="color: {'#1E3A8A' if staged_count > 0 else '#64748B'};">{staged_count}</div>
        <div class="metric-label">📄 Extracted PDFs</div>
    </div>
    """, unsafe_allow_html=True)

with col_m3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-value">{len(ALL_COLUMNS)}</div>
        <div class="metric-label">📑 Standard Fields</div>
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
        <span style="font-size: 19px; font-weight: 800; color: #0F172A;">Upload Incoming PDFs</span>
    </div>
    <p style="color: #475569; font-size: 14px; margin-top: 0; margin-bottom: 14px;">
        Drag & drop 1 file or multiple Site Request Form (SRF) PDFs below. You can verify the file names before proceeding with extraction.
    </p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Drop your PDF Site Request Form(s) here:",
    type=["pdf"],
    accept_multiple_files=True,
    help="Drag and drop PDF files into this area."
)

# Verification Card & Proceed Button Before Extraction
if uploaded_files:
    unextracted_files = [uf for uf in uploaded_files if uf.name not in st.session_state.staged_records]
    
    if unextracted_files:
        st.markdown(f"""
        <div class="verify-card">
            <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; margin-bottom: 12px;">
                <div style="display: flex; align-items: center; gap: 8px;">
                    <span style="background: #2563EB; color: white; padding: 4px 12px; border-radius: 14px; font-weight: 800; font-size: 12px;">VERIFICATION</span>
                    <span style="font-weight: 800; font-size: 18px; color: #0F172A;">Verify Uploaded Files Before Extraction</span>
                </div>
                <span style="background: #EFF6FF; color: #1D4ED8; font-weight: 700; font-size: 13px; padding: 4px 12px; border-radius: 12px; border: 1px solid #BFDBFE;">
                    {len(unextracted_files)} File{'s' if len(unextracted_files) > 1 else ''} Ready
                </span>
            </div>
            <p style="color: #475569; font-size: 14px; margin: 0 0 14px 0; font-weight: 500;">
                Review the uploaded files below. When you are ready, click <b>Proceed with Extraction</b> to extract all 101 fields into memory.
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
            with st.spinner(f"Extracting all 101 fields from {len(unextracted_files)} PDF(s)..."):
                for uf in unextracted_files:
                    bytes_data = uf.getvalue()
                    data = extract_pdf_data(bytes_data, filename=uf.name)
                    st.session_state.staged_records[uf.name] = data
            st.toast(f"Extracted {len(unextracted_files)} PDF(s) successfully!", icon="✅")
            st.rerun()
    else:
        st.caption("✓ All selected files have been extracted below. Drop more files above if you have more incoming PDFs.")

# STEP 2 FLOATING CARD (Direct Download from Extracted Data)
if st.session_state.staged_records:
    rec_count = len(st.session_state.staged_records)
    records_list = list(st.session_state.staged_records.values())
    excel_bytes = generate_excel_bytes(records_list)
    
    st.markdown(f"""
    <div class="step2-floating-container">
        <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; margin-bottom: 14px;">
            <div style="display: flex; align-items: center; gap: 10px;">
                <span class="step-badge" style="background: #1D4ED8; font-size: 14px; padding: 6px 16px;">STEP 2</span>
                <span style="font-size: 21px; font-weight: 800; color: #0F172A;">Extracted Data</span>
            </div>
            <span class="step-badge-ready">✓ {rec_count} Record(s) Extracted & Ready</span>
        </div>
        <p style="color: #334155; font-size: 14px; margin-top: 0; margin-bottom: 16px; font-weight: 500;">
            Review the extracted fields in the table below. When ready, click the blue 
            <b style="color: #1D4ED8;">Download Extracted Excel File</b> button to save your spreadsheet.
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
    <div style="background-color: #F1F5F9; border: 2px dashed #CBD5E1; border-radius: 14px; padding: 22px; text-align: center; margin-top: 20px;">
        <div style="display: inline-flex; align-items: center; gap: 8px; margin-bottom: 6px;">
            <span style="background: #94A3B8; color: white; padding: 3px 10px; border-radius: 12px; font-weight: 700; font-size: 12px;">STEP 2</span>
            <span style="color: #475569; font-weight: 700; font-size: 16px;">Extracted Data & Download</span>
        </div>
        <p style="color: #64748B; font-size: 13px; margin: 0;">
            Drop your PDF Site Request Form(s) in Step 1 above. The extracted table and download button will appear here automatically.
        </p>
    </div>
    """, unsafe_allow_html=True)
