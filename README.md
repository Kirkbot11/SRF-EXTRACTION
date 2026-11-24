# Frontier Towers Philippines - Automated SRF Extractor & Excel Generator

A lightweight, non-technical, browser-based enterprise application built for **Frontier Tower Associates Philippines (FTAP)** to extract the 14 core telecom fields from Site Request Form (SRF) PDFs directly into beautifully formatted Microsoft Excel spreadsheets matching `Site_Request_Summary.xlsx`.

## Features
- **Official FTAP Brand Design**: Styled with Frontier Towers Philippines corporate crimson palette (`#C00000`), clear badges, and company header logo.
- **Pure In-Browser Processing**: Uploaded PDFs are processed strictly in RAM—no files or temporary documents are stored on disk.
- **Pre-Extraction Verification**: Review uploaded filenames and sizes before confirming extraction.
- **14 Core Fields Extracted (1:1 with Python Code)**:
  1. Smart Site ID
  2. Smart Site Name
  3. Site Proposal Type
  4. Site Proposal Date
  5. Tower Co. Name
  6. Site Classification (Uptime SLA)
  7. Site Type
  8. Tower Type
  9. Tower Height (m)
  10. Building Height (m)
  11. Latitude
  12. Longitude
  13. Site Address
  14. Ref. No.
  *(plus Source File tracking)*
- **Direct Excel Download**: Instant export to a styled `.xlsx` file with FTAP brand-coded headers and custom column widths.
- **High-Contrast, Modern UI**: Designed with 3D floating card elements, clear badges, and soft canvas to eliminate screen glare.

## Quick Start (Windows)
Simply double-click **`Launch_App.bat`** to start the app and open it in your browser.

Or run manually from terminal:
```bash
pip install -r requirements.txt
streamlit run app.py
``
