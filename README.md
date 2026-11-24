# Frontier Towers Philippines - Automated SRF Extractor & Excel Generator

A lightweight, non-technical, browser-based enterprise application built for **Frontier Tower Associates Philippines (FTAP)** to extract 101 standard telecom fields from Site Request Form (SRF) PDFs directly into beautifully formatted Microsoft Excel spreadsheets.

## Features
- **Official FTAP Brand Design**: Styled with Frontier Towers Philippines corporate crimson palette (`#C00000`), clear badges, and company header logo.
- **Pure In-Browser Processing**: Uploaded PDFs are processed strictly in RAM—no files or temporary documents are stored on disk.
- **Pre-Extraction Verification**: Review uploaded filenames and sizes before confirming extraction.
- **101 Fields Extracted**: General site info, infrastructure details, dismantling & additional equipment, RF antenna specs, power metrics, transmission solution, and sequential approvals.
- **Direct Excel Download**: Instant export to a styled `.xlsx` file with FTAP brand-coded headers, custom column widths, and blank cells for unpopulated fields.
- **High-Contrast, Modern UI**: Designed with 3D floating card elements, clear badges, and soft canvas to eliminate screen glare.

## Quick Start (Windows)
Simply double-click **`Launch_App.bat`** to start the app and open it in your browser.

Or run manually from terminal:
```bash
pip install -r requirements.txt
streamlit run app.py
``
