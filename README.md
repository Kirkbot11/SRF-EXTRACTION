# Frontier Towers Philippines - Automated SRF Extractor & Excel Generator

A lightweight, non-technical, browser-based enterprise application built for **Frontier Tower Associates Philippines (FTAP)** to extract all 101 standard fields from Telecom Site Request Form (SRF) PDFs directly into beautifully formatted Microsoft Excel spreadsheets.

## Features
- **Official FTAP Brand Design**: Styled with Frontier Towers Philippines corporate crimson palette (`#C00000`), clear badges, and company header logo.
- **Pure In-Browser Processing**: Uploaded PDFs are processed strictly in RAM—no files or temporary documents are stored on disk.
- **Pre-Extraction Verification**: Review uploaded filenames and sizes before confirming extraction.
- **Batch Appending & Late SRF Support**: Upload PDFs in multiple batches throughout the day. Late-arriving SRFs seamlessly append into the active dataset without losing previous records.
- **101 Fields Extracted (Clean Blank Returns)**: General site info, infrastructure details, dismantling & additional equipment, RF antenna specs, power metrics, transmission solution, and sequential approvals. Any unpopulated field returns clean blank cells.
- **Direct Excel Download**: Instant export to a styled `.xlsx` file with FTAP brand-coded headers and custom column widths.
- **High-Contrast, Modern UI**: Designed with 3D floating card elements, clear badges, and soft canvas to eliminate screen glare.

## Quick Start (Windows)
Simply double-click **`Launch_App.bat`** to start the app and open it in your browser.

Or run manually from terminal:
```bash
pip install -r requirements.txt
streamlit run app.py
``
