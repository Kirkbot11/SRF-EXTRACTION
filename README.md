# SRF Extraction & Daily Excel Generator

A lightweight, non-technical, browser-based application built with Python and Streamlit to extract 101 standard fields from Telecom Site Request Form (SRF) PDFs directly into formatted Microsoft Excel spreadsheets.

## Features
- **Pure In-Browser Processing**: Uploaded PDFs are processed strictly in RAM—no files or temporary documents are stored on disk.
- **Pre-Extraction Verification**: Review uploaded filenames and sizes before confirming extraction.
- **101 Fields Extracted**: General site info, infrastructure details, dismantling & additional equipment, RF antenna specs, power metrics, transmission solution, and sequential approvals.
- **Direct Excel Download**: Instant export to a styled `.xlsx` file with color-coded headers, custom column widths, and blank cells for unpopulated fields.
- **High-Contrast, Modern UI**: Designed with floating card elements, clear badges, and soft slate background to eliminate screen glare.

## Quick Start (Windows)
Simply double-click **`Launch_App.bat`** to start the app and open it in your browser.

Or run manually from terminal:
```bash
pip install -r requirements.txt
streamlit run app.py
```
Open [http://localhost:8501](http://localhost:8501) in your browser.
