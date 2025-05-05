# Invoice-generator
This Python desktop application streamlines vendor invoicing and statement-of-account (SOA)
## Invoice & Statement Management Application

**Overview:**

This Python desktop application streamlines vendor invoicing and statement-of-account (SOA) processes by combining:

* **Excel data ingestion**: Automatically detect headers and sum line items.
* **SQLite database**: Persistently store vendor profiles and invoice metadata.
* **PDF generation**: Render professional invoices and SOA reports with dynamic tables, headers, footers, signature images, and aging summaries.
* **Graphical User Interface**: Built using CustomTkinter for a modern, responsive look and leveraging tkcalendar for date selection.

---

## Table of Contents

1. [Key Features](#key-features)
2. [Technology Stack & Dependencies](#technology-stack--dependencies)
3. [Installation & Setup](#installation--setup)
4. [Directory Structure](#directory-structure)
5. [Database Schema](#database-schema)
6. [Core Modules & Functions](#core-modules--functions)

   * Excel Processing
   * PDF Generation (Invoices & SOA)
   * Aging Calculation
   * Excel Row Editor
7. [GUI Workflow](#gui-workflow)
8. [Usage Guide](#usage-guide)
9. [Customization & Extensibility](#customization--extensibility)
10. [Troubleshooting & FAQs](#troubleshooting--faqs)
11. [SEO Keywords](#seo-keywords)

---

## Key Features

* **Automated Header Detection**: Scans Excel sheets for rows containing both “name” and “amount” to set column headers dynamically.
* **Total Calculation**: Converts and aggregates numeric "amount" column values, handling missing or malformed data gracefully.
* **Vendor & Invoice Registry**: Maintains `vendors` and `invoices` tables in a local SQLite database, enabling multiple transaction cycles.
* **Professional PDF Output**: Uses ReportLab to generate A4 invoices and SOA PDFs with:

  * Page header/footer images
  * Tabular presentation with running balances
  * Aging buckets (current, 1–4+ months)
  * Embedded signature and seal images
* **Interactive Excel Editor**: Pop‑up grid to review and delete unwanted rows before final PDF generation.
* **User‑Friendly GUI**: Offers file dialogs, date pickers, and themed controls via CustomTkinter and ttk.

---

## Technology Stack & Dependencies

* Python 3.8+
* GUI:

  * `customtkinter`
  * `tkcalendar`
  * `ttk`
* Data:

  * `pandas` (Excel reading/manipulation)
  * `sqlite3` (built‑in DB)
* PDF:

  * `reportlab`
  * `num2words`
* Others:

  * `openpyxl` (optional for advanced Excel)
  * `Pillow` (image handling)

Install with:

```bash
pip install customtkinter tkcalendar pandas reportlab num2words pillow
```

---

## Installation & Setup

1. **Clone or copy** the project folder.
2. Ensure **header.png**, **footer.png**, **signeture.jpg**, **ss.jpg**, and **seal.png** exist in the working directory.
3. Install dependencies: see above.
4. Run the main GUI script:

   ```bash
   python invoice_manager.py
   ```

No external config required; database file `app.db` auto‑generates on first launch.

---

## Directory Structure

```
/InvoiceApp/
├── invoice_manager.py        # Main application script
├── app.db                    # SQLite database (auto-created)
├── header.png                # PDF header image
├── footer.png                # PDF footer image
├── signeture.jpg             # Signature image
├── ss.jpg                    # Secondary signature
├── seal.png                  # Company seal
├── README.md                 # This documentation
└── requirements.txt          # pip install list
```

---

## Database Schema

### Table: `vendors`

| Column          | Type    | Description                |
| --------------- | ------- | -------------------------- |
| id              | INTEGER | Auto-increment primary key |
| vendor\_id      | TEXT    | Unique vendor identifier   |
| vendor\_name    | TEXT    | Full vendor name           |
| vendor\_address | TEXT    | Mailing or billing address |
| po\_number      | TEXT    | Purchase order reference   |

### Table: `invoices`

| Column        | Type    | Description                                     |
| ------------- | ------- | ----------------------------------------------- |
| id            | INTEGER | Auto-increment primary key                      |
| vendor\_id    | TEXT    | Foreign key to `vendors.vendor_id`              |
| invoice\_no   | TEXT    | Invoice number                                  |
| invoice\_date | TEXT    | Date string (YYYY-MM-DD)                        |
| invoice\_type | TEXT    | "Debit" or "Credit"                             |
| po\_mr\_no    | TEXT    | Related PO/MR reference                         |
| excel\_file   | TEXT    | Path to original Excel sheet for record-keeping |

---

## Core Modules & Functions

### 1. Excel Processing (`process_excel_file`)

* **Input:** Path to `.xlsx` file
* **Workflow:**

  1. Read without headers
  2. Detect header row containing “name” & “amount”
  3. Re-read with that header
  4. Locate “amount” column, convert to numeric
  5. Sum amounts, return `(DataFrame, total)`
* **Errors:** Raises descriptive `ValueError` on missing header or column

### 2. Aging Calculation (`compute_aging`)

* **Input:** List of tuples `(invoice_date, _, _, debit, credit)`
* **Output:** Dict of buckets `{current, 1month, 2months, 3months, 4plus, total}`
* **Logic:** Days since invoice grouped into 0–30, 31–60, … >120

### 3. PDF Table Builder (`create_report_table_pdf`)

* **Input:**

  * `output_path`: PDF filename
  * `title`: Report title
  * `data_rows`: List of `(date, inv_no, name, debit, credit)`
  * `balance_bf`: Starting balance
  * `aging_summary`: Optional aging data
* **Features:**

  * Adds balance b/f and sub-total lines
  * Formats columns, right-aligns numbers
  * Embeds footer/header via callback

### 4. Invoice PDF (`create_invoice_pdf_modified`)

* Combines vendor info, invoice metadata, and a one‑row table of processed Excel total
* Uses `create_report_table_pdf` for the table portion
* Renders form & banker details tables side by side

### 5. Full Invoice & SOA PDF (`create_invoice_pdf`, `create_soa_pdf_modified`)

* Extended versions that embed:

  * Excel line‑item table with wrapped cell text
  * Multi‑column totals with amount in words
  * Optionally include company seal image

### 6. Excel Row Editor (`open_excel_editor`)

* Pop‑up `Toplevel` window presenting Excel rows in a `ttk.Treeview`
* Allows multi‑row deletion before final save

---

## GUI Workflow

1. **Vendor Entry:** Fill vendor ID, name, address, PO number → “Add Vendor”.
2. **Invoice Import:** Select vendor and Excel file → system auto‑processes and displays total.
3. **Invoice Metadata:** Enter invoice number, date, type, related PO/MR → “Save Invoice”.
4. **Preview & Edit:** Optionally open Excel editor to remove unwanted lines.
5. **PDF Generation:** Click “Generate PDF” → choose save location → receive formatted invoice/SOA.
6. **History:** All vendors and invoices listed in SQLite for lookup and re-generation.

---

## Usage Guide

* **Add New Vendor**: Click ▶️, enter fields, click “Save Vendor.”
* **Load Excel**: Browse file dialog, ensure columns `Name` & `Amount` exist.
* **Edit Data**: Click “Edit Excel” to drop unwanted rows.
* **Generate PDF**: Click “Create Invoice PDF” or “Create SOA PDF.”
* **View Records**: Invoice and vendor lists accessible via menu.

---

## Customization & Extensibility

* **Themes**: Customize CustomTkinter appearance.
* **DB Path**: Change `DB_FILE` constant.
* **PDF Layout**: Modify ReportLab styles or replace header/footer images.
* **Currency Words**: Swap out `num2words` language parameter.
* **Advanced Excel**: Extend `process_excel_file` to support multiple currencies or sheets.

---

## Troubleshooting & FAQs

* **Missing Header Error**: Ensure Excel contains columns labelled “Name” and “Amount.”
* **Image Not Found**: Place `header.png`, `footer.png`, etc., in same folder or adjust paths.
* **SQLite Locked**: Close other connections or delete `app.db` to reset.
* **Long Tables**: For large Excel files, consider chunking or increasing PDF margins.

---

## SEO Keywords

```
invoice generation python
excel to pdf automation
customtkinter invoice app
sqlite invoice management
reportlab pdf invoice
aging summary python
vendor invoice tool
num2words invoice
tkcalendar date picker
ttk treeview sqlite
excel header detection python
invoice generation python
excel to pdf automation
customtkinter invoice app
sqlite invoice management
reportlab pdf invoice
aging summary python
vendor invoice tool
num2words invoice
tkcalendar date picker
ttk treeview sqlite
excel header detection python
```

---

**Author:** Your Name – © 2025

**License:** MIT
