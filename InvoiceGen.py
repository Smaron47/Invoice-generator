import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
import pandas as pd
import sqlite3
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from num2words import num2words
import os


# ----------------------------------------------------
# Database Setup (SQLite)
# ----------------------------------------------------
DB_FILE = "app.db"
conn = sqlite3.connect(DB_FILE)
cursor = conn.cursor()

# Vendors table
cursor.execute("""
    CREATE TABLE IF NOT EXISTS vendors (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vendor_id TEXT,
        vendor_name TEXT,
        vendor_address TEXT,
        po_number TEXT
    )
""")

# Invoices table – note: for simplicity, we store the original Excel file path.
cursor.execute("""
    CREATE TABLE IF NOT EXISTS invoices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vendor_id TEXT,
        invoice_no TEXT,
        invoice_date TEXT,
        invoice_type TEXT,
        po_mr_no TEXT,
        excel_file TEXT
    )
""")
conn.commit()

# ----------------------------------------------------
# Utility: Process Excel File
# ----------------------------------------------------
def process_excel_file(excel_path):
    """
    Reads the Excel file, finds a row that contains both 'name' and 'amount' (case-insensitive),
    and then uses that row as the header. It then converts the column named "amount" to numeric,
    sums its values, and returns the DataFrame along with the total sum.
    """
    try:
        # Read the Excel file without headers
        df = pd.read_excel(excel_path, header=None)
        header_row = None
        # Iterate through rows to find a header row that contains both 'name' and 'amount'
        for i, row in df.iterrows():
            # Convert each cell to a lower-case string
            row_str = row.astype(str).str.lower()
            if 'name' in row_str.values and 'amount' in row_str.values:
                header_row = i
                break

        if header_row is None:
            raise ValueError("No header row found containing both 'name' and 'amount'.")

        # Re-read the Excel file using the found header row
        df = pd.read_excel(excel_path, header=header_row)
        # Search for the column named "amount" (case-insensitive)
        df = df.iloc[:-1]
        amount_col = None
        for col in df.columns:
            if str(col).strip().lower() == 'amount':
                amount_col = col
                break

        if amount_col is None:
            raise ValueError("No column named 'amount' found.")

        df = df.drop_duplicates()
        # Convert the 'amount' column to numeric and fill NaN with 0
        df[amount_col] = pd.to_numeric(df[amount_col], errors='coerce').fillna(0)
        print(df)
        total = df[amount_col].sum()

        return df, total

    except Exception as e:
        raise ValueError(f"Error processing Excel file: {e}")
# ----------------------------------------------------
# PDF Generation Helpers – New Table Format
# ----------------------------------------------------
def add_page_header_footer(canvas, doc):
    page_width, page_height = A4
    try:
        canvas.drawImage("header.png", 0, page_height - 80, width=page_width, height=80)
    except:
        canvas.drawString(10, page_height - 50, "[Header Image Missing]")
    try:
        canvas.drawImage("footer.png", 0, 0, width=page_width, height=80)
    except:
        canvas.drawString(10, 30, "[Footer Image Missing]")

def create_report_table_pdf(output_path, title, data_rows, balance_bf=0.0, aging_summary=None):
    """
    Generates a PDF with a table having columns:
    Date | Invoice # | Name | Debit | Credit | Balance
    It adds:
      - A "Balance b/f" row,
      - Each data row (with running balance computed),
      - A "Sub-Total" row,
      - And an optional aging summary.
    """
    doc = SimpleDocTemplate(output_path, pagesize=A4, topMargin=90, bottomMargin=90)
    elements = []
    styles = getSampleStyleSheet()
    page_width, page_height = A4
    normal_style = ParagraphStyle(
        'normal_style',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=12
    )

    # Title
    elements.append(Paragraph(title, styles['Title']))
    elements.append(Spacer(1, 15))
    # elements.append(Paragraph(, ParagraphStyle('Title', parent=styles['Title'], alignment=1)))
    elements.append(Spacer(1, 15))

    # Build table data
    table_data = []
    header = ["Date", "Invoice #", "Name", "Debit", "Credit", "Balance"]
    table_data.append(header)
    
    # Balance b/f row
    table_data.append(["", "", Paragraph("<b>Balance b/f</b>", normal_style), "", "", f"{balance_bf:,.2f}"])
    
    running_balance = balance_bf
    total_debit = 0.0
    total_credit = 0.0
    
    # Each row in data_rows is expected to be a tuple:
    # (invoice_date, invoice_no, Name, debit, credit)
    for row in data_rows:
        inv_date, inv_no, Name, debit, credit = row
        debit = float(debit) if debit else 0.0
        credit = float(credit) if credit else 0.0
        running_balance += (debit - credit)
        total_debit += debit
        total_credit += credit
        table_data.append([
            inv_date, inv_no, Name,
            f"{debit:,.2f}" if debit else "",
            f"{credit:,.2f}" if credit else "",
            f"{running_balance:,.2f}"
        ])
    
    # Sub-Total row
    table_data.append(["", "", Paragraph("<b>Sub-Total</b>", normal_style),
                        f"{total_debit:,.2f}",
                        f"{total_credit:,.2f}",
                        f"{running_balance:,.2f}"])
    
    t = Table(table_data, colWidths=[70, 80, 100, 60, 60, 70])
    t.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('ALIGN', (3,1), (4,-1), 'RIGHT'),
        ('ALIGN', (5,1), (5,-1), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('BOTTOMPADDING', (0,0), (-1,0), 5),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 12))

    if aging_summary:
        aging_header = ["Current Month", "1 Month", "2 Months", "3 Months", "4 Months & Above", "Total"]
        aging_data = [
            aging_header,
            [f"{aging_summary.get('current', 0):,.2f}",
             f"{aging_summary.get('1month', 0):,.2f}",
             f"{aging_summary.get('2months', 0):,.2f}",
             f"{aging_summary.get('3months', 0):,.2f}",
             f"{aging_summary.get('4plus', 0):,.2f}",
             f"{aging_summary.get('total', 0):,.2f}"]
        ]
        aging_table = Table(aging_data, colWidths=[80, 60, 60, 60, 80, 60])
        aging_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('BACKGROUND', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,1), (-1,1), 'RIGHT'),
        ]))
        elements.append(aging_table)
    try:
        signature_img = "signeture.jpg"
        elements.append(Image(signature_img, width=page_width * 0.95, height=50))
    except Exception as e:
        elements.append(Paragraph("[Signature Image Missing]", normal_style))
    elements.append(Spacer(1, 20))
    try:
        second_signature = "ss.jpg"
        elements.append(Image(second_signature, width=page_width * 0.95, height=10))
    except Exception as e:
        elements.append(Paragraph("[Second Signature Image Missing]", normal_style))
    elements.append(Spacer(1, 20))
    doc.build(elements, onFirstPage=add_page_header_footer, onLaterPages=add_page_header_footer)

def compute_aging(data_rows):
    """
    Computes an aging summary from data_rows (list of tuples: (invoice_date, ..., debit, credit)).
    Buckets (example):
      - current: 0-30 days,
      - 1month: 31-60,
      - 2months: 61-90,
      - 3months: 91-120,
      - 4plus: over 120 days.
    Sums the net (debit - credit) for each bucket.
    """
    today = datetime.now().date()
    buckets = {"current": 0.0, "1month": 0.0, "2months": 0.0, "3months": 0.0, "4plus": 0.0}
    for row in data_rows:
        inv_date_str, _, _, debit, credit = row
        try:
            d = datetime.strptime(inv_date_str, "%Y-%m-%d").date()
        except:
            d = today
        delta = (today - d).days
        net = (float(debit) if debit else 0.0) - (float(credit) if credit else 0.0)
        if delta <= 30:
            buckets["current"] += net
        elif delta <= 60:
            buckets["1month"] += net
        elif delta <= 90:
            buckets["2months"] += net
        elif delta <= 120:
            buckets["3months"] += net
        else:
            buckets["4plus"] += net
    buckets["total"] = sum(buckets.values())
    return buckets

# ----------------------------------------------------
# PDF Generation Functions for Invoices & SOA (Modified)
# ----------------------------------------------------




def create_invoice_pdf_modified(output_path, input_details, total_amount):
    """
    Creates an invoice PDF that includes header sections (vendor details, form & banker details)
    and then a table built in the new format. The table will have one row (the processed Excel total).
    Based on invoice type, the amount is placed in Debit (if invoice type is "Credit")
    or in Credit (if invoice type is "Debit"). Running balance starts from zero.
    """
    # Build the header details as before:
    doc = SimpleDocTemplate(output_path, pagesize=A4, topMargin=90, bottomMargin=90)
    elements = []
    styles = getSampleStyleSheet()
    page_width, page_height = A4

    normal_style = ParagraphStyle('normal_style', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=12)

    # Title
    elements.append(Paragraph("Invoice", styles['Title']))
    elements.append(Spacer(1, 20))

    # Vendor & Invoice Details
    header_style = ParagraphStyle('header_style', parent=normal_style, alignment=1, backColor=colors.lightblue, fontName='Helvetica-Bold')
    detail_style = ParagraphStyle('detail_style', parent=normal_style, alignment=0)

    left_data = [
        [Paragraph("<b>VENDOR DETAILS (TO)</b>", header_style)],
        [Paragraph(f"<b>{input_details.get('vendor_name', '')}</b><br/>{input_details.get('vendor_address', '')}<br/><br/><br/>", detail_style)]
    ]
    left_table = Table(left_data, colWidths=[page_width * 0.45])
    left_table.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 1, colors.black),
                                    ('LEFTPADDING', (0,0), (-1,-1), 8),
                                    ('RIGHTPADDING', (0,0), (-1,-1), 8),
                                    ('TOPPADDING', (0,0), (-1,-1), 6),
                                    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                                    ('VALIGN', (0,0), (-1,-1), 'TOP')]))
    right_data = [
        [Paragraph("<b>INVOICE DETAILS</b>", header_style)],
        [Paragraph(f"INVOICE TYPE: {input_details.get('invoice_type', '')}<br/>"
                   f"INVOICE NO: {input_details.get('invoice_no', '')}<br/>"
                   f"PO/MR No: {input_details.get('vendor_po', '')}<br/>"
                   f"DATE: {input_details.get('invoice_date', '')}", detail_style)]
    ]
    right_table = Table(right_data, colWidths=[page_width * 0.45])
    right_table.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 1, colors.black),
                                     ('LEFTPADDING', (0,0), (-1,-1), 8),
                                     ('RIGHTPADDING', (0,0), (-1,-1), 8),
                                     ('TOPPADDING', (0,0), (-1,-1), 6),
                                     ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                                     ('VALIGN', (0,0), (-1,-1), 'TOP')]))
    container = Table([[left_table, Spacer(1,10), right_table]], colWidths=[page_width*0.45, 10, page_width*0.45])
    container.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    elements.append(container)
    elements.append(Spacer(1, 15))

    # Form & Banker Details (hard-coded as before)
    left_data = [
        [Paragraph("<b>FORM</b>", header_style)],
        [Paragraph("BOX NO: 80697<br/>NO: 182-WIDAM BUILDING<br/>ABU HAMOUR –DOHA", detail_style)]
    ]
    left_table = Table(left_data, colWidths=[page_width * 0.45])
    left_table.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 1, colors.black),
                                    ('LEFTPADDING', (0,0), (-1,-1), 8),
                                    ('RIGHTPADDING', (0,0), (-1,-1), 8),
                                    ('TOPPADDING', (0,0), (-1,-1), 6),
                                    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                                    ('VALIGN', (0,0), (-1,-1), 'TOP')]))
    right_data = [
        [Paragraph("<b>BANKER DETAILS</b>", header_style)],
        [Paragraph("TRADE NAME: DOCMED SERVICES<br/>Account No: 0250561138001<br/>BANK: QNB –AIN KHALED<br/>IBAN: QA98QNBA000000000250561138001", detail_style)]
    ]
    right_table = Table(right_data, colWidths=[page_width * 0.45])
    right_table.setStyle(TableStyle([('BOX', (0, 0), (-1, -1), 1, colors.black),
                                     ('LEFTPADDING', (0,0), (-1,-1), 8),
                                     ('RIGHTPADDING', (0,0), (-1,-1), 8),
                                     ('TOPPADDING', (0,0), (-1,-1), 6),
                                     ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                                     ('VALIGN', (0,0), (-1,-1), 'TOP')]))
    container = Table([[left_table, Spacer(1,10), right_table]], colWidths=[page_width*0.45, 10, page_width*0.45])
    container.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    elements.append(container)
    elements.append(Spacer(1, 15))

    # New Table: Build a one-row table from processed Excel total.
    invoice_date = input_details.get("invoice_date", "")
    invoice_no = input_details.get("invoice_no", "")
    inv_type = input_details.get("invoice_type", "").lower()
    # For our table, if type is "credit" then total goes to Debit column; if "debit" then to Credit column.
    if inv_type == "credit":
        debit_val = total_amount
        credit_val = 0.0
    else:
        debit_val = 0.0
        credit_val = total_amount

    # Data row for our table
    data_rows = [(invoice_date, invoice_no, "", debit_val, credit_val)]
    # Generate PDF table with new format using our helper (balance b/f=0)
    aging = compute_aging(data_rows)
    # Title for invoice report PDF
    new_title = f"Invoice Report for Invoice #{invoice_no}"
    # Use our report table function to build the table part.
    # (This function adds a header row, a "Balance b/f" row, each data row with running balance, and a sub-total.)
    create_report_table_pdf(output_path, new_title, data_rows, balance_bf=0.0, aging_summary=aging)
    # After table, we could append additional images if desired.
    # For brevity, we assume header/footer images are added by add_page_header_footer.




def wrap_cell_text(text, style):
    if not isinstance(text, str):
        text = str(text)
    return Paragraph(text, style)

# ---------------------------
# PDF Creation Functions
# ---------------------------
# def add_page_header_footer(canvas, doc):
#     page_width, page_height = A4
#     try:
#         canvas.drawImage("header.png", 0, page_height - 80, width=page_width, height=80)
#     except Exception as e:
#         canvas.drawString(10, page_height - 50, "[Header Image Missing]")
#     try:
#         canvas.drawImage("footer.png", 0, 0, width=page_width, height=80)
#     except Exception as e:
#         canvas.drawString(10, 30, "[Footer Image Missing]")


def open_excel_editor(excel_path, parent):
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load Excel file: {e}")
        return None

    editor_win = tk.Toplevel(parent)
    editor_win.title("Excel Editor - Delete Rows")
    editor_win.geometry("800x400")
    
    tree = ttk.Treeview(editor_win, show="headings")
    tree.pack(fill="both", expand=True)
    
    tree["columns"] = list(df.columns)
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100, anchor="center")
    
    for index, row in df.iterrows():
        tree.insert("", "end", iid=index, values=list(row))
    
    def delete_selected():
        selected = tree.selection()
        for item in selected:
            tree.delete(item)
        nonlocal df
        df = df.drop(index=[int(item) for item in selected])
        df.reset_index(drop=True, inplace=True)
    
    def save_and_close():
        editor_win.destroy()
    
    btn_frame = tk.Frame(editor_win)
    btn_frame.pack(pady=5)
    tk.Button(btn_frame, text="Delete Selected Row(s)", command=delete_selected).pack(side="left", padx=5)
    tk.Button(btn_frame, text="Save & Close", command=save_and_close).pack(side="left", padx=5)
    
    editor_win.grab_set()
    editor_win.wait_window()
    return df

def create_invoice_pdf(output_path, input_details, excel_df,amount, include_seal=True):
    doc = SimpleDocTemplate(output_path, pagesize=A4, topMargin=90, bottomMargin=90)
    elements = []
    styles = getSampleStyleSheet()
    page_width, page_height = A4

    normal_style = ParagraphStyle(
        'normal_style',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=12
    )
    wrap_style = ParagraphStyle(
        'wrap_style',
        parent=normal_style,
        wordWrap='CJK'
    )

    # Title
    elements.append(Paragraph("Invoice", styles['Title']))
    elements.append(Spacer(1, 20))

    # ---------------------------------------------------------------
    # 1. Top Row – Vendor & Invoice Details
    # ---------------------------------------------------------------
    header_style = ParagraphStyle(
        'header_style',
        parent=normal_style,
        alignment=1,
        backColor=colors.lightblue,
        fontName='Helvetica-Bold'
    )
    detail_style = ParagraphStyle(
        'detail_style',
        parent=normal_style,
        alignment=0
    )

    left_data = [
        [Paragraph("<b>VENDOR DETAILS (TO)</b>", header_style)],
        [Paragraph(f"<b>{input_details.get('vendor_name', '')}</b><br/>{input_details.get('vendor_address', '')}<br/><br/><br/>", detail_style)]
    ]
    left_table = Table(left_data, colWidths=[page_width * 0.45])
    left_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))

    right_data = [
        [Paragraph("<b>INVOICE DETAILS</b>", header_style)],
        [Paragraph(
            f"INVOICE TYPE: {input_details.get('invoice_type', '')}<br/>"
            f"INVOICE NO: {input_details.get('invoice_no', '')}<br/>"
            f"PO/MR No: {input_details.get('vendor_po', '')}<br/>"
            f"DATE: {input_details.get('invoice_date', '')}",
            detail_style)
        ]
    ]
    right_table = Table(right_data, colWidths=[page_width * 0.45])
    right_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    container = Table([[left_table, Spacer(1,10), right_table]], colWidths=[page_width*0.45, 10, page_width*0.45])
    container.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    elements.append(container)
    elements.append(Spacer(1, 15))

    # ---------------------------------------------------------------
    # 2. Second Row – Form & Bank Details
    # ---------------------------------------------------------------
    left_data = [
        [Paragraph("<b>FORM</b>", header_style)],
        [Paragraph(
            f"<b>DOCMED SERVICES</b><br/>"
            f"BOX NO: 80697<br/>"
            f"NO: 182-WIDAM BUILDING<br/>"
            f"ABU HAMOUR –DOHA",
            detail_style)]
    ]
    left_table = Table(left_data, colWidths=[page_width * 0.45])
    left_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    right_data = [
        [Paragraph("<b>BANKER DETAILS</b>", header_style)],
        [Paragraph(
            f"TRADE NAME :  DOCMED SERVICES<br/>"
            f"Account No :  0250561138001<br/>"
            f"BANK      :   QNB –AIN KHALED<br/>"
            f"IBAN    :  QA98QNBA000000000250561138001",
            detail_style)
        ]
    ]
    right_table = Table(right_data, colWidths=[page_width * 0.45])
    right_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    container = Table([[left_table, Spacer(1,10), right_table]], colWidths=[page_width*0.45, 10, page_width*0.45])
    container.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    elements.append(container)
    elements.append(Spacer(1, 15))

    # ---------------------------------------------------------------
    # 3. Section Title for Invoice Items
    # ---------------------------------------------------------------
    elements.append(Paragraph("Invoice Details: Pre-medical Employment", styles['Title']))
    elements.append(Spacer(1, 10))

    # ---------------------------------------------------------------
    # 4. Excel Data Table with Wrapped Text
    # ---------------------------------------------------------------
    if excel_df is not None and not excel_df.empty:
        data = []
        header_row = [wrap_cell_text(col, wrap_style) for col in excel_df.columns]
        data.append(header_row)
        for row in excel_df.values:
            wrapped_row = [wrap_cell_text("" if pd.isnull(cell) else str(cell), wrap_style) for cell in row]
            data.append(wrapped_row)
       
        num_cols = len(data[0])
        col_widths_ratio = []
        for col in range(num_cols):
            max_len = max([len(str(excel_df.iloc[r, col])) for r in range(len(excel_df))] + [len(str(excel_df.columns[col]))])
            if col == 0:
                max_len = max(max_len, 5)
            col_widths_ratio.append(max_len)
        total_ratio = sum(col_widths_ratio)
        table_width = page_width * 0.95
        colWidths = [table_width * (ratio / total_ratio) for ratio in col_widths_ratio]

        excel_table = Table(data, colWidths=colWidths)
        excel_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))
        elements.append(excel_table)
        elements.append(Spacer(1, 15))

        # ---------------------------------------------------------------
        # 5. Total Calculation with Custom Currency Conversion
        # ---------------------------------------------------------------


        total_amount = amount

        total_str = f"Total: {total_amount}"
        amount_int = int(round(total_amount))
        total_words = num2words(amount_int, lang='en').title() + " Riyals Only"
        total_data = [
            [Paragraph("<b>Total:</b>", normal_style), Paragraph(total_str, normal_style), Paragraph(total_words, normal_style)]
        ]
        total_table = Table(total_data, colWidths=[table_width*0.15, table_width*0.40, table_width*0.40])
        total_table.setStyle(TableStyle([
            ('BOX', (0,0), (-1,-1), 1, colors.black),
            ('BACKGROUND', (0,0), (0,0), colors.lightsteelblue),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ]))
        elements.append(total_table)
        elements.append(Spacer(1, 20))

    # ---------------------------------------------------------------
    # 6. Signature Images Section
    # ---------------------------------------------------------------
    try:
        signature_img = "signeture.jpg"
        elements.append(Image(signature_img, width=page_width * 0.95, height=50))
    except Exception as e:
        elements.append(Paragraph("[Signature Image Missing]", normal_style))
    elements.append(Spacer(1, 20))
    try:
        second_signature = "ss.jpg"
        elements.append(Image(second_signature, width=page_width * 0.95, height=10))
    except Exception as e:
        elements.append(Paragraph("[Second Signature Image Missing]", normal_style))
    elements.append(Spacer(1, 20))
    if include_seal:
        try:
            seal_img = "seal.png"
            elements.append(Image(seal_img, width=page_width * 0.10, height=50))
        except Exception as e:
            elements.append(Paragraph("[Seal Image Missing]", normal_style))
    elements.append(Spacer(1, 20))

    doc.build(elements, onFirstPage=add_page_header_footer, onLaterPages=add_page_header_footer)





def create_soa_pdf_modified(output_path, soa_info, invoices_data):
    """
    Generates an SOA PDF similar to invoice PDF but with SOA header details.
    invoices_data is a list of tuples: (invoice_date, invoice_no, Name, debit, credit)
    The table is built with our new format.
    """
    doc = SimpleDocTemplate(output_path, pagesize=A4, topMargin=90, bottomMargin=90)
    elements = []
    styles = getSampleStyleSheet()
    normal_style = ParagraphStyle('normal_style', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=12)
    
    # Header with statement details
    header_text = f"""
    <para align=center>
    <b>Statement Date:</b> {soa_info.get('statement_date','')} &nbsp;&nbsp;&nbsp;
    <b>Due Date:</b> {soa_info.get('due_date','')}
    </para>
    """
    elements.append(Paragraph(header_text, normal_style))
    elements.append(Spacer(1, 10))
    company_details = f"""
    <para align=center>
    <font color="darkblue"><b>STATEMENT OF ACCOUNT</b><br/>
    {soa_info.get('company_name','')}<br/>
    {soa_info.get('company_address','')}</font>
    </para>
    """
    elements.append(Paragraph(company_details, styles['Title']))
    elements.append(Spacer(1, 15))
    
    # Build table from invoices_data (same new table format)
    aging = compute_aging(invoices_data)
    create_report_table_pdf(output_path, "Statement of Account", invoices_data, balance_bf=0.0, aging_summary=aging)
    # Note: In a complete solution, you might want to merge multiple invoices into one table.
    # Here we assume invoices_data is already the merged list.
    
# ----------------------------------------------------
# Main Application (Single Window with Frames)
# ----------------------------------------------------
class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("DocMed Qatar - Single Window")
        self.geometry("1100x700")

        # Top Header (Logo, Title, Search)
        header_frame = ctk.CTkFrame(self, corner_radius=0)
        header_frame.pack(side="top", fill="x")
        self.logo_label = ctk.CTkLabel(header_frame, text="DocMed Qatar", font=("Arial", 20, "bold"))
        self.logo_label.pack(side="left", padx=20, pady=10)
        self.search_var = tk.StringVar()
        self.search_entry = ctk.CTkEntry(header_frame, textvariable=self.search_var, placeholder_text="Search...")
        self.search_entry.pack(side="right", padx=20, pady=10)
        
        # Navigation Bar
        nav_frame = ctk.CTkFrame(self, corner_radius=0)
        nav_frame.pack(side="top", fill="x")
        self.btn_supplier = ctk.CTkButton(nav_frame, text="Supplier Creation", command=self.show_supplier_frame)
        self.btn_supplier.pack(side="left", padx=5, pady=5)
        self.btn_invoice = ctk.CTkButton(nav_frame, text="Invoice Creation", command=self.show_invoice_frame)
        self.btn_invoice.pack(side="left", padx=5, pady=5)
        self.btn_report = ctk.CTkButton(nav_frame, text="Invoice Reports", command=self.show_report_frame)
        self.btn_report.pack(side="left", padx=5, pady=5)
        self.btn_soa = ctk.CTkButton(nav_frame, text="SOA Reports", command=self.show_soa_frame)
        self.btn_soa.pack(side="left", padx=5, pady=5)

        # Main Content Area (Frames)
        self.content_frame = ctk.CTkFrame(self, corner_radius=0)
        self.content_frame.pack(side="top", fill="both", expand=True)
        self.supplier_frame = ctk.CTkFrame(self.content_frame)
        self.invoice_frame = ctk.CTkFrame(self.content_frame)
        self.report_frame = ctk.CTkFrame(self.content_frame)
        self.soa_frame = ctk.CTkFrame(self.content_frame)
        for f in (self.supplier_frame, self.invoice_frame, self.report_frame, self.soa_frame):
            f.place(in_=self.content_frame, x=0, y=0, relwidth=1, relheight=1)

        # Build Frames
        self.build_supplier_frame()
        self.build_invoice_frame()
        self.build_report_frame()
        self.build_soa_frame()

        self.show_supplier_frame()

    def show_frame(self, frame: ctk.CTkFrame):
        frame.lift()

    def show_supplier_frame(self):
        self.show_frame(self.supplier_frame)

    def show_invoice_frame(self):
        self.show_frame(self.invoice_frame)

    def show_report_frame(self):
        self.show_frame(self.report_frame)

    def show_soa_frame(self):
        self.show_frame(self.soa_frame)

    # -----------------------------
    # 1) Supplier Creation Frame
    # -----------------------------
    def build_supplier_frame(self):
        tk.Label(self.supplier_frame, text="Supplier Creation", font=("Arial", 18, "bold")).pack(pady=10)
        form_frame = tk.Frame(self.supplier_frame)
        form_frame.pack(pady=10)
        tk.Label(form_frame, text="Vendor ID:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.vendor_id_entry = tk.Entry(form_frame, width=30)
        self.vendor_id_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Label(form_frame, text="Vendor Name:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.vendor_name_entry = tk.Entry(form_frame, width=30)
        self.vendor_name_entry.grid(row=1, column=1, padx=5, pady=5)
        tk.Label(form_frame, text="Vendor Address:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.vendor_address_entry = tk.Entry(form_frame, width=30)
        self.vendor_address_entry.grid(row=2, column=1, padx=5, pady=5)
        tk.Label(form_frame, text="PO/MR Number:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.vendor_po_entry = tk.Entry(form_frame, width=30)
        self.vendor_po_entry.grid(row=3, column=1, padx=5, pady=5)
        tk.Button(self.supplier_frame, text="Add Vendor", command=self.add_vendor).pack(pady=10)

    def add_vendor(self):
        vid = self.vendor_id_entry.get().strip()
        vname = self.vendor_name_entry.get().strip()
        vaddr = self.vendor_address_entry.get().strip()
        vpo = self.vendor_po_entry.get().strip()
        if not vid or not vname or not vaddr or not vpo:
            messagebox.showerror("Error", "All fields are required.")
            return
        cursor.execute("INSERT INTO vendors (vendor_id, vendor_name, vendor_address, po_number) VALUES (?,?,?,?)",
                       (vid, vname, vaddr, vpo))
        conn.commit()
        messagebox.showinfo("Success", "Vendor added successfully.")
        self.vendor_id_entry.delete(0, tk.END)
        self.vendor_name_entry.delete(0, tk.END)
        self.vendor_address_entry.delete(0, tk.END)
        self.vendor_po_entry.delete(0, tk.END)

    # -----------------------------
    # 2) Invoice Creation Frame
    # -----------------------------
    def build_invoice_frame(self):
        tk.Label(self.invoice_frame, text="Invoice Creation", font=("Arial", 18, "bold")).pack(pady=10)
        cursor.execute("SELECT vendor_name, vendor_id, vendor_address, po_number FROM vendors")
        self.vendors = cursor.fetchall()
        vendor_names = [v[0] for v in self.vendors]
        form_frame = tk.Frame(self.invoice_frame)
        form_frame.pack(pady=10)
        tk.Label(form_frame, text="Select Vendor:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.vendor_var = tk.StringVar()
        self.vendor_dropdown = ctk.CTkOptionMenu(form_frame, values=vendor_names, variable=self.vendor_var)
        self.vendor_dropdown.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.vendor_var.trace("w", self.fill_vendor_details)
        tk.Label(form_frame, text="Vendor Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.invoice_vendor_name_var = tk.StringVar()
        tk.Entry(form_frame, textvariable=self.invoice_vendor_name_var, width=40, state="readonly").grid(row=1, column=1, padx=5, pady=5)
        tk.Label(form_frame, text="Vendor ID:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.invoice_vendor_id_var = tk.StringVar()
        tk.Entry(form_frame, textvariable=self.invoice_vendor_id_var, width=40, state="readonly").grid(row=2, column=1, padx=5, pady=5)
        tk.Label(form_frame, text="Vendor Address:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.invoice_vendor_address_var = tk.StringVar()
        tk.Entry(form_frame, textvariable=self.invoice_vendor_address_var, width=40, state="readonly").grid(row=3, column=1, padx=5, pady=5)
        tk.Label(form_frame, text="Invoice Type (Credit/Debit):").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.invoice_type_entry = tk.Entry(form_frame, width=40)
        self.invoice_type_entry.grid(row=4, column=1, padx=5, pady=5)
        tk.Label(form_frame, text="Invoice No:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
        self.invoice_no_entry = tk.Entry(form_frame, width=40)
        self.invoice_no_entry.grid(row=5, column=1, padx=5, pady=5)
        tk.Label(form_frame, text="Invoice Date:").grid(row=7, column=0, padx=5, pady=5, sticky="w")
        self.invoice_date_entry = DateEntry(form_frame, width=37, date_pattern="yyyy-mm-dd")
        self.invoice_date_entry.grid(row=7, column=1, padx=5, pady=5)
        self.include_seal_var = tk.BooleanVar(value=True)
        tk.Checkbutton(form_frame, text="Include Seal", variable=self.include_seal_var).grid(row=8, column=1, padx=5, pady=5, sticky="w")
        tk.Label(form_frame, text="Excel File:").grid(row=9, column=0, padx=5, pady=5, sticky="w")
        excel_file_var = tk.StringVar()
        def browse_excel():
            path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
            if path:
                excel_file_var.set(path)
                edited_df = open_excel_editor(path, form_frame)
                if edited_df is not None:
                    excel_file_var.edited_df = edited_df
        tk.Label(form_frame, text="Excel File:").grid(row=9, column=0, padx=5, pady=5, sticky="w")
        tk.Entry(form_frame, textvariable=excel_file_var, width=40).grid(row=9, column=1, padx=5, pady=5)
        tk.Button(form_frame, text="Browse", command=browse_excel).grid(row=9, column=2, padx=5, pady=5)
        self.progress_label = tk.Label(self.invoice_frame, text="", fg="green", font=("Helvetica", 10))
        self.progress_label.pack(pady=5)
        tk.Button(self.invoice_frame, text="Generate Invoice", command=self.generate_invoice).pack(pady=15)

    def fill_vendor_details(self, *args):
        selected = self.vendor_var.get()
        for v in self.vendors:
            if v[0] == selected:
                self.invoice_vendor_name_var.set(v[0])
                self.invoice_vendor_id_var.set(v[1])
                self.invoice_vendor_address_var.set(v[2])
                break

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            self.excel_file_var.set(path)
            # Optionally, you can allow editing via a pop-up editor if needed.
            # For now we simply store the file path.
    
    def generate_invoice(self):
        self.progress_label.config(text="Generating Invoice...")
        self.update_idletasks()
        vendor_name = self.invoice_vendor_name_var.get()
        vendor_id = self.invoice_vendor_id_var.get()
        vendor_address = self.invoice_vendor_address_var.get()
        cursor.execute("SELECT po_number FROM vendors WHERE vendor_id=?", (vendor_id,))
        row = cursor.fetchone()
        vendor_po = row[0] if row else ""
        invoice_type = self.invoice_type_entry.get().strip()
        invoice_no = self.invoice_no_entry.get().strip()
        invoice_date = self.invoice_date_entry.get()
        if not vendor_name or not invoice_no or not self.excel_file_var.get():
            messagebox.showerror("Error", "Vendor, Invoice No, and Excel file are required.")
            self.progress_label.config(text="")
            return
        try:
            df_processed, total_amount = process_excel_file(self.excel_file_var.get())
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.progress_label.config(text="")
            return
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not output_path:
            self.progress_label.config(text="")
            return
        input_details = {
            "vendor_name": vendor_name,
            "vendor_address": vendor_address,
            "vendor_po": vendor_po,
            "invoice_type": invoice_type,
            "invoice_no": invoice_no,
            "invoice_date": invoice_date
        }
        try:
            create_invoice_pdf(output_path, input_details,df_processed, total_amount)
            cursor.execute("INSERT INTO invoices (vendor_id, invoice_no, invoice_date, invoice_type, po_mr_no, excel_file) VALUES (?,?,?,?,?,?)",
                           (vendor_id, invoice_no, invoice_date, invoice_type, vendor_po, self.excel_file_var.get()))
            conn.commit()
            self.progress_label.config(text="Invoice Generated Successfully!")
            messagebox.showinfo("Success", "Invoice PDF generated and saved.")
        except Exception as e:
            self.progress_label.config(text="")
            messagebox.showerror("Error", f"Failed to generate invoice: {e}")

    # -----------------------------
    # 3) Invoice Reports Frame
    # -----------------------------
    # 3) Invoice Reports Frame
    # -----------------------------
    def build_report_frame(self):
        tk.Label(self.report_frame, text="Invoice Reports", font=("Arial", 18, "bold")).pack(pady=10)
        
        # Filter area with calendar-type input for date
        filter_frame = tk.Frame(self.report_frame)
        filter_frame.pack(pady=10)
        tk.Label(filter_frame, text="Invoice No:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.report_invoice_no_var = tk.StringVar()
        tk.Entry(filter_frame, textvariable=self.report_invoice_no_var, width=20).grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(filter_frame, text="Vendor Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.report_vendor_name_var = tk.StringVar()
        tk.Entry(filter_frame, textvariable=self.report_vendor_name_var, width=20).grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(filter_frame, text="Date (YYYY-MM-DD):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        # Using DateEntry from tkcalendar for calendar-type input:
        self.report_date_entry = DateEntry(filter_frame, width=20, date_pattern="yyyy-mm-dd")
        self.report_date_entry.grid(row=2, column=1, padx=5, pady=5)
        
        tk.Button(filter_frame, text="Search", command=self.search_invoices).grid(row=3, column=1, padx=5, pady=5, sticky="e")
        tk.Button(filter_frame, text="Select Transactions", command=self.select_transactions).grid(row=4, column=1, padx=5, pady=15)
        tk.Button(filter_frame, text="Generate Invoice Report PDF", command=self.generate_invoice_report_pdf).grid(row=7, column=1, padx=5, pady=15)
        
        # Treeview for search results (populated by search_invoices)
        tk.Label(self.report_frame, text="Search Results:").pack()
        self.report_tree = ttk.Treeview(self.report_frame, columns=("id", "vendor_id", "invoice_no", "invoice_date", "invoice_type", "po_mr_no", "excel_file"), show="headings")
        self.report_tree.pack(fill="both", expand=True)
        for col in ("id", "vendor_id", "invoice_no", "invoice_date", "invoice_type", "po_mr_no", "excel_file"):
            self.report_tree.heading(col, text=col.capitalize())
            self.report_tree.column(col, width=100)
        
        # Always-visible treeview for selected transactions
        tk.Label(self.report_frame, text="Selected Transactions for Report:").pack(pady=5)
        self.selected_report_tree = ttk.Treeview(self.report_frame, columns=("date", "invoice_no", "Name", "debit", "credit"), show="headings")
        self.selected_report_tree.pack(fill="both", expand=True, padx=10, pady=10)
        for col in ("date", "invoice_no", "Name", "debit", "credit"):
            self.selected_report_tree.heading(col, text=col.capitalize())
            self.selected_report_tree.column(col, width=100)

    def search_invoices(self):
        # Clear previous search results
        for row in self.report_tree.get_children():
            self.report_tree.delete(row)
        invoice_no = self.report_invoice_no_var.get().strip()
        vendor_name = self.report_vendor_name_var.get().strip()
        date_filter = self.report_date_entry.get().strip()
        query = "SELECT i.id, i.vendor_id, i.invoice_no, i.invoice_date, i.invoice_type, i.po_mr_no, i.excel_file FROM invoices i"
        params = []
        conditions = []
        if invoice_no:
            conditions.append("i.invoice_no LIKE ?")
            params.append(f"%{invoice_no}%")
        if vendor_name:
            query += " JOIN vendors v ON i.vendor_id = v.vendor_id"
            conditions.append("v.vendor_name LIKE ?")
            params.append(f"%{vendor_name}%")
        if date_filter:
            conditions.append("i.invoice_date = ?")
            params.append(date_filter)
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        cursor.execute(query, params)
        rows = cursor.fetchall()
        for r in rows:
            self.report_tree.insert("", tk.END, values=r)

    def select_transactions(self):
        """
        Opens a pop-up window showing the search results (from self.report_tree).
        For each row, it processes the associated Excel file to obtain the transaction total.
        It then looks up the vendor name (via vendor_id) to fill in the Name column.
        The user selects the rows they want to include, and these rows are added to the 
        main Selected Transactions tree (self.selected_report_tree).
        """
        search_rows = self.report_tree.get_children()
        if not search_rows:
            messagebox.showerror("Error", "No search results available.")
            return

        # Build a temporary list of processed rows from search results.
        processed_rows = []
        for item in search_rows:
            vals = self.report_tree.item(item, "values")
            # Expected format: (id, vendor_id, invoice_no, invoice_date, invoice_type, po_mr_no, excel_file)
            inv_id, vendor_id, inv_no, inv_date, inv_type, _, excel_file = vals
            try:
                # Process the Excel file to get the total amount.
                _, total_amt = process_excel_file(excel_file)
            except Exception as e:
                continue  # Skip rows that cannot be processed
            
            # Lookup the vendor name from the vendors table
            cursor.execute("SELECT vendor_name FROM vendors WHERE vendor_id=?", (vendor_id,))
            vendor_row = cursor.fetchone()
            vendor_name = vendor_row[0] if vendor_row else ""
            
            # Based on invoice type, assign amount to Debit or Credit.
            if inv_type.lower() == "credit":
                debit_val = total_amt
                credit_val = 0.0
            else:
                debit_val = 0.0
                credit_val = total_amt
            
            # Now, use vendor_name as the Name column.
            processed_rows.append((inv_date, inv_no, vendor_name, debit_val, credit_val))
        
        # Create a pop-up for selection.
        popup = tk.Toplevel(self)
        popup.title("Select Transactions")
        popup.geometry("600x300")
        sel_tree = ttk.Treeview(popup, columns=("date", "invoice_no", "Name", "debit", "credit"), show="headings", selectmode="extended")
        sel_tree.pack(fill="both", expand=True, padx=10, pady=10)
        for col in ("date", "invoice_no", "Name", "debit", "credit"):
            sel_tree.heading(col, text=col.capitalize())
            sel_tree.column(col, width=100)
        for row in processed_rows:
            sel_tree.insert("", tk.END, values=row)
        
        def add_selection():
            selected = sel_tree.selection()
            if not selected:
                messagebox.showerror("Error", "No rows selected.")
                return
            for item in selected:
                row_data = sel_tree.item(item, "values")
                self.selected_report_tree.insert("", tk.END, values=row_data)
            popup.destroy()
        
        tk.Button(popup, text="Add Selected Rows", command=add_selection).pack(pady=10)


    def generate_invoice_report_pdf(self):
        """
        Generates the Invoice Report PDF based on the transactions in the Selected Transactions table.
        (This function calls your helper create_report_table_pdf which you should have defined.)
        """
        rows = self.selected_report_tree.get_children()
        if not rows:
            messagebox.showerror("Error", "No transactions selected for report.")
            return
        report_data = []
        for item in rows:
            vals = self.selected_report_tree.item(item, "values")
            # Expected format: (date, invoice_no, Name, debit, credit)
            report_data.append(vals)
        aging = compute_aging(report_data)
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not output_path:
            return
        try:
            create_report_table_pdf(output_path, "Invoice Report", report_data, balance_bf=0.0, aging_summary=aging)
            messagebox.showinfo("Success", "Invoice Report PDF generated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate PDF: {e}")

    # -----------------------------
    # 4) SOA Reports Frame
    # -----------------------------
    def build_soa_frame(self):
        tk.Label(self.soa_frame, text="SOA Reports", font=("Arial", 18, "bold")).pack(pady=10)
        cursor.execute("SELECT vendor_name, vendor_id, vendor_address FROM vendors")
        self.soa_vendors = cursor.fetchall()
        vendor_names = [v[0] for v in self.soa_vendors]
        form_frame = tk.Frame(self.soa_frame)
        form_frame.pack(pady=10)
        tk.Label(form_frame, text="Select Vendor:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.soa_vendor_var = tk.StringVar()
        self.soa_vendor_dropdown = ctk.CTkOptionMenu(form_frame, values=vendor_names, variable=self.soa_vendor_var)
        self.soa_vendor_dropdown.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.filter_method = tk.StringVar(value="date")
        tk.Label(form_frame, text="Filter By:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        tk.Radiobutton(form_frame, text="Date Range", variable=self.filter_method, value="date").grid(row=1, column=1, sticky="w")
        tk.Radiobutton(form_frame, text="Invoice #", variable=self.filter_method, value="invoice").grid(row=1, column=2, sticky="w")
        tk.Radiobutton(form_frame, text="Count", variable=self.filter_method, value="count").grid(row=1, column=3, sticky="w")
        tk.Label(form_frame, text="From Date:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.soa_from_date_entry = DateEntry(form_frame, width=20, date_pattern="yyyy-mm-dd")
        self.soa_from_date_entry.grid(row=2, column=1, padx=5, pady=5)
        tk.Label(form_frame, text="To Date:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.soa_to_date_entry = DateEntry(form_frame, width=20, date_pattern="yyyy-mm-dd")
        self.soa_to_date_entry.grid(row=3, column=1, padx=5, pady=5)
        tk.Label(form_frame, text="Invoice # (comma separated):").grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.soa_invoice_nums = tk.Entry(form_frame, width=30)
        self.soa_invoice_nums.grid(row=4, column=1, padx=5, pady=5)
        tk.Label(form_frame, text="Invoice Count:").grid(row=5, column=0, padx=5, pady=5, sticky="w")
        self.soa_invoice_count = tk.Entry(form_frame, width=10)
        self.soa_invoice_count.grid(row=5, column=1, padx=5, pady=5, sticky="w")
        self.soa_include_seal_var = tk.BooleanVar(value=True)
        tk.Checkbutton(form_frame, text="Include Seal", variable=self.soa_include_seal_var).grid(row=6, column=1, padx=5, pady=5, sticky="w")
        tk.Button(form_frame, text="Generate SOA", command=self.generate_soa).grid(row=7, column=1, padx=5, pady=15)

    def generate_soa(self):
        selected_vendor = self.soa_vendor_var.get()
        if not selected_vendor:
            messagebox.showerror("Error", "Please select a vendor.")
            return
        vendor_id = None
        vendor_address = ""
        for v in self.soa_vendors:
            if v[0] == selected_vendor:
                vendor_id = v[1]
                vendor_address = v[2]
                break
        if not vendor_id:
            messagebox.showerror("Error", "Vendor not found in DB.")
            return
        soa_info = {
            "statement_date": self.soa_from_date_entry.get(),
            "due_date": self.soa_to_date_entry.get(),
            "company_name": selected_vendor,
            "company_address": vendor_address
        }
        filter_type = self.filter_method.get()
        if filter_type == "date":
            from_date = self.soa_from_date_entry.get()
            to_date = self.soa_to_date_entry.get()
            try:
                datetime.strptime(from_date, "%Y-%m-%d")
                datetime.strptime(to_date, "%Y-%m-%d")
            except:
                messagebox.showerror("Error", "Invalid date format.")
                return
            cursor.execute(
                "SELECT invoice_no, invoice_date, invoice_type, excel_file FROM invoices WHERE vendor_id=? AND invoice_date BETWEEN ? AND ?",
                (vendor_id, from_date, to_date)
            )
        elif filter_type == "invoice":
            invoice_nums = self.soa_invoice_nums.get().strip()
            if not invoice_nums:
                messagebox.showerror("Error", "Please enter invoice number(s).")
                return
            invoice_list = tuple(item.strip() for item in invoice_nums.split(",") if item.strip())
            placeholders = ",".join("?" * len(invoice_list))
            query = f"SELECT invoice_no, invoice_date, invoice_type, excel_file FROM invoices WHERE vendor_id=? AND invoice_no IN ({placeholders})"
            cursor.execute(query, (vendor_id, *invoice_list))
        elif filter_type == "count":
            count_str = self.soa_invoice_count.get().strip()
            if not count_str.isdigit():
                messagebox.showerror("Error", "Please enter a valid invoice count.")
                return
            count = int(count_str)
            cursor.execute(
                "SELECT invoice_no, invoice_date, invoice_type, excel_file FROM invoices WHERE vendor_id=? ORDER BY invoice_date DESC LIMIT ?",
                (vendor_id, count)
            )
        else:
            messagebox.showerror("Error", "Invalid filter method.")
            return
        invoices = cursor.fetchall()
        if not invoices:
            messagebox.showinfo("Info", "No invoices found for the selected criteria.")
            return
        # For each invoice, process its Excel file to get the total amount row.
        soa_rows = []
        for inv in invoices:
            inv_no, inv_date, inv_type, excel_path = inv
            try:
                df_proc, total_amt = process_excel_file(excel_path)
            except:
                continue
            if inv_type.lower() == "credit":
                debit_val = total_amt
                credit_val = 0.0
            else:
                debit_val = 0.0
                credit_val = total_amt
            soa_rows.append((inv_date, inv_no, "", debit_val, credit_val))
        # Generate SOA PDF with new table format.
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not output_path:
            return
        try:
            create_soa_pdf_modified(output_path, soa_info, soa_rows)
            messagebox.showinfo("Success", "SOA PDF generated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate SOA: {e}")

# ----------------------------------------------------
# Run the App
# ----------------------------------------------------
if __name__ == "__main__":
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("dark-blue")
    app = MainApp()
    app.mainloop()
