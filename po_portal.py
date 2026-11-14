import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_purchase_order_and_log(po_data, items):
    # Styling Palette (Clean Slate & Teal)
    DARK_SLATE = "1E3A4C"
    MEDIUM_TEAL = "2C5E7A"
    SUBHEADER_BG = "D9E4EC"
    ZEBRA_BG = "F6F9FB"
    TOTAL_BG = "E6EEF4"
    BORDER_COLOR = "CBD5E1"
    WHITE = "FFFFFF"

    font_title = Font(name="Calibri", size=14, bold=True, color=WHITE)
    font_sec = Font(name="Calibri", size=10, bold=True, color=DARK_SLATE)
    font_hdr = Font(name="Calibri", size=10, bold=True, color=WHITE)
    font_lbl = Font(name="Calibri", size=10, bold=True, color="1E293B")
    font_data = Font(name="Calibri", size=10, color="1E293B")
    font_total = Font(name="Calibri", size=10, bold=True, color=DARK_SLATE)

    fill_banner = PatternFill("solid", fgColor=DARK_SLATE)
    fill_hdr = PatternFill("solid", fgColor=MEDIUM_TEAL)
    fill_sub = PatternFill("solid", fgColor=SUBHEADER_BG)
    fill_zebra = PatternFill("solid", fgColor=ZEBRA_BG)
    fill_total = PatternFill("solid", fgColor=TOTAL_BG)

    thin = Side(border_style="thin", color=BORDER_COLOR)
    border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ---------------------------------------------------------
    # STEP 1: GENERATE INDIVIDUAL PURCHASE ORDER EXCEL FILE
    # ---------------------------------------------------------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Purchase Order"

    # Banner
    ws.merge_cells("A1:G1")
    ws["A1"] = "PURCHASE ORDER"
    ws["A1"].font = font_title
    ws["A1"].fill = fill_banner
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Vendor & PO Metadata Header
    meta_rows = [
        ("Vendor Name:", po_data["vendor_name"], "PO Number:", po_data["po_number"]),
        ("Vendor GSTIN:", po_data["vendor_gstin"], "PO Date:", po_data["po_date"]),
        ("Vendor Contact:", po_data["vendor_contact"], "Delivery Date:", po_data["delivery_date"]),
        ("Ship To:", "Production Plant / Store", "PO Type:", po_data["po_type"]),
        ("Payment Terms:", po_data["payment_terms"], "Currency:", po_data["currency"]),
    ]

    for idx, (l1, v1, l2, v2) in enumerate(meta_rows, start=3):
        ws.cell(row=idx, column=1, value=l1).font = font_lbl
        ws.merge_cells(start_row=idx, start_column=2, end_row=idx, end_column=3)
        ws.cell(row=idx, column=2, value=v1).font = font_data

        ws.cell(row=idx, column=5, value=l2).font = font_lbl
        ws.merge_cells(start_row=idx, start_column=6, end_row=idx, end_column=7)
        ws.cell(row=idx, column=6, value=v2).font = font_data

        for c in (1, 2, 3, 5, 6, 7):
            ws.cell(row=idx, column=c).border = border_cell
            if c in (1, 5):
                ws.cell(row=idx, column=c).fill = fill_sub
        ws.row_dimensions[idx].height = 20

    # Table Headers (Row 9)
    headers = ["S.No", "Description / Item", "HSN/SAC", "Qty", "Unit Price", "GST %", "Amount"]
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=9, column=c_idx, value=h)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_cell
    ws.row_dimensions[9].height = 24

    # Populate Line Items
    start_row = 10
    subtotal_val = 0
    total_gst_val = 0

    for i, item in enumerate(items, start=1):
        r = start_row + i - 1
        desc, hsn, qty, price, gst_rate = item
        base_amt = qty * price
        gst_amt = base_amt * gst_rate
        subtotal_val += base_amt
        total_gst_val += gst_amt

        ws.cell(row=r, column=1, value=i).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=desc).alignment = Alignment(horizontal="left")
        ws.cell(row=r, column=3, value=hsn).alignment = Alignment(horizontal="center")

        c_qty = ws.cell(row=r, column=4, value=qty)
        c_qty.number_format = "#,##0"
        c_qty.alignment = Alignment(horizontal="right")

        c_price = ws.cell(row=r, column=5, value=price)
        c_price.number_format = "₹#,##0.00"
        c_price.alignment = Alignment(horizontal="right")

        c_gst = ws.cell(row=r, column=6, value=gst_rate)
        c_gst.number_format = "0%"
        c_gst.alignment = Alignment(horizontal="right")

        c_amt = ws.cell(row=r, column=7, value=f"=D{r}*E{r}")
        c_amt.number_format = "₹#,##0.00"
        c_amt.alignment = Alignment(horizontal="right")

        for c in range(1, 8):
            cell = ws.cell(row=r, column=c)
            cell.font = font_data
            cell.border = border_cell
            if i % 2 == 0:
                cell.fill = fill_zebra
        ws.row_dimensions[r].height = 20

    end_row = start_row + len(items) - 1
    sub_row = end_row + 2
    gst_row = sub_row + 1
    grand_row = gst_row + 1
    words_row = grand_row + 1

    # Totals Block
    totals = [
        (sub_row, "Subtotal", f"=SUM(G{start_row}:G{end_row})"),
        (gst_row, "CGST + SGST", f"=SUMPRODUCT(G{start_row}:G{end_row}, F{start_row}:F{end_row})"),
        (grand_row, "Grand Total", f"=G{sub_row}+G{gst_row}"),
    ]
    for r, lbl, formula in totals:
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
        ws.cell(row=r, column=5, value=lbl).font = font_total
        ws.cell(row=r, column=5).alignment = Alignment(horizontal="right")
        val_cell = ws.cell(row=r, column=7, value=formula)
        val_cell.font = font_total
        val_cell.number_format = "₹#,##0.00"
        val_cell.alignment = Alignment(horizontal="right")
        for c in range(5, 8):
            ws.cell(row=r, column=c).fill = fill_total
            ws.cell(row=r, column=c).border = border_cell
        ws.row_dimensions[r].height = 22

    # Amount in Words
    ws.cell(row=words_row, column=1, value="Amount in Words:").font = font_lbl
    ws.merge_cells(start_row=words_row, start_column=2, end_row=words_row, end_column=7)
    ws.cell(row=words_row, column=2, value=po_data["amount_in_words"]).font = font_data
    for c in range(1, 8):
        ws.cell(row=words_row, column=c).border = border_cell
        ws.cell(row=words_row, column=c).fill = fill_sub
    ws.row_dimensions[words_row].height = 22

    # Signatory Section (Prepared By, Checked By, Authorised Signatory)
    sig_start = words_row + 2
    sigs = [
        ("A", "B", "Prepared By", po_data["prepared_by"], po_data["po_date"]),
        ("C", "D", "Checked By", po_data["checked_by"], ""),
        ("E", "G", "Authorised Signatory", po_data["authorised_by"], ""),
    ]
    ws.row_dimensions[sig_start].height = 22
    ws.row_dimensions[sig_start + 1].height = 32
    ws.row_dimensions[sig_start + 2].height = 20
    ws.row_dimensions[sig_start + 3].height = 20

    for c1, c2, role_title, person_name, dt in sigs:
        ws.merge_cells(f"{c1}{sig_start}:{c2}{sig_start}")
        ws[f"{c1}{sig_start}"] = role_title
        ws[f"{c1}{sig_start}"].font = font_hdr
        ws[f"{c1}{sig_start}"].fill = fill_hdr
        ws[f"{c1}{sig_start}"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells(f"{c1}{sig_start+1}:{c2}{sig_start+1}")
        ws[f"{c1}{sig_start+1}"] = "(Signature)"
        ws[f"{c1}{sig_start+1}"].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells(f"{c1}{sig_start+2}:{c2}{sig_start+2}")
        ws[f"{c1}{sig_start+2}"] = f"Name: {person_name}"
        ws[f"{c1}{sig_start+2}"].font = font_data

        ws.merge_cells(f"{c1}{sig_start+3}:{c2}{sig_start+3}")
        ws[f"{c1}{sig_start+3}"] = f"Date: {dt if dt else '____________________'}"
        ws[f"{c1}{sig_start+3}"].font = font_data

    for r in range(sig_start, sig_start + 4):
        for c in range(1, 8):
            ws.cell(row=r, column=c).border = border_cell

    widths = [18, 36, 14, 10, 16, 10, 18]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    po_filename = f"{po_data['po_number']}.xlsx"
    wb.save(po_filename)
    print(f"[1/2] Generated Purchase Order: {po_filename}")

    # ---------------------------------------------------------
    # STEP 2: AUTO-LOG INTO MASTER VENDOR & PO TRACKING PORTAL
    # ---------------------------------------------------------
    master_file = "Vendor_PO_Master_Log.xlsx"
    grand_total_val = subtotal_val + total_gst_val

    if not os.path.exists(master_file):
        m_wb = openpyxl.Workbook()
        m_ws = m_wb.active
        m_ws.title = "PO_Tracking_Portal"
        log_headers = [
            "PO Number", "PO Date", "Delivery Date", "Vendor Name",
            "PO Type", "Payment Terms", "Subtotal (₹)", "GST (₹)",
            "Grand Total (₹)", "Delivery Status", "Payment Status"
        ]
        for idx, h in enumerate(log_headers, start=1):
            c = m_ws.cell(row=1, column=idx, value=h)
            c.font = font_hdr
            c.fill = fill_hdr
            c.alignment = Alignment(horizontal="center")
            m_ws.column_dimensions[get_column_letter(idx)].width = 18
        m_wb.save(master_file)

    m_wb = openpyxl.load_workbook(master_file)
    m_ws = m_wb["PO_Tracking_Portal"]
    next_row = m_ws.max_row + 1

    log_entry = [
        po_data["po_number"], po_data["po_date"], po_data["delivery_date"],
        po_data["vendor_name"], po_data["po_type"], po_data["payment_terms"],
        subtotal_val, total_gst_val, grand_total_val, "PO Issued", "Pending"
    ]
    for col_idx, val in enumerate(log_entry, start=1):
        cell = m_ws.cell(row=next_row, column=col_idx, value=val)
        if col_idx in (7, 8, 9):
            cell.number_format = "₹#,##0.00"

    m_wb.save(master_file)
    print(f"[2/2] Updated Master Tracking Portal: {master_file}")


# =========================================================
# RUN THE AUTOMATION (Edit details below for any new PO)
# =========================================================
if __name__ == "__main__":
    current_po = {
        "po_number": "PO-2026-001",
        "po_date": "14-Nov-2026",
        "delivery_date": "24-Nov-2026",
        "po_type": "Standard Goods PO",
        "payment_terms": "50% Advance, 50% on Delivery",
        "currency": "INR (₹)",
        "vendor_name": "Mean Well India",
        "vendor_gstin": "33AABCM1234F1Z5",
        "vendor_contact": "+91 98400 00000",
        "amount_in_words": "INR Ninety-Five Thousand Two Hundred Eighty-Five Only",
        "prepared_by": "Thanigai Malai",
        "checked_by": "Bharathi",
        "authorised_by": "Rudhran",
    }

    # Format: (Description, HSN/SAC, Qty, Unit Price, GST Rate)
    order_items = [
        ("Industrial DIN-Rail SMPS (24V 10A)", "85044090", 5, 3150.00, 0.18),
        ("Custom STM32 Interface PCB", "85340000", 5, 2200.00, 0.18),
        ("24V Geared BLDC Motor w/ Encoder", "85013119", 10, 5400.00, 0.18),
    ]

    create_purchase_order_and_log(current_po, order_items)