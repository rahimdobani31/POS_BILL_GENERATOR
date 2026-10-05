import os
import sys
import ctypes

# Automatically hide the Windows Command Prompt / Terminal window on app start
if sys.platform == "win32":
    try:
        console_window = ctypes.windll.kernel32.GetConsoleWindow()
        if console_window != 0:
            ctypes.windll.user32.ShowWindow(console_window, 0)  # SW_HIDE
    except Exception:
        pass

import re
import time
import io
import tempfile
import pdfplumber
import pymupdf
from PIL import Image as PILImage

from reportlab.lib.pagesizes import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QFileDialog, QTextEdit,
    QGroupBox, QSplitter, QComboBox, QMessageBox, QScrollArea, QFrame,
    QStackedWidget, QListWidget, QListWidgetItem, QGraphicsDropShadowEffect
)
from PyQt6.QtGui import QFont, QImage, QPixmap, QColor

# Windows Native Printing Imports
try:
    import win32print
    import win32ui
    import win32con
    import win32gui
    import PIL.ImageWin
    WIN32_PRINT_AVAILABLE = True
except ImportError:
    WIN32_PRINT_AVAILABLE = False


# ==============================================================================
# PDF GENERATOR ENGINE
# ==============================================================================

FONT_NAME = "Helvetica"
FONT_BOLD = "Helvetica-Bold"
FONT_SHOP_HEADER = "Helvetica-Bold"

LOGO_IMAGE_PATH = "logo.png"
WHATSAPP_QR_PATH = "whatapp qr.png"

windows_font_path = r"C:\Windows\Fonts\BRITANIC.TTF"
if os.path.exists(windows_font_path):
    try:
        pdfmetrics.registerFont(TTFont('BritannicBold', windows_font_path))
        FONT_SHOP_HEADER = 'BritannicBold'
    except Exception:
        pass


def format_smart_amount(val_str, append_suffix=False):
    """Formats raw numerical string safely, preserving negative amounts and removing trailing decimals if whole."""
    if not val_str:
        return "0/=" if append_suffix else "0"
    
    is_negative = '-' in val_str
    
    cleaned = re.sub(r'[^\d\.]', '', val_str)
    if not cleaned:
        return "0/=" if append_suffix else "0"
    
    try:
        val_float = float(cleaned)
        if is_negative:
            val_float = -val_float
            
        if val_float.is_integer():
            fmt = f"{int(val_float):,}"
        else:
            fmt = f"{val_float:,.2f}"
            
        if append_suffix:
            return f"{fmt}/="
        return fmt
    except ValueError:
        if append_suffix and val_str.strip() and not val_str.endswith('/='):
            return f"{val_str.strip()}/="
        return val_str


def parse_noorani_invoice(pdf_path):
    """Parses Noorani Sanitary invoices dynamically supporting optional Discount column detection."""
    invoice_data = {
        "doc_type": "ESTIMATE", 
        "shop_name": "NOORANI SANITARY",
        "shop_address": "Shop#9 WaterPump Block 16/16 F.b Area Karachi<br/>Near UBL Bank<br/>Mob: 0336-2580863",
        "estimate_no": "N/A",
        "date": "N/A",
        "customer": "walk in customer",
        "contact": "",
        "address": "",
        "shipping_address": "",
        "has_discount": False,
        "items": [],
        "subtotal": "0",
        "cartage": "0",
        "total": "0",
        "paid": "0",
        "balance": "0",
        "total_qty": "0",
        "remarks": ""
    }
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            full_text = page.extract_text()
            if not full_text:
                continue
                
            text_lines = full_text.split('\n')
            full_text_upper = full_text.upper()
            
            if "CHECKING LIST" in full_text_upper:
                invoice_data["doc_type"] = "CHECKING LIST"
            elif "SALES QUOTE" in full_text_upper:
                invoice_data["doc_type"] = "SALES QUOTE"
            elif "CREDIT NOTE" in full_text_upper:
                invoice_data["doc_type"] = "CREDIT NOTE"
            elif "ESTIMATE" in full_text_upper:
                invoice_data["doc_type"] = "ESTIMATE"
            
            remarks_lines = []
            capture_remarks = False
            
            for line in text_lines:
                cleaned_line = line.strip()
                if not cleaned_line:
                    continue
                
                if "ESTIMATE #" in cleaned_line:
                    parts = cleaned_line.split("ESTIMATE #")
                    if len(parts) > 1:
                        invoice_data["estimate_no"] = parts[1].replace('"', '').replace(',', '').strip().split()[0]
                elif "Order #" in cleaned_line:
                    parts = cleaned_line.split("Order #")
                    if len(parts) > 1:
                        invoice_data["estimate_no"] = parts[1].replace('"', '').replace(',', '').strip().split()[0]
                elif "Quote #" in cleaned_line:
                    parts = cleaned_line.split("Quote #")
                    if len(parts) > 1:
                        invoice_data["estimate_no"] = parts[1].replace('"', '').replace(',', '').strip().split()[0]
                
                if "Date" in cleaned_line:
                    date_match = re.search(r'\d{1,2}-[A-Za-z]{3}-\d{2,4}', cleaned_line)
                    if date_match:
                        invoice_data["date"] = date_match.group(0)

                if "Sub-Total" in cleaned_line:
                    invoice_data["subtotal"] = cleaned_line.split("Sub-Total")[-1]
                elif "Cartage" in cleaned_line:
                    raw_cartage = cleaned_line.split("Cartage")[-1]
                    clean_cartage_val = re.sub(r'^[^\d\-]+', '', raw_cartage)
                    invoice_data["cartage"] = clean_cartage_val
                elif "Total" in cleaned_line and "Sub" not in cleaned_line and "Qty" not in cleaned_line:
                    invoice_data["total"] = cleaned_line.split("Total")[-1]
                elif "Credit" in cleaned_line and invoice_data["doc_type"] == "CREDIT NOTE":
                    invoice_data["total"] = cleaned_line.split("Credit")[-1]
                elif "Paid" in cleaned_line:
                    invoice_data["paid"] = cleaned_line.split("Paid")[-1]
                elif "Balance" in cleaned_line:
                    invoice_data["balance"] = cleaned_line.split("Balance")[-1]
                elif "Total Qty" in cleaned_line:
                    invoice_data["total_qty"] = re.sub(r'[^\d\.]', '', cleaned_line.split("Total Qty")[-1]).strip()

                if "Remarks" in cleaned_line:
                    capture_remarks = True
                    cleaned = re.sub(r'^Remarks\s*[.\s:]*', '', cleaned_line, flags=re.IGNORECASE)
                    cleaned = re.sub(r'Sub-Total.*$', '', cleaned, flags=re.IGNORECASE).strip()
                    if cleaned:
                        remarks_lines.append(cleaned)
                    continue
                
                if capture_remarks:
                    if "15% Deduction" in cleaned_line or "Otherwise No Return" in cleaned_line:
                        capture_remarks = False
                    else:
                        cleaned = re.sub(r'(Sub-Total|Total|Paid|Balance|Fee|Refunded|Credit).*$', '', cleaned_line, flags=re.IGNORECASE).strip()
                        cleaned = re.sub(r'\b(Rs)?[\d,]+\.\d{2}\s*[\-]?$', '', cleaned).strip()
                        cleaned = re.sub(r'\b[\d,]+\s*[\-]?$', '', cleaned).strip()
                        if cleaned:
                            remarks_lines.append(cleaned)

            if remarks_lines:
                invoice_data["remarks"] = "<br/>".join(remarks_lines)

            billing_match = re.search(r'Billing\s+(.*?)(?=Address|Contact|Shipping|$)', full_text, re.DOTALL | re.IGNORECASE)
            if billing_match:
                cust_candidate = re.sub(r'^[.\s:]+', '', billing_match.group(1).strip()).strip()
                if cust_candidate and cust_candidate != "." and "walk in" not in cust_candidate.lower():
                    if not any(hdr in cust_candidate for hdr in ["Description", "Quantity", "Unit Price", "Sub-Total"]):
                        invoice_data["customer"] = cust_candidate

            address_match = re.search(r'Address\s+(.*?)(?=Contact|Billing|Shipping|\n\n|$)', full_text, re.DOTALL | re.IGNORECASE)
            if address_match:
                addr_candidate = re.sub(r'^[.\s:]+', '', address_match.group(1).split('\n')[0].strip()).strip()
                if addr_candidate and addr_candidate != "." and len(addr_candidate) > 1:
                    if not any(hdr in addr_candidate for hdr in ["Description", "Quantity", "Unit Price", "Sub-Total", "Address"]):
                        invoice_data["address"] = addr_candidate

            contact_match = re.search(r'Contact\s+(.*?)(?=Billing|Address|Shipping|\n\n|$)', full_text, re.DOTALL | re.IGNORECASE)
            if contact_match:
                con_candidate = contact_match.group(1).strip()
                phone_find = re.search(r'(\d{4}-\d{7}|\d{11})', con_candidate)
                if phone_find:
                    invoice_data["contact"] = phone_find.group(1)
                else:
                    con_candidate = re.sub(r'^[.\s:]+', '', con_candidate).strip().split('\n')[0]
                    if con_candidate and con_candidate != "." and len(con_candidate) > 5:
                        if not any(hdr in con_candidate for hdr in ["Description", "Quantity", "Unit Price", "Sub-Total", "Address"]):
                            invoice_data["contact"] = con_candidate

            shipping_match = re.search(r'Shipping\s+(?:Address\s+)?(.*?)(?=Description|Quantity|Unit|Item|$)', full_text, re.DOTALL | re.IGNORECASE)
            if shipping_match:
                ship_candidate = re.sub(r'^[.\s:]+', '', shipping_match.group(1).split('\n')[0].strip()).strip()
                if invoice_data["doc_type"] == "CHECKING LIST" and ship_candidate:
                    invoice_data["shipping_address"] = ship_candidate
                    if invoice_data["customer"] == "walk in customer" and len(ship_candidate) > 2:
                        invoice_data["customer"] = ship_candidate
                elif ship_candidate and ship_candidate != ".":
                    if not any(hdr in ship_candidate for hdr in ["Description", "Quantity", "Unit Price", "Sub-Total"]):
                        invoice_data["shipping_address"] = ship_candidate

            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or len(row) < 2:
                        continue
                        
                    col0 = str(row[0]).strip() if row[0] else ""
                    
                    if invoice_data["doc_type"] == "CHECKING LIST":
                        if len(row) >= 3:
                            desc = str(row[1]).strip() if row[1] else ""
                            qty = str(row[2]).strip() if row[2] else ""
                            
                            if "Description" in desc or not desc or "Total Qty" in col0 or "Total" in desc:
                                continue
                                
                            qty_raw = qty.replace('\n', ' ').replace(',', '').strip()
                            qty_match = re.match(r'^(\d+(?:\.\d+)?)\s*(.*)$', qty_raw)
                            
                            qty_num, qty_unit = (qty_match.group(1), qty_match.group(2).strip()) if qty_match else ("1", "Pcs")
                            if not qty_unit:
                                qty_unit = "Bottle" if "BOTTLE" in desc.upper() else "SET" if "SET" in desc.upper() else "PAIR" if "PAIR" in desc.upper() else "Pcs"
                                
                            if "BOTTLE" in qty_unit.upper(): qty_unit = "Bottle"
                            elif "PCS" in qty_unit.upper(): qty_unit = "Pcs"
                            elif "FEET" in qty_unit.upper(): qty_unit = "FEET"
                                
                            if '.' in qty_num:
                                try:
                                    f_qty = float(qty_num)
                                    if f_qty.is_integer(): qty_num = str(int(f_qty))
                                except ValueError: pass
                                
                            invoice_data["items"].append([format_smart_amount(qty_num), qty_unit, desc, "0", "", "0"])
                    else:
                        if len(row) >= 5:
                            invoice_data["has_discount"] = True
                            desc = col0
                            qty = str(row[1]).strip() if row[1] else ""
                            unit_p = str(row[2]).strip() if row[2] else ""
                            disc_val = str(row[3]).strip() if row[3] else ""
                            sub_t = str(row[4]).strip() if row[4] else ""
                            
                            if "Description" in desc or not desc or "Sub-Total" in desc or "Cartage" in desc or "Total" in desc or "Paid" in desc or "Balance" in desc or "Credit" in desc:
                                continue
                            if desc.strip() == "CHANGE)" or "RETURN" in desc and (not qty and not unit_p):
                                continue
                                
                            if sub_t or unit_p:
                                desc = desc.replace("\n", " ")
                                desc = re.sub(r'\s+', ' ', desc).strip()
                                
                                qty_raw = qty.replace('\n', ' ').replace(',', '').strip()
                                qty_match = re.match(r'^(\d+(?:\.\d+)?)\s*(.*)$', qty_raw)
                                
                                qty_num, qty_unit = (qty_match.group(1), qty_match.group(2).strip()) if qty_match else ("1", "Pcs")
                                qty_unit = re.sub(r'^[B|O|T|L|E|P|C|S|\s\.\n\s]+$', '', qty_unit)
                                if not qty_unit:
                                    qty_unit = "Bottle" if "BOTTLE" in desc else "SET" if "SET" in desc else "PAIR" if "PAIR" in desc else "Pcs"
                                    
                                if "BOTTLE" in qty_unit.upper(): qty_unit = "Bottle"
                                elif "PCS" in qty_unit.upper(): qty_unit = "Pcs"
                                elif "FEET" in qty_unit.upper(): qty_unit = "FEET"
                                
                                if '.' in qty_num:
                                    try:
                                        f_qty = float(qty_num)
                                        if f_qty.is_integer(): qty_num = str(int(f_qty))
                                    except ValueError: pass
                                    
                                price_clean = format_smart_amount(unit_p.split('\n')[0], append_suffix=False)
                                disc_clean = disc_val.replace('\n', ' ').strip()
                                if disc_clean and not disc_clean.endswith('%') and disc_clean.isdigit():
                                    disc_clean += "%"
                                    
                                total_clean = format_smart_amount(sub_t.split('\n')[0], append_suffix=False)
                                
                                if qty_num == "" and price_clean == "0" and total_clean == "0":
                                    continue
                                    
                                invoice_data["items"].append([format_smart_amount(qty_num), qty_unit, desc, price_clean, disc_clean, total_clean])

                        elif len(row) == 4:
                            desc = col0
                            qty = str(row[1]).strip() if row[1] else ""
                            unit_p = str(row[2]).strip() if row[2] else ""
                            sub_t = str(row[3]).strip() if row[3] else ""
                            
                            if "Description" in desc or not desc or "Sub-Total" in desc or "Cartage" in desc or "Total" in desc or "Paid" in desc or "Balance" in desc or "Credit" in desc:
                                continue
                            if desc.strip() == "CHANGE)" or "RETURN" in desc and (not qty and not unit_p):
                                continue
                                
                            if sub_t or unit_p:
                                desc = desc.replace("\n", " ")
                                desc = re.sub(r'\s+', ' ', desc).strip()
                                
                                qty_raw = qty.replace('\n', ' ').replace(',', '').strip()
                                qty_match = re.match(r'^(\d+(?:\.\d+)?)\s*(.*)$', qty_raw)
                                
                                qty_num, qty_unit = (qty_match.group(1), qty_match.group(2).strip()) if qty_match else ("1", "Pcs")
                                qty_unit = re.sub(r'^[B|O|T|L|E|P|C|S|\s\.\n\s]+$', '', qty_unit)
                                if not qty_unit:
                                    qty_unit = "Bottle" if "BOTTLE" in desc else "SET" if "SET" in desc else "PAIR" if "PAIR" in desc else "Pcs"
                                    
                                if "BOTTLE" in qty_unit.upper(): qty_unit = "Bottle"
                                elif "PCS" in qty_unit.upper(): qty_unit = "Pcs"
                                elif "FEET" in qty_unit.upper(): qty_unit = "FEET"
                                
                                if '.' in qty_num:
                                    try:
                                        f_qty = float(qty_num)
                                        if f_qty.is_integer(): qty_num = str(int(f_qty))
                                    except ValueError: pass
                                    
                                price_clean = format_smart_amount(unit_p.split('\n')[0], append_suffix=False)
                                total_clean = format_smart_amount(sub_t.split('\n')[0], append_suffix=False)
                                
                                if qty_num == "" and price_clean == "0" and total_clean == "0":
                                    continue
                                    
                                invoice_data["items"].append([format_smart_amount(qty_num), qty_unit, desc, price_clean, "", total_clean])

    raw_paid = format_smart_amount(invoice_data["paid"], append_suffix=False)
    raw_balance = format_smart_amount(invoice_data["balance"], append_suffix=False)
    invoice_data["hide_paid_balance"] = (raw_paid in ["0", "0.00"]) and (raw_balance in ["0", "0.00"])

    invoice_data["subtotal"] = format_smart_amount(invoice_data["subtotal"], append_suffix=True)
    invoice_data["cartage"] = format_smart_amount(invoice_data["cartage"], append_suffix=True)
    invoice_data["total"] = format_smart_amount(invoice_data["total"], append_suffix=True)
    invoice_data["paid"] = format_smart_amount(invoice_data["paid"], append_suffix=True)
    invoice_data["balance"] = format_smart_amount(invoice_data["balance"], append_suffix=True)
    
    if '.' in invoice_data["total_qty"]:
        try:
            f_tot_qty = float(invoice_data["total_qty"])
            if f_tot_qty.is_integer(): invoice_data["total_qty"] = str(int(f_tot_qty))
            else: invoice_data["total_qty"] = f"{f_tot_qty:g}"
        except ValueError: pass

    return invoice_data


def generate_3inch_receipt(data, output_pdf_path):
    """Renders data into a clean 2.83-inch thermal layout supporting dynamic discount column."""
    page_width = 2.83 * inch
    
    total_item_lines = 0
    char_limit = 14 if data["has_discount"] else 18
    for item in data["items"]:
        desc_len = len(item[2])
        lines = max(1, (desc_len + (char_limit - 1)) // char_limit)
        total_item_lines += lines

    remarks_lines = data["remarks"].count("<br/>") + 1 if data["remarks"] else 0
    remarks_offset = 0.4 + (remarks_lines * 0.18) if remarks_lines > 0 else 0.2
    qr_offset = 1.35 if os.path.exists(WHATSAPP_QR_PATH) else 0.0
    
    estimated_height = 5.2 + (total_item_lines * 0.25) + remarks_offset + qr_offset
    page_height = max(8.0, estimated_height) * inch
    
    doc = SimpleDocTemplate(
        output_pdf_path, 
        pagesize=(page_width, page_height), 
        rightMargin=4, 
        leftMargin=4, 
        topMargin=8, 
        bottomMargin=8
    )
    styles = getSampleStyleSheet()
    
    style_shop = ParagraphStyle('ShopName', parent=styles['Normal'], fontName=FONT_SHOP_HEADER, fontSize=14, leading=16, alignment=TA_CENTER)
    style_shop_addr = ParagraphStyle('ShopAddr', parent=styles['Normal'], fontName=FONT_NAME, fontSize=8, leading=11, alignment=TA_CENTER)
    style_center = ParagraphStyle('CenterTxt', parent=styles['Normal'], fontName=FONT_NAME, fontSize=8, leading=11, alignment=TA_CENTER)
    style_left = ParagraphStyle('LeftTxt', parent=styles['Normal'], fontName=FONT_NAME, fontSize=8.5, leading=11, alignment=TA_LEFT)
    style_hdr = ParagraphStyle('HdrTxt', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=8, leading=10, alignment=TA_CENTER)
    style_hdr_left = ParagraphStyle('HdrTxtLf', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=8, leading=10, alignment=TA_LEFT)
    style_hdr_right = ParagraphStyle('HdrTxtRt', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=8, leading=10, alignment=TA_RIGHT)
    style_qty_num = ParagraphStyle('QtyNum', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=8.5, leading=10, alignment=TA_CENTER)
    style_item = ParagraphStyle('ItemTxt', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=7.5, leading=9.5, alignment=TA_LEFT)
    style_total = ParagraphStyle('TotalTxt', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=7.5, leading=9.5, alignment=TA_RIGHT)
    style_total_lbl = ParagraphStyle('TotalLbl', parent=styles['Normal'], fontName=FONT_NAME, fontSize=9, leading=12, alignment=TA_RIGHT)
    style_total_val = ParagraphStyle('TotalVal', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=9.5, leading=12, alignment=TA_RIGHT)
    style_remarks = ParagraphStyle('RemarksTxt', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=8.2, leading=11, alignment=TA_LEFT)

    story = []
    
    if os.path.exists(LOGO_IMAGE_PATH):
        try:
            logo_img = RLImage(LOGO_IMAGE_PATH, width=0.75*inch, height=0.85*inch)
            logo_img.hAlign = 'CENTER'
            story.append(logo_img)
            story.append(Spacer(1, 3))
        except Exception: pass
    
    story.append(Paragraph(data["shop_name"], style_shop))
    story.append(Spacer(1, 2))
    story.append(Paragraph(data["shop_address"], style_shop_addr))
    story.append(Spacer(1, 3))
    
    if data["doc_type"] == "CREDIT NOTE":
        label_decor = "*** GOODS RETURN ***"
        num_prefix = "CREDIT NO:"
    elif data["doc_type"] == "CHECKING LIST":
        label_decor = "*** CHECKING LIST ***"
        num_prefix = "ORDER #:"
    elif data["doc_type"] == "SALES QUOTE":
        label_decor = "*** SALES QUOTE ***"
        num_prefix = "QUOTE #:"
    else:
        label_decor = "*** ESTIMATE ***"
        num_prefix = "ESTIMATE #:"
        
    story.append(Paragraph(label_decor, style_center))
    story.append(Paragraph("-" * 52, style_center))
    story.append(Spacer(1, 2))
        
    story.append(Paragraph(f"<b>{num_prefix}</b> {data['estimate_no']}", style_left))
    story.append(Paragraph(f"<b>Date:</b> {data['date']}", style_left))
    story.append(Paragraph(f"<b>Customer Name:</b> {data['customer']}", style_left))
    
    if data["contact"].strip():
        story.append(Paragraph(f"<b>Contact:</b> {data['contact']}", style_left))
    if data["address"].strip():
        story.append(Paragraph(f"<b>Customer Address:</b> {data['address']}", style_left))
    if data["shipping_address"].strip() and data["shipping_address"].lower() != data["customer"].lower() and "walk in" not in data["shipping_address"].lower():
        story.append(Paragraph(f"<b>Shipping Address:</b> {data['shipping_address']}", style_left))
        
    story.append(Spacer(1, 2))
    story.append(Paragraph("=" * 44, style_center))
    story.append(Spacer(1, 3))
    
    if data["doc_type"] == "CHECKING LIST":
        col_widths = [40, 155]
        table_content = [[Paragraph("Qty", style_hdr), Paragraph("Item Description", style_hdr_left)]]
        for item in data["items"]:
            qty_block = f"<para align=center>{item[0]}<br/><font size=6.5 color=black>{item[1]}</font></para>"
            table_content.append([Paragraph(qty_block, style_qty_num), Paragraph(item[2], style_item)])
            
        items_table = Table(table_content, colWidths=col_widths)
        items_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 2), ('RIGHTPADDING', (0,0), (-1,-1), 2),
            ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('LINEBELOW', (0,0), (-1,0), 1.2, 'black'), 
            ('LINEBELOW', (0,1), (-1,-1), 1.0, 'black'),  
            ('LINEBEFORE', (1,0), (1,-1), 1.0, 'black'), 
        ]))
        story.append(items_table)
        story.append(Spacer(1, 4))
        
        totals_table = Table([[Paragraph("<b>TOTAL QTY:</b>", style_total_lbl), Paragraph(f"<b>{format_smart_amount(data['total_qty'])}</b>", style_total_val)]], colWidths=[115, 80])
        totals_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2), ('TOPPADDING', (0,0), (-1,-1), 2), ('LINEABOVE', (0,0), (1,0), 1.2, 'black'),
        ]))
        story.append(totals_table)
    else:
        if data["has_discount"]:
            col_widths = [26, 74, 33, 27, 35]
            table_content = [[
                Paragraph("Qty", style_hdr), 
                Paragraph("Item Description", style_hdr_left), 
                Paragraph("Price", style_hdr_right), 
                Paragraph("Disc", style_hdr_right), 
                Paragraph("Total", style_hdr_right)
            ]]
            for item in data["items"]:
                qty_block = f"<para align=center>{item[0]}<br/><font size=6.5 color=black>{item[1]}</font></para>"
                table_content.append([
                    Paragraph(qty_block, style_qty_num), 
                    Paragraph(item[2], style_item), 
                    Paragraph(item[3], style_total), 
                    Paragraph(item[4], style_total), 
                    Paragraph(item[5], style_total)
                ])
                
            items_table = Table(table_content, colWidths=col_widths)
            items_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('LEFTPADDING', (0,0), (-1,-1), 1), ('RIGHTPADDING', (0,0), (-1,-1), 1),
                ('TOPPADDING', (0,0), (-1,-1), 3), ('BOTTOMPADDING', (0,0), (-1,-1), 3),
                ('LINEBELOW', (0,0), (-1,0), 1.2, 'black'), 
                ('LINEBELOW', (0,1), (-1,-1), 1.0, 'black'),  
                ('LINEBEFORE', (1,0), (1,-1), 1.0, 'black'), 
                ('LINEBEFORE', (2,0), (2,-1), 1.0, 'black'), 
                ('LINEBEFORE', (3,0), (3,-1), 1.0, 'black'), 
                ('LINEBEFORE', (4,0), (4,-1), 1.0, 'black'), 
            ]))
        else:
            col_widths = [30, 95, 35, 35]
            table_content = [[
                Paragraph("Qty", style_hdr), 
                Paragraph("Item Description", style_hdr_left), 
                Paragraph("Price", style_hdr_right), 
                Paragraph("Total", style_hdr_right)
            ]]
            for item in data["items"]:
                qty_block = f"<para align=center>{item[0]}<br/><font size=6.5 color=black>{item[1]}</font></para>"
                table_content.append([
                    Paragraph(qty_block, style_qty_num), 
                    Paragraph(item[2], style_item), 
                    Paragraph(item[3], style_total), 
                    Paragraph(item[5], style_total)
                ])
                
            items_table = Table(table_content, colWidths=col_widths)
            items_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('LEFTPADDING', (0,0), (-1,-1), 2), ('RIGHTPADDING', (0,0), (-1,-1), 2),
                ('TOPPADDING', (0,0), (-1,-1), 4), ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                ('LINEBELOW', (0,0), (-1,0), 1.2, 'black'), 
                ('LINEBELOW', (0,1), (-1,-1), 1.0, 'black'),  
                ('LINEBEFORE', (1,0), (1,-1), 1.0, 'black'), 
                ('LINEBEFORE', (2,0), (2,-1), 1.0, 'black'), 
                ('LINEBEFORE', (3,0), (3,-1), 1.0, 'black'), 
            ]))
            
        story.append(items_table)
        story.append(Spacer(1, 4))
        
        grand_total_label = "<b>TOTAL CREDIT:</b>" if data["doc_type"] == "CREDIT NOTE" else "<b>GRAND TOTAL:</b>"
        totals_data = [
            [Paragraph("Sub-Total:", style_total_lbl), Paragraph(f"{data['subtotal']}", style_total_val)],
            [Paragraph("Cartage/P.B:", style_total_lbl), Paragraph(f"{data['cartage']}", style_total_val)],
            [Paragraph(grand_total_label, style_total_lbl), Paragraph(f"<b>{data['total']}</b>", style_total_val)]
        ]
        
        if not data.get("hide_paid_balance", False) and data["doc_type"] != "CREDIT NOTE":
            totals_data.append([Paragraph("Paid:", style_total_lbl), Paragraph(f"{data['paid']}", style_total_val)])
            totals_data.append([Paragraph("<b>Balance:</b>", style_total_lbl), Paragraph(f"<b>{data['balance']}</b>", style_total_val)])
            
        totals_table = Table(totals_data, colWidths=[115, 80])
        totals_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2), ('TOPPADDING', (0,0), (-1,-1), 2), ('LINEABOVE', (0,2), (1,2), 1.2, 'black'),
        ]))
        story.append(totals_table)
    
    if data["remarks"]:
        story.append(Spacer(1, 4))
        story.append(Paragraph("-" * 52, style_center))
        story.append(Paragraph(f"<b>Remarks:</b><br/>{data['remarks']}", style_remarks))
        story.append(Paragraph("-" * 52, style_center))
        
    story.append(Spacer(1, 8))
    story.append(Paragraph("15% Deduction On Goods Return Within 10 Days", style_center))
    story.append(Paragraph("On Presentation Of Original Bill.", style_center))
    story.append(Spacer(1, 3))
    story.append(Paragraph("<b>Thank You!</b>", style_shop))
    
    if os.path.exists(WHATSAPP_QR_PATH):
        try:
            story.append(Spacer(1, 4))
            story.append(Paragraph("<b>Follow us on FaceBook</b>", style_center))
            story.append(Spacer(1, 2))
            qr_img = RLImage(WHATSAPP_QR_PATH, width=1.1*inch, height=1.1*inch)
            qr_img.hAlign = 'CENTER'
            story.append(qr_img)
        except Exception:
            pass
    
    doc.build(story)


# ==============================================================================
# BACKGROUND WATCHER THREAD
# ==============================================================================

class WatcherWorker(QThread):
    log_signal = pyqtSignal(str)
    file_processed_signal = pyqtSignal(str, dict)

    def __init__(self, watch_dir, output_dir):
        super().__init__()
        self.watch_dir = watch_dir
        self.output_dir = output_dir
        self.running = True

    def run(self):
        self.log_signal.emit(f"[INFO] Watching folder: {self.watch_dir}")
        while self.running:
            try:
                if os.path.exists(self.watch_dir):
                    files = [f for f in os.listdir(self.watch_dir) if f.lower().endswith('.pdf')]
                    for file in files:
                        if not self.running:
                            break
                        target_path = os.path.join(self.watch_dir, file)
                        self.log_signal.emit(f"[PROCESSING] Reading file: {file}")
                        time.sleep(0.15)

                        output_pdf = os.path.join(self.output_dir, f"receipt_{file}")
                        try:
                            dataset = parse_noorani_invoice(target_path)
                            try:
                                generate_3inch_receipt(dataset, output_pdf)
                            except PermissionError:
                                output_pdf = os.path.join(self.output_dir, f"receipt_{int(time.time())}_{file}")
                                generate_3inch_receipt(dataset, output_pdf)

                            self.log_signal.emit(f"[SUCCESS] Converted -> {os.path.basename(output_pdf)}")
                            self.file_processed_signal.emit(output_pdf, dataset)

                        except Exception as e:
                            self.log_signal.emit(f"[ERROR] Parsing failed for {file}: {str(e)}")

                        try:
                            os.remove(target_path)
                            self.log_signal.emit(f"[CLEANUP] Removed source invoice file.")
                        except Exception as e:
                            self.log_signal.emit(f"[WARNING] Could not delete source file: {str(e)}")

                time.sleep(0.35)
            except Exception as outer_e:
                self.log_signal.emit(f"[CRITICAL ERROR] Watcher loop: {str(outer_e)}")
                time.sleep(1)

    def stop(self):
        self.running = False
        self.log_signal.emit("[INFO] Automation watcher stopped.")


# ==============================================================================
# MAIN WORKSTATION GUI
# ==============================================================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("NOORANI SANITARY RECEIPT GENERATOR")
        self.resize(1260, 820)
        self.watcher_thread = None
        self.last_generated_pdf = None
        self.current_pixmap = None
        self.is_billing_active = False
        
        self.zoom_factor = 0.73
        self.fit_width_mode = False

        self.init_ui()
        self.populate_printers()
        self.apply_stylesheet()

        # AUTO-START WATCHER ON APP OPEN
        self.start_watcher()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(16)

        # ----------------------------------------------------------------------
        # LEFT NAVIGATION RAIL
        # ----------------------------------------------------------------------
        nav_rail = QFrame()
        nav_rail.setObjectName("BrandFrame")
        nav_rail.setFixedWidth(240)
        nav_layout = QVBoxLayout(nav_rail)
        nav_layout.setContentsMargins(14, 18, 14, 18)
        nav_layout.setSpacing(14)

        # App Brand Header
        brand_box = QHBoxLayout()
        logo_badge = QLabel("N")
        logo_badge.setObjectName("LogoBadge")
        logo_badge.setFixedSize(38, 38)
        logo_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)

        brand_titles = QVBoxLayout()
        brand_titles.setSpacing(1)
        app_title = QLabel("NOORANI SANITARY")
        app_title.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        app_title.setStyleSheet("color: #0f172a; font-weight: 700;")

        sub_title = QLabel("Receipt Generator")
        sub_title.setFont(QFont("Segoe UI", 8))
        sub_title.setStyleSheet("color: #64748b;")

        brand_titles.addWidget(app_title)
        brand_titles.addWidget(sub_title)

        brand_box.addWidget(logo_badge)
        brand_box.addSpacing(8)
        brand_box.addLayout(brand_titles)
        brand_box.addStretch()

        nav_layout.addLayout(brand_box)
        nav_layout.addSpacing(8)

        # Navigation Menu Items
        self.nav_list = QListWidget()
        self.nav_list.setObjectName("NavList")
        
        nav_font = QFont("Segoe UI", 10, QFont.Weight.Bold)
        
        item_print = QListWidgetItem("Print Station")
        item_print.setFont(nav_font)
        
        item_settings = QListWidgetItem("Settings & Logs")
        item_settings.setFont(nav_font)
        
        self.nav_list.addItem(item_print)
        self.nav_list.addItem(item_settings)
        self.nav_list.setCurrentRow(0)
        self.nav_list.currentRowChanged.connect(self.switch_view)

        nav_layout.addWidget(self.nav_list)
        nav_layout.addStretch()

        # LIVE STATUS INDICATOR BADGE
        self.status_card = QFrame()
        self.status_card.setObjectName("StatusCardIdle")
        status_card_layout = QHBoxLayout(self.status_card)
        status_card_layout.setContentsMargins(12, 10, 12, 10)

        self.status_dot = QLabel("●")
        self.status_dot.setStyleSheet("color: #64748b; font-size: 14px;")
        self.status_text = QLabel("Watcher Idle")
        self.status_text.setStyleSheet("font-weight: 700; color: #475569; font-size: 11px;")

        status_card_layout.addWidget(self.status_dot)
        status_card_layout.addWidget(self.status_text)
        status_card_layout.addStretch()

        nav_layout.addWidget(self.status_card)

        main_layout.addWidget(nav_rail)

        # ----------------------------------------------------------------------
        # MAIN STACKED WORKSPACE
        # ----------------------------------------------------------------------
        self.stacked_widget = QStackedWidget()

        # VIEW 1: PRINT STATION PAGE
        self.print_view_page = QWidget()
        print_view_layout = QHBoxLayout(self.print_view_page)
        print_view_layout.setContentsMargins(0, 0, 0, 0)
        print_view_layout.setSpacing(16)

        # Controls Side Panel
        controls_panel = QFrame()
        controls_panel.setObjectName("CardFrame")
        controls_panel.setFixedWidth(280)
        controls_layout = QVBoxLayout(controls_panel)
        controls_layout.setContentsMargins(16, 16, 16, 16)
        controls_layout.setSpacing(12)

        controls_title = QLabel("Print Controls")
        controls_title.setStyleSheet("color: #1e293b; font-size: 14px; font-weight: 700; font-family: 'Segoe UI';")
        controls_layout.addWidget(controls_title)

        controls_layout.addWidget(QLabel("Target Printer"))
        printer_row = QHBoxLayout()
        self.printer_combo = QComboBox()
        refresh_printers_btn = QPushButton("Refresh")
        refresh_printers_btn.clicked.connect(self.populate_printers)
        printer_row.addWidget(self.printer_combo, 1)
        printer_row.addWidget(refresh_printers_btn)
        controls_layout.addLayout(printer_row)

        # PRINTER PREFERENCES DROPDOWNS
        pref_card = QFrame()
        pref_card.setObjectName("LightMetaCard")
        pref_card_layout = QVBoxLayout(pref_card)
        pref_card_layout.setContentsMargins(12, 10, 12, 10)
        pref_card_layout.setSpacing(8)

        pref_hdr = QLabel("Print Preferences")
        pref_hdr.setStyleSheet("color: #0f172a; font-size: 11px; font-weight: 800; font-family: 'Segoe UI';")
        pref_card_layout.addWidget(pref_hdr)

        orient_layout = QHBoxLayout()
        orient_lbl = QLabel("Orientation:")
        orient_lbl.setStyleSheet("font-size: 11px; font-weight: 600; color: #334155;")
        self.orient_combo = QComboBox()
        self.orient_combo.addItems(["Auto", "Portrait", "Landscape"])
        orient_layout.addWidget(orient_lbl)
        orient_layout.addWidget(self.orient_combo, 1)
        pref_card_layout.addLayout(orient_layout)

        scale_layout = QHBoxLayout()
        scale_lbl = QLabel("Page Scaling:")
        scale_lbl.setStyleSheet("font-size: 11px; font-weight: 600; color: #334155;")
        self.scale_combo = QComboBox()
        self.scale_combo.addItems(["Actual Size (1:1)", "Auto Fit Width"])
        scale_layout.addWidget(scale_lbl)
        scale_layout.addWidget(self.scale_combo, 1)
        pref_card_layout.addLayout(scale_layout)

        controls_layout.addWidget(pref_card)

        controls_layout.addSpacing(4)

        # START / STOP BILLING SOFTWARE TOGGLE BUTTON
        self.toggle_billing_btn = QPushButton("Start Billing Software")
        self.toggle_billing_btn.setObjectName("StartBtn")
        self.toggle_billing_btn.setFixedHeight(44)
        self.toggle_billing_btn.clicked.connect(self.toggle_billing_software)

        controls_layout.addWidget(self.toggle_billing_btn)

        controls_layout.addSpacing(6)

        # ORDER SUMMARY META CARD
        meta_card = QFrame()
        meta_card.setObjectName("LightMetaCard")
        meta_card_layout = QVBoxLayout(meta_card)
        meta_card_layout.setContentsMargins(16, 16, 16, 16)
        meta_card_layout.setSpacing(12)

        meta_header = QHBoxLayout()
        meta_title = QLabel("Order Summary")
        meta_title.setStyleSheet("color: #0f172a; font-size: 13px; font-weight: 800; font-family: 'Segoe UI';")
        meta_header.addWidget(meta_title)
        meta_header.addStretch()
        meta_card_layout.addLayout(meta_header)

        divider1 = QFrame()
        divider1.setFrameShape(QFrame.Shape.HLine)
        divider1.setStyleSheet("background-color: #e2e8f0; max-height: 1px; border: none;")
        meta_card_layout.addWidget(divider1)

        id_box = QHBoxLayout()
        id_lbl = QLabel("Order #")
        id_lbl.setStyleSheet("color: #334155; font-size: 12px; font-weight: 700; font-family: 'Segoe UI';")
        self.lbl_cust_id = QLabel("N/A")
        self.lbl_cust_id.setStyleSheet("color: #0f172a; font-size: 12px; font-weight: 700; font-family: 'Segoe UI';")
        self.lbl_cust_id.setAlignment(Qt.AlignmentFlag.AlignRight)
        id_box.addWidget(id_lbl)
        id_box.addWidget(self.lbl_cust_id)
        meta_card_layout.addLayout(id_box)

        name_box = QHBoxLayout()
        name_lbl = QLabel("Customer")
        name_lbl.setStyleSheet("color: #334155; font-size: 12px; font-weight: 700; font-family: 'Segoe UI';")
        self.lbl_cust_name = QLabel("N/A")
        self.lbl_cust_name.setStyleSheet("color: #0f172a; font-size: 12px; font-weight: 700; font-family: 'Segoe UI';")
        self.lbl_cust_name.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.lbl_cust_name.setWordWrap(True)
        name_box.addWidget(name_lbl)
        name_box.addWidget(self.lbl_cust_name)
        meta_card_layout.addLayout(name_box)

        date_box = QHBoxLayout()
        date_lbl = QLabel("Billing Date")
        date_lbl.setStyleSheet("color: #334155; font-size: 12px; font-weight: 700; font-family: 'Segoe UI';")
        self.lbl_bill_date = QLabel("N/A")
        self.lbl_bill_date.setStyleSheet("color: #0f172a; font-size: 12px; font-weight: 700; font-family: 'Segoe UI';")
        self.lbl_bill_date.setAlignment(Qt.AlignmentFlag.AlignRight)
        date_box.addWidget(date_lbl)
        date_box.addWidget(self.lbl_bill_date)
        meta_card_layout.addLayout(date_box)

        divider2 = QFrame()
        divider2.setFrameShape(QFrame.Shape.HLine)
        divider2.setStyleSheet("background-color: #e2e8f0; max-height: 1px; border: none;")
        meta_card_layout.addWidget(divider2)

        amount_box = QFrame()
        amount_box.setObjectName("ChargeBlock")
        amount_box_layout = QHBoxLayout(amount_box)
        amount_box_layout.setContentsMargins(14, 12, 14, 12)

        charge_lbl = QLabel("Charge")
        charge_lbl.setStyleSheet("color: #ffffff; font-size: 13px; font-weight: 800; font-family: 'Segoe UI';")

        self.lbl_bill_amount = QLabel("Rs. 0/=")
        self.lbl_bill_amount.setStyleSheet("color: #ffffff; font-size: 14px; font-weight: 800; font-family: 'Segoe UI';")
        self.lbl_bill_amount.setAlignment(Qt.AlignmentFlag.AlignRight)

        amount_box_layout.addWidget(charge_lbl)
        amount_box_layout.addStretch()
        amount_box_layout.addWidget(self.lbl_bill_amount)

        meta_card_layout.addWidget(amount_box)

        controls_layout.addWidget(meta_card)
        
        controls_layout.addSpacing(10)

        # MANUAL PRINT BUTTON
        self.manual_print_btn = QPushButton("Print Bill")
        self.manual_print_btn.setObjectName("PrintBtn")
        self.manual_print_btn.setFixedHeight(44)
        self.manual_print_btn.clicked.connect(self.print_current_receipt)
        controls_layout.addWidget(self.manual_print_btn)

        controls_layout.addStretch()

        print_view_layout.addWidget(controls_panel)

        # PDF Canvas Display
        canvas_container = QWidget()
        canvas_layout = QVBoxLayout(canvas_container)
        canvas_layout.setContentsMargins(0, 0, 0, 0)
        canvas_layout.setSpacing(12)

        top_bar = QFrame()
        top_bar.setObjectName("CardFrame")
        top_bar_layout = QHBoxLayout(top_bar)
        top_bar_layout.setContentsMargins(16, 10, 16, 10)

        preview_header = QLabel("Receipt Viewer")
        preview_header.setStyleSheet("color: #1e293b; font-size: 13px; font-weight: 700; font-family: 'Segoe UI';")

        zoom_in_btn = QPushButton("+")
        zoom_in_btn.setObjectName("ZoomBtn")
        zoom_in_btn.setFixedSize(32, 32)
        zoom_in_btn.clicked.connect(self.zoom_in)

        zoom_out_btn = QPushButton("−")
        zoom_out_btn.setObjectName("ZoomBtn")
        zoom_out_btn.setFixedSize(32, 32)
        zoom_out_btn.clicked.connect(self.zoom_out)

        default_73_btn = QPushButton("73%")
        default_73_btn.clicked.connect(self.set_default_73_zoom)

        fit_width_btn = QPushButton("Fit Width")
        fit_width_btn.clicked.connect(self.set_fit_width)

        reset_zoom_btn = QPushButton("100%")
        reset_zoom_btn.clicked.connect(self.reset_zoom)

        top_bar_layout.addWidget(preview_header)
        top_bar_layout.addStretch()
        top_bar_layout.addWidget(zoom_in_btn)
        top_bar_layout.addWidget(zoom_out_btn)
        top_bar_layout.addWidget(default_73_btn)
        top_bar_layout.addWidget(fit_width_btn)
        top_bar_layout.addWidget(reset_zoom_btn)

        canvas_layout.addWidget(top_bar)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setObjectName("CanvasScrollArea")

        self.pdf_render_label = QLabel("No active receipt loaded.")
        self.pdf_render_label.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        self.pdf_render_label.setStyleSheet("color: #64748b; font-size: 14px; font-weight: 600; margin-top: 20px;")

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(25)
        shadow.setColor(QColor(0, 0, 0, 120))
        shadow.setXOffset(0)
        shadow.setYOffset(6)
        self.pdf_render_label.setGraphicsEffect(shadow)

        self.scroll_area.setWidget(self.pdf_render_label)
        canvas_layout.addWidget(self.scroll_area, 1)

        print_view_layout.addWidget(canvas_container, 1)

        # VIEW 2: SETTINGS & LOGS PAGE
        self.settings_view_page = QWidget()
        settings_layout = QVBoxLayout(self.settings_view_page)
        settings_layout.setContentsMargins(0, 0, 0, 0)
        settings_layout.setSpacing(16)

        config_card = QFrame()
        config_card.setObjectName("CardFrame")
        config_card_layout = QVBoxLayout(config_card)
        config_card_layout.setContentsMargins(20, 18, 20, 18)
        config_card_layout.setSpacing(12)

        config_header = QLabel("Directory & Automation Settings")
        config_header.setStyleSheet("color: #1e293b; font-size: 14px; font-weight: 700; font-family: 'Segoe UI';")
        config_card_layout.addWidget(config_header)

        config_card_layout.addWidget(QLabel("Watch Directory"))
        watch_row = QHBoxLayout()
        self.watch_input = QLineEdit(os.path.abspath("./inflow_invoices"))
        watch_btn = QPushButton("Browse")
        watch_btn.clicked.connect(lambda: self.browse_folder(self.watch_input))
        watch_row.addWidget(self.watch_input)
        watch_row.addWidget(watch_btn)
        config_card_layout.addLayout(watch_row)

        config_card_layout.addWidget(QLabel("Output Directory"))
        out_row = QHBoxLayout()
        self.out_input = QLineEdit(os.path.abspath("./thermal_receipts"))
        out_btn = QPushButton("Browse")
        out_btn.clicked.connect(lambda: self.browse_folder(self.out_input))
        out_row.addWidget(self.out_input)
        out_row.addWidget(out_btn)
        config_card_layout.addLayout(out_row)

        settings_layout.addWidget(config_card)

        log_card = QFrame()
        log_card.setObjectName("CardFrame")
        log_card_layout = QVBoxLayout(log_card)
        log_card_layout.setContentsMargins(20, 18, 20, 18)
        log_card_layout.setSpacing(10)

        log_header = QLabel("Activity Terminal Console")
        log_header.setStyleSheet("color: #1e293b; font-size: 14px; font-weight: 700; font-family: 'Segoe UI';")
        log_card_layout.addWidget(log_header)

        self.log_console = QTextEdit()
        self.log_console.setReadOnly(True)
        log_card_layout.addWidget(self.log_console, 1)

        settings_layout.addWidget(log_card, 1)

        self.stacked_widget.addWidget(self.print_view_page)
        self.stacked_widget.addWidget(self.settings_view_page)

        main_layout.addWidget(self.stacked_widget, 1)

    def switch_view(self, index):
        self.stacked_widget.setCurrentIndex(index)

    def populate_printers(self):
        """Fetches active printers installed on Windows."""
        self.printer_combo.clear()
        printers = []

        if WIN32_PRINT_AVAILABLE:
            try:
                raw_printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
                for printer in raw_printers:
                    printers.append(printer[2])
                default_printer = win32print.GetDefaultPrinter()
            except Exception:
                default_printer = None
        else:
            default_printer = None

        if not printers:
            printers = ["Microsoft Print to PDF", "Save as PDF"]

        self.printer_combo.addItems(printers)

        if default_printer and default_printer in printers:
            self.printer_combo.setCurrentText(default_printer)

        self.append_log(f"[PRINTER] Active selection: '{self.printer_combo.currentText()}'")

    def render_pdf_to_view(self, pdf_path):
        """Renders PDF image into QPixmap cleanly forcing RGB without transparent alpha dark layers."""
        try:
            doc = pymupdf.open(pdf_path)
            page = doc.load_page(0)

            # Force alpha=False to prevent black background rendering on other systems
            pix = page.get_pixmap(dpi=150, alpha=False)
            
            qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888)

            self.current_pixmap = QPixmap.fromImage(qimg)
            self.fit_width_mode = False
            
            self.zoom_factor = 0.73
            self.update_preview_scaling()

            doc.close()
            self.append_log(f"[PREVIEW] Rendered preview for {os.path.basename(pdf_path)} (Default 73% Zoom)")
        except Exception as e:
            self.pdf_render_label.setText(f"Preview rendering error:\n{e}")
            self.append_log(f"[ERROR] Could not render PDF preview: {e}")

    def update_preview_scaling(self):
        """Scales receipt image preserving strict aspect ratio."""
        if not self.current_pixmap:
            return

        if self.fit_width_mode:
            viewport_width = self.scroll_area.viewport().width() - 32
            if viewport_width > 100:
                scaled_pixmap = self.current_pixmap.scaledToWidth(
                    viewport_width,
                    Qt.TransformationMode.SmoothTransformation
                )
                self.pdf_render_label.setPixmap(scaled_pixmap)
        else:
            target_width = int(self.current_pixmap.width() * self.zoom_factor)
            scaled_pixmap = self.current_pixmap.scaledToWidth(
                target_width,
                Qt.TransformationMode.SmoothTransformation
            )
            self.pdf_render_label.setPixmap(scaled_pixmap)

    def zoom_in(self):
        if not self.current_pixmap:
            return
        if self.fit_width_mode:
            self.zoom_factor = (self.scroll_area.viewport().width() - 32) / self.current_pixmap.width()
            self.fit_width_mode = False
        self.zoom_factor = min(3.0, self.zoom_factor + 0.10)
        self.update_preview_scaling()

    def zoom_out(self):
        if not self.current_pixmap:
            return
        if self.fit_width_mode:
            self.zoom_factor = (self.scroll_area.viewport().width() - 32) / self.current_pixmap.width()
            self.fit_width_mode = False
        self.zoom_factor = max(0.2, self.zoom_factor - 0.10)
        self.update_preview_scaling()

    def set_default_73_zoom(self):
        if not self.current_pixmap:
            return
        self.fit_width_mode = False
        self.zoom_factor = 0.73
        self.update_preview_scaling()

    def set_fit_width(self):
        if not self.current_pixmap:
            return
        self.fit_width_mode = True
        self.update_preview_scaling()

    def reset_zoom(self):
        if not self.current_pixmap:
            return
        self.fit_width_mode = False
        self.zoom_factor = 1.0
        self.update_preview_scaling()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.fit_width_mode:
            self.update_preview_scaling()

    def send_to_printer(self, pdf_path, printer_name):
        """Renders PDF to a clean 1-bit B&W image without heavy font bolding."""
        abs_path = os.path.abspath(pdf_path)
        if not os.path.exists(abs_path):
            self.append_log(f"[PRINT ERROR] File missing: {abs_path}")
            return False

        if not WIN32_PRINT_AVAILABLE:
            self.append_log("[PRINT ERROR] pywin32 library missing!")
            return False

        orient_choice = self.orient_combo.currentText()

        try:
            doc = pymupdf.open(abs_path)
            page = doc.load_page(0)

            if orient_choice == "Auto":
                if page.rect.width > page.rect.height:
                    page.set_rotation(90)
            elif orient_choice == "Landscape":
                page.set_rotation(90)
            elif orient_choice == "Portrait":
                page.set_rotation(0)

            # Native 203 DPI raster render
            pix = page.get_pixmap(dpi=203)
            doc.close()

            # Create PIL RGB Image
            img_rgb = PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Target exact 576 dot head width
            target_width = 576
            aspect_ratio = img_rgb.height / img_rgb.width
            target_height = int(target_width * aspect_ratio)

            # 1. High quality BILINEAR resize to prevent font line thickening
            img_resized = img_rgb.resize((target_width, target_height), PILImage.Resampling.BILINEAR)

            # 2. Convert to Grayscale
            img_gray = img_resized.convert("L")

            # 3. Higher threshold (130) makes text thinner and prevents bolding effect
            threshold = 130
            img_mono = img_gray.point(lambda p: 255 if p > threshold else 0, mode='1')

            # Re-convert back to RGB format for win32ui DC compatibility
            img_final = img_mono.convert("RGB")

            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(printer_name)
            
            hDC.StartDoc(os.path.basename(abs_path))
            hDC.StartPage()

            dib_pil = PIL.ImageWin.Dib(img_final)
            dib_pil.draw(hDC.GetHandleOutput(), (0, 0, target_width, target_height))

            hDC.EndPage()
            hDC.EndDoc()

            self.append_log(f"[PRINT SUCCESS] Crisp text printed on '{printer_name}' (576 Dots)")
            return True

        except Exception as e:
            self.append_log(f"[PRINT CRITICAL ERROR] {str(e)}")
            QMessageBox.critical(self, "Printing Failed", f"Printer error on '{printer_name}':\n{str(e)}")
            return False

    def browse_folder(self, line_edit):
        folder = QFileDialog.getExistingDirectory(self, "Select Directory", line_edit.text())
        if folder:
            line_edit.setText(os.path.abspath(folder))

    def append_log(self, text):
        self.log_console.append(text)

    def handle_processed_file(self, pdf_path, invoice_data):
        """Renders preview and updates order info ONLY (AUTO-PRINT DISABLED)."""
        self.last_generated_pdf = pdf_path
        self.render_pdf_to_view(pdf_path)

        self.lbl_cust_id.setText(str(invoice_data.get("estimate_no", "N/A")))
        self.lbl_cust_name.setText(str(invoice_data.get("customer", "N/A")))
        self.lbl_bill_date.setText(str(invoice_data.get("date", "N/A")))
        
        tot_val = invoice_data.get("total", "0/=")
        if not str(tot_val).startswith("Rs"):
            tot_val = f"Rs. {tot_val}"
        self.lbl_bill_amount.setText(tot_val)

    def print_current_receipt(self):
        """Triggered ONLY when user manually clicks 'Print Bill' button."""
        if not self.last_generated_pdf or not os.path.exists(self.last_generated_pdf):
            QMessageBox.warning(self, "No Receipt Loaded", "Please wait for a receipt to be generated before printing.")
            return

        selected_printer = self.printer_combo.currentText()
        self.send_to_printer(self.last_generated_pdf, selected_printer)

    def toggle_billing_software(self):
        if not self.is_billing_active:
            self.start_watcher()
        else:
            self.stop_watcher()

    def start_watcher(self):
        watch_dir = self.watch_input.text().strip()
        out_dir = self.out_input.text().strip()

        os.makedirs(watch_dir, exist_ok=True)
        os.makedirs(out_dir, exist_ok=True)

        self.watcher_thread = WatcherWorker(watch_dir, out_dir)
        self.watcher_thread.log_signal.connect(self.append_log)

        if self.watcher_thread:
            try:
                self.watcher_thread.file_processed_signal.disconnect()
            except Exception:
                pass
            self.watcher_thread.file_processed_signal.connect(self.handle_processed_file)

        self.watcher_thread.start()

        self.is_billing_active = True
        self.toggle_billing_btn.setText("Stop Billing Software")
        self.toggle_billing_btn.setObjectName("StopBtn")
        self.toggle_billing_btn.setStyle(self.toggle_billing_btn.style())

        self.watch_input.setEnabled(False)
        self.out_input.setEnabled(False)

        self.status_card.setObjectName("StatusCardActive")
        self.status_dot.setStyleSheet("color: #10b981; font-size: 14px;")
        self.status_text.setText("Watcher Active")
        self.status_text.setStyleSheet("font-weight: 700; color: #065f46; font-size: 11px;")
        self.status_card.setStyle(self.status_card.style())

    def stop_watcher(self):
        if self.watcher_thread and self.watcher_thread.isRunning():
            self.watcher_thread.stop()
            self.watcher_thread.wait()

        self.is_billing_active = False
        self.toggle_billing_btn.setText("Start Billing Software")
        self.toggle_billing_btn.setObjectName("StartBtn")
        self.toggle_billing_btn.setStyle(self.toggle_billing_btn.style())

        self.watch_input.setEnabled(True)
        self.out_input.setEnabled(True)

        self.status_card.setObjectName("StatusCardIdle")
        self.status_dot.setStyleSheet("color: #64748b; font-size: 14px;")
        self.status_text.setText("Watcher Idle")
        self.status_text.setStyleSheet("font-weight: 700; color: #475569; font-size: 11px;")
        self.status_card.setStyle(self.status_card.style())

    def closeEvent(self, event):
        self.stop_watcher()
        event.accept()

    def apply_stylesheet(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f7fb;
            }
            #BrandFrame, #CardFrame {
                background-color: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
            }
            #LightMetaCard {
                background-color: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 12px;
            }
            #ChargeBlock {
                background-color: #2563eb;
                border-radius: 8px;
            }
            #LogoBadge {
                background-color: #4f46e5;
                color: #ffffff;
                border-radius: 10px;
                font-size: 18px;
                font-weight: bold;
            }
            #StatusCardIdle {
                background-color: #f1f5f9;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
            }
            #StatusCardActive {
                background-color: #d1fae5;
                border: 1px solid #6ee7b7;
                border-radius: 8px;
            }
            #NavList {
                background-color: transparent;
                border: none;
                outline: none;
                font-family: 'Segoe UI';
                font-weight: 700;
            }
            #NavList::item {
                padding: 10px 14px;
                border-radius: 8px;
                font-size: 13px;
                font-weight: 700;
                color: #64748b;
                margin-bottom: 4px;
            }
            #NavList::item:selected {
                background-color: #4f46e5;
                color: #ffffff;
                font-weight: 700;
            }
            #NavList::item:hover:!selected {
                background-color: #e0e7ff;
                color: #4338ca;
                font-weight: 700;
            }
            QLabel {
                font-size: 12px;
                color: #475569;
            }
            QLineEdit {
                padding: 7px 12px;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                background-color: #ffffff;
                font-size: 12px;
                color: #0f172a;
            }
            QLineEdit:focus {
                border: 1.5px solid #4f46e5;
                background-color: #ffffff;
            }
            QComboBox {
                padding: 6px 10px;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                background-color: #ffffff;
                font-size: 11px;
                font-weight: 600;
                color: #0f172a;
            }
            QComboBox:focus {
                border: 1.5px solid #4f46e5;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border: none;
            }
            QComboBox QAbstractItemView {
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                background-color: #ffffff;
                color: #0f172a;
                selection-background-color: #e0e7ff;
                selection-color: #4338ca;
            }
            QComboBox QAbstractItemView::item {
                color: #0f172a;
                background-color: #ffffff;
                min-height: 24px;
            }
            QComboBox QAbstractItemView::item:selected {
                color: #4338ca;
                background-color: #e0e7ff;
            }
            QPushButton {
                padding: 7px 14px;
                border-radius: 8px;
                border: 1px solid #cbd5e1;
                background-color: #ffffff;
                font-weight: 600;
                font-size: 12px;
                color: #334155;
            }
            QPushButton:hover {
                background-color: #f8fafc;
                border-color: #94a3b8;
            }
            
            #ZoomBtn {
                padding: 0px;
                font-size: 16px;
                font-weight: 800;
                color: #0f172a;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                background-color: #ffffff;
            }
            #ZoomBtn:hover {
                background-color: #f1f5f9;
                border-color: #4f46e5;
                color: #4f46e5;
            }

            #StartBtn {
                background-color: #10b981;
                color: #ffffff;
                border: none;
                border-radius: 8px;
                font-weight: bold;
                font-size: 13px;
            }
            #StartBtn:hover {
                background-color: #059669;
            }
            #StopBtn {
                background-color: #ef4444;
                color: #ffffff;
                border: none;
                border-radius: 8px;
                font-weight: bold;
                font-size: 13px;
            }
            #StopBtn:hover {
                background-color: #dc2626;
            }
            #PrintBtn {
                background-color: #0f172a;
                color: #ffffff;
                border: none;
                border-radius: 8px;
                font-weight: bold;
                font-size: 13px;
            }
            #PrintBtn:hover {
                background-color: #1e293b;
            }

            #CanvasScrollArea {
                background-color: #f1f5f9;
                border-radius: 12px;
                border: 1px solid #cbd5e1;
            }
            #CanvasScrollArea QWidget#qt_scrollarea_viewport {
                background-color: #f1f5f9;
                border-radius: 11px;
            }
            QScrollBar:vertical {
                border: none;
                background: transparent;
                width: 10px;
                margin: 6px 2px 6px 0px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background: #cbd5e1;
                min-height: 20px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #94a3b8;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
                background: none;
                height: 0px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }

            QTextEdit {
                background-color: #0f172a;
                color: #38bdf8;
                font-family: Consolas, Monospace;
                border-radius: 10px;
                padding: 10px;
                border: 1px solid #1e293b;
            }
        """)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.showMaximized()
    sys.exit(app.exec())