import pdfplumber
import openpyxl
import os
import re

input_folder = r"D:\InvoiceBot\input"
output_file = r"D:\InvoiceBot\output\invoice_report.xlsx"

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Invoices"
ws.append(["Invoice Number", "Vendor", "Date", "Amount"])

def extract_invoice_data(text):
    # INVOICE NUMBER — handles INV-001, #000000005, #9000000001
    inv_no = ""
    patterns = [
        r"Invoice\s*Number[\s:]*([A-Z0-9\-#]+)",
        r"INVOICE\s*NUMBER[\s:#]*([A-Z0-9\-]+)",
        r"INVOICE\s*#([0-9]+)",
        r"Invoice\s*#([A-Z0-9\-]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            inv_no = match.group(1).strip()
            break

    # VENDOR NAME
    vendor = ""
    patterns = [
        r"From:\s*\n([A-Za-z\s&.,\-]+)",
        r"DEMO\s*-\s*([A-Za-z\s]+)",
        r"^([A-Za-z]+Shop|[A-Za-z]+Store|[A-Za-z]+Invoices)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            vendor = match.group(1).strip()
            vendor = re.sub(r"^(Vendor:|From:)\s*", "", vendor, flags=re.IGNORECASE).strip()
            break
    if not vendor:
        # fallback — first meaningful line
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        for line in lines[:5]:
            if not any(x in line.lower() for x in ["invoice", "#", "order", "date"]):
                vendor = line
                break

    # DATE
    date = ""
    patterns = [
        r"Invoice\s*Date[\s:]*([A-Za-z0-9 ,]+)",
        r"Order\s*[Dd]ate[\s:]*([A-Za-z0-9 ,:\s]+?)(?:\n|AM|PM)",
        r"Date[\s:]*(\d{2}-\d{2}-\d{4})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            date = match.group(1).strip()
            break

    # AMOUNT — handles Total Due, Grand Total, GRAND TOTAL, Amount
    amount = ""
    patterns = [
        r"Grand\s*Total[\s:]*\$?([\d,\.]+)",
        r"GRAND\s*TOTAL[\s:]*\$?([\d,\.]+)",
        r"Total\s*Due[\s:]*\$?([\d,\.]+)",
        r"Total[\s:]*\$?([\d,\.]+)",
        r"Amount[\s:]*(\d+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            amount = "$" + match.group(1).strip()
            break

    return inv_no, vendor, date, amount

for filename in os.listdir(input_folder):
    if filename.endswith(".pdf"):
        filepath = os.path.join(input_folder, filename)
        with pdfplumber.open(filepath) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""

            if not text.strip():
                print(f"Skipped (no text): {filename}")
                continue

            inv_no, vendor, date, amount = extract_invoice_data(text)
            ws.append([inv_no, vendor, date, amount])
            print(f"✓ {filename} → {inv_no} | {vendor} | {date} | {amount}")

wb.save(output_file)
print("\nExcel report saved!")