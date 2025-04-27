from flask import Flask, request, send_file, jsonify
from notion_client import Client as NotionClient
import requests
import os
import io
import tempfile
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font, Alignment
import smtplib
from email.message import EmailMessage
from PIL import Image
import pdfkit
from xlsx2html import xlsx2html
import fitz

app = Flask(__name__)

# === ENVIRONMENT VARIABLES ===
NOTION_TOKEN = os.getenv("NOTION_TOKEN")
DATABASE_ID = os.getenv("NOTION_DATABASE_ID")
TEMPLATE_PATH = "invoice-watermarked.xlsx"
notion = NotionClient(auth=NOTION_TOKEN)

# === UTILITIES ===
def send_email(recipient_email, subject, body, attachment_paths):
    """Send an email with one or multiple attachments."""
    smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
    smtp_port = int(os.getenv('SMTP_PORT', 587))
    smtp_user = os.getenv('SMTP_USER')
    smtp_pass = os.getenv('SMTP_PASS')

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = smtp_user
    msg['To'] = recipient_email
    msg.set_content(body)

    for attachment_path in attachment_paths:
        with open(attachment_path, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(attachment_path)

        # Decide attachment MIME type
        if attachment_path.endswith('.xlsx'):
            maintype = 'application'
            subtype = 'vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        elif attachment_path.endswith('.pdf'):
            maintype = 'application'
            subtype = 'pdf'
        else:
            maintype = 'application'
            subtype = 'octet-stream'

        msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=file_name)

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)


def remove_background(image_file, tolerance=30):
    """Remove background color based on pixel (0,0) and make transparent."""
    img = Image.open(image_file).convert("RGBA")
    datas = img.getdata()

    new_data = []
    background_color = datas[0]  # top-left pixel
    for item in datas:
        if all(abs(item[i] - background_color[i]) <= tolerance for i in range(3)):
            new_data.append((255, 255, 255, 0))  # Transparent pixel
        else:
            new_data.append(item)
    img.putdata(new_data)
    return img

def protect_workbook(workbook, password='etsysc123'):
    """Protect all sheets in workbook."""
    for sheet in workbook.worksheets:
        sheet.protection.sheet = True
        sheet.protection.password = password

def insert_logo(ws, image_bytes):
    """Insert processed logo into cell A1, snapping exactly to A1 size."""
    temp_logo = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    temp_logo.write(image_bytes)
    temp_logo.flush()
    temp_logo.close()

    img = OpenpyxlImage(temp_logo.name)

    # Read A1 dimensions
    col_width = ws.column_dimensions['A'].width or 8
    row_height = ws.row_dimensions[1].height or 15

    col_pixels = int(col_width * 7.5)
    row_pixels = int(row_height * 1.33)

    img.width = col_pixels
    img.height = row_pixels

    img.anchor = 'A1'

    ws.add_image(img)

def insert_watermark_background(ws):
    """Insert watermark.png as a background image (undeletable by users)."""
    if os.path.exists("watermark.png"):
        with open("watermark.png", 'rb') as img_file:
            ws._background = img_file.read()

def insert_logo_into_pdf(original_pdf_path, logo_bytes, output_pdf_path):
    doc = fitz.open(original_pdf_path)

    # Load logo as image
    logo_img = fitz.Pixmap(fitz.open("png", logo_bytes))

    page = doc[0]  # First page

    # Insert logo at (x, y) position you want
    rect = fitz.Rect(50, 700, 250, 800)  # (left, top, right, bottom) coordinates
    page.insert_image(rect, pixmap=logo_img)

    # Save modified PDF
    doc.save(output_pdf_path)
    
# === WEBHOOK ENDPOINTS ===
@app.route("/preview_webhook", methods=["POST"])
def handle_preview_request():
    """Process incoming Tally form submission for preview"""
    data = request.json
    print("🚀 Raw incoming data from Tally:", data)

    fields_list = data.get('data', {}).get('fields', [])
    fields = {}

    for field in fields_list:
        label = field.get('label')
        value = field.get('value')

        if isinstance(value, list):
            if value and isinstance(value[0], dict) and 'url' in value[0]:
                value = value[0]['url']
            else:
                value = None

        fields[label] = value

    business_name = fields.get('Company Name', 'Your Business')
    address1 = fields.get('Address', '')
    address2 = fields.get('City, State ZIP', '')
    phone = fields.get('Phone', '')
    email = fields.get('Email', '')
    tax_percentage = fields.get('Tax %', '7')
    currency = fields.get('Currency', 'USD')
    logo_url = fields.get('Upload your logo', '')

    if not logo_url:
        return jsonify({"error": "Logo upload missing!"}), 400

    # Download and process logo
    logo_response = requests.get(logo_url)
    if logo_response.status_code != 200:
        return jsonify({"error": "Failed to download logo"}), 400

    processed_logo = remove_background(io.BytesIO(logo_response.content))

    # Save processed logo to bytes
    logo_bytes = io.BytesIO()
    processed_logo.save(logo_bytes, format="PNG")
    logo_bytes.seek(0)

    # Load template
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # Insert watermark as background
    insert_watermark_background(ws)

    # Insert business info
    ws['A2'] = business_name
    ws['A3'] = address1
    ws['A4'] = address2
    ws['A5'] = phone
    ws['A6'] = email

    # Insert tax info
    ws['D30'] = f"Tax ({tax_percentage}%)"
    ws['E30'] = f'=IF(NOT(IsGoogleSheets),E29*{float(tax_percentage)}/100,"GOOGLE SHEETS DETECTED")'

    # Insert currency info, then merge and center
    ws['C32'] = f"All amounts shown in {currency}"
    ws.merge_cells('C32:E32')
    ws['C32'].alignment = Alignment(horizontal='center', vertical='center')

    # Insert logo at A1
    insert_logo(ws, logo_bytes.read())

    # Protect workbook
    protect_workbook(wb)

    # Save to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        wb.save(tmp.name)
        tmp_path = tmp.name

    # Convert .xlsx to .html
    html_file_path = tmp_path.replace(".xlsx", ".html")
    with open(html_file_path, "w", encoding="utf-8") as f:
        xlsx2html(tmp_path, f)

    # Convert .html to .pdf
    pdf_file_path = tmp_path.replace(".xlsx", ".pdf")
    pdfkit.from_file(html_file_path, pdf_file_path)

    #try to insert the missing logo
    try:
        insert_logo_into_pdf(
        original_pdf_path=pdf_path_no_logo,
        logo_bytes=logo_bytes.getvalue(),   # <-- logo in bytes
        output_pdf_path=final_pdf_path
        )
    except Exception as e:
        print('failed to insert the logo')
    # Send email (even if email fails, still return file)
    try:
        send_email(
            recipient_email=email,
            subject="Your Custom Invoice is Ready!",
            body="Please find attached both the Excel and PDF versions.",
            attachment_paths=[tmp_path, pdf_file_path]  # <-- multiple attachments!
        )
    except Exception as e:
        print(f"⚠️ Failed to send email: {e}")


    return '',204

if __name__ == "__main__":
    app.run(debug=True)
