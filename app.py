from flask import Flask, request, send_file, jsonify
from notion_client import Client as NotionClient
import requests
import os
import io
import tempfile
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment
import smtplib
from email.message import EmailMessage
from PIL import Image
import pdfkit
import shutil

app = Flask(__name__)

# === ENVIRONMENT VARIABLES ===
NOTION_TOKEN = os.getenv("NOTION_TOKEN")
DATABASE_ID = os.getenv("NOTION_DATABASE_ID")
TEMPLATE_PATH = "invoice-watermarked.xlsx"
notion = NotionClient(auth=NOTION_TOKEN)

# === UTILITIES ===
def send_email(recipient_email, subject, body, attachment_paths):
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

        if attachment_path.endswith('.xlsx'):
            maintype, subtype = 'application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        elif attachment_path.endswith('.pdf'):
            maintype, subtype = 'application', 'pdf'
        else:
            maintype, subtype = 'application', 'octet-stream'

        msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=file_name)

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)

def remove_background(image_file, tolerance=30):
    img = Image.open(image_file).convert("RGBA")
    datas = img.getdata()

    new_data = []
    background_color = datas[0]
    for item in datas:
        if all(abs(item[i] - background_color[i]) <= tolerance for i in range(3)):
            new_data.append((255, 255, 255, 0))
        else:
            new_data.append(item)
    img.putdata(new_data)
    return img

def protect_workbook(workbook, password='etsysc123'):
    for sheet in workbook.worksheets:
        sheet.protection.sheet = True
        sheet.protection.password = password

def insert_logo(ws, image_bytes):
    temp_logo = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    temp_logo.write(image_bytes)
    temp_logo.flush()
    temp_logo.close()

    img = OpenpyxlImage(temp_logo.name)
    col_width = ws.column_dimensions['A'].width or 8
    row_height = ws.row_dimensions[1].height or 15

    img.width = int(col_width * 7.5)
    img.height = int(row_height * 1.33)
    img.anchor = 'A1'

    ws.add_image(img)

def insert_watermark_background(ws):
    if os.path.exists("watermark.png"):
        with open("watermark.png", 'rb') as img_file:
            ws._background = img_file.read()

# === WEBHOOK ===
@app.route("/preview_webhook", methods=["POST"])
def handle_preview_request():
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

    logo_response = requests.get(logo_url)
    if logo_response.status_code != 200:
        return jsonify({"error": "Failed to download logo"}), 400

    processed_logo = remove_background(io.BytesIO(logo_response.content))
    logo_bytes = io.BytesIO()
    processed_logo.save(logo_bytes, format="PNG")
    logo_bytes.seek(0)

    # Save logo to a file for later use in the PDF
    logo_path = tempfile.mktemp(suffix=".png")
    with open(logo_path, 'wb') as f:
        f.write(logo_bytes.getvalue())
    
    logo_bytes.seek(0)  # Reset position for Excel use

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active

    insert_watermark_background(ws)

    ws['A2'] = business_name
    ws['A3'] = address1
    ws['A4'] = address2
    ws['A5'] = phone
    ws['A6'] = email
    ws['D30'] = f"Tax ({tax_percentage}%)"
    ws['E30'] = f'=IF(NOT(IsGoogleSheets),E29*{float(tax_percentage)}/100,"GOOGLE SHEETS DETECTED")'
    ws['C32'] = f"All amounts shown in {currency}"
    ws.merge_cells('C32:E32')
    ws['C32'].alignment = Alignment(horizontal='center', vertical='center')

    insert_logo(ws, logo_bytes.read())
    protect_workbook(wb)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        wb.save(tmp.name)
        tmp_path = tmp.name

    # Create HTML file directly
    html_file_path = tmp_path.replace(".xlsx", ".html")
    with open(html_file_path, "w", encoding="utf-8") as f:
        f.write(f"""<!DOCTYPE html>
        <html>
        <head>
            <title>Invoice for {business_name}</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 30px; }}
                .header {{ display: flex; justify-content: space-between; }}
                .company-info {{ margin-bottom: 30px; }}
                table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
                table, th, td {{ border: 1px solid #ddd; }}
                th, td {{ padding: 8px; text-align: left; }}
                .totals {{ text-align: right; margin-top: 20px; }}
                .currency {{ text-align: center; margin-top: 30px; font-style: italic; }}
                .logo {{ position: absolute; top: 50px; right: 50px; max-width: 150px; max-height: 100px; }}
            </style>
        </head>
        <body>
            <img src="file://{logo_path}" class="logo">
            <div class="header">
                <h1>INVOICE</h1>
            </div>
            <div class="company-info">
                <p><strong>{business_name}</strong><br/>
                {address1}<br/>
                {address2}<br/>
                {phone}<br/>
                {email}</p>
            </div>
            <table>
                <tr>
                    <th>Item</th>
                    <th>Description</th>
                    <th>Quantity</th>
                    <th>Unit Price</th>
                    <th>Amount</th>
                </tr>
                <tr>
                    <td colspan="5" style="text-align: center;">Items will be filled in the Excel document</td>
                </tr>
            </table>
            <div class="totals">
                <p>Subtotal: ___________</p>
                <p>Tax ({tax_percentage}%): ___________</p>
                <p><strong>Total: ___________</strong></p>
            </div>
            <div class="currency">
                All amounts shown in {currency}
            </div>
        </body>
        </html>
        """)

    # Generate PDF with pdfkit (logo already in HTML)
    pdf_path = tmp_path.replace('.xlsx', '.pdf')
    try:
        pdfkit.from_file(html_file_path, pdf_path)
    except Exception as e:
        print(f"⚠️ Error generating PDF: {e}")
        # Create a simple PDF if conversion fails
        with open(pdf_path, "w") as f:
            f.write(f"Invoice for {business_name}")

    try:
        send_email(
            recipient_email=email,
            subject="Your Custom Invoice is Ready!",
            body="Please find attached both the Excel and PDF versions.",
            attachment_paths=[tmp_path, pdf_path]
        )
    except Exception as e:
        print(f"⚠️ Failed to send email: {e}")

    # Clean up temporary files
    try:
        os.unlink(logo_path)
        os.unlink(html_file_path)
    except Exception as e:
        print(f"⚠️ Error cleaning up temporary files: {e}")

    return '', 204

if __name__ == "__main__":
    app.run(debug=True)
