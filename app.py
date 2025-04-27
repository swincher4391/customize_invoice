from flask import Flask, request, send_file, jsonify
from notion_client import Client as NotionClient
import requests
import os
import io
import tempfile
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image, ImageOps

app = Flask(__name__)

# === ENVIRONMENT VARIABLES ===
NOTION_TOKEN = os.getenv("NOTION_TOKEN")
DATABASE_ID = os.getenv("NOTION_DATABASE_ID")
TEMPLATE_PATH = "invoice-watermarked.xlsx"
notion = NotionClient(auth=NOTION_TOKEN)

# === UTILITIES ===
def send_email(recipient_email, subject, body, attachment_path):
    """Send an email with the given attachment."""
    smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
    smtp_port = int(os.getenv('SMTP_PORT', 587))
    smtp_user = os.getenv('SMTP_USER')
    smtp_pass = os.getenv('SMTP_PASS')

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = smtp_user
    msg['To'] = recipient_email
    msg.set_content(body)

    with open(attachment_path, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(attachment_path)

    msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
        
def remove_background(image_path, tolerance=30):
    """Remove background color based on pixel (0,0) and make transparent."""
    img = Image.open(image_path).convert("RGBA")
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
    """Insert processed logo into A1."""
    temp_logo = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    temp_logo.write(image_bytes)
    temp_logo.flush()
    temp_logo.close()

    img = OpenpyxlImage(temp_logo.name)
    img.width = 200  # Resize logo to a clean standard size (can adjust)
    img.height = 200
    img.anchor = 'A1'

    ws.add_image(img)

# === WEBHOOK ENDPOINTS ===

@app.route("/preview_webhook", methods=["POST"])
def handle_preview_request():
    """Process incoming Tally form submission for preview"""
    data = request.json
    print("🚀 Raw incoming data from Tally:", data)   # <-- Add this line!
    # Extract fields
    fields = {field["label"]: field["value"] for field in data.get("data", {}).get("fields", [])}
    business_name = fields.get('Business Name', 'Your Business')
    address1 = fields.get('Business Address Line 1', '')
    address2 = fields.get('Business Address Line 2', '')
    phone = fields.get('Phone Number', '')
    email = fields.get('Email Address', '')
    tax_percentage = fields.get('Tax Percentage', '7')
    currency = fields.get('Currency', 'USD')
    notes = fields.get('Notes (Optional)', '')
    logo_url = fields.get('Upload a Logo')

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

    # Insert business info
    ws['A2'] = business_name
    ws['A3'] = address1
    ws['A4'] = address2
    ws['A5'] = phone
    ws['A6'] = email

    # Insert tax info
    ws['D30'] = f"Tax ({tax_percentage}%)"
    ws['E30'] = f'=IF(NOT(IsGoogleSheets),E29*{float(tax_percentage)}/100,"GOOGLE SHEETS DETECTED")'

    # Insert currency info
    ws['C32'] = f"All amounts shown in {currency}"

    # Insert logo at A1
    insert_logo(ws, logo_bytes.read())

    # Protect workbook
    protect_workbook(wb)

    # Save to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        wb.save(tmp.name)
        tmp_path = tmp.name

    # Send email
    send_email(
        recipient_email=email,
        subject="Your Custom Invoice Preview is Ready!",
        body=f"Hello {business_name},\n\nThank you for your order! Please find your customized invoice attached.\n\nBest regards,\nSwincher Creative",
        attachment_path=tmp_path
    )

    # 🛠 CORRECTLY INDENTED:
    return send_file(tmp_path, as_attachment=True, download_name=f"{business_name}_invoice_preview.xlsx")


if __name__ == "__main__":
    app.run(debug=True)
