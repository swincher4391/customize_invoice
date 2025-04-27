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
import time
from collections import OrderedDict
from datetime import datetime, timezone
from dateutil import parser
import uuid

app = Flask(__name__)

# === ENVIRONMENT VARIABLES ===
NOTION_TOKEN = os.getenv("NOTION_TOKEN")
DATABASE_ID = os.getenv("NOTION_DATABASE_ID")
TEMPLATE_PATH = "invoice-watermarked.xlsx"
notion = NotionClient(auth=NOTION_TOKEN)

# Simple in-memory cache of processed events
# Using OrderedDict to limit memory usage (keeps only most recent 100 events)
PROCESSED_EVENTS = OrderedDict()
MAX_CACHE_SIZE = 100

# === UTILITIES ===
def send_email(recipient_email, subject, body, attachment_paths, business_name=''):
    smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
    smtp_port = int(os.getenv('SMTP_PORT', 587))
    smtp_user = os.getenv('SMTP_USER')
    smtp_pass = os.getenv('SMTP_PASS')
    sender_name = os.getenv('SENDER_NAME', 'Invoice Generator')

    # Create a more sophisticated email
    msg = EmailMessage()
    
    # Better subject line with personalization
    personalized_subject = f"{business_name} - {subject}" if business_name else subject
    msg['Subject'] = personalized_subject
    
    # Add proper From header with sender name
    msg['From'] = f'"{sender_name}" <{smtp_user}>'
    msg['To'] = recipient_email
    
    # Add more headers to improve deliverability
    # Add a unique Message-ID
    domain = smtp_user.split('@')[-1]
    msg['Message-ID'] = f"<{uuid.uuid4()}@{domain}>"
    
    # Add Date header
    msg['Date'] = datetime.now().strftime("%a, %d %b %Y %H:%M:%S %z")
    
    # Add X-Mailer header 
    msg['X-Mailer'] = 'InvoiceCustomizer Service'
    
    # Add a List-Unsubscribe header (helps with spam prevention)
    msg['List-Unsubscribe'] = f'<mailto:{smtp_user}?subject=Unsubscribe>'
    
    # Create personalized HTML content
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>{personalized_subject}</title>
        <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; padding: 20px; }}
            .container {{ max-width: 600px; margin: 0 auto; }}
            .header {{ background-color: #f8f9fa; padding: 15px; border-bottom: 1px solid #e9ecef; }}
            .content {{ padding: 20px 0; }}
            .footer {{ font-size: 12px; color: #6c757d; padding-top: 20px; border-top: 1px solid #e9ecef; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h2>Your Custom Invoice is Ready!</h2>
            </div>
            <div class="content">
                <p>Hello{' ' + business_name if business_name else ''},</p>
                
                <p>Thank you for using our Invoice Generator service! Your custom Excel invoice template is now ready.</p>
                
                <p>We've attached the spreadsheet to this email. You can fill in the item details, and the calculations will be performed automatically.</p>
                
                <p>The invoice includes:</p>
                <ul>
                    <li>Your business information</li>
                    <li>Customized tax rate</li>
                    <li>Your logo</li>
                    <li>Protected formulas to prevent accidental changes</li>
                </ul>
                
                <p>If you have any questions or need assistance, please reply to this email.</p>
            </div>
            <div class="footer">
                <p>This is a transactional email sent to you because you submitted the Invoice Customization Form.</p>
                <p>To unsubscribe from future emails, please reply with "Unsubscribe" in the subject line.</p>
            </div>
        </div>
    </body>
    </html>
    """
    
    # Set both plain text and HTML content
    msg.set_content(body)
    msg.add_alternative(html_content, subtype='html')

    # Add attachments
    for attachment_path in attachment_paths:
        with open(attachment_path, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(attachment_path)

        if attachment_path.endswith('.xlsx'):
            maintype = 'application'
            subtype = 'vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        else:
            maintype = 'application'
            subtype = 'octet-stream'

        msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=file_name)

    # Send the email with a retry mechanism
    max_retries = 3
    retry_delay = 2  # seconds
    
    for attempt in range(max_retries):
        try:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(smtp_user, smtp_pass)
                server.send_message(msg)
                print(f"✅ Email sent successfully to {recipient_email}")
                break
        except Exception as e:
            if attempt < max_retries - 1:
                print(f"⚠️ Email attempt {attempt + 1} failed. Retrying in {retry_delay} seconds: {e}")
                time.sleep(retry_delay)
                retry_delay *= 2  # Exponential backoff
            else:
                print(f"❌ All email attempts failed: {e}")
                raise

def remove_background(image_file, tolerance=50):  # Increased from 30 to 50
    img = Image.open(image_file).convert("RGBA")
    datas = img.getdata()

    new_data = []
    background_color = datas[0]
    print(f"Detected background color: {background_color}")  # Debug output
    
    for item in datas:
        if all(abs(item[i] - background_color[i]) <= tolerance for i in range(3)):
            new_data.append((255, 255, 255, 0))
        else:
            new_data.append(item)
    img.putdata(new_data)
    return img
def update_notion_database(fields, event_id=None):
    """Updates the Notion database with preview information using existing fields"""
    if not NOTION_TOKEN or not DATABASE_ID:
        print("⚠️ Notion credentials not set, skipping database update")
        return
    
    try:
        # Extract relevant information
        business_name = fields.get('Company Name', 'Unknown')
        email = fields.get('Email', '')
        timestamp = datetime.now().isoformat()
        
        # Prepare the data for Notion using only existing fields
        properties = {
            "Name": {"title": [{"text": {"content": business_name}}]},
            "Email": {"email": email},
            "Timestamp": {"date": {"start": timestamp}},
            "Validated": {"checkbox": True},
            "Excel Sent": {"checkbox": False}  # Will be updated to True when email is sent
        }
        
        # If we have a company name, add it to the Company field
        if fields.get('Company Name'):
            properties["Company"] = {"rich_text": [{"text": {"content": fields.get('Company Name', '')}}]}
        
        # Check if this email already exists in the database
        existing_pages = notion.databases.query(
            database_id=DATABASE_ID,
            filter={
                "property": "Email",
                "email": {
                    "equals": email
                }
            }
        ).get("results", [])
        
        if existing_pages:
            # Update existing page
            notion.pages.update(
                page_id=existing_pages[0]["id"],
                properties=properties
            )
            page_id = existing_pages[0]["id"]
            print(f"✅ Updated existing Notion entry for {business_name}")
        else:
            # Create new page
            response = notion.pages.create(
                parent={"database_id": DATABASE_ID},
                properties=properties
            )
            page_id = response["id"]
            print(f"✅ Created new Notion entry for {business_name}")
        
        return page_id
            
    except Exception as e:
        print(f"⚠️ Error updating Notion database: {e}")
        return None
        
def protect_workbook(workbook, password='etsysc123'):
    for sheet in workbook.worksheets:
        # Enable protection with specific options
        sheet.protection.sheet = True
        sheet.protection.password = password
        
        # Disable object editing/deletion
        sheet.protection.objects = True
        sheet.protection.scenarios = True
        
        # Other protection options
        sheet.protection.selectLockedCells = False
        sheet.protection.selectUnlockedCells = False
        sheet.protection.formatCells = False
        sheet.protection.formatColumns = False
        sheet.protection.formatRows = False
        sheet.protection.insertColumns = False
        sheet.protection.insertRows = False
        sheet.protection.insertHyperlinks = False
        sheet.protection.deleteColumns = False
        sheet.protection.deleteRows = False
        sheet.protection.sort = False
        sheet.protection.autoFilter = False
        sheet.protection.pivotTables = False
        sheet.protection.drawings = True  # Specifically protects drawings/images

def insert_logo(ws, image_bytes):
    # Create a temporary file that won't be deleted until program exit
    temp_logo = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    temp_logo.write(image_bytes)
    temp_logo.close()  # Close the file but don't delete it

    try:
        img = OpenpyxlImage(temp_logo.name)
        col_width = ws.column_dimensions['A'].width or 8
        row_height = ws.row_dimensions[1].height or 15

        img.width = int(col_width * 7.5)
        img.height = int(row_height * 1.33)
        img.anchor = 'A1'

        ws.add_image(img)
        
        # Don't delete the temp file here - it's needed later when saving
        # Return the filename so we can delete it after the workbook is saved
        return temp_logo.name
    except Exception as e:
        print(f"⚠️ Error adding logo: {e}")
        # If there's an error, try to clean up the file
        try:
            os.unlink(temp_logo.name)
        except:
            pass
        return None

def insert_watermark_background(ws):
    if os.path.exists("watermark.png"):
        with open("watermark.png", 'rb') as img_file:
            ws._background = img_file.read()

def is_stale_event(event_time, max_age_minutes=5):
    """Check if an event is too old to process"""
    try:
        event_datetime = parser.parse(event_time)
        current_time = datetime.now(timezone.utc)
        time_diff = (current_time - event_datetime).total_seconds() / 60
        return time_diff > max_age_minutes, time_diff
    except Exception as e:
        print(f"⚠️ Error parsing event time: {e}")
        return False, 0  # If we can't parse the time, assume it's not stale

# === WEBHOOK ===
# Fix the indentation in the webhook handler function
@app.route("/preview_webhook", methods=["POST"])
def handle_preview_request():
    temp_files = []  # List to keep track of temporary files to clean up
    event_id = None
    
    try:
        data = request.json
        print("🚀 Raw incoming data from Tally:", data)

        # Extract event ID and creation time for idempotency check
        event_id = data.get('eventId')
        event_time = data.get('createdAt')

        # Check if we've already processed this event SUCCESSFULLY
        if event_id in PROCESSED_EVENTS and PROCESSED_EVENTS[event_id].get("processed", False):
            print(f"✅ Skipping already processed event: {event_id}")
            return '', 204
        
        # If we received it before but didn't process it successfully, we'll try again
        if event_id in PROCESSED_EVENTS:
            print(f"⚠️ Retrying previously failed event: {event_id}")
        
        # Check if the event is too old (stale)
        is_stale, time_diff = is_stale_event(event_time)
        if is_stale:
            print(f"⚠️ Skipping stale event from {time_diff:.1f} minutes ago: {event_id}")
            
            # Mark as received but not successfully processed
            if len(PROCESSED_EVENTS) >= MAX_CACHE_SIZE:
                PROCESSED_EVENTS.popitem(last=False)
            PROCESSED_EVENTS[event_id] = {"timestamp": time.time(), "processed": False}
            
            return '', 204

        # Extract fields from the webhook payload
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

        # Insert logo and get the temp file path
        logo_temp_path = insert_logo(ws, logo_bytes.read())
        if logo_temp_path:
            temp_files.append(logo_temp_path)

        protect_workbook(wb)

        # Save the workbook to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            wb.save(tmp.name)
            temp_files.append(tmp.name)  # Add to cleanup list

        notion_page_id = update_notion_database(fields=fields)

        # Send the email
        try:
            send_email(
                recipient_email=email,
                subject="Your Custom Invoice is Ready!",
                body="Please find attached your custom Excel invoice template. You can fill in the item details and the calculations will be performed automatically.",
                attachment_paths=[tmp.name],
                business_name=business_name  # Pass the business name for personalization
            )
            
            # Mark this event as successfully processed
            if len(PROCESSED_EVENTS) >= MAX_CACHE_SIZE:
                PROCESSED_EVENTS.popitem(last=False)
            PROCESSED_EVENTS[event_id] = {"timestamp": time.time(), "processed": True}
            
            print(f"✅ Successfully processed event: {event_id}")
            # Then after successfully sending the email, update the Excel Sent field
            if notion_page_id:
                try:
                    notion.pages.update(
                        page_id=notion_page_id,
                        properties={
                            "Excel Sent": {"checkbox": True}
                        }
                    )
                    print(f"✅ Updated Excel Sent status in Notion")
                except Exception as e:
                    print(f"⚠️ Error updating Excel Sent status: {e}")
        except Exception as e:
            # Mark as received but not successfully processed
            if len(PROCESSED_EVENTS) >= MAX_CACHE_SIZE:
                PROCESSED_EVENTS.popitem(last=False)
            PROCESSED_EVENTS[event_id] = {"timestamp": time.time(), "processed": False}
            
            print(f"⚠️ Failed to send email: {e}")
            return jsonify({"error": "Failed to send email"}), 500

        return '', 204
        
    except Exception as e:
        # Mark the event as received but not processed if we have an event_id
        if event_id:
            if len(PROCESSED_EVENTS) >= MAX_CACHE_SIZE:
                PROCESSED_EVENTS.popitem(last=False)
            PROCESSED_EVENTS[event_id] = {"timestamp": time.time(), "processed": False}
        
        print(f"⚠️ Unhandled error: {e}")
        return jsonify({"error": str(e)}), 500
        
    finally:
        # Clean up all temporary files
        for file_path in temp_files:
            try:
                if file_path and os.path.exists(file_path):
                    os.unlink(file_path)
                    print(f"✅ Removed temporary file: {file_path}")
            except Exception as e:
                print(f"⚠️ Error removing temporary file {file_path}: {e}")

@app.route("/health", methods=["GET"])
def health_check():
    """Simple health check endpoint"""
    return jsonify({
        "status": "ok",
        "processed_events": len(PROCESSED_EVENTS),
        "processed_details": {k: v.get("processed", False) for k, v in PROCESSED_EVENTS.items()},
        "timestamp": datetime.now().isoformat()
    })

if __name__ == "__main__":
    app.run(debug=True)
