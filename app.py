from flask import Flask, request, send_file, jsonify
from notion_client import Client as NotionClient
import requests
from requests_oauthlib import OAuth1
import os
import io
import json
import tempfile
import uuid
import random
import string
import time
import logging
from datetime import datetime, timezone, timedelta
from dateutil import parser
from collections import OrderedDict
from apscheduler.schedulers.background import BackgroundScheduler
from retry import retry

# Image processing
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment
from PIL import Image

# Email
import smtplib
from email.message import EmailMessage

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Initialize Flask app
app = Flask(__name__)

# === ENVIRONMENT VARIABLES ===
# Load environment variables if .env file exists
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Notion credentials
NOTION_TOKEN = os.getenv("NOTION_TOKEN")
BRANDID_DATABASE_ID = os.getenv("NOTION_DATABASE_ID")
ORDERS_DATABASE_ID = os.getenv("NOTION_ORDERS_DATABASE_ID")

# Etsy API credentials
ETSY_API_KEY = os.getenv("ETSY_API_KEY")
ETSY_API_SECRET = os.getenv("ETSY_API_SECRET")
ETSY_ACCESS_TOKEN = os.getenv("ETSY_ACCESS_TOKEN")
ETSY_ACCESS_TOKEN_SECRET = os.getenv("ETSY_ACCESS_TOKEN_SECRET")
ETSY_SHOP_ID = os.getenv("ETSY_SHOP_ID")

# Email configuration
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))
SMTP_USER = os.getenv('SMTP_USER')
SMTP_PASS = os.getenv('SMTP_PASS')
SENDER_NAME = os.getenv('SENDER_NAME', 'Invoice Generator')

# Template configuration
TEMPLATE_PATH = os.getenv('TEMPLATE_PATH', 'invoice-watermarked.xlsx')
LICENSE_TEMPLATE_PATH = os.getenv('LICENSE_TEMPLATE_PATH', 'invoice-licensed.xlsx')
WATERMARK_PATH = os.getenv('WATERMARK_PATH', 'watermark.png')

# Initialize clients
notion = NotionClient(auth=NOTION_TOKEN) if NOTION_TOKEN else None

# Simple in-memory cache of processed events
# Using OrderedDict to limit memory usage (keeps only most recent events)
PROCESSED_EVENTS = OrderedDict()
MAX_CACHE_SIZE = 100

# === UTILITY FUNCTIONS ===

def generate_brand_id():
    """Generate a guaranteed unique Brand ID"""
    max_attempts = 10  # Prevent infinite loops
    
    for attempt in range(max_attempts):
        # Generate a candidate ID
        prefix = "BRAND"
        random_component = ''.join(random.choices(string.digits, k=2))
        uuid_component = str(uuid.uuid4())[:4]
        candidate_id = f"{prefix}-{random_component}-{uuid_component}"
        
        # Check if this ID already exists in Notion
        if notion and BRANDID_DATABASE_ID:
            existing = notion.databases.query(
                database_id=BRANDID_DATABASE_ID,
                filter={
                    "property": "BrandID",
                    "rich_text": {
                        "equals": candidate_id
                    }
                }
            ).get("results", [])
            
            # If no matching records found, this ID is unique
            if not existing:
                logger.info(f"Generated unique Brand ID: {candidate_id}")
                return candidate_id
            else:
                logger.warning(f"Brand ID collision detected: {candidate_id}, retrying...")
        else:
            # If Notion isn't configured, we can't check, so just return the ID
            return candidate_id
    
    # If we get here, we've hit max attempts without finding a unique ID
    # Use a longer UUID component to make collision practically impossible
    prefix = "BRAND"
    random_component = ''.join(random.choices(string.digits, k=2))
    uuid_component = str(uuid.uuid4())[:8]  # Using 8 chars instead of 4
    fallback_id = f"{prefix}-{random_component}-{uuid_component}"
    logger.warning(f"Used fallback method for Brand ID generation: {fallback_id}")
    return fallback_id

@retry(tries=3, delay=2, backoff=2)
def send_email(recipient_email, subject, body, attachment_paths=None, business_name='', template_type='invoice'):
    """
    Send email with attachments and HTML formatting
    
    Args:
        recipient_email: Email address of recipient
        subject: Email subject
        body: Plain text body (fallback)
        attachment_paths: List of file paths to attach
        business_name: Customer's business name for personalization
        template_type: Type of template (invoice, receipt, etc.)
    """
    if attachment_paths is None:
        attachment_paths = []

    # Create email
    msg = EmailMessage()
    
    # Personalized subject line
    personalized_subject = f"{business_name} - {subject}" if business_name else subject
    msg['Subject'] = personalized_subject
    
    # Add proper From header with sender name
    msg['From'] = f'"{SENDER_NAME}" <{SMTP_USER}>'
    msg['To'] = recipient_email
    
    # Add headers to improve deliverability
    domain = SMTP_USER.split('@')[-1]
    msg['Message-ID'] = f"<{uuid.uuid4()}@{domain}>"
    msg['Date'] = datetime.now().strftime("%a, %d %b %Y %H:%M:%S %z")
    msg['X-Mailer'] = 'TemplateCustomizer Service'
    msg['List-Unsubscribe'] = f'<mailto:{SMTP_USER}?subject=Unsubscribe>'
    
    # Create HTML content
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
            .button {{ display: inline-block; background-color: #007bff; color: white; padding: 10px 20px; 
                      text-decoration: none; border-radius: 4px; margin-top: 15px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h2>Your Custom {template_type.title()} is Ready!</h2>
            </div>
            <div class="content">
                <p>Hello{' ' + business_name if business_name else ''},</p>
                
                <p>Thank you for using our Template Generator service! Your custom {template_type} template is now ready.</p>
                
                <p>We've attached the spreadsheet to this email. You can fill in the item details, and the calculations will be performed automatically.</p>
                
                <p>The template includes:</p>
                <ul>
                    <li>Your business information</li>
                    <li>Customized tax rate</li>
                    <li>Your logo</li>
                    <li>Protected formulas to prevent accidental changes</li>
                </ul>
                
                <p>If you have any questions or need assistance, please reply to this email.</p>
            </div>
            <div class="footer">
                <p>This is a transactional email sent to you because you submitted the Template Customization Form.</p>
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

    # Send the email with retry mechanism
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)
            logger.info(f"✅ Email sent successfully to {recipient_email}")
    except Exception as e:
        logger.error(f"❌ Email sending failed: {e}")
        raise

def remove_background(image_file, tolerance=50):
    """
    Remove background from logo image
    
    Args:
        image_file: PIL Image or file-like object
        tolerance: Color difference tolerance (higher = more aggressive)
        
    Returns:
        PIL Image with transparent background
    """
    img = Image.open(image_file).convert("RGBA")
    datas = img.getdata()

    new_data = []
    background_color = datas[0]
    logger.debug(f"Detected background color: {background_color}")
    
    for item in datas:
        if all(abs(item[i] - background_color[i]) <= tolerance for i in range(3)):
            new_data.append((255, 255, 255, 0))
        else:
            new_data.append(item)
    img.putdata(new_data)
    return img

# Removed S3 upload function as we're using Tally URLs directly

def protect_workbook(workbook, password='etsy123'):
    """
    Apply protection to Excel workbook to prevent formula tampering
    
    Args:
        workbook: openpyxl Workbook object
        password: Protection password
    """
    for sheet in workbook.worksheets:
        # Enable protection with specific options
        sheet.protection.sheet = True
        sheet.protection.password = password
        
        # Disable object editing/deletion
        sheet.protection.objects = True
        sheet.protection.scenarios = True
        
        # Other protection options
        sheet.protection.selectLockedCells = False
        sheet.protection.selectUnlockedCells = True
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
    """
    Insert logo into Excel worksheet
    
    Args:
        ws: openpyxl Worksheet object
        image_bytes: Bytes of the image
        
    Returns:
        Path to temporary logo file
    """
    # Create a temporary file
    temp_logo = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    temp_logo.write(image_bytes)
    temp_logo.close()

    try:
        img = OpenpyxlImage(temp_logo.name)
        col_width = ws.column_dimensions['A'].width or 8
        row_height = ws.row_dimensions[1].height or 15

        img.width = int(col_width * 7.5)
        img.height = int(row_height * 1.33)
        img.anchor = 'A1'

        ws.add_image(img)
        return temp_logo.name
    except Exception as e:
        logger.error(f"⚠️ Error adding logo: {e}")
        try:
            os.unlink(temp_logo.name)
        except:
            pass
        return None

def insert_watermark_background(ws):
    """Add watermark to worksheet background"""
    if os.path.exists(WATERMARK_PATH):
        with open(WATERMARK_PATH, 'rb') as img_file:
            ws._background = img_file.read()

def is_stale_event(event_time, max_age_minutes=5):
    """
    Check if an event is too old to process
    
    Args:
        event_time: ISO format timestamp
        max_age_minutes: Maximum age in minutes
        
    Returns:
        (is_stale, time_diff_minutes) tuple
    """
    try:
        event_datetime = parser.parse(event_time)
        current_time = datetime.now(timezone.utc)
        time_diff = (current_time - event_datetime).total_seconds() / 60
        return time_diff > max_age_minutes, time_diff
    except Exception as e:
        logger.error(f"⚠️ Error parsing event time: {e}")
        return False, 0

def get_brand_id_by_email(email):
    """
    Find Brand ID by customer email
    
    Args:
        email: Customer email address
        
    Returns:
        Brand ID or None if not found
    """
    if not notion or not BRANDID_DATABASE_ID:
        logger.warning("⚠️ Notion not configured, can't lookup Brand ID")
        return None
        
    try:
        pages = notion.databases.query(
            database_id=BRANDID_DATABASE_ID,
            filter={
                "property": "Email",
                "email": {
                    "equals": email
                }
            }
        ).get("results", [])
        
        if pages:
            # Get the Brand ID from the first matching page
            brand_id = pages[0]["properties"]["BrandID"]["rich_text"][0]["text"]["content"]
            logger.info(f"✅ Found Brand ID {brand_id} for email {email}")
            return brand_id
        else:
            logger.info(f"❓ No Brand ID found for email {email}")
            return None
    except Exception as e:
        logger.error(f"❌ Error looking up Brand ID: {e}")
        return None

def get_brand_details(brand_id):
    """
    Get all brand details by Brand ID
    
    Args:
        brand_id: The Brand ID to look up
        
    Returns:
        Dictionary of brand details or None if not found
    """
    if not notion or not BRANDID_DATABASE_ID:
        logger.warning("⚠️ Notion not configured, can't get brand details")
        return None
        
    try:
        pages = notion.databases.query(
            database_id=BRANDID_DATABASE_ID,
            filter={
                "property": "BrandID",
                "rich_text": {
                    "equals": brand_id
                }
            }
        ).get("results", [])
        
        if not pages:
            logger.warning(f"⚠️ No brand found with ID {brand_id}")
            return None
            
        page = pages[0]
        properties = page["properties"]
        
        # Extract all relevant fields
        details = {
            "brand_id": brand_id,
            "company_name": properties.get("Company", {}).get("title", [{}])[0].get("text", {}).get("content", ""),
            "email": properties.get("Email", {}).get("email", ""),
            "logo_url": properties.get("LogoURL", {}).get("url", ""),
            "phone": properties.get("Phone", {}).get("phone_number", ""),
            "address": properties.get("Address", {}).get("rich_text", [{}])[0].get("text", {}).get("content", ""),
            "city_state_zip": properties.get("CityStateZip", {}).get("rich_text", [{}])[0].get("text", {}).get("content", "")
        }
        
        logger.info(f"✅ Retrieved details for Brand ID {brand_id}")
        return details
    except Exception as e:
        logger.error(f"❌ Error retrieving brand details: {e}")
        return None

def create_notion_order(order_details):
    """
    Create a new order entry in the Orders database
    
    Args:
        order_details: Dictionary with order information
        
    Returns:
        ID of created page or None on failure
    """
    if not notion or not ORDERS_DATABASE_ID:
        logger.warning("⚠️ Notion not configured, can't create order")
        return None
        
    try:
        # Prepare the data for Notion
        properties = {
            "Etsy Email Address": {"title": [{"text": {"content": order_details["email"]}}]},
            "Etsy Order #": {"rich_text": [{"text": {"content": str(order_details["order_id"])}}]},
            "Timestamp": {"date": {"start": datetime.now().isoformat()}},
            "Email Sent": {"checkbox": False}
        }
        
        # Add Brand ID if available
        if order_details.get("brand_id"):
            # This assumes BrandId is a relation property in your Notion database
            # You might need to adjust this based on your actual database structure
            properties["BrandId"] = {
                "relation": [
                    {"id": order_details["brand_id"]}
                ]
            }
        
        # Create page
        response = notion.pages.create(
            parent={"database_id": ORDERS_DATABASE_ID},
            properties=properties
        )
        
        logger.info(f"✅ Created Notion order for Etsy Order #{order_details['order_id']}")
        return response["id"]
    except Exception as e:
        logger.error(f"❌ Error creating Notion order: {e}")
        return None

def update_notion_database(fields, event_id=None):
    """
    Create or update a brand entry in the Brand ID database
    
    Args:
        fields: Dictionary of form fields
        event_id: Optional event ID for logging
        
    Returns:
        Page ID of created/updated page or None on failure
    """
    if not notion or not BRANDID_DATABASE_ID:
        logger.warning("⚠️ Notion credentials not set, skipping database update")
        return None
    
    try:
        # Extract relevant information
        business_name = fields.get('Company Name', 'Unknown')
        email = fields.get('Email', '')
        timestamp = datetime.now().isoformat()
        
        # Generate a Brand ID if this is a new entry
        brand_id = fields.get('BrandID') or generate_brand_id()
        
        # Prepare the data for Notion
        properties = {
            "Company": {"title": [{"text": {"content": business_name}}]},
            "Email": {"email": email},
            "Timestamp": {"date": {"start": timestamp}},
            "Email Sent": {"checkbox": False},
            "BrandID": {"rich_text": [{"text": {"content": brand_id}}]}
        }
        
        # Add other fields if available
        if fields.get('LogoURL'):
            properties["LogoURL"] = {"url": fields.get('LogoURL')}
            
        if fields.get('Phone'):
            properties["Phone"] = {"phone_number": fields.get('Phone')}
            
        if fields.get('Address'):
            properties["Address"] = {"rich_text": [{"text": {"content": fields.get('Address')}}]}
            
        if fields.get('City, State ZIP'):
            properties["CityStateZip"] = {"rich_text": [{"text": {"content": fields.get('City, State ZIP')}}]}
        
        # Check if this email already exists in the database
        existing_pages = notion.databases.query(
            database_id=BRANDID_DATABASE_ID,
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
            logger.info(f"✅ Updated existing Notion entry for {business_name}")
        else:
            # Create new page
            response = notion.pages.create(
                parent={"database_id": BRANDID_DATABASE_ID},
                properties=properties
            )
            page_id = response["id"]
            logger.info(f"✅ Created new Notion entry for {business_name}")
        
        return page_id
            
    except Exception as e:
        logger.error(f"⚠️ Error updating Notion database: {e}")
        return None

def update_order_email_sent(page_id):
    """Mark an order as having email sent in Notion"""
    if not notion:
        logger.warning("⚠️ Notion not configured, can't update order")
        return
        
    try:
        notion.pages.update(
            page_id=page_id,
            properties={
                "Email Sent": {"checkbox": True}
            }
        )
        logger.info(f"✅ Updated order email status")
    except Exception as e:
        logger.error(f"❌ Error updating order email status: {e}")
        
def generate_template(business_name, address1, address2, phone, email, tax_percentage, 
                     currency, logo_bytes, template_path, add_watermark=True):
    """
    Generate a customized Excel template
    
    Args:
        business_name: Company name
        address1: Street address
        address2: City, state, ZIP
        phone: Phone number
        email: Email address
        tax_percentage: Tax rate as string
        currency: Currency code
        logo_bytes: Logo image bytes
        template_path: Path to base template
        add_watermark: Whether to add watermark
        
    Returns:
        (workbook, temp_files) tuple
    """
    temp_files = []  # List to track temp files for cleanup
    
    # Create a BytesIO object from logo bytes
    logo_bytes_io = io.BytesIO(logo_bytes)
    
    # Process logo to remove background
    processed_logo = remove_background(logo_bytes_io)
    
    # Convert processed logo back to bytes
    logo_bytes_processed = io.BytesIO()
    processed_logo.save(logo_bytes_processed, format="PNG")
    logo_bytes_processed.seek(0)
    
    # Load template
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    
    # Add watermark if requested
    if add_watermark:
        insert_watermark_background(ws)
    
    # Insert business information
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
    
    # Insert logo
    logo_temp_path = insert_logo(ws, logo_bytes_processed.read())
    if logo_temp_path:
        temp_files.append(logo_temp_path)
    
    # Apply protection
    protect_workbook(wb)
    
    return wb, temp_files

# === ETSY API FUNCTIONS ===

def get_etsy_auth():
    """Create OAuth1 auth object for Etsy API"""
    return OAuth1(
        ETSY_API_KEY,
        client_secret=ETSY_API_SECRET,
        resource_owner_key=ETSY_ACCESS_TOKEN,
        resource_owner_secret=ETSY_ACCESS_TOKEN_SECRET
    )

def get_recent_orders():
    """
    Poll Etsy API for recent orders
    
    Returns:
        List of order dictionaries
    """
    if not all([ETSY_API_KEY, ETSY_API_SECRET, ETSY_ACCESS_TOKEN, ETSY_ACCESS_TOKEN_SECRET, ETSY_SHOP_ID]):
        logger.warning("⚠️ Etsy API credentials not configured")
        return []
    
    # Get yesterday's date for the min_created parameter
    yesterday = datetime.now() - timedelta(days=1)
    min_created = int(yesterday.timestamp())
    
    url = f"https://openapi.etsy.com/v2/shops/{ETSY_SHOP_ID}/receipts"
    params = {
        "min_created": min_created,
        "limit": 100,
        "includes": "Transactions,Country"
    }
    
    try:
        auth = get_etsy_auth()
        response = requests.get(url, params=params, auth=auth)
        
        if response.status_code != 200:
            logger.error(f"❌ Etsy API error: {response.status_code} - {response.text}")
            return []
            
        data = response.json()
        return data.get("results", [])
    except Exception as e:
        logger.error(f"❌ Error fetching Etsy orders: {e}")
        return []

def extract_brand_id_from_notes(message):
    """
    Extract Brand ID from order notes using regex
    
    Args:
        message: Message from buyer
        
    Returns:
        Brand ID or None if not found
    """
    import re
    if not message:
        return None
        
    pattern = r"BrandID:\s*([A-Z0-9-]+)"
    match = re.search(pattern, message, re.IGNORECASE)
    
    if match:
        return match.group(1)
    return None
    
def process_etsy_orders():
    """
    Check for new Etsy orders and process them
    """
    logger.info("🔍 Checking for new Etsy orders...")
    
    # Get recent orders
    orders = get_recent_orders()
    
    if not orders:
        logger.info("✅ No new orders found")
        return
        
    logger.info(f"🚀 Processing {len(orders)} orders")
    
    for order in orders:
        # Skip if we've already processed this order
        order_id = order.get("receipt_id")
        if order_id in PROCESSED_EVENTS and PROCESSED_EVENTS[order_id].get("processed", False):
            logger.info(f"✅ Skipping already processed order: {order_id}")
            continue
            
        # Extract buyer information
        buyer_email = order.get("buyer_email")
        message_from_buyer = order.get("message_from_buyer", "")
        
        # Try to extract Brand ID from message
        brand_id = extract_brand_id_from_notes(message_from_buyer)
        
        # If not found in message, try to look up by email
        if not brand_id and buyer_email:
            brand_id = get_brand_id_by_email(buyer_email)
            
        if not brand_id:
            logger.warning(f"⚠️ Could not find Brand ID for order {order_id}, skipping")
            continue
            
        # Get brand details
        brand_details = get_brand_details(brand_id)
        if not brand_details:
            logger.warning(f"⚠️ Could not find brand details for ID {brand_id}, skipping")
            continue
            
        # Process the order
        process_paid_order(order, brand_details)
        
        # Mark as processed
        if len(PROCESSED_EVENTS) >= MAX_CACHE_SIZE:
            PROCESSED_EVENTS.popitem(last=False)
        PROCESSED_EVENTS[order_id] = {"timestamp": time.time(), "processed": True}

def process_paid_order(order, brand_details):
    """
    Process a paid Etsy order
    
    Args:
        order: Etsy order dictionary
        brand_details: Brand details dictionary
    """
    order_id = order.get("receipt_id")
    logger.info(f"🔄 Processing paid order {order_id} for {brand_details['company_name']}")
    
    temp_files = []  # Track temp files for cleanup
    
    try:
        # Create an entry in the Orders database
        order_data = {
            "order_id": order_id,
            "email": brand_details["email"],
            "brand_id": brand_details["brand_id"]
        }
        
        notion_page_id = create_notion_order(order_data)
        
        # Get the logo
        logo_response = requests.get(brand_details["logo_url"])
        if logo_response.status_code != 200:
            logger.error(f"❌ Failed to download logo: {logo_response.status_code}")
            return
            
        # Generate the licensed template
        business_name = brand_details["company_name"]
        address1 = brand_details["address"]
        address2 = brand_details["city_state_zip"]
        phone = brand_details["phone"]
        email = brand_details["email"]
        
        # Use default values for missing info
        tax_percentage = "7.0"  # Default tax rate
        currency = "USD"  # Default currency
        
        # Generate the workbook without watermark
        wb, temp_file_list = generate_template(
            business_name, address1, address2, phone, email,
            tax_percentage, currency, logo_response.content,
            LICENSE_TEMPLATE_PATH, add_watermark=False  # No watermark for paid version
        )
        
        temp_files.extend(temp_file_list)
        
        # Save workbook to temp file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            wb.save(tmp.name)
            temp_files.append(tmp.name)
            
        # Send email with the template
        try:
            send_email(
                recipient_email=email,
                subject="Your Licensed Template is Ready!",
                body=f"Thank you for your purchase! Your licensed template is attached. Your order number is {order_id}.",
                attachment_paths=[tmp.name],
                business_name=business_name,
                template_type="invoice"  # TODO: Make this dynamic based on purchase
            )
            logger.info(f"✅ Sent licensed template to {email}")
            
            # Update order status in Notion
            if notion_page_id:
                update_order_email_sent(notion_page_id)
                
        except Exception as e:
            logger.error(f"❌ Error sending email: {e}")
            
    except Exception as e:
        logger.error(f"❌ Error processing order: {e}")
        
    finally:
        # Clean up temp files
        for file_path in temp_files:
            try:
                if os.path.exists(file_path):
                    os.unlink(file_path)
            except Exception as e:
                logger.error(f"❌ Error cleaning up temp file: {e}")

# === WEBHOOK ENDPOINTS ===
@app.route("/preview_webhook", methods=["POST"])
def handle_preview_request():
    """Handle preview requests from Tally form submissions"""
    temp_files = []  # List to keep track of temporary files to clean up
    event_id = None
    
    try:
        data = request.json
        logger.info("🚀 Raw incoming data from Tally")
        
        # Extract event ID and creation time for idempotency check
        event_id = data.get('eventId')
        event_time = data.get('createdAt')
        
        # Check if we've already processed this event
        if event_id in PROCESSED_EVENTS and PROCESSED_EVENTS[event_id].get("processed", False):
            logger.info(f"✅ Skipping already processed event: {event_id}")
            return '', 204
        
        # Check if the event is too old (stale)
        is_stale, time_diff = is_stale_event(event_time)
        if is_stale:
            logger.info(f"⚠️ Skipping stale event from {time_diff:.1f} minutes ago: {event_id}")
            
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
            
        # Extract relevant fields
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
            
        # Download the logo
        logo_response = requests.get(logo_url)
        if logo_response.status_code != 200:
            return jsonify({"error": "Failed to download logo"}), 400
            
        # Store the Tally logo URL directly
        permanent_logo_url = logo_url
        fields["LogoURL"] = permanent_logo_url
        
        # Generate a new Brand ID and add to fields
        fields["BrandID"] = generate_brand_id()
        
        # Create/update Notion database entry
        notion_page_id = update_notion_database(fields=fields)
        
        # Generate preview template
        wb, temp_file_list = generate_template(
            business_name, address1, address2, phone, email,
            tax_percentage, currency, logo_response.content,
            TEMPLATE_PATH, add_watermark=True  # Add watermark for preview
        )
        
        temp_files.extend(temp_file_list)
        
        # Save the workbook to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            wb.save(tmp.name)
            temp_files.append(tmp.name)
            
        # Send the email
        try:
            send_email(
                recipient_email=email,
                subject="Your Preview Template is Ready!",
                body="Please find attached your preview template. To get the full version without watermark, visit our Etsy store.",
                attachment_paths=[tmp.name],
                business_name=business_name
            )
            
            # Mark this event as successfully processed
            if len(PROCESSED_EVENTS) >= MAX_CACHE_SIZE:
                PROCESSED_EVENTS.popitem(last=False)
            PROCESSED_EVENTS[event_id] = {"timestamp": time.time(), "processed": True}
            
            logger.info(f"✅ Successfully processed event: {event_id}")
            
            # Update Notion after successfully sending the email
            if notion_page_id:
                try:
                    notion.pages.update(
                        page_id=notion_page_id,
                        properties={
                            "Email Sent": {"checkbox": True}
                        }
                    )
                    logger.info("✅ Updated Email Sent status in Notion")
                except Exception as e:
                    logger.error(f"⚠️ Error updating Email Sent status: {e}")
        except Exception as e:
            # Mark as received but not successfully processed
            if len(PROCESSED_EVENTS) >= MAX_CACHE_SIZE:
                PROCESSED_EVENTS.popitem(last=False)
            PROCESSED_EVENTS[event_id] = {"timestamp": time.time(), "processed": False}
            
            logger.error(f"⚠️ Failed to send email: {e}")
            return jsonify({"error": "Failed to send email"}), 500
            
        return jsonify({
            "success": True, 
            "message": "Preview generated and sent",
            "brand_id": fields["BrandID"]
        }), 200
        
    except Exception as e:
        # Mark the event as received but not processed if we have an event_id
        if event_id:
            if len(PROCESSED_EVENTS) >= MAX_CACHE_SIZE:
                PROCESSED_EVENTS.popitem(last=False)
            PROCESSED_EVENTS[event_id] = {"timestamp": time.time(), "processed": False}
        
        logger.error(f"⚠️ Unhandled error: {e}")
        return jsonify({"error": str(e)}), 500
        
    finally:
        # Clean up all temporary files
        for file_path in temp_files:
            try:
                if file_path and os.path.exists(file_path):
                    os.unlink(file_path)
                    logger.debug(f"✅ Removed temporary file: {file_path}")
            except Exception as e:
                logger.error(f"⚠️ Error removing temporary file {file_path}: {e}")

@app.route("/purchase_webhook", methods=["POST"])
def handle_purchase():
    """Handle manual purchase notifications (fallback for Etsy API)"""
    try:
        data = request.json
        logger.info("🚀 Received purchase webhook")
        
        order_id = data.get("order_id")
        brand_id = data.get("brand_id")
        
        if not order_id or not brand_id:
            return jsonify({"error": "Missing required fields"}), 400
            
        # Check if we've already processed this order
        if order_id in PROCESSED_EVENTS and PROCESSED_EVENTS[order_id].get("processed", False):
            logger.info(f"✅ Skipping already processed order: {order_id}")
            return jsonify({"message": "Order already processed"}), 200
            
        # Get brand details
        brand_details = get_brand_details(brand_id)
        if not brand_details:
            return jsonify({"error": f"Brand ID {brand_id} not found"}), 404
            
        # Process the order with minimal order data
        mock_order = {
            "receipt_id": order_id,
            "buyer_email": brand_details["email"]
        }
        
        process_paid_order(mock_order, brand_details)
        
        # Mark as processed
        if len(PROCESSED_EVENTS) >= MAX_CACHE_SIZE:
            PROCESSED_EVENTS.popitem(last=False)
        PROCESSED_EVENTS[order_id] = {"timestamp": time.time(), "processed": True}
        
        return jsonify({"message": "Order processed successfully"}), 200
        
    except Exception as e:
        logger.error(f"❌ Error processing purchase: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/health", methods=["GET"])
def health_check():
    """Simple health check endpoint"""
    return jsonify({
        "status": "ok",
        "processed_events": len(PROCESSED_EVENTS),
        "timestamp": datetime.now().isoformat(),
        "version": "1.0.0",
        "etsy_configured": all([ETSY_API_KEY, ETSY_API_SECRET, ETSY_ACCESS_TOKEN, ETSY_ACCESS_TOKEN_SECRET]),
        "notion_configured": bool(NOTION_TOKEN)
    })

# === SCHEDULER SETUP ===
scheduler = BackgroundScheduler()
scheduler.add_job(process_etsy_orders, 'interval', minutes=5)

# Start the scheduler when the app starts
@app.before_first_request
def init_scheduler():
    scheduler.start()
    logger.info("🚀 Started Etsy order polling scheduler")

# Shutdown the scheduler when the app stops
@app.teardown_appcontext
def shutdown_scheduler(exception=None):
    scheduler.shutdown()
    logger.info("✅ Shutdown Etsy order polling scheduler")

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
