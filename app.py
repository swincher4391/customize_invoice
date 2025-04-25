from flask import Flask, request, jsonify
from notion_client import Client as NotionClient
import requests
import time
import os

app = Flask(__name__)

# === ENVIRONMENT VARIABLES ===
NOTION_TOKEN = os.getenv("NOTION_TOKEN")
DATABASE_ID = os.getenv("NOTION_DATABASE_ID")
ETSY_API_KEY = os.getenv("ETSY_API_KEY")
ETSY_ACCESS_TOKEN = os.getenv("ETSY_ACCESS_TOKEN")
SHOP_ID = os.getenv("ETSY_SHOP_ID")
TEMPLATE_PATH=""
notion = NotionClient(auth=NOTION_TOKEN)

# === LICENSE KEY GENERATION ===
def generate_license_key(order_number, timestamp=None):
    if timestamp is None:
        timestamp = int(time.time())
    numeric_order = int(''.join(filter(str.isdigit, order_number)))
    combined = numeric_order + (timestamp % 10000)
    multiplier = 7919
    xor_value = 61681
    result = (combined * multiplier) ^ xor_value
    return f"EXCEL-{order_number}-{(result % 9000) + 1000}"

# === ETSY ORDER VALIDATION ===
def validate_etsy_order(order_number):
    url = f"https://openapi.etsy.com/v3/application/shops/{SHOP_ID}/receipts/{order_number}"
    headers = {
        "x-api-key": ETSY_API_KEY,
        "Authorization": f"Bearer {ETSY_ACCESS_TOKEN}"
    }
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return True
    except requests.exceptions.RequestException as e:
        print(f"Etsy validation failed: {e}")
        return False

# === UPDATE NOTION ===
def update_notion(order_number, email, company, license_key, validated=True):
    print("🚨 DATABASE_ID from env:", repr(DATABASE_ID))
    notion.pages.create(
        parent={"database_id": DATABASE_ID},
        properties={
            "Order Number": {"rich_text": [{"text": {"content": order_number}}]},
            "Email": {"email": email},
            "Company": {"rich_text": [{"text": {"content": company}}]},
            "License Key": {"rich_text": [{"text": {"content": license_key}}]},
            "Validated": {"checkbox": validated},
            "Excel Sent": {"checkbox": False}
        }
    )

# === WEBHOOK ENDPOINT ===
@app.route("/webhook", methods=["POST"])
def handle_webhook():
    data = request.json
    print("Received webhook:", data)

    # Get the 'fields' list
    fields = data.get("data", {}).get("fields", [])

    # Extract values by label
    field_map = {field["label"]: field["value"] for field in fields}
    order_number = field_map.get("Etsy Order Number")
    email = field_map.get("Email")
    company = field_map.get("Company Name")

    if not all([order_number, email, company]):
        return jsonify({"error": "Missing required fields"}), 400

    is_valid = True  # Or call validate_etsy_order(order_number)
    license_key = generate_license_key(order_number)

    update_notion(order_number, email, company, license_key, validated=is_valid)

    return jsonify({"message": "Success", "license_key": license_key}), 200

@app.route('/preview_webhook', methods=['POST'])
def handle_preview_request():
    """Process incoming preview requests from Tally.so form"""
    data = request.json
    
    # Extract form data
    business_name = data.get('business_name', 'Your Business')
    template_type = data.get('template_type', 'invoice')
    email = data.get('email')
    
    # Create a temporary file for the customized template
    with tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False) as temp_file:
        temp_path = temp_file.name
    
    # Generate preview template
    output_path = generate_preview_template(
        business_name=business_name,
        output_path=temp_path
    )
    
    # In a real application, you'd now:
    # 1. Email the template to the user
    # 2. Or upload it to cloud storage and generate a download link
    # 3. Log the activity in your database
    
    # For demo purposes, just send the file directly
    return send_file(output_path, as_attachment=True, 
                    download_name=f"{business_name}_invoice_preview.xlsm")

@app.route('/purchase_webhook', methods=['POST'])
def handle_purchase():
    """Process incoming Etsy purchase notifications"""
    data = request.json
    
    # Extract order details
    order_id = data.get('order_id')
    customer_email = data.get('buyer_email')
    license_key = f"EXCEL-{order_id}-{datetime.now().strftime('%Y%m%d')}"
    
    # Create a temporary file for the licensed template
    with tempfile.NamedTemporaryFile(suffix='.xlsm', delete=False) as temp_file:
        temp_path = temp_file.name
    
    # Generate licensed template
    output_path = generate_licensed_template(
        business_name="Customer Business", 
        output_path=temp_path, 
        license_key=license_key
    )
    
    # In a real application, you'd now:
    # 1. Email the template to the customer
    # 2. Store the license key in your database
    # 3. Log the purchase activity
    
    return jsonify({
        "status": "success",
        "license_key": license_key
    })

def generate_preview_template(business_name, output_path,type="invoice"):
    if type == "invoice":
        TEMPLATE_PATH = "invoice-watermarked.xlsm"
    else:
        TEMPLATE_PATH = "invoice-watermarked.xlsm"
    """Creates a watermarked preview version of the invoice template"""
    # Load the pre-built template (with VBA already included)
    wb = openpyxl.load_workbook(TEMPLATE_PATH, keep_vba=True)
    ws = wb["Invoice"]
    
    # Update business info
    ws['A1'] = business_name
    
    # Ensure watermarks are visible (license is inactive)
    license_sheet = wb["License"]
    license_sheet['B2'] = "Inactive"
    
    # Save the customized template
    wb.save(output_path)
    
    return output_path

def generate_licensed_template(business_name, output_path, license_key,type="invoice"):
    if type == "invoice":
        TEMPLATE_PATH = "invoice-watermarked.xlsm"
    else:
        TEMPLATE_PATH = "invoice-watermarked.xlsm"
    """Creates a licensed version of the invoice template"""
    # Load the pre-built template (with VBA already included)
    wb = openpyxl.load_workbook(TEMPLATE_PATH, keep_vba=True)
    ws = wb["Invoice"]
    
    # Update business info
    ws['A1'] = business_name
    
    # Update license information
    license_sheet = wb["License"]
    license_sheet['B1'] = license_key
    license_sheet['B2'] = "Active"  # Set to active immediately
    
    # Save the customized template
    wb.save(output_path)
    
    return output_path
    
if __name__ == "__main__":
    app.run(debug=True)
