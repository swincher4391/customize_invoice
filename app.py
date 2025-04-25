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

if __name__ == "__main__":
    app.run(debug=True)
