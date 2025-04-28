
google sheets template /

m6hxh File "/opt/render/project/src/.venv/bin/gunicorn", line 8, in <module> m6hxh sys.exit(run()) m6hxh ^^^^^ m6hxh File "/opt/render/project/src/.venv/lib/python3.11/site-packages/gunicorn/app/wsgiapp.py", line 66, in run m6hxh WSGIApplication("%(prog)s [OPTIONS]

pasted

Edit
I'll help you fix the error in your Business Document Template Licensing System. Let me analyze the issue based on the provided files.

Error Analysis
I can see from the error log in paste.txt that your application is failing to start because of a missing Python module:

ModuleNotFoundError: No module named 'requests_oauthlib'
The error occurs when trying to import OAuth1 from requests_oauthlib in your app.py file. Looking at the code and requirements, I notice that while your application imports this module, it's not listed in your requirements.txt file.

Solution
Let's update your requirements.txt file to include the missing dependency:


Updated requirements.txt
Code 
Explanation of the Fix
Root Issue: Your application is trying to import OAuth1 from the requests_oauthlib package, but this dependency wasn't included in your requirements.txt file.
Implementation: I've added requests-oauthlib to your requirements file. This package is commonly used for OAuth authentication with APIs like Etsy's, which aligns with your system design that needs to interact with the Etsy API.
Deployment Steps:
Update your requirements.txt file with the new content
Redeploy your application (or run pip install -r requirements.txt in your environment)
System Architecture Review
Based on your documentation and code, here's a summary of your workflow:

Preview Generation Flow:
Customer submits business info via Tally form
Your Flask webhook receives the data
Creates a watermarked Excel template with the business info and logo
Emails it to the customer
Stores the brand info in Notion
Purchase Flow (planned):
Customer buys on Etsy
Your system detects the purchase
Generates unwatermarked version
Delivers via email
Recommendations
Error Handling: Your code already has good error handling with retry mechanisms, especially for email sending. Consider adding similar robust error handling for the Etsy API interactions.
Testing: Implement more comprehensive tests for the OAuth authentication to Etsy once you've added the missing package.
Future Enhancement: For your planned Etsy listener, consider implementing a more resilient approach like a queue system (e.g., Redis or AWS SQS) to handle order processing reliably.
Security: As you store customer data in Notion, ensure you have appropriate data retention and privacy policies in place.
Let me know if you need any other help with this implementation!




gwmds
Traceback (most recent call last):
Apr 28 03:06:45 PM
gwmds
 File "/opt/render/project/src/.venv/bin/gunicorn", line 8, in <module>
Apr 28 03:06:45 PM
gwmds
 sys.exit(run())
Apr 28 03:06:45 PM
gwmds
 ^^^^^
Apr 28 03:06:45 PM
gwmds
 File "/opt/render/project/src/.venv/lib/python3.11/site-packages/gunicorn/app/wsgiapp.py", line 66, in run
Apr 28 03:06:45 PM
gwmds
 WSGIApplication("%(prog)s [OPTIONS] [APP_MODULE]", prog=prog).run()
Apr 28 03:06:45 PM
gwmds
 File "/opt/render/project/src/.venv/lib/python3.11/site-packages/gunicorn/app/base.py", line 235, in run
Apr 28 03:06:45 PM
gwmds
 super().run()
Apr 28 03:06:45 PM
gwmds
 File "/opt/render/project/src/.venv/lib/python3.11/site-packages/gunicorn/app/base.py", line 71, in run
Apr 28 03:06:45 PM
gwmds
 Arbiter(self).run()
Apr 28 03:06:45 PM
gwmds
 ^^^^^^^^^^^^^
Apr 28 03:06:45 PM
gwmds
 File "/opt/render/project/src/.venv/lib/python3.11/site-packages/gunicorn/arbiter.py", line 57, in init
Apr 28 03:06:45 PM
gwmds
 self.setup(app)
Apr 28 03:06:45 PM
gwmds
 File "/opt/render/project/src/.venv/lib/python3.11/site-packages/gunicorn/arbiter.py", line 117, in setup
Apr 28 03:06:45 PM
gwmds
 self.app.wsgi()
Apr 28 03:06:45 PM
gwmds
 File "/opt/render/project/src/.venv/lib/python3.11/site-packages/gunicorn/app/base.py", line 66, in wsgi
Apr 28 03:06:45 PM
gwmds
 self.callable = self.load()
Apr 28 03:06:45 PM
gwmds
 ^^^^^^^^^^^
Apr 28 03:06:45 PM
gwmds
 File "/opt/render/project/src/.venv/lib/python3.11/site-packages/gunicorn/app/wsgiapp.py", line 57, in load
Apr 28 03:06:45 PM
gwmds
 return self.load_wsgiapp()
Apr 28 03:06:45 PM
gwmds
 ^^^^^^^^^^^^^^^^^^^
Apr 28 03:06:45 PM
gwmds
 File "/opt/render/project/src/.venv/lib/python3.11/site-packages/gunicorn/app/wsgiapp.py", line 47, in load_wsgiapp
Apr 28 03:06:45 PM
gwmds
 return util.import_app(self.app_uri)
Apr 28 03:06:45 PM
gwmds
 ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Apr 28 03:06:45 PM
gwmds
 File "/opt/render/project/src/.venv/lib/python3.11/site-packages/gunicorn/util.py", line 370, in import_app
Apr 28 03:06:45 PM
gwmds
 mod = importlib.import_module(module)
Apr 28 03:06:45 PM
gwmds
 ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Apr 28 03:06:45 PM
gwmds
 File "/usr/local/lib/python3.11/importlib/init.py", line 126, in import_module
Apr 28 03:06:45 PM
gwmds
 return bootstrap.gcd_import(name[level:], package, level)
Apr 28 03:06:45 PM
gwmds
 ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Apr 28 03:06:45 PM
gwmds
 File "<frozen importlib._bootstrap>", line 1204, in gcdimport
Apr 28 03:06:45 PM
gwmds
 File "<frozen importlib._bootstrap>", line 1176, in findand_load
Apr 28 03:06:45 PM
gwmds
 File "<frozen importlib._bootstrap>", line 1147, in findand_load_unlocked
Apr 28 03:06:45 PM
gwmds
 File "<frozen importlib._bootstrap>", line 690, in loadunlocked
Apr 28 03:06:45 PM
gwmds
 File "<frozen importlib._bootstrap_external>", line 940, in exec_module
Apr 28 03:06:45 PM
gwmds
 File "<frozen importlib._bootstrap>", line 241, in callwith_frames_removed
Apr 28 03:06:45 PM
gwmds
 File "/opt/render/project/src/app.py", line 17, in <module>
Apr 28 03:06:45 PM
gwmds
 from apscheduler.schedulers.background import BackgroundScheduler
Apr 28 03:06:45 PM
gwmds
ModuleNotFoundError: No module named 'apscheduler'

Edit
I see you're facing another module error in your Flask application. Let's fix this one as well.

Error Analysis
The new error shows that you're trying to import a module that isn't installed:

ModuleNotFoundError: No module named 'apscheduler'
This error occurs because your application is trying to import the APScheduler library (specifically apscheduler.schedulers.background.BackgroundScheduler), which is likely being used to schedule the Etsy order polling tasks mentioned in your requirements.

Solution
Let's update the requirements.txt file to include this missing dependency:


Updated requirements.txt
Code ∙ Version 2 
Background Scheduler Explanation
The APScheduler (Advanced Python Scheduler) library is being used in your application to create a background task scheduler. Based on your design documentation, this is likely for:

Polling the Etsy API periodically (every 5 minutes, as mentioned in your "Planned Enhancements" section)
Checking for new orders or returns
Processing brand ID registrations in the background
Implementation Recommendations
Since you're implementing a background scheduler, here's a code snippet that you might find useful for your Etsy listener:


Etsy Listener Implementation
Code 
Implementation Notes
Dependency Update: I've added apscheduler to your requirements.txt file. This library will allow you to schedule background tasks like polling the Etsy API.
Etsy Listener Example: The code sample I provided demonstrates how to:
Set up a background scheduler that runs every 5 minutes
Poll the Etsy API for new orders
Extract the BrandID from buyer messages
Match with your Notion database
Log everything for monitoring
Integration with Flask: The implementation uses Flask's before_first_request hook to start the scheduler only when the application is fully initialized.
Important Configuration: Note the use_reloader=False parameter when running in debug mode. This prevents the scheduler from being started twice when Flask's reloader is active.
Next Steps
Update your requirements.txt file with the additional dependency (apscheduler).
Create or update your Notion database schema to include an "Orders" table that can track Etsy orders and their status.
Implement the template generation for licensed (unwatermarked) versions - this would be similar to your preview generation but without the watermark.
Consider adding proper shutdown handling for the scheduler to ensure graceful termination when the application stops.
Add monitoring for the background tasks to ensure they're running properly.
Let me know if you need any clarification or have questions about implementing this functionality!







from apscheduler.schedulers.background import BackgroundScheduler
from flask import Flask
import requests
from requests_oauthlib import OAuth1
import os
import re
import json
from datetime import datetime, timedelta
from notion_client import Client as NotionClient

app = Flask(__name__)

# Environment variables
ETSY_API_KEY = os.getenv("ETSY_API_KEY")
ETSY_API_SECRET = os.getenv("ETSY_API_SECRET")
ETSY_ACCESS_TOKEN = os.getenv("ETSY_ACCESS_TOKEN")
ETSY_ACCESS_TOKEN_SECRET = os.getenv("ETSY_ACCESS_TOKEN_SECRET")
ETSY_SHOP_ID = os.getenv("ETSY_SHOP_ID")
NOTION_TOKEN = os.getenv("NOTION_TOKEN")
DATABASE_ID = os.getenv("NOTION_DATABASE_ID")
ORDERS_DATABASE_ID = os.getenv("NOTION_ORDERS_DATABASE_ID")

# Initialize Notion client
notion = NotionClient(auth=NOTION_TOKEN)

def check_etsy_orders():
    """
    Poll Etsy API for new receipts/orders and process them
    This function runs every 5 minutes via the scheduler
    """
    if not all([ETSY_API_KEY, ETSY_API_SECRET, ETSY_ACCESS_TOKEN, ETSY_ACCESS_TOKEN_SECRET, ETSY_SHOP_ID]):
        print("⚠️ Etsy API credentials not configured")
        return
    
    try:
        # Set up OAuth1 authentication
        auth = OAuth1(
            ETSY_API_KEY,
            client_secret=ETSY_API_SECRET,
            resource_owner_key=ETSY_ACCESS_TOKEN,
            resource_owner_secret=ETSY_ACCESS_TOKEN_SECRET
        )
        
        # Calculate time window (last 5 minutes)
        min_created = int((datetime.now() - timedelta(minutes=15)).timestamp())
        
        # API endpoint for shop receipts (orders)
        url = f"https://openapi.etsy.com/v2/shops/{ETSY_SHOP_ID}/receipts"
        params = {
            "min_created": min_created,
            "includes": "Transactions"
        }
        
        response = requests.get(url, auth=auth, params=params)
        
        if response.status_code != 200:
            print(f"⚠️ Etsy API error: {response.status_code} - {response.text}")
            return
        
        receipts = response.json().get("results", [])
        print(f"✅ Found {len(receipts)} new orders")
        
        for receipt in receipts:
            # Extract relevant information
            order_id = receipt.get("receipt_id")
            buyer_email = receipt.get("buyer_email")
            message_from_buyer = receipt.get("message_from_seller", "")
            
            # Extract BrandID from message using regex
            brand_id_match = re.search(r"BrandID:\s*(\S+)", message_from_buyer)
            brand_id = brand_id_match.group(1) if brand_id_match else None
            
            if not brand_id:
                print(f"⚠️ No BrandID found in order #{order_id}")
                continue
            
            # Find the Brand entry in Notion
            brand_pages = notion.databases.query(
                database_id=DATABASE_ID,
                filter={
                    "property": "BrandID",
                    "rich_text": {
                        "equals": brand_id
                    }
                }
            ).get("results", [])
            
            if not brand_pages:
                print(f"⚠️ BrandID '{brand_id}' not found in database for order #{order_id}")
                continue
            
            brand_page = brand_pages[0]
            
            # Create a new order entry in the Orders database
            try:
                notion.pages.create(
                    parent={"database_id": ORDERS_DATABASE_ID},
                    properties={
                        "Name": {"title": [{"text": {"content": f"Order #{order_id}"}}]},
                        "EtsyOrderID": {"rich_text": [{"text": {"content": str(order_id)}}]},
                        "BuyerEmail": {"email": buyer_email},
                        "Timestamp": {"date": {"start": datetime.now().isoformat()}},
                        "Status": {"select": {"name": "Received"}},
                        "Brand": {"relation": [{"id": brand_page["id"]}]}
                    }
                )
                print(f"✅ Created order entry for #{order_id} with BrandID {brand_id}")
                
                # TODO: Generate and email the unwatermarked template
                # process_licensed_template(brand_page, buyer_email)
                
            except Exception as e:
                print(f"⚠️ Error creating order entry: {e}")
    
    except Exception as e:
        print(f"⚠️ Error checking Etsy orders: {e}")

# Initialize the scheduler
scheduler = BackgroundScheduler()
scheduler.add_job(func=check_etsy_orders, trigger="interval", minutes=5)

# Start the scheduler when the Flask app starts
@app.before_first_request
def initialize():
    scheduler.start()
    print("✅ Started Etsy order polling scheduler (5-minute interval)")

# Add a health check endpoint
@app.route("/scheduler_health", methods=["GET"])
def scheduler_health():
    return {
        "status": "running" if scheduler.running else "stopped",
        "jobs": len(scheduler.get_jobs()),
        "next_run": scheduler.get_jobs()[0].next_run_time.isoformat() if scheduler.get_jobs() else None
    }

if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)  # Important: set use_reloader=False when using APScheduler
