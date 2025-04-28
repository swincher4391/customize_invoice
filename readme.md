Business Document Template System
A Flask-based system for generating, customizing, and delivering business document templates with both preview and licensed versions.

📋 System Overview
This system enables:

Preview Generation: Users submit business information and receive watermarked document previews
Purchase and Licensing: After purchasing on Etsy, users receive unwatermarked versions
Brand Identity Storage: Users create one persistent "Brand ID" that stores logo and company info
Automated Delivery: All files are generated and delivered automatically
🔄 System Flow
1. User submits form → Flask app generates preview → Email delivery
2. User purchases on Etsy → Order details captured → Licensed file delivered
🛠️ Technical Components
Backend (Flask)
/preview_webhook: Processes form submissions, generates watermarked previews
/purchase_webhook: Processes Etsy purchases, generates licensed versions
process_etsy_orders: Background job that polls Etsy API for new orders
Storage (Notion)
Brand ID Database: Stores customer brand information (logo, company name, etc.)
Orders Database: Tracks order status and links to Brand IDs
Asset Storage (Tally URLs)
Uses the image URLs provided by Tally form
Simple solution for MVP (no additional storage needed)
Future enhancement: migrate to permanent storage like S3
🚀 Installation
Clone this repository
Install dependencies:
pip install -r requirements.txt
Create a .env file with your credentials (see Environment Variables below)
Run the app:
python app.py
For production deployment, we recommend using Gunicorn:

gunicorn app:app
🔐 Environment Variables
bash
# Notion
NOTION_TOKEN=
NOTION_DATABASE_ID=           # Brand ID database
NOTION_ORDERS_DATABASE_ID=    # Orders database

# Etsy (OAuth1)
ETSY_API_KEY=
ETSY_API_SECRET=
ETSY_ACCESS_TOKEN=
ETSY_ACCESS_TOKEN_SECRET=
ETSY_SHOP_ID=

# Email
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=
SMTP_PASS=
SENDER_NAME="Template Generator"

# No additional storage configuration required for MVP
# Uses Tally image URLs directly

# Template Paths
TEMPLATE_PATH=invoice-watermarked.xlsx
LICENSE_TEMPLATE_PATH=invoice-licensed.xlsx
WATERMARK_PATH=watermark.png
📊 API Endpoints
POST /preview_webhook: Receives Tally form submissions
POST /purchase_webhook: Manual purchase processing (fallback for Etsy API)
GET /health: Health check endpoint
🏗️ Project Structure
├── app.py                     # Main Flask application
├── requirements.txt           # Python dependencies
├── invoice-watermarked.xlsx   # Preview template with watermark
├── invoice-licensed.xlsx      # Licensed template without watermark
├── watermark.png              # "PREVIEW ONLY" background image
├── .env                       # Environment variables (not in repo)
└── README.md                  # This file
🔄 Data Flow Diagram
┌─────────────┐        ┌─────────────┐        ┌─────────────┐
│   Tally     │───────▶│    Flask    │───────▶│   Notion    │
│  Form       │        │   Server    │        │  Database   │
└─────────────┘        └──────┬──────┘        └─────────────┘
                              │
                              ▼
                       ┌─────────────┐
                       │   Email     │
                       │  Delivery   │
                       └─────────────┘
                              ▲
                              │
┌─────────────┐        ┌─────┴──────┐
│    Etsy     │───────▶│  Etsy API  │
│  Purchase   │        │  Listener  │
└─────────────┘        └─────────────┘
📈 Future Enhancements
Template Variety: Add more document types (quotes, receipts, letterhead)
Customization Options: Color schemes, font choices, layout options
Self-Service Portal: Allow customers to manage templates in one place
Subscription Model: Offer access to all templates for a monthly fee
Analytics Dashboard: Track conversions and template popularity
📜 License
This project is for internal use only.

🙏 Contributors
Your Name - Project Lead
