# 🚀 Automated PageSpeed Insights Fetcher for Google Sheets

A fully automated PageSpeed Insights batch processor built using Google Apps Script.
This script fetches detailed Lighthouse metrics for multiple URLs, processes them in time-safe batches, stores intermediate results in Script Properties, and finally writes all consolidated metrics into your Google Sheet — automatically, reliably, and without hitting API time limits.

This tool is ideal for:
SEO teams
Competitor analysis
Page experience audits
Large-scale PageSpeed monitoring
Weekly or monthly performance benchmarking

# 📌 Features
✅ Fully Automated Batch Processing
Processes URLs in batches (default: 5 per batch) to prevent script timeouts and API quota issues.

✅ Supports Both Mobile & Desktop
For every URL, the script fetches two lighthouse reports:
Mobile strategy
Desktop strategy

✅ Retrieves 7 Key Lighthouse Metrics
For both mobile & desktop:
Performance score
First Contentful Paint (FCP)
Speed Index (SI)
Total Blocking Time (TBT)
Largest Contentful Paint (LCP)
Cumulative Layout Shift (CLS)
Interaction to Next Paint (INP)

✅ Complete Automation Flow
One-click Start Update
Automatic Trigger-based batching
Automatic final writing into sheet
Optional email notification on completion
Full cleanup of triggers & temporary data

✅ Error-Safe & Time-Safe
Avoids exceeding the 6-minute Apps Script runtime
Gracefully handles API errors or missing URLs
Ensures complete data collection before writing to sheet

# 📁 Sheet Structure Requirements
Your Google Sheet should include:
A sheet named exactly: Page Speed
Row 3 onward should contain URLs horizontally starting from Column C
(Example: C3, D3, E3, F3 …)
The script automatically appends a new block of results below existing data.

⚙️ Configuration
You can adjust these values inside the script:
const SHEET_NAME = 'Page Speed';  
const API_KEY_PROPERTY_NAME = 'PAGESPEED_API_KEY';  
const BATCH_SIZE = 5;  
const NOTIFICATION_EMAIL = 'abhishek.yadav@infidigit.com';

👉 Before Running for the First Time
Set your API key in Script Properties:
Extensions → Apps Script
Left menu → Project Properties
Script Properties → Add
Key: PAGESPEED_API_KEY
Value: your Google PageSpeed API key

# 🧩 How the Script Works?

1️⃣ onOpen() – Adds Custom Menu
Adds a menu inside Google Sheets:

🚀 Automate Page Speed
  ├── Start Full Update
  └── Cancel Running Update

2️⃣ startUpdateProcess() – Initializes Everything
Clears old data & triggers
Resets index
Creates empty arrays for each metric (mobile & desktop)
Starts the first batch after 1 second
Shows confirmation popup

3️⃣ processBatch() – Batch Processor
Loads next set of URLs
Calls PageSpeed API for mobile & desktop
Saves metric results into Script Properties
Respects time limits (stops after ~4.5 minutes)
If more URLs left → creates another time-based trigger
If all done → calls finalizeUpdate()

4️⃣ finalizeUpdate() – Writes Final Output
Pulls all stored metrics
Appends them into the Google Sheet in a clean, structured table
Sends completion email
Cleans up triggers & properties

5️⃣ cancelUpdateProcess() – Safety Reset
Helps stop the process any time.
🔗 API Used
Google PageSpeed Insights v5 API:
https://www.googleapis.com/pagespeedonline/v5/runPagespeed

Each call includes:
URL
Strategy (mobile/desktop)
Category: PERFORMANCE

🧪 Error Handling
If a URL is invalid or API fails, the script inserts:
'-'
Instead of breaking the automation.

📬 Email Notification
At the end of execution, you receive:
Subject: Google Sheet PageSpeed Update Complete
Message: Automated competitor PageSpeed analysis completed.

🚦 Trigger Logic
The script uses time-based triggers to avoid runtime limits:

Stage	Trigger Delay	Purpose
First batch	1 sec	Start quickly
Subsequent batches	1 min	Avoid API quota issues
Auto-stop	After all URLs processed	Cleanup
🛑 When to Use "Cancel Running Update"

Use it when:
You want to restart the process
You added new URLs
Something got stuck
This resets everything instantly.

🧾 Code Included
This repository includes the full Google Apps Script:
Batch processing logic
PageSpeed API fetch logic
Metrics extraction
Sheet writing
Error handling
Trigger cleanup

Custom menu creation
📣 Final Notes
This automation is designed to handle large competitor lists safely without timeouts, making PageSpeed benchmarking faster and completely hands-free.
