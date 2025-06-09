# HireTrack-AI
HireTrack AI is a smart automation tool built with Python that connects to your Gmail inbox, classifies job application emails (e.g., Applied, Interview, Rejected), and logs them into an organized Excel file. It helps job seekers effortlessly track their job application statuses using NLP, email parsing, and Gmail API.

## 🚀 Key Features
- 🔐 Gmail OAuth Integration for secure, read-only email access
- 🧠 NLP with spaCy & Regex to classify emails into: Applied, Interview, Rejected
- 📩 Filters out spammy alerts (e.g., job newsletters, duplicate alerts)
- 🧾 Auto-extracts Company Name, Job Role, Status, Date & Time from email content
- 📊 Excel Logging with colored status indicators and dropdown filtering
- ⏱️ Scheduled Execution: Runs every 30 minutes to keep the Excel sheet updated

## 🧠 How It Works
- Authenticates with Gmail API using OAuth (credentials.json)
- Scans recent emails with keywords like application, rejected, interview
- Extracts relevant data using NLP + Regex from subject and body
- Writes results to Excel with:
- Color codes (green = applied, yellow = interview, red = rejected)
- Auto-sorting by date and time
- Schedules itself to run every 30 minutes via schedule module

## 🗂️ Project Structure
```
HireTrackAI/
├── HiretrackAI.py              # Main script with all functionality
├── token.json                  # Generated after first Gmail OAuth login
├── credentials.json            # Gmail API credentials file
├── job_applications_hireai.xlsx # Output file with tracked applications
```
## 📦 Requirements
Install dependencies:

bash
```
pip install -r requirements.txt
```
Typical libraries used: pandas, openpyxl, google-auth, google-api-python-client,vbs4, spacy, schedule, re, datetime, pytz

Also run:
```
python -m spacy download en_core_web_sm
```
## ▶️How to Run
1. Add credentials.json to your project root (from Google Cloud Console)
2. Run the script:
bash
```
python HiretrackAI.py
```
3. On first run, it will ask for Gmail authorization via URL
4. Once authenticated, tracking begins and runs every 30 minutes
5. The output will be saved in job_applications_hireai.xlsx

## 📊 Excel Output Sample

| Company | Job Role          | Status    | Classification Phrase | Date Applied | Time Received |
| ------- | ----------------- | --------- | --------------------- | ------------ | ------------- |
| Google  | Software Engineer | Applied   | your application to   | 2025-06-05   | 09:30         |
| Amazon  | Data Analyst      | Rejected  | unfortunately         | 2025-06-02   | 14:15         |
| OpenAI  | ML Intern         | Interview | interview scheduled   | 2025-06-01   | 11:00         |

## 📌 Notes
- Only real application status emails are considered (LinkedIn job alerts, newsletters, etc. are excluded)
- Adds dropdown validation for status column
- Handles first run (2-year range) and subsequent runs (30-minute lookback)
- Smart fallback using spaCy for company/role when regex fails
