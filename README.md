# 🤖 AI Opportunity Agent

An autonomous Python agent that monitors a Gmail inbox, identifies career opportunities (internships, research, jobs), extracts key details using AI, and organizes them in a Google Sheet. It sends real-time notifications for high-relevance opportunities to a Telegram group.

---
## ✨ Features

* **Smart Filtering:** Automatically ignores personal emails, promotional content, and irrelevant announcements by checking sender domains, keywords, and using an AI classifier.
* **Intelligent Extraction:** Uses Google's Gemini 1.5 Flash to parse email bodies and attachments (PDF/DOCX) to extract 18 key data points, including role, company, deadline, and stipend.
* **Personalized Relevance Scoring:** Analyzes your resume to provide a custom 1-10 relevance score for each opportunity.
* **Automated Data Entry:** Populates a Google Sheet with all extracted information, serving as a centralized tracking dashboard.
* **Thread-Aware Updates:** Intelligently handles follow-up emails in the same thread, updating existing records with new information (e.g., deadline changes) and highlighting the changes in red.
* **Real-time Notifications:** Sends formatted summary messages for new or updated high-relevance opportunities to a Telegram chat or group.
* **Autonomous Operation:** Runs automatically every hour using a free GitHub Actions workflow.

---
## 🛠️ Technical Workflow

The agent operates on a schedule, performing a series of checks and actions.



1.  **Schedule Trigger:** A GitHub Actions cron job wakes the agent up every hour.
2.  **Fetch Emails:** The agent connects to the Gmail API and fetches emails from the last 2 hours.
3.  **Filter Pipeline:** Each email is passed through a series of filters:
    * **Processed Check:** Skips emails it has seen before (using `processed_emails.json`).
    * **Personal Check:** Skips emails sent to a small number of recipients.
    * **Domain Check:** Skips emails not from trusted domains (e.g., `@iitm.ac.in`).
    * **Keyword Check:** Skips emails that don't contain opportunity-related keywords.
    * **AI Classifier:** A final, quick AI check to confirm it's an opportunity.
4.  **Extract or Update:**
    * If the email's thread ID is new, the agent performs a full extraction.
    * If the thread ID already exists, the agent performs an update.
5.  **Log to Google Sheets:** The data is written to the central Google Sheet.
6.  **Notify:** High-relevance opportunities trigger a formatted message to a Telegram chat.

---
## 🚀 Setup Guide

### **1. Clone the Repository**
Create a new **private** repository on GitHub and clone it to your local machine.

### **2. Set Up Local Environment**
```bash
# Create and activate a virtual environment
python -m venv venv
source venv/bin/activate

# Install required libraries
pip install -r requirements.txt
```

### **3. Google Cloud & API Setup**
1.  Create a project in the [Google Cloud Console](https://console.cloud.google.com/).
2.  Enable the **Gmail API** and **Google Sheets API**.
3.  Create an **OAuth 2.0 Client ID** for a **Desktop app**.
4.  Download the credentials and save the file as `credentials.json` in your project root.

### **4. Initial Authentication**
Run the `auth.py` script once to generate your `token.json` file.
```bash
python auth.py
```
A browser window will open. Log in and grant the requested permissions. A `token.json` file will be created.

### **5. Configuration**
1.  **`resume.txt`:** Paste the full text of your resume into this file.
2.  **Google Sheet:**
    * Create a new Google Sheet.
    * Set up the required column headers in the first row.
    * Copy the **Sheet ID** from its URL.
3.  **Telegram Bot:**
    * Use `@BotFather` on Telegram to create a new bot and get your **Bot Token**.
    * Use a bot like `@userinfobot` (for personal chat) or `@get_id_bot` (for a group) to get your **Chat ID**.
4.  **`.env` file:** Create a file named `.env` in the project root and add your secrets:
    ```
    GEMINI_API_KEY="YOUR_GEMINI_API_KEY"
    TELEGRAM_BOT_TOKEN="YOUR_TELEGRAM_BOT_TOKEN"
    TELEGRAM_CHAT_ID="YOUR_TELEGRAM_CHAT_ID"
    SPREADSHEET_ID="YOUR_GOOGLE_SHEET_ID"
    GOOGLE_SHEET_LINK="YOUR_GOOGLE_SHEET_SHARE_LINK"
    ```

### **6. GitHub Secrets for Automation**
Go to your GitHub repo's **Settings > Secrets and variables > Actions** and create the following repository secrets:
* `GEMINI_API_KEY`
* `GCP_CREDENTIALS_JSON` (Paste the entire content of `credentials.json`)
* `GCP_TOKEN_JSON` (Paste the entire content of `token.json`)
* `TELEGRAM_BOT_TOKEN`
* `TELEGRAM_CHAT_ID`
* `SPREADSHEET_ID`
* `GOOGLE_SHEET_LINK`

---
## ▶️ How to Run

### **Local Testing**
To test the agent on your local machine, simply run:
```bash
python main.py
```

### **Automated Run**
The agent is configured to run automatically every hour via the `.github/workflows/run_agent.yml` file once the code is pushed to your GitHub repository. You can also trigger a manual run from the **Actions** tab on your GitHub repo page.
