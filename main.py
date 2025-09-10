
import os
import base64
import time
import json
from dotenv import load_dotenv
import io
import requests


# Google API Libraries
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Google Sheets Library
import gspread

# Google Gemini AI Library
import google.generativeai as genai

# Attachment Reading Libraries
import PyPDF2
import docx

# --- CONFIGURATION ---
load_dotenv()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
GOOGLE_SHEET_LINK = os.getenv("GOOGLE_SHEET_LINK")
PROCESSED_EMAILS_FILE = "processed_emails.json"
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/spreadsheets", # This allows writing values and formatting
]
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")
MY_EMAIL_ADDRESS = "ed22b004@smail.iitm.ac.in" 

# --- FILTERS ---
TRUSTED_DOMAINS = ["iitm.ac.in"] # Add any other trusted domains
OPPORTUNITY_KEYWORDS = ["internship", "hiring", "research fellowship", "research internship", "fellowship", "recruiting", "job alert", "job opportunity"]

# --- SETUP AND AUTHENTICATION ---

def authenticate_google_services():
    """Authenticates using OAuth with refresh token for automation."""
    try:
        # Try to use existing token first
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
            if creds and creds.valid:
                print("✅ Using valid existing token")
                return build("gmail", "v1", credentials=creds), gspread.authorize(creds), creds
                
        # If no valid token, use refresh token from environment
        refresh_token = os.getenv('GOOGLE_REFRESH_TOKEN')
        client_id = os.getenv('GOOGLE_CLIENT_ID')
        client_secret = os.getenv('GOOGLE_CLIENT_SECRET')
        
        if not all([refresh_token, client_id, client_secret]):
            raise Exception("Missing OAuth credentials. Please set GOOGLE_REFRESH_TOKEN, GOOGLE_CLIENT_ID, and GOOGLE_CLIENT_SECRET environment variables.")
        
        print("🔄 Refreshing token using refresh token...")
        creds = Credentials(
            token=None,
            refresh_token=refresh_token,
            token_uri="https://oauth2.googleapis.com/token",
            client_id=client_id,
            client_secret=client_secret,
            scopes=SCOPES
        )
        
        # Refresh the token
        from google.auth.transport.requests import Request
        creds.refresh(Request())
        
        # Save the new token for next time
        with open("token.json", "w") as token:
            token.write(creds.to_json())
        
        print("✅ Successfully refreshed token")
        gmail_service = build("gmail", "v1", credentials=creds)
        gc = gspread.authorize(creds)
        return gmail_service, gc, creds
        
    except Exception as e:
        print(f"❌ Authentication error: {e}")
        raise


def configure_gemini():
    """Configures the Gemini AI model."""
    genai.configure(api_key=GEMINI_API_KEY)
    return genai.GenerativeModel('gemini-2.0-flash')

# --- STATE MANAGEMENT (with self-pruning) ---

def get_email_age_days(message_id):
    try:
        # This is the corrected value
        timestamp_ms = int(message_id, 16) >> 16 
        timestamp_s = timestamp_ms / 1000.0
        return (time.time() - timestamp_s) / (24 * 60 * 60)
    except (ValueError, TypeError):
        return 0

def load_processed_emails(retention_days=7):
    """Loads processed email IDs and prunes entries older than the retention period."""
    if not os.path.exists(PROCESSED_EMAILS_FILE):
        return set()
    
    with open(PROCESSED_EMAILS_FILE, "r") as f:
        try:
            processed_ids = json.load(f)
        except json.JSONDecodeError:
            return set()

    recent_ids = {
        msg_id for msg_id in processed_ids 
        if get_email_age_days(msg_id) < retention_days
    }
    
    if len(recent_ids) != len(processed_ids):
        print(f"🧹 Cleaned up {len(processed_ids) - len(recent_ids)} old email IDs from the tracking file.")
        with open(PROCESSED_EMAILS_FILE, "w") as f:
            json.dump(list(recent_ids), f)
            
    return recent_ids

def save_processed_email(email_id, processed_ids):
    """Adds a new email ID to the set and saves it to the file."""
    processed_ids.add(email_id)
    with open(PROCESSED_EMAILS_FILE, "w") as f:
        json.dump(list(processed_ids), f)

# --- EMAIL PROCESSING ---

def get_emails(service, search_query="newer_than:1h"):
    """Fetches a list of email message IDs, excluding sent mail."""
    # Add the exclusion filter for your own email address
    full_query = f"{search_query} -from:{MY_EMAIL_ADDRESS}"
    print(f"Using Gmail search query: '{full_query}'") # Helpful for debugging
    try:
        response = service.users().messages().list(userId="me", q=full_query).execute()
        return response.get("messages", [])
    except HttpError as error:
        print(f"An error occurred fetching emails: {error}")
        return []

def get_email_metadata(service, msg_id):
    """Fetches basic email details without parsing attachments. Returns attachment info for later."""
    try:
        message = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
        payload = message["payload"]
        headers = payload["headers"]

        sender_raw = next(h["value"] for h in headers if h["name"].lower() == "from")
        sender_email = sender_raw
        if '<' in sender_raw and '>' in sender_raw:
            sender_email = sender_raw.split('<')[1].split('>')[0]

        # Initialize list to hold attachment info (id, filename)
        attachment_list = []
        body_text = ""
        parts = [payload]
        while parts:
            part = parts.pop(0)
            if "parts" in part:
                parts.extend(part["parts"])

            mime_type = part.get("mimeType")
            if mime_type == "text/plain":
                encoded_body = part.get("body", {}).get("data", "")
                if encoded_body:
                    body_text += base64.urlsafe_b64decode(encoded_body).decode("utf-8")
            # Check for attachments, but only store their IDs, don't decode them yet.
            if "filename" in part and part["filename"]:
                filename = part["filename"].lower()
                body_data = part.get("body", {})
                if body_data and "attachmentId" in body_data:
                    # Just store the info for later use
                    attachment_list.append({
                        "id": body_data["attachmentId"],
                        "filename": filename
                    })

        details = {
            "id": message["id"],
            "threadId": message["threadId"],
            "sender_raw": sender_raw,
            "sender_email": sender_email.lower(),
            "subject": next(h["value"] for h in headers if h["name"].lower() == "subject"),
            "body": body_text.strip(),
            "attachment_ids": attachment_list # Now this is a list of dicts, not text
        }
        return details

    except Exception as e:
        print(f"Could not parse email {msg_id}: {e}")
        return None

def fetch_and_parse_attachments(service, msg_id, attachment_list):
    """Fetches and parses attachments based on the provided list."""
    full_attachment_text = ""
    for att_info in attachment_list:
        att_id = att_info["id"]
        filename = att_info["filename"]
        if not filename.endswith((".pdf", ".docx")):
            continue # Skip non-parsable attachments

        print(f"📄 Parsing attachment: {filename}")
        try:
            attachment = service.users().messages().attachments().get(
                userId="me", messageId=msg_id, id=att_id
            ).execute()
            file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
            file_stream = io.BytesIO(file_data)

            if filename.endswith(".pdf"):
                reader = PyPDF2.PdfReader(file_stream)
                for page in reader.pages:
                    full_attachment_text += page.extract_text() or ""
            elif filename.endswith(".docx"):
                doc = docx.Document(file_stream)
                for para in doc.paragraphs:
                    full_attachment_text += para.text + "\n"
        except Exception as e:
            print(f"Could not read attachment {filename}: {e}")
    return full_attachment_text.strip()

# --- AI INTERACTION ---

def is_opportunity_ai_check(model, email_text):
    """Uses AI to classify if an email is an opportunity."""
    prompt = f"Analyze the following email text. Is this a career opportunity (internship, job, research, fellowship)? | Strcitly avoid Events , Shows and Talks. Respond with only the word YES or NO.\n\nEMAIL TEXT:\n---\n{email_text}\n---"
    response = model.generate_content(prompt)
    return "yes" in response.text.lower()

def parse_ai_response(response_text):
    """Cleans and parses the JSON from the AI's response."""
    try:
        json_start = response_text.find('{')
        json_end = response_text.rfind('}') + 1
        if json_start == -1 or json_end == 0: return None
        json_text = response_text[json_start:json_end]
        return json.loads(json_text)
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON from AI response: {e}\nRaw AI response was:\n{response_text}")
        return None

def extract_initial_details_with_ai(model, email_details, resume_text):
    prompt = f"""
    Analyze the email and any attached document text. 
    Extract the fields in a valid JSON format. 
    For fields that contain a list of items (like 'Required Skills' or points in a 'Job Description') format the string with bullet points (•) to ensure readability.
    If a field is not mentioned, use "N/A". 
    Do not add any text outside the JSON object.

    The required fields are: "Application Deadline", "Institution/Company", "Eligibility" ,"Role Title", "Opportunity Type", "Role Field", "Location", "Work Mode", "Duration", "Time Commitment", "Stipend Details", "Required Skills", "Job Description (JD)", "Application Link", "Relevance Score (1-10)".


    **FIELD DEFINITIONS AND FORMATTING RULES:**
    - "Application Deadline": Extract the exact date. If a time like "EOD" or "11:59 PM" is mentioned, include it.
    - "Company/Institution": The name of the organization.
    - "Eligibility": Summarize who can apply. Examples: "B.Tech/M.Tech final year students", "2025 Graduates only", "All students". If not mentioned, **default to "All"**.
    - "Role Title": The official title of the position.
    - "Opportunity Type": Categorize into ONE of: Internship, Research Internship, Full-time, Part-time, Fellowship, Contest , Institute Student Body positions.
    - "Role Field": Categorize the specific domain. Examples: Software Engineering, Data Science, Mechanical, Operations, Finance.
    - "Location": The city and country. E.g., "Bengaluru, India".
    - "Work Mode": Specify ONE of: On-site, Remote, Hybrid.
    - "Duration": State the length of the opportunity. E.g., "6 Months", "May - August 2025".
    - "Time Commitment": Specify the expected hours. E.g., "Full-time", "Part-time (20 hrs/week)".
    - "Stipend Details": Extract the exact numbers or range, including currency.
    - "Required Skills": Create a bulleted (•) list of key technical skills or qualifications.
    - "Job Description (JD)": Create a concise, bulleted (•) summary of the main responsibilities.
    - "Application Link": Find the direct URL. If it's an email link, write "Reply to email".
    
    CRITICAL INSTRUCTION FOR RELEVANCE SCORING:
    - Be VERY strict when calculating the "Relevance Score (1-10)". This score should reflect how well the opportunity matches the candidate's profile.
    - Analyze the candidate's resume thoroughly and compare it with the opportunity requirements.
    - Score based on: field of study match, skills match, experience level, and overall fit.
    - A score of 10 means perfect match (all requirements met, ideal field).
    - A score of 7-9 means strong match (most requirements met, relevant field).
    - A score of 4-6 means moderate match (some requirements met, somewhat related field).
    - A score of 1-3 means weak match (few requirements met, unrelated field).
    - A thermal plant mechanical internship for a computer science student should score 2-3.
    - A data science internship for a computer science student with ML skills should score 8-9.
    - A Software development internship for a datascience student with low level skill/expirence in development should score 5-7.
    - BE CONSERVATIVE - default to lower scores when in doubt.

    
    CANDIDATE'S RESUME: --- {resume_text} --- [truncated if too long]
    EMAIL CONTENT: --- Subject: {email_details['subject']}; From: {email_details['sender_raw']}; Body: {email_details['body']} 
    ATTACHED DOCUMENT TEXT: --- {email_details['attachment_text'][:4000]} --- [truncated if too long]

    Provide ONLY the JSON output with no additional text:
    {{
        "Application Deadline": "value",
        "Institution/Company": "value",
        "Eligibility":"value"
        "Role Title": "value",
        "Opportunity Type": "value",
        "Role Field": "value",
        "Location": "value",
        "Work Mode": "value",
        "Duration": "value",
        "Time Commitment": "value",
        "Stipend Details": "value",
        "Required Skills": "value",
        "Job Description (JD)": "value",
        "Application Link": "value",
        "Relevance Score (1-10)": "value"
    }}
    """
    
    try:
        response = model.generate_content(prompt)
        extracted_data = parse_ai_response(response.text)
        
        # Additional validation for relevance score
        if extracted_data and "Relevance Score (1-10)" in extracted_data:
            try:
                score = extracted_data["Relevance Score (1-10)"]
                # Ensure it's a number between 1-10
                if isinstance(score, str):
                    if '/' in score:
                        score = score.split('/')[0]
                    score = float(score.strip())
                score = max(1, min(10, round(score)))
                extracted_data["Relevance Score (1-10)"] = str(score)
            except (ValueError, TypeError):
                # If AI gives invalid score, default to conservative 3
                extracted_data["Relevance Score (1-10)"] = "3"
                print("⚠️  Invalid relevance score received, defaulting to 3")
        
        return extracted_data
        
    except Exception as e:
        print(f"Error in AI extraction: {e}")
        return None



# In main.py, replace the old update_details_with_ai function with this one.

def update_details_with_ai(model, existing_data_json, new_email_details):
    """
    Uses Gemini to extract ONLY the updated details from a follow-up email.
    """
    prompt = f"""
    You are an AI data extraction assistant.
    An opportunity has already been recorded. A new email has arrived in the same thread.
    Analyze ONLY the new email content and its attachment below.

    Your task is to extract ONLY the fields that are explicitly mentioned in this new email.
    - If the new email mentions a new deadline, return a JSON with just the "Application Deadline" key.
    - If the new email contains no new information about any of the possible fields, return an empty JSON object {{}}.
    - Do not include fields that are not mentioned in the new email.

    POSSIBLE FIELDS: "Application Deadline", "Institution/Company", "Eligibility" ,  "Role Title", "Opportunity Type", "Role Field", "Location", "Work Mode", "Duration", "Time Commitment", "Stipend Details", "Required Skills", "Job Description (JD)", "Application Link".
    
    NEW EMAIL CONTENT:
    ---
    Subject: {new_email_details['subject']}
    Body: {new_email_details['body']}
    ---
    
    NEW ATTACHMENT TEXT:
    ---
    {new_email_details['attachment_text']}
    ---

    JSON Output (only new information):
    """
    response = model.generate_content(prompt)
    return parse_ai_response(response.text)

# --- GOOGLE SHEETS INTERACTION ---

def format_row_from_json(thread_id, email_details, extracted_data):
    """Formats the extracted data into a list for a sheet row."""
    return [
        thread_id, time.strftime("%Y-%m-%d %H:%M:%S"),
        extracted_data.get("Application Deadline", "N/A"), email_details["sender_raw"],
        extracted_data.get("Institution/Company", "N/A"), extracted_data.get("Eligibility", "All"),
        extracted_data.get("Role Title", "N/A"),
        extracted_data.get("Opportunity Type", "N/A"), extracted_data.get("Role Field", "N/A"),
        extracted_data.get("Location", "N/A"), extracted_data.get("Work Mode", "N/A"),
        extracted_data.get("Duration", "N/A"), extracted_data.get("Time Commitment", "N/A"),
        extracted_data.get("Stipend Details", "N/A"), extracted_data.get("Required Skills", "N/A"),
        extracted_data.get("Job Description (JD)", "N/A"), extracted_data.get("Application Link", "N/A"),
        extracted_data.get("Relevance Score (1-10)", "N/A"), email_details["subject"],
    ]
    

def format_opportunity_for_telegram(data_dict):
    """Formats a new opportunity message, fully sanitized for MarkdownV2."""
    # Sanitize all the raw text from the dictionary first
    title = sanitize_telegram_markdown(data_dict.get("Role Title", "N/A"))
    company = sanitize_telegram_markdown(data_dict.get("Institution/Company", "N/A"))
    deadline = sanitize_telegram_markdown(data_dict.get("Application Deadline", "N/A"))
    location = sanitize_telegram_markdown(data_dict.get("Location", "N/A"))
    mode = sanitize_telegram_markdown(data_dict.get("Work Mode", "N/A"))
    commitment = sanitize_telegram_markdown(data_dict.get("Time Commitment", "N/A"))
    stipend = sanitize_telegram_markdown(data_dict.get("Stipend Details", "N/A"))
    eligibility = sanitize_telegram_markdown(data_dict.get("Eligibility", "All"))


    # Now, build the message using our own formatting
    return (f"*Opportunity: {title}*\n"
            f"*Company:* {company}\n\n"
            f"🎓 *Eligibility:* {eligibility}\n"
            f"📅 *Deadline:* {deadline}\n"
            f"📍 *Location:* {location} \\({mode}\\)\n"
            f"💼 *Commitment:* {commitment}\n"
            f"💰 *Stipend:* {stipend}")

def format_update_for_telegram(original_data, merged_data):
    """Formats an update message, fully sanitized for MarkdownV2."""
    # Sanitize the main identifiers
    title = sanitize_telegram_markdown(merged_data.get("Role Title", "N/A"))
    company = sanitize_telegram_markdown(merged_data.get("Institution/Company", "N/A"))
    
    changes = []
    for key, new_value in merged_data.items():
        # Compare against the original data to find what changed
        if key in original_data and new_value != original_data.get(key):
            clean_key = sanitize_telegram_markdown(key.replace('_', ' ').title())
            clean_value = sanitize_telegram_markdown(new_value)
            changes.append(f"\\- *{clean_key}:* {clean_value}")
            
    if not changes:
        return None

    change_summary = "\n".join(changes)
    update_header = sanitize_telegram_markdown("The following details were updated:")
    
    return (f"*Opportunity: {title}*\n"
            f"*Company:* {company}\n\n"
            f"{update_header}\n"
            f"{change_summary}")

def send_telegram_message(message):
    """Sends a message to your Telegram bot using MarkdownV2."""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("Telegram token or chat ID not set. Skipping notification.")
        return
    
    api_url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {'chat_id': TELEGRAM_CHAT_ID, 'text': message, 'parse_mode': 'MarkdownV2'}
    try:
        response = requests.post(api_url, json=payload)
        if response.status_code == 200:
            print("📬 Notification sent successfully!")
        else:
            print(f"Failed to send notification: {response.text}")
    except Exception as e:
        print(f"An error occurred sending notification: {e}")




def sanitize_telegram_markdown(text: str) -> str:
    """Escapes all special characters for Telegram's MarkdownV2."""
    if not isinstance(text, str):
        text = str(text)
    
    escape_chars = r'_*[]()~`>#+-=|{}.!'
    return ''.join(['\\' + char if char in escape_chars else char for char in text])




# --- MAIN EXECUTION ---
def main():
    print(f"🚀 Agent started at {time.strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Authenticate once at the start to get all necessary service objects and credentials
    gmail_service, gspread_client, creds = authenticate_google_services()
    gemini_model = configure_gemini()
    new_opportunities_to_notify = []
    updated_opportunities_to_notify = []
    
    worksheet = gspread_client.open_by_key(SPREADSHEET_ID).sheet1
    
    with open("resume.txt", "r", encoding="utf-8") as f:
        resume_text = f.read()
        
    # 1. Load both memories: processed message IDs and existing thread IDs
    processed_email_ids = load_processed_emails()
    # This line assumes your headers are on row 3, as per your sheet's structure
    existing_thread_ids = set(worksheet.col_values(1)[3:]) 
    print(f"Found {len(processed_email_ids)} recent processed emails and {len(existing_thread_ids)} records in the sheet.")
    
    # 2. Get all emails from the last 24 hours
    messages = get_emails(gmail_service, search_query="newer_than:1h")
    if not messages:
        print("No new emails in the last 1 hours. Exiting.")
        return
        
    print(f"Found {len(messages)} emails from the last 6 hours to check.")
    
    # 3. Loop through all fetched emails
    for msg in messages:
        # --- PRIMARY CHECK: Have we already processed this specific message? ---
        if msg["id"] in processed_email_ids:
            continue

        # If the message is new, get its METADATA only (fast)
        email_metadata = get_email_metadata(gmail_service, msg["id"])
        if not email_metadata:
            save_processed_email(msg["id"], processed_email_ids)
            continue

        # --- FAST FILTERING LOGIC (using only metadata) ---
        if not any(domain in email_metadata["sender_email"] for domain in TRUSTED_DOMAINS):
            print(f"🚫 Skipping external email from {email_metadata['sender_email']}: '{email_metadata['subject']}'")
            save_processed_email(msg["id"], processed_email_ids)
            continue
            
        email_text_for_search = (email_metadata['subject'] + ' ' + email_metadata['body']).lower()
        if not any(keyword in email_text_for_search for keyword in OPPORTUNITY_KEYWORDS):
            print(f"⏭️  Skipping email: '{email_metadata['subject']}' (No keywords found).")
            save_processed_email(msg["id"], processed_email_ids)
            continue
            
        if not is_opportunity_ai_check(gemini_model, email_text_for_search):
            print(f"🤖 AI classified as NOT an opportunity: '{email_metadata['subject']}'")
            save_processed_email(msg["id"], processed_email_ids)
            continue
            
        print(f"✅ Email passed initial filters: '{email_metadata['subject']}'")
        
        # --- ONLY NOW: Fetch and parse attachments (slow) ---
        attachment_text = ""
        if email_metadata['attachment_ids']:
            print("📄 Checking for parsable attachments...")
            attachment_text = fetch_and_parse_attachments(gmail_service, msg["id"], email_metadata['attachment_ids'])
        
        # Build complete email details for extraction
        email_details = {
            "id": email_metadata["id"],
            "threadId": email_metadata["threadId"],
            "sender_raw": email_metadata["sender_raw"],
            "sender_email": email_metadata["sender_email"],
            "subject": email_metadata["subject"],
            "body": email_metadata["body"],
            "attachment_text": attachment_text
        }
        
        # --- DECISION LOGIC: Is this an update or a new entry? ---
        current_thread_id = email_details["threadId"]
        
        if current_thread_id in existing_thread_ids:
            print(f"🔄 Found new email in existing thread. Running update.")
            try:
                cell = worksheet.find(current_thread_id)
                if not cell: continue
                
                # This now correctly assumes headers are on row 3
                column_headers = worksheet.row_values(3)
                row_values = worksheet.row_values(cell.row)
                
                original_data_dict = {header: (row_values[i] if i < len(row_values) else "") for i, header in enumerate(column_headers)}

                updated_data_json = update_details_with_ai(gemini_model, original_data_dict, email_details)
                
                if updated_data_json:
                    has_changed = False
                    merged_data = original_data_dict.copy()
                    cells_to_format = [] 

                    for key, new_value in updated_data_json.items():
                        if key in merged_data and new_value != "N/A" and merged_data.get(key) != new_value:
                            print(f"    -> Updating '{key}' from '{merged_data.get(key)}' to '{new_value}'")
                            merged_data[key] = new_value
                            has_changed = True
                            try:
                                col_index = column_headers.index(key) + 1
                                cells_to_format.append((cell.row, col_index))
                            except ValueError:
                                print(f"    [Warning] Column '{key}' not found in sheet headers.")
                    
                    if has_changed:
                        print("✅ Changes found. Updating sheet and highlighting.")
                        final_row_values = [merged_data.get(h, "") for h in column_headers]
                        worksheet.update(values=[final_row_values], range_name=f'A{cell.row}')
                        
                        updated_opportunities_to_notify.append({
                                'original': original_data_dict,
                                'merged': merged_data
                            })

                        
                        # --- FORMATTING LOGIC ---
                        spreadsheet_id = worksheet.spreadsheet.id
                        requests = []
                        for row, col in cells_to_format:
                            requests.append({
                                "repeatCell": {
                                    "range": {
                                        "sheetId": worksheet.id,
                                        "startRowIndex": row - 1, "endRowIndex": row,
                                        "startColumnIndex": col - 1, "endColumnIndex": col
                                    },
                                    "cell": { "userEnteredFormat": { "textFormat": { "foregroundColor": { "red": 1.0 }}}},
                                    "fields": "userEnteredFormat(textFormat)"
                                }
                            })
                        
                        if requests:
                            sheets_service = build("sheets", "v4", credentials=creds)
                            sheets_service.spreadsheets().batchUpdate(
                                spreadsheetId=spreadsheet_id, body={"requests": requests}
                            ).execute()
                    else:
                        print("No new information to update.")
                else:
                    print("AI did not return a valid update.")

            except Exception as e:
                print(f"An error occurred while updating row: {e}")
        else:
            # It's a new opportunity
            print(f"✨ Processing new opportunity: '{email_details['subject']}'")
            extracted_data = extract_initial_details_with_ai(gemini_model, email_details, resume_text)
            
            if extracted_data:
                print("✅ Successfully extracted details. Appending to Sheet.")
                row_to_append = format_row_from_json(current_thread_id, email_details, extracted_data)
                worksheet.append_row(row_to_append, value_input_option='USER_ENTERED')
                existing_thread_ids.add(current_thread_id)
                new_opportunities_to_notify.append(extracted_data)

            else:
                print("❌ Failed to extract details from this email.")
                
        save_processed_email(msg["id"], processed_email_ids)
        time.sleep(5)
    if new_opportunities_to_notify or updated_opportunities_to_notify:
        # Sanitize the static parts of the message shell
        header = sanitize_telegram_markdown("🔔 AI Agent Alert!")
        summary_message = f"*{header}*\n\n"

        if new_opportunities_to_notify:
            new_opp_header = sanitize_telegram_markdown("Found these new opportunities for you:")
            summary_message += f"{new_opp_header}\n"
            for opp in new_opportunities_to_notify:
                summary_message += format_opportunity_for_telegram(opp) + "\n\n"

        if updated_opportunities_to_notify:
            update_header = sanitize_telegram_markdown("The following opportunities were updated:")
            summary_message += f"{update_header}\n"
            for update_info in updated_opportunities_to_notify:
                update_message = format_update_for_telegram(update_info['original'], update_info['merged'])
                if update_message:
                    summary_message += update_message + "\n\n"

        footer = sanitize_telegram_markdown(f"Find out more details at:")
        link = sanitize_telegram_markdown(GOOGLE_SHEET_LINK)
        summary_message += f"{footer}\n{link}"

        send_telegram_message(summary_message.strip())
    else:
        print("No new or updated opportunities found.")
        
    print("🎉 All new emails checked. Script finished.")


if __name__ == "__main__":
    main()









