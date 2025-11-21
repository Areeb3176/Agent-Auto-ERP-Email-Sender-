import pandas as pd
import os
import re
import requests
import msal
import google.generativeai as genai
from dotenv import load_dotenv
from datetime import datetime
import json
import time
from pathlib import Path

# ---------------- CONFIG ----------------
cert_path = r"C:\Users\muhammad.areeb\AppData\Local\.certifi\cacert.pem"
os.environ.setdefault('SSL_CERT_FILE', cert_path)
os.environ.setdefault('REQUESTS_CA_BUNDLE', cert_path)
os.environ.setdefault('GRPC_DEFAULT_SSL_ROOTS_FILE_PATH', cert_path)

load_dotenv()

# Paths
INPUT_DATA_PATH = os.getenv("INPUT_DATA_PATH", r"C:\ERP_Agent\input")
OUTPUT_DATA_PATH = os.getenv("OUTPUT_DATA_PATH", r"C:\ERP_Agent\output")
LOG_PATH = os.getenv("LOG_PATH", r"C:\ERP_Agent\logs")

# API Keys
GEMINI_API_KEY = os.getenv("GOOGLE_API_KEY")
genai.configure(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None
CHAT_MODEL = "models/gemini-2.5-flash"

# Outlook / Microsoft config
CLIENT_ID = os.getenv("CLIENT_ID")
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPE = ["Mail.Send", "User.Read"]
GRAPH_API = "https://graph.microsoft.com/v1.0"

# Create directories if not exist
Path(INPUT_DATA_PATH).mkdir(parents=True, exist_ok=True)
Path(OUTPUT_DATA_PATH).mkdir(parents=True, exist_ok=True)
Path(LOG_PATH).mkdir(parents=True, exist_ok=True)

# ---------------- LOGGING ----------------
def log_message(message, level="INFO"):
    """Log messages to file and console"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] [{level}] {message}"
    print(log_entry)
    
    log_file = os.path.join(LOG_PATH, f"agent_log_{datetime.now().strftime('%Y%m%d')}.txt")
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(log_entry + "\n")

# ---------------- EMAIL LOG ----------------
def log_email(supplier_name, email, subject, status, note=""):
    """Log email activity to CSV"""
    log_file = os.path.join(LOG_PATH, "email_log.csv")
    
    log_entry = {
        "timestamp": datetime.now().isoformat(),
        "supplier": supplier_name,
        "email": email,
        "subject": subject,
        "status": status,
        "note": note
    }
    
    df_log = pd.DataFrame([log_entry])
    
    if os.path.exists(log_file):
        df_log.to_csv(log_file, mode='a', header=False, index=False)
    else:
        df_log.to_csv(log_file, index=False)
    
    log_message(f"Email logged: {supplier_name} - {status}")

# ---------------- HELPERS ----------------
def ask_gemini(prompt):
    """Call Gemini API"""
    if not GEMINI_API_KEY:
        log_message("Gemini API key not configured", "ERROR")
        return None
    try:
        model = genai.GenerativeModel(CHAT_MODEL)
        resp = model.generate_content(prompt)
        return resp.text if getattr(resp, "text", None) else None
    except Exception as e:
        log_message(f"Gemini error: {str(e)}", "ERROR")
        return None

def get_access_token():
    """Get Microsoft access token for Outlook"""
    if not CLIENT_ID:
        raise Exception("CLIENT_ID not configured in environment (.env).")
    
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    
    # Try to get token from cache first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPE, account=accounts[0])
        if result and "access_token" in result:
            log_message("Token acquired from cache")
            return result["access_token"]
    
    # Device flow for authentication
    flow = app.initiate_device_flow(scopes=SCOPE)
    if "user_code" not in flow:
        raise Exception("Failed to create device flow")
    
    log_message(f"Please visit: {flow['verification_uri']}", "INFO")
    log_message(f"Enter code: {flow['user_code']}", "INFO")
    log_message("Waiting for authentication...", "INFO")
    
    result = app.acquire_token_by_device_flow(flow)
    
    if "access_token" in result:
        log_message("Authentication successful", "INFO")
        return result["access_token"]
    else:
        raise Exception(f"Token error: {result.get('error_description', result)}")

def send_email(token, to_email, subject, body):
    """Send email via Microsoft Graph API"""
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": to_email}}],
        },
        "saveToSentItems": "true"
    }
    
    try:
        resp = requests.post(f"{GRAPH_API}/me/sendMail", headers=headers, json=payload)
        return resp
    except Exception as e:
        log_message(f"Email send error: {str(e)}", "ERROR")
        return None

def generate_email_content(row_data, supplier_name):
    """Generate email subject and body using Gemini"""
    row_ctx = " | ".join([f"{col}: {val}" for col, val in row_data.items() if pd.notna(val)])
    
    prompt = f"""
    Generate a professional reminder email for an OVERDUE order/delivery.
    
    Supplier: {supplier_name}
    Details: {row_ctx}
    
    Return ONLY valid JSON with:
    {{
        "subject": "brief professional subject line",
        "body": "polite 3-4 sentence email body requesting update on overdue item"
    }}
    """
    
    response = ask_gemini(prompt)
    
    if response:
        try:
            text = response.strip()
            text = re.sub(r"^```json|```$", "", text.strip(), flags=re.IGNORECASE).strip()
            parsed = json.loads(text)
            return parsed.get("subject"), parsed.get("body")
        except:
            pass
    
    # Fallback
    return (
        f"Reminder: Overdue Delivery - {supplier_name}",
        f"Dear {supplier_name},\n\nWe noticed that the delivery is overdue. Please provide an update at your earliest convenience.\n\nBest regards."
    )

def find_latest_file(directory):
    """Find the most recent file in directory"""
    files = [os.path.join(directory, f) for f in os.listdir(directory) 
             if f.endswith(('.xlsx', '.xls', '.csv'))]
    
    if not files:
        return None
    
    return max(files, key=os.path.getmtime)

def process_data_file(file_path):
    """Process a single data file"""
    log_message(f"Processing file: {file_path}")
    
    try:
        # Read file
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
        
        log_message(f"Loaded {len(df)} records")
        
        # Find required columns
        status_col = None
        email_col = None
        name_col = None
        
        for col in df.columns:
            if 'status' in col.lower():
                status_col = col
            if 'email' in col.lower():
                email_col = col
            if 'name' in col.lower():
                name_col = col
        
        if not all([status_col, email_col, name_col]):
            log_message(f"Missing required columns. Found: {list(df.columns)}", "ERROR")
            return None
        
        log_message(f"Using columns - Status: {status_col}, Email: {email_col}, Name: {name_col}")
        
        # Add Email Sent column if doesn't exist
        if 'Email Sent' not in df.columns:
            df['Email Sent'] = 'No'
        
        # Filter OVERDUE records that haven't been emailed
        overdue_df = df[
            (df[status_col].astype(str).str.contains('overdue', case=False, na=False)) &
            (df['Email Sent'] != 'Yes')
        ]
        
        log_message(f"Found {len(overdue_df)} OVERDUE records without emails")
        
        if len(overdue_df) == 0:
            log_message("No overdue records to process")
            return df
        
        # Get authentication token once
        log_message("Authenticating with Microsoft...")
        token = get_access_token()
        
        # Process each overdue record
        sent_count = 0
        failed_count = 0
        
        for idx, row in overdue_df.iterrows():
            supplier_name = str(row[name_col])
            supplier_email = str(row[email_col])
            
            log_message(f"Processing: {supplier_name} ({supplier_email})")
            
            # Generate email content
            subject, body = generate_email_content(row.to_dict(), supplier_name)
            
            # Send email
            resp = send_email(token, supplier_email, subject, body)
            
            if resp and resp.status_code == 202:
                log_message(f"✅ Email sent to {supplier_email}", "SUCCESS")
                df.at[idx, 'Email Sent'] = 'Yes'
                log_email(supplier_name, supplier_email, subject, "sent", "Success")
                sent_count += 1
            else:
                error_msg = resp.text if resp else "Unknown error"
                log_message(f"❌ Failed to send to {supplier_email}: {error_msg}", "ERROR")
                log_email(supplier_name, supplier_email, subject, "failed", error_msg)
                failed_count += 1
            
            # Small delay to avoid rate limiting
            time.sleep(2)
        
        log_message(f"Email summary - Sent: {sent_count}, Failed: {failed_count}", "INFO")
        
        return df
        
    except Exception as e:
        log_message(f"Error processing file: {str(e)}", "ERROR")
        return None

def save_updated_file(df, original_file):
    """Save updated dataframe to output folder"""
    filename = os.path.basename(original_file)
    base_name, ext = os.path.splitext(filename)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"{base_name}_updated_{timestamp}{ext}"
    output_path = os.path.join(OUTPUT_DATA_PATH, output_filename)
    
    try:
        if ext == '.csv':
            df.to_csv(output_path, index=False)
        else:
            df.to_excel(output_path, index=False)
        
        log_message(f"✅ Updated file saved: {output_path}", "SUCCESS")
        return output_path
    except Exception as e:
        log_message(f"Error saving file: {str(e)}", "ERROR")
        return None

def run_agent():
    """Main agent loop"""
    log_message("=" * 50)
    log_message("ERP EMAIL AGENT STARTED")
    log_message("=" * 50)
    
    log_message(f"Input folder: {INPUT_DATA_PATH}")
    log_message(f"Output folder: {OUTPUT_DATA_PATH}")
    log_message(f"Log folder: {LOG_PATH}")
    
    # Find latest file
    latest_file = find_latest_file(INPUT_DATA_PATH)
    
    if not latest_file:
        log_message("No data files found in input folder", "WARNING")
        return
    
    log_message(f"Found file: {latest_file}")
    
    # Process file
    updated_df = process_data_file(latest_file)
    
    if updated_df is not None:
        # Save updated file
        output_path = save_updated_file(updated_df, latest_file)
        
        if output_path:
            log_message("Agent task completed successfully", "SUCCESS")
        else:
            log_message("Failed to save output file", "ERROR")
    else:
        log_message("Failed to process data file", "ERROR")
    
    log_message("=" * 50)
    log_message("AGENT FINISHED")
    log_message("=" * 50)

if __name__ == "__main__":
    try:
        run_agent()
    except KeyboardInterrupt:
        log_message("Agent stopped by user", "WARNING")
    except Exception as e:
        log_message(f"Fatal error: {str(e)}", "ERROR")
        import traceback
        log_message(traceback.format_exc(), "ERROR")