import os
from dotenv import load_dotenv
import pandas as pd
from functions import (
    log_message,
    get_access_token,
    find_latest_file,
    process_email_replies,
    process_data_file,
    save_updated_file
)


# ---------------- CONFIG ----------------
load_dotenv()
INPUT_DATA_PATH = os.getenv("INPUT_DATA_PATH")

# SSL Certificate 
cert_path = os.getenv("CERT_PATH", r"C:\Users\muhammad.areeb\AppData\Local\.certifi\cacert.pem") 
os.environ.setdefault('SSL_CERT_FILE', cert_path) 
os.environ.setdefault('REQUESTS_CA_BUNDLE', cert_path) 
os.environ.setdefault('GRPC_DEFAULT_SSL_ROOTS_FILE_PATH', cert_path)

# Claude API Configuration
CLAUDE_API_KEY = os.getenv("CLAUDE_API_KEY")
CLAUDE_MODEL = os.getenv("CLAUDE_MODEL", "claude-sonnet-4-20250514")

# Microsoft Graph Configuration
CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID", "common")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = [
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/Mail.Read"
]
GRAPH_API = "https://graph.microsoft.com/v1.0"
TOKEN_CACHE_FILE = os.path.join(INPUT_DATA_PATH, "token_cache.bin")

# Email Processing Configuration
STATUS_KEYWORDS = [kw.strip() for kw in os.getenv("STATUS_KEYWORDS", "overdue").split(',') if kw.strip()]
REPLY_KEYWORDS = [kw.strip() for kw in os.getenv("REPLY_KEYWORDS", "delivered,shipped,dispatched").split(',') if kw.strip()]
REPLY_CHECK_HOURS = int(os.getenv("REPLY_CHECK_HOURS", "24"))


def run_agent():
    print("=" * 60)
    print("🚀 ERP EMAIL AGENT STARTED (Claude AI)")
    print("=" * 60)
    
    # Ensure input directory exists
    os.makedirs(INPUT_DATA_PATH, exist_ok=True)
    
    # Debug: Check token cache status
    log_message(f"🔍 Checking authentication status...", "INFO")
    log_message(f"   Token cache: {TOKEN_CACHE_FILE}", "INFO")
    
    if os.path.exists(TOKEN_CACHE_FILE):
        size = os.path.getsize(TOKEN_CACHE_FILE)
        writable = os.access(TOKEN_CACHE_FILE, os.W_OK)
        
        if size > 50:  # Valid cache file
            log_message(f"   ✅ Token cache found ({size} bytes)", "SUCCESS")
            log_message(f"   ✅ Should NOT need authentication!", "SUCCESS")
        else:
            log_message(f"   ⚠️ Token cache empty or corrupt", "WARNING")
            log_message(f"   Will need fresh authentication", "INFO")
        
        if not writable:
            log_message(f"   ❌ Token cache NOT writable!", "ERROR")
            log_message(f"   Please check file permissions", "ERROR")
    else:
        log_message(f"   ℹ️ No token cache (first run)", "INFO")
        log_message(f"   One-time authentication required", "INFO")

    # Authenticate
    log_message("\n🔐 Authenticating...", "INFO")
    token = get_access_token(CLIENT_ID, AUTHORITY, SCOPE, TOKEN_CACHE_FILE)
    
    if not token:
        log_message("\n❌ Authentication failed - cannot proceed", "ERROR")
        log_message("💡 Try deleting token cache and re-authenticating:", "INFO")
        log_message(f"   rm {TOKEN_CACHE_FILE}", "INFO")
        return
    
    log_message("✅ Ready to process emails!\n", "SUCCESS")

    # Find latest file
    latest = find_latest_file(INPUT_DATA_PATH)
    if not latest:
        log_message("No Excel/CSV file found in input folder", "WARNING")
        return

    log_message(f"Using file: {os.path.basename(latest)}")

    # Config dict with Claude API
    config = {
        'claude_api_key': CLAUDE_API_KEY,
        'claude_model': CLAUDE_MODEL,
        'graph_api': GRAPH_API,
        'cert_path': cert_path
    }

    # Load data
    df = pd.read_excel(latest) if latest.endswith('.xlsx') else pd.read_csv(latest)
    email_col = next((c for c in df.columns if 'email' in c.lower()), None)
    status_col = next((c for c in df.columns if 'status' in c.lower()), None)
    
    if not email_col:
        log_message("Email column not found in file!", "ERROR")
        return
    
    if not status_col:
        log_message("⚠️ Status column not found - reply analysis will be limited", "WARNING")

    # ====================================================
    # STEP 1: Check for supplier replies and update file
    # ====================================================
    log_message("=" * 60, "INFO")
    log_message("STEP 1: Checking Outlook inbox for supplier replies...", "INFO")
    log_message("=" * 60, "INFO")
    
    # Show current status
    if 'Email Sent' in df.columns:
        sent_count = (df['Email Sent'].astype(str).str.lower() == 'yes').sum()
        replied_count = (df['Email Reply'].astype(str).str.lower() == 'yes').sum() if 'Email Reply' in df.columns else 0
        log_message(f"📊 Current status:", "INFO")
        log_message(f"   - Emails sent: {sent_count}", "INFO")
        log_message(f"   - Replies received: {replied_count}", "INFO")
        log_message(f"   - Awaiting reply: {sent_count - replied_count}", "INFO")
    
    # Process replies
    updated_df = process_email_replies(token, df, email_col, status_col, REPLY_KEYWORDS, STATUS_KEYWORDS, config, latest, REPLY_CHECK_HOURS)
    # File is saved after each status update now
    
    # Show updated status
    if 'Email Reply' in updated_df.columns:
        new_replied_count = (updated_df['Email Reply'].astype(str).str.lower() == 'yes').sum()
        log_message(f"✅ Reply check completed - Total replies now: {new_replied_count}", "SUCCESS")

    # ====================================================
    # STEP 2: Send reminders to overdue entries
    # (First time: send initial email)
    # (Follow-up: send to those who didn't reply)
    # ====================================================
    log_message("=" * 60, "INFO")
    log_message("STEP 2: Processing overdue entries and sending reminders...", "INFO")
    log_message("=" * 60, "INFO")
    
    process_data_file(latest, token, config, STATUS_KEYWORDS)
    
    log_message("=" * 60, "SUCCESS")
    log_message("✅ Agent completed successfully", "SUCCESS")
    log_message("=" * 60, "SUCCESS")

    print("\n" + "=" * 60)
    print("🟢 AGENT FINISHED")
    print("=" * 60)


if __name__ == "__main__":
    try:
        run_agent()
    except KeyboardInterrupt:
        print("\n⚠️ Stopped manually by user")
    except Exception as e:
        import traceback
        print(f"❌ Fatal Error: {e}")
        print(traceback.format_exc())