import pandas as pd
import os
import re
import requests
import msal
from datetime import datetime, timedelta
import json
import time
import webbrowser
import anthropic

# =====================================================
# 🧱 LOGGING
# =====================================================
def log_message(message, level="INFO"):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    prefix = {
        "INFO": "ℹ️",
        "SUCCESS": "✅",
        "WARNING": "⚠️",
        "ERROR": "❌"
    }.get(level, "📝")
    print(f"{prefix} [{ts}] {message}")


# =====================================================
# 🔐 AUTHENTICATION (MSAL)
# =====================================================
def load_cache(token_cache_file):
    cache = msal.SerializableTokenCache()
    if os.path.exists(token_cache_file):
        with open(token_cache_file, 'r') as f:
            cache.deserialize(f.read())
    return cache


def save_cache(cache, token_cache_file):
    if cache.has_state_changed:
        with open(token_cache_file, 'w') as f:
            f.write(cache.serialize())


def get_access_token(client_id, authority, scope, token_cache_file):
    """
    Microsoft authentication with persistent token caching
    """
    cache = load_cache(token_cache_file)
    app = msal.PublicClientApplication(client_id, authority=authority, token_cache=cache)
    
    # First try: Silent authentication (using cached token)
    accounts = app.get_accounts()
    if accounts:
        log_message(f"🔍 Found cached account: {accounts[0].get('username', 'Unknown')}", "INFO")
        result = app.acquire_token_silent(scope, account=accounts[0])
        if result and "access_token" in result:
            log_message("✅ Authentication successful (using cached token)", "SUCCESS")
            save_cache(cache, token_cache_file)
            return result["access_token"]
        else:
            log_message("⚠️ Cached token expired or invalid, refreshing...", "WARNING")
    else:
        log_message("ℹ️ No cached account found, starting fresh authentication", "INFO")

    # Second try: Interactive authentication
    log_message("🔓 Starting device code authentication...", "INFO")
    try:
        flow = app.initiate_device_flow(scopes=scope)
        
        if "user_code" not in flow:
            log_message(f"❌ Failed to create device flow: {flow.get('error_description', 'Unknown error')}", "ERROR")
            return None
        
        print("\n" + "="*60)
        print("🔐 AUTHENTICATION REQUIRED")
        print("="*60)
        print(flow['message'])
        print("="*60)
        print("⏳ Waiting for authentication... (expires in 15 minutes)")
        print("="*60 + "\n")
        
        # Auto-open browser
        try:
            webbrowser.open(flow['verification_uri'])
            log_message("🌐 Browser opened automatically", "INFO")
        except:
            log_message("⚠️ Could not open browser automatically, please open manually", "WARNING")
        
        # Wait for authentication
        result = app.acquire_token_by_device_flow(flow)
        
        if "access_token" in result:
            log_message(f"✅ Authentication successful! Account: {result.get('account', {}).get('username', 'Unknown')}", "SUCCESS")
            save_cache(cache, token_cache_file)
            log_message(f"💾 Token cached to: {token_cache_file}", "INFO")
            log_message("ℹ️ Next time you won't need to authenticate again!", "INFO")
            return result["access_token"]
        else:
            error_msg = result.get("error_description", result.get("error", "Unknown error"))
            log_message(f"❌ Authentication failed: {error_msg}", "ERROR")
            return None
            
    except Exception as e:
        log_message(f"❌ Authentication error: {str(e)}", "ERROR")
        return None


# =====================================================
# 🤖 CLAUDE AI (Email Content Generation & Reply Analysis)
# =====================================================
def ask_claude(prompt, api_key, model="claude-sonnet-4-20250514"):
    """
    Claude API se response lena
    """
    if not api_key:
        log_message("Claude API key missing", "ERROR")
        return None

    try:
        client = anthropic.Anthropic(api_key=api_key)
        
        message = client.messages.create(
            model=model,
            max_tokens=1024,
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        
        return message.content[0].text if message.content else None
        
    except Exception as e:
        log_message(f"Claude API Error: {e}", "ERROR")
        return None


def extract_ets_from_reply(reply_text, api_key, model):
    """
    Reply se ETS (Expected Time of Shipment/Delivery) extract karta hai
    """
    prompt = f"""Analyze this supplier's email reply and extract the ETS (Expected Time of Shipment/Delivery date).

Reply: {reply_text}

Look for:
- Dates mentioned (e.g., "January 15", "15-01-2025", "next Monday")
- ETS/ETD/ETA mentions
- Delivery timeline ("will deliver in 5 days", "shipping next week")

Return ONLY the date in format: DD-MM-YYYY
If no date found, return "Not specified"

Examples:
- "Will ship on 25th January" → 25-01-2025
- "Delivery in 3 days" → [calculate date 3 days from now]
- "No specific date mentioned" → Not specified"""
    
    result = ask_claude(prompt, api_key, model)
    if result:
        # Clean up the response
        ets = result.strip().strip('"').strip("'")
        return ets
    
    return "Not specified"


def analyze_supplier_reply(reply_text, api_key, model):
    """
    Supplier ke reply ko analyze karke best status extract karta hai
    """
    prompt = f"""Analyze this supplier's email reply and extract the delivery status.

Reply: {reply_text}

Based on the reply, determine the best status update. Choose from:
- "Delivered" (if already delivered)
- "In Transit" (if on the way)
- "Dispatched" (if shipped/dispatched)
- "Delayed - [reason]" (if delayed with reason)
- "Expected in X days" (if timeline given)
- "Pending" (if still waiting)
- "Confirmed" (if acknowledged)

Return only the status text, nothing else. Keep it short (max 30 characters).
If no clear status found, return "Reply Received"."""
    
    result = ask_claude(prompt, api_key, model)
    if result:
        # Clean up the response
        status = result.strip().strip('"').strip("'")
        # Limit length
        if len(status) > 50:
            status = status[:47] + "..."
        return status
    
    return "Reply Received"


def generate_consolidated_email(orders_data, supplier_name, api_key, model):
    """
    Ek supplier ke liye multiple orders ko consolidate karke single email generate karta hai
    """
    # Build subject line with all orders
    # Format: CustomerName-Country-Order # 12345-Quote # 789, CustomerName2-Country2-Order # 456-Quote # 101
    subject_parts = []
    for order in orders_data:
        customer = order.get('Customer Name', 'N/A')
        country = order.get('Country', 'N/A')
        order_num = order.get('Order Number', 'N/A')
        quote_num = order.get('Quote Number', 'N/A')
        
        subject_parts.append(f"{customer}-{country}-Order # {order_num}-Quote # {quote_num}")
    
    # Join all orders with comma+space
    subject = ", ".join(subject_parts)
    
    # Build order details for email body
    order_details = []
    for i, order in enumerate(orders_data, 1):
        customer = order.get('Customer Name', 'N/A')
        country = order.get('Country', 'N/A')
        order_num = order.get('Order Number', 'N/A')
        quote_num = order.get('Quote Number', 'N/A')
        
        order_details.append(f"{i}. {customer} - {country} - Order: {order_num}, Quote: {quote_num}")
    
    orders_text = "\n".join(order_details)
    
    prompt = f"""Generate a professional reminder email body for overdue deliveries.
Supplier: {supplier_name}

Multiple orders are overdue:
{orders_text}

Requirements:
1. Email should politely remind about ALL orders listed above
2. Ask for ETS (Expected Time of Shipment) for each order
3. Keep tone professional and courteous
4. Keep it concise (4-6 sentences)

Return ONLY the email body text, no subject line needed."""
    
    result = ask_claude(prompt, api_key, model)
    if result:
        body = result.strip()
    else:
        # Fallback email body
        body = f"""Dear {supplier_name},

We hope this email finds you well. This is a kind reminder regarding {len(orders_data)} overdue orders:

{orders_text}

Could you please provide the ETS (Expected Time of Shipment) for each order and update us on their current status?

We appreciate your prompt response.

Best regards,"""
    
    return subject, body


# =====================================================
# 📧 EMAIL SENDING (Microsoft Graph)
# =====================================================
def send_email(token, to_email, subject, body, graph_api, cert_path):
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    msg = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": to_email}}]
        },
        "saveToSentItems": "true"
    }
    resp = requests.post(f"{graph_api}/me/sendMail", headers=headers, json=msg, verify=cert_path)
    if resp.status_code == 202:
        log_message(f"Email sent to {to_email}", "SUCCESS")
    else:
        log_message(f"Email failed: {resp.status_code}", "ERROR")
    return resp


# =====================================================
# 📂 FILE HANDLING
# =====================================================
def find_latest_file(directory):
    files = [os.path.join(directory, f) for f in os.listdir(directory)
             if f.endswith(('.xlsx', '.xls', '.csv')) and not f.startswith('~')]
    return max(files, key=os.path.getmtime) if files else None


def save_updated_file(df, original_file):
    """
    File ko save karta hai aur verify karta hai ke properly save hui
    """
    try:
        if original_file.endswith('.csv'):
            df.to_csv(original_file, index=False)
        else:
            # Excel ke liye engine specify karo
            df.to_excel(original_file, index=False, engine='openpyxl')
        
        log_message(f"💾 File saved: {os.path.basename(original_file)}", "SUCCESS")
        
        # Wait a bit for file system to sync
        time.sleep(0.5)
        
        return True
    except Exception as e:
        log_message(f"❌ Error saving file: {str(e)}", "ERROR")
        return False


# =====================================================
# 📬 INBOX REPLY HANDLING
# =====================================================
def check_inbox_for_replies(token, graph_api, cert_path, hours_back=24):
    """
    Outlook inbox se emails fetch karta hai (Graph API use karke)
    """
    since = (datetime.utcnow() - timedelta(hours=hours_back)).strftime('%Y-%m-%dT%H:%M:%SZ')
    
    # Inbox messages fetch karo
    url = f"{graph_api}/me/mailFolders/inbox/messages"
    params = {
        '$filter': f'receivedDateTime ge {since}',
        '$select': 'from,subject,body,receivedDateTime,isRead',
        '$top': 100,  # Last 100 emails
        '$orderby': 'receivedDateTime desc'
    }
    
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    
    try:
        r = requests.get(url, headers=headers, params=params, verify=cert_path)
        
        if r.status_code != 200:
            log_message(f"Failed to fetch emails: {r.status_code}", "ERROR")
            log_message(f"Response: {r.text[:200]}", "ERROR")
            return []

        data = r.json().get('value', [])
        log_message(f"📥 {len(data)} emails fetched from Outlook Inbox (last {hours_back} hours)", "SUCCESS")
        
        # Debug: Show first few emails
        if data:
            log_message("📧 Recent emails:", "INFO")
            for i, email in enumerate(data[:5]):  # Show first 5
                sender = email.get('from', {}).get('emailAddress', {}).get('address', 'Unknown')
                subject = email.get('subject', 'No subject')
                received = email.get('receivedDateTime', '')
                log_message(f"   {i+1}. From: {sender}", "INFO")
                log_message(f"      Subject: {subject[:60]}", "INFO")
                log_message(f"      Received: {received}", "INFO")
        
        return data
        
    except Exception as e:
        log_message(f"Error fetching inbox: {str(e)}", "ERROR")
        return []


def process_email_replies(token, df, email_col, status_col, REPLY_KEYWORDS, STATUS_KEYWORDS, config, latest_file, hours_back=24):
    """
    Outlook inbox mein supplier replies check karta hai aur status + ETS update karta hai
    """
    emails = check_inbox_for_replies(token, config['graph_api'], config['cert_path'], hours_back)

    if not emails:
        log_message("⚠️ No emails found in inbox to process", "WARNING")
        return df

    # ETS column dhundo ya banao
    ets_col = next((c for c in df.columns if 'ets' in c.lower()), None)
    if not ets_col:
        log_message("⚠️ ETS column not found - will add one", "WARNING")
        df['ETS'] = ''
        ets_col = 'ETS'

    matched = 0
    log_message(f"🔍 Processing {len(emails)} emails for supplier replies...", "INFO")
    log_message("-" * 60, "INFO")
    
    # Create a dictionary to store processed emails to avoid duplicates
    processed_emails = {}
    updates_made = False
    
    # First, process all emails
    for email in emails:
        sender_email = (email.get('from', {}).get('emailAddress', {}).get('address') or '').lower().strip()
        if not sender_email:
            continue
            
        # Find ALL matching overdue orders for this supplier
        matching_rows = df[
            (df[email_col].str.lower().str.strip() == sender_email) &
            (df[status_col].astype(str).str.contains('|'.join(STATUS_KEYWORDS), case=False, na=False))
        ]
        
        if len(matching_rows) == 0:
            continue
        
        # Skip if we already processed an email from this supplier
        if sender_email in processed_emails:
            continue
            
        supplier_name = str(matching_rows.iloc[0].get('Supplier Name', sender_email))
        log_message(f"\n🔎 Processing email from: {supplier_name} ({sender_email})", "INFO")
        log_message(f"   📦 Found {len(matching_rows)} overdue orders from this supplier", "INFO")
        
        # Process this email
        subject = (email.get('subject') or '').lower()
        raw_body = (email.get('body', {}) or {}).get('content', '')
        clean_body = re.sub(r'<[^>]+>', '', raw_body)
        received_time = email.get('receivedDateTime', '')
        
        log_message(f"   📧 Email details:", "INFO")
        log_message(f"      Subject: {subject[:60]}", "INFO")
        log_message(f"      Received: {received_time}", "INFO")
        
        # Debug: Show body preview
        log_message(f"      Body preview: {clean_body[:200]}...", "INFO")
        
        # Check for reply keywords
        matched_keywords = [kw for kw in REPLY_KEYWORDS if kw.lower() in clean_body.lower() or kw.lower() in subject.lower()]
        has_keyword = len(matched_keywords) > 0
        
        if matched_keywords:
            log_message(f"      ✅ Found keywords: {', '.join(matched_keywords)}", "SUCCESS")
        else:
            log_message(f"      ⚠️ No reply keywords found, skipping...", "WARNING")
            continue
        
        # Agar reply keywords match hote hain
        if has_keyword:
            matched += 1
            
            # 🤖 EXTRACT ETS FROM REPLY
            log_message(f"      🤖 Extracting ETS with Claude AI...", "INFO")
            new_ets = extract_ets_from_reply(
                clean_body, 
                config['claude_api_key'], 
                config['claude_model']
            )
            log_message(f"      📅 ETS extracted: {new_ets}", "SUCCESS")
            
            # 🤖 ANALYZE REPLY AND UPDATE STATUS
            log_message(f"      🤖 Analyzing reply status with Claude AI...", "INFO")
            new_status = analyze_supplier_reply(
                clean_body, 
                config['claude_api_key'], 
                config['claude_model']
            )
            
            # UPDATE ALL MATCHING ORDERS
            for idx in matching_rows.index:
                try:
                    old_status = str(df.at[idx, status_col])
                    old_ets = str(df.at[idx, ets_col]) if ets_col else "N/A"
                    
                    # Update status
                    df.at[idx, status_col] = new_status
                    
                    # Update ETS
                    if ets_col:
                        df.at[idx, ets_col] = new_ets
                    
                    updates_made = True
                    
                    order_num = df.at[idx, 'Order Number'] if 'Order Number' in df.columns else idx
                    log_message(f"      📝 Order {order_num}: Status '{old_status}' → '{new_status}'", "SUCCESS")
                    if ets_col:
                        log_message(f"         ETS '{old_ets}' → '{new_ets}'", "SUCCESS")
                    
                except Exception as e:
                    log_message(f"      ❌ Error updating row {idx}: {str(e)}", "ERROR")
            
            log_message(f"      ✅ REPLY PROCESSED - Updated {len(matching_rows)} orders!", "SUCCESS")
            
            # Mark this email as processed for this supplier
            processed_emails[sender_email] = True

    # Save file ONCE at the end if any updates were made
    if updates_made:
        log_message("\n💾 Saving all updates to file...", "INFO")
        if save_updated_file(df, latest_file):
            # Verify save
            time.sleep(1)  # Wait for file to be fully written
            try:
                check_df = pd.read_excel(latest_file) if latest_file.endswith('.xlsx') else pd.read_csv(latest_file)
                log_message(f"✅ File saved successfully - verified {len(check_df)} rows", "SUCCESS")
            except Exception as e:
                log_message(f"⚠️ Could not verify file save: {str(e)}", "WARNING")
        else:
            log_message(f"❌ Failed to save file!", "ERROR")
    
    log_message("-" * 60, "INFO")
    log_message(f"📊 Total suppliers processed: {matched}", "SUCCESS" if matched > 0 else "INFO")
    
    return df


# =====================================================
# 🚀 SEND CONSOLIDATED REMINDERS
# =====================================================
def process_data_file(file_path, token, config, STATUS_KEYWORDS):
    # CRITICAL: Read fresh data from file after reply processing
    log_message(f"📂 Reading FRESH data from file...", "INFO")
    time.sleep(1)  # Small delay to ensure file is fully written
    
    df = pd.read_excel(file_path) if file_path.endswith('.xlsx') else pd.read_csv(file_path)
    
    log_message(f"📊 Loaded {len(df)} total rows from file", "INFO")

    status_col = next((c for c in df.columns if 'status' in c.lower()), None)
    email_col = next((c for c in df.columns if 'email' in c.lower()), None)
    name_col = next((c for c in df.columns if 'name' in c.lower()), None)

    if not status_col or not email_col or not name_col:
        log_message("Missing one of the required columns (status, email, name)", "ERROR")
        return df

    # Find overdue rows that haven't received a reply yet
    pattern = '|'.join(STATUS_KEYWORDS)
    
    # EXCLUDE rows that have reply-related statuses
    reply_statuses = ['Reply Received', 'Delivered', 'In Transit', 'Dispatched', 'Delayed', 'Expected', 'Confirmed']
    
    log_message(f"🔍 Checking for overdue orders without replies...", "INFO")
    log_message(f"   Searching for status containing: {STATUS_KEYWORDS}", "INFO")
    log_message(f"   Excluding statuses: {reply_statuses}", "INFO")
    
    overdue = df[
        (df[status_col].astype(str).str.contains(pattern, case=False, na=False))
    ]
    
    log_message(f"   Found {len(overdue)} rows with overdue status", "INFO")
    
    # Now filter out those that have reply statuses
    overdue_no_reply = overdue[
        ~overdue[status_col].astype(str).str.contains('|'.join(reply_statuses), case=False, na=False)
    ]
    
    log_message(f"   After excluding replies: {len(overdue_no_reply)} rows need emails", "INFO")

    if len(overdue_no_reply) == 0:
        log_message("✅ No pending overdue orders - all have received replies!", "SUCCESS")
        return df

    # Group by supplier email
    supplier_groups = overdue_no_reply.groupby(email_col)
    
    log_message(f"📊 Found {len(supplier_groups)} suppliers needing reminder emails", "INFO")
    
    emails_sent = 0
    total_orders = 0
    
    for supplier_email, group in supplier_groups:
        supplier_email = str(supplier_email).strip()
        
        if not supplier_email or '@' not in supplier_email:
            log_message(f"Skipping invalid email: {supplier_email}", "WARNING")
            continue
        
        supplier_name = str(group.iloc[0][name_col]).strip()
        num_orders = len(group)
        total_orders += num_orders
        
        log_message(f"\n📧 Preparing email for {supplier_name} ({supplier_email})", "INFO")
        log_message(f"   📦 {num_orders} overdue orders to include", "INFO")
        
        # Show which orders are being included
        for idx, row in group.iterrows():
            order_num = row.get('Order Number', 'N/A')
            current_status = row.get(status_col, 'N/A')
            log_message(f"      • Order {order_num}: {current_status}", "INFO")
        
        # Prepare order data for email generation
        orders_data = []
        for idx, row in group.iterrows():
            order_info = {
                'Customer Name': row.get('Customer Name', 'N/A'),
                'Country': row.get('Country', 'N/A'),
                'Order Number': row.get('Order Number', 'N/A'),
                'Quote Number': row.get('Quote Number', 'N/A')
            }
            orders_data.append(order_info)
        
        # Generate consolidated email using Claude
        subject, body = generate_consolidated_email(
            orders_data,
            supplier_name,
            config['claude_api_key'],
            config['claude_model']
        )
        
        log_message(f"   📝 Subject: {subject}", "INFO")
        
        # Send the email
        response = send_email(token, supplier_email, subject, body, config['graph_api'], config['cert_path'])
        
        if response.status_code == 202:
            emails_sent += 1
            log_message(f"   ✅ Consolidated email sent successfully!", "SUCCESS")
        else:
            log_message(f"   ❌ Failed to send email", "ERROR")
            
        time.sleep(2)  # Rate limiting

    log_message(f"\n✅ Summary: {emails_sent} emails sent covering {total_orders} orders.", "SUCCESS")
    
    return df