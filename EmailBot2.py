import os
import sys
import time
import json
import base64
import uuid
import logging
import requests
import msal
import html2text
import re
import google.generativeai as genai
from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv

# --- Step 0: Load Environment Variables ---
load_dotenv()

# === 1. Logging Setup ===
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)

# === 2. Configuration ===
CLIENT_ID = os.getenv("CLIENT_ID", "") # You will have to add your own client id
TENANT_ID = os.getenv("TENANT_ID", "") # You will have to add your own tenant id
EXCEL_SHARE_LINK = os.getenv("EXCEL_SHARE_LINK", "") ## You will have you own EXCEL link, the readme includes what to put in it
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "") # You will have to generate your own gemini API key.
SHEET_OPPORTUNITIES = "OpportunitiesMaster"
SHEET_INTERACTIONS = "InteractionLog"
TOKEN_CACHE_FILE = "msal_token_cache.bin"
TIMESTAMP_FILE = "last_run_timestamp.txt" 
PROCESSED_EMAILS_FILE = "processed_emails.json"  # Track processed emails to prevent duplicates
SCOPES = ["User.Read", "Mail.Read", "Files.ReadWrite.All"] # You will have to allow these in microsoft AZURE. If you dont do that then it will not work as it needs it to read your mail and extract the data from it.

# === 3. Helper Functions ===
html_converter = html2text.HTML2Text()
html_converter.ignore_links = True
html_converter.body_width = 0

def get_access_token(client_id, tenant_id):
    """Handles MSAL authentication and token acquisition."""
    token_cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        with open(TOKEN_CACHE_FILE, "r") as f: token_cache.deserialize(f.read())
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.PublicClientApplication(client_id, authority=authority, token_cache=token_cache)
    accounts = app.get_accounts()
    token_response = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
    if not token_response:
        logging.info("🔐 No cached token found. Launching device login...")
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow: raise ValueError("Failed to initiate device flow.")
        logging.info(f"🔐 Go to {flow['verification_uri']} and enter code: {flow['user_code']}")
        token_response = app.acquire_token_by_device_flow(flow)
    if token_cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f: f.write(token_cache.serialize())
    if "access_token" not in token_response:
        raise ConnectionError(f"Token error: {token_response.get('error_description')}")
    logging.info("✅ Access token acquired.")
    return {"Authorization": f"Bearer {token_response['access_token']}"}

def get_excel_file_id(share_link, headers):
    """Converts a SharePoint share link to a drive item ID."""
    encoded_bytes = base64.b64encode(share_link.encode('utf-8'))
    share_id = f"u!{encoded_bytes.decode('utf-8').replace('+', '-').replace('/', '_').rstrip('=')}"
    logging.info("🔎 Resolving SharePoint link to file ID...")
    api_url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem"
    response = requests.get(api_url, headers=headers)
    response.raise_for_status()
    logging.info("✅ Successfully resolved file ID.")
    return response.json()['id']

def load_processed_emails():
    """Load the set of already processed email IDs."""
    try:
        with open(PROCESSED_EMAILS_FILE, 'r') as f:
            return set(json.load(f))
    except FileNotFoundError:
        return set()

def save_processed_emails(processed_emails):
    """Save the set of processed email IDs."""
    with open(PROCESSED_EMAILS_FILE, 'w') as f:
        json.dump(list(processed_emails), f)

def get_all_historical_emails(headers, months_back=6):
    """Fetch all emails from the specified months back for comprehensive matching."""
    cutoff_date = (datetime.now(timezone.utc) - timedelta(days=months_back * 30)).strftime('%Y-%m-%dT%H:%M:%SZ')
    
    logging.info(f"📚 Fetching historical emails from {cutoff_date} for comprehensive matching...")
    
    graph_url = (
        f"https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?"
        f"$filter=receivedDateTime gt {cutoff_date}&"
        "$orderby=receivedDateTime desc&"
        "$top=1000"  # Increase limit for historical data
    )
    
    all_emails = []
    while graph_url:
        response = requests.get(graph_url, headers=headers)
        response.raise_for_status()
        data = response.json()
        emails = data.get("value", [])
        
        # Filter out internal emails early
        filtered_emails = []
        for email in emails:
            sender_email = email.get("from", {}).get("emailAddress", {}).get("address", "").lower()
            if "@eucloid.com" not in sender_email and "noreply" not in sender_email:
                filtered_emails.append({
                    'id': email.get('id'),
                    'subject': email.get('subject', 'No Subject'),
                    'body': html_converter.handle(email.get('body', {}).get('content', '')),
                    'sender_email': sender_email,
                    'sender_name': email.get("from", {}).get("emailAddress", {}).get("name", sender_email),
                    'received_date': email.get('receivedDateTime'),
                    'conversation_id': email.get('conversationId')
                })
        
        all_emails.extend(filtered_emails)
        graph_url = data.get("@odata.nextLink")  # Handle pagination
    
    logging.info(f"📚 Retrieved {len(all_emails)} historical emails for matching.")
    return all_emails

def parse_email_for_opportunities(subject, body, sender_email):
    """Uses Gemini to extract a list of opportunities from an email."""
    if not GEMINI_API_KEY or "YOUR_GEMINI_API_KEY" in GEMINI_API_KEY:
        raise ValueError("GEMINI_API_KEY is not set in configuration.")
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel('gemini-1.5-flash')
    prompt = f"""
You are a CRM assistant. Given the email below, extract all distinct sales opportunities. For each opportunity, return: title, summary, action_item, contact_name, contact_company, and contact_email. If no opportunities are found, return an empty list: []

A sales opportunity is defined as:
- A potential business deal or project
- Request for proposal/quote
- Product inquiry with commercial intent
- Service request that could lead to revenue
- Partnership discussion with business potential

Exclude:
- General inquiries without clear commercial intent
- Support requests
- Administrative communications
- Social/networking emails

Respond ONLY in valid JSON format: [{{...}}]

Email Content:
Subject: {subject}
Sender: {sender_email}
Body: {body[:2000]}
"""
    try:
        response = model.generate_content(prompt)
        clean_response = response.text.strip().replace("```json", "").replace("```", "")
        return json.loads(clean_response)
    except Exception as e:
        logging.error(f"❌ Gemini parsing failed: {e}"); return []

def get_existing_opportunities_for_ai(headers, file_id):
    """Fetches existing opportunities for the AI contextual match."""
    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets('{SHEET_OPPORTUNITIES}')/usedRange(valuesOnly=true)"
    try:
        res = requests.get(url, headers=headers)
        res.raise_for_status()
        values = res.json().get("values", [])
        opportunity_list = []
        for row in values[1:]:  # Skip header
            if len(row) > 9:
                opp_id, company, title, summary = row[0], row[2], row[5], row[9]
                opportunity_list.append({
                    "id": opp_id, 
                    "summary": summary, 
                    "title": title, 
                    "company": company
                })
        logging.info(f"🧾 Found {len(opportunity_list)} existing opportunities for AI matching.")
        return opportunity_list
    except Exception as e:
        logging.error(f"❌ Error fetching from Excel: {e}"); 
        return []

def find_related_opportunity_with_ai(new_opportunity, existing_opportunities, historical_emails):
    """Uses AI to determine if a new opportunity is a follow-up to an existing one using comprehensive data."""
    if not existing_opportunities and not historical_emails: 
        return None, None
    
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    # Prepare existing opportunities context
    existing_list_str = ""
    if existing_opportunities:
        existing_list_str = "EXISTING OPPORTUNITIES:\n" + "\n".join([
            f"- ID: {opp['id']}, Company: {opp['company']}, Title: {opp['title']}, Summary: {opp['summary'][:200]}"
            for opp in existing_opportunities[:20]  # Limit to prevent token overflow
        ])
    
    # Prepare historical emails context (focus on relevant ones)
    historical_context = ""
    relevant_historical = []
    if historical_emails:
        # Filter historical emails that might be relevant based on sender or keywords
        sender_company = (new_opportunity.get('contact_company') or '').lower()
        sender_email = new_opportunity.get('contact_email', '').lower()
        title_keywords = new_opportunity.get('title', '').lower().split()
        
        for email in historical_emails[:50]:  # Limit for performance
            email_content = f"{email['subject']} {email['body']}".lower()
            email_sender = email['sender_email'].lower()
            
            # Check for relevance
            if (sender_company and sender_company in email_content) or \
               (sender_email and sender_email == email_sender) or \
               any(keyword in email_content for keyword in title_keywords if len(keyword) > 3):
                relevant_historical.append(email)
        
        if relevant_historical:
            historical_context = "\n\nRELEVANT HISTORICAL EMAILS:\n" + "\n".join([
                f"- Date: {email['received_date'][:10]}, From: {email['sender_name']}, Subject: {email['subject']}, Preview: {email['body'][:200]}..."
                for email in relevant_historical[:10]  # Further limit
            ])
    
    prompt = f"""
You are a highly intelligent CRM de-duplication assistant. Your goal is to prevent duplicate opportunities and correctly link related communications.

CRITICAL MATCHING RULES - READ CAREFULLY:
1. **CONTENT MUST EXPLICITLY MENTION THE SAME PROJECT/PRODUCT**: The email must contain specific keywords, project names, product names, or technical details that directly reference the existing opportunity
2. **SENDER IDENTITY ALONE IS NEVER ENOUGH**: Just because someone sent emails about Project A before does NOT mean their new generic email is about Project A
3. **GENERIC COMMUNICATION = NO MATCH**: Generic messages like "thanks", "quick question", "status update", "did you get my message" should NEVER match unless they explicitly mention the project details
4. **BURDEN OF PROOF**: You must find SPECIFIC CONTENT OVERLAP, not just assume relationship

MANDATORY CONTENT ANALYSIS:
- Does the new email mention specific project names, products, or technical terms from the existing opportunity?
- Does it reference specific companies, contracts, or business details?
- Is there clear topical connection beyond just being from the same person?

NEW EMAIL TO ANALYZE:
- Title: "{new_opportunity.get('title', 'NA')}"
- Content: "{new_opportunity.get('summary', 'NA')}"
- From: "{new_opportunity.get('contact_email', 'NA')}"

{existing_list_str}

{historical_context}

STRICT INSTRUCTION: Only return match=true if you can identify specific content words/phrases in the new email that directly relate to an existing opportunity's project/product. Generic communication from known senders should be treated as separate opportunities.

Respond ONLY with valid JSON: {{"match": true/false, "opportunity_id": "The ID of the best match or null", "confidence": 0.0-1.0, "reason": "Brief explanation focusing on specific content analysis"}}
"""
    
    try:
        logging.info("🤔 Performing comprehensive AI match analysis...")
        response = model.generate_content(prompt)
        clean_response = response.text.strip().replace("```json", "").replace("```", "")
        result = json.loads(clean_response)
        
        if result.get("match") and result.get("confidence", 0) >= 0.9:
            logging.info(f"✅ High confidence match found: {result.get('reason', 'No reason provided')}")
            return result.get("opportunity_id"), relevant_historical
        elif result.get("match"):
            logging.info(f"⚠️ Low confidence match rejected: {result.get('reason', 'No reason provided')}")
        
        return None, relevant_historical
    except Exception as e:
        logging.error(f"❌ AI contextual match failed: {e}"); 
        return None, relevant_historical

def find_earliest_mention(opportunity_data, relevant_historical_emails):
    """Finds the earliest mention of this opportunity in historical emails using AI."""
    if not relevant_historical_emails:
        return None
    
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    # Sort historical emails by date (oldest first)
    sorted_emails = sorted(relevant_historical_emails, key=lambda x: x['received_date'])
    
    # Prepare email context for AI analysis
    emails_context = "\n".join([
        f"Email {i+1}: Date: {email['received_date']}, Subject: {email['subject']}, From: {email['sender_name']}, Content: {email['body'][:300]}..."
        for i, email in enumerate(sorted_emails[:15])  # Limit to prevent token overflow
    ])
    
    prompt = f"""
You are analyzing historical emails to find the FIRST mention of a specific business opportunity.

CURRENT OPPORTUNITY:
- Company: {opportunity_data.get('contact_company', 'NA')}
- Title: {opportunity_data.get('title', 'NA')}
- Summary: {opportunity_data.get('summary', 'NA')}

HISTORICAL EMAILS (sorted by date, oldest first):
{emails_context}

Your task: Identify which email (if any) represents the FIRST mention of this specific opportunity/project.

Rules:
1. Look for the same project/product/service discussion
2. Same company context
3. Must be about the SAME business opportunity, not just general communication
4. Return the email number (1-based) of the FIRST relevant mention
5. If no clear first mention found, return null

Respond ONLY with valid JSON: {{"first_mention_email_number": number_or_null, "confidence": 0.0-1.0, "reason": "Brief explanation"}}
"""
    
    try:
        logging.info("🕐 Analyzing historical emails to find earliest mention...")
        response = model.generate_content(prompt)
        clean_response = response.text.strip().replace("```json", "").replace("```", "")
        result = json.loads(clean_response)
        
        email_number = result.get("first_mention_email_number")
        confidence = result.get("confidence", 0)
        
        if email_number and confidence >= 0.7 and email_number <= len(sorted_emails):
            earliest_email = sorted_emails[email_number - 1]  # Convert to 0-based index
            logging.info(f"🕐 Found earliest mention on {earliest_email['received_date'][:10]} with {confidence:.1%} confidence")
            return earliest_email['received_date']
        else:
            logging.info("🕐 No clear earliest mention found in historical emails")
            return None
            
    except Exception as e:
        logging.error(f"❌ Error finding earliest mention: {e}")
        return None

def read_last_run_timestamp():
    """Always process emails from the last 24 hours to ensure consistency."""
    # Always look back 24 hours to ensure we don't miss emails
    return (datetime.now(timezone.utc) - timedelta(hours=24)).strftime('%Y-%m-%dT%H:%M:%SZ')

def write_last_run_timestamp(timestamp):
    """Writes the timestamp of the current run to a file."""
    with open(TIMESTAMP_FILE, 'w') as f:
        f.write(timestamp)
    logging.info(f"✅ Timestamp {timestamp} saved for next run.")

def append_rows_to_excel(rows, table_name, sheet_name, file_id, headers):
    """Inserts new rows at the top of a specified table in an Excel sheet."""
    if not rows: return
    
    logging.info(f"📝 Inserting {len(rows)} new row(s) at the top of table '{table_name}'...")
    
    # Reverse the list of rows so the newest email ends up at the very top (row 0)
    for row_data in reversed(rows):
        url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets('{sheet_name}')/tables('{table_name}')/rows/add"
        
        # The 'index: 0' tells the API to insert this row at the top
        data = {
            "values": [row_data],
            "index": 0
        }
        
        res = requests.post(url, headers=headers, json=data)
        
        if res.status_code != 201:
            logging.error(f"❌ Failed to insert row into {table_name}: {res.text}")
        else:
            logging.info(f"✅ Successfully inserted 1 row into {table_name}.")

# === MAIN WORKFLOW ===
def main():
    """Main execution function with enhanced duplicate prevention and comprehensive matching."""
    current_run_timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    
    try:
        # Load processed emails to prevent duplicates
        processed_emails = load_processed_emails()
        logging.info(f"📋 Loaded {len(processed_emails)} previously processed email IDs.")
        
        headers = get_access_token(CLIENT_ID, TENANT_ID)
        excel_file_id = get_excel_file_id(EXCEL_SHARE_LINK, headers)
        
        # Get existing opportunities from Excel
        existing_opportunities_list = get_existing_opportunities_for_ai(headers, excel_file_id)
        
        # Get comprehensive historical email data for better matching
        historical_emails = get_all_historical_emails(headers, months_back=6)
        
        # Get emails from last 24 hours for processing
        time_24_hours_ago = (datetime.now(timezone.utc) - timedelta(hours=24)).strftime('%Y-%m-%dT%H:%M:%SZ')
        graph_url = (
        f"https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?"
        f"$filter=receivedDateTime ge {time_24_hours_ago}&"
        "$orderby=receivedDateTime desc"
        )           
        response = requests.get(graph_url, headers=headers)
        response.raise_for_status()
        messages = response.json().get("value", [])
        logging.info(f"📨 Found {len(messages)} emails from last 24 hours.")

        # Filter out already processed emails and internal emails
        new_messages = []
        for msg in messages:
            msg_id = msg.get('id')
            sender_email = msg.get("from", {}).get("emailAddress", {}).get("address", "").lower()
            
            if msg_id in processed_emails:
                continue  # Skip already processed
                
            #if "@eucloid.com" in sender_email or "noreply" in sender_email:
             #   processed_emails.add(msg_id)  # Mark as processed but skip
              #  continue
                
            new_messages.append(msg)

        logging.info(f"📨 {len(new_messages)} new emails to process after filtering.")

        if not new_messages:
            logging.info("✅ No new emails to process.")
            save_processed_emails(processed_emails)
            write_last_run_timestamp(current_run_timestamp)
            return

        new_opportunity_rows = []
        interaction_rows = []

        for msg in new_messages:
            msg_id = msg.get('id')
            subject = msg.get("subject", "No Subject")
            sender_obj = msg.get("from", {}).get("emailAddress", {})
            sender_email = sender_obj.get("address", "").lower()
            sender_name = sender_obj.get("name", sender_email)
            received_dt = msg.get("receivedDateTime")
            conv_id = msg.get("conversationId")

            logging.info(f"\n📨 Processing email: '{subject}' from {sender_name}")

            body_html = msg.get("body", {}).get("content", "")
            body_text = html_converter.handle(body_html)
            
            # Parse for opportunities
            opportunities = parse_email_for_opportunities(subject, body_text, sender_email)
            
            if opportunities:
                logging.info(f"✅ Found {len(opportunities)} opportunities in '{subject}'.")
                for opp in opportunities:
                    # Enhanced opportunity object for matching
                    enhanced_opp = {
                        **opp,
                        'contact_email': opp.get('contact_email') or sender_email,
                        'sender_name': sender_name,
                        'email_subject': subject
                    }
                    
                    opp_id, relevant_emails = find_related_opportunity_with_ai(
                        enhanced_opp, 
                        existing_opportunities_list, 
                        historical_emails
                    )
                    
                    if opp_id:
                        logging.info(f"🤖 Matched to existing Opportunity ID '{opp_id}' via comprehensive AI analysis.")
                        interaction_rows.append([
                            opp_id, received_dt, "Follow-up", "Email", sender_name, 
                            opp.get("summary", "N/A")[:500], opp.get("action_item", "N/A"), ""
                        ])
                    else:
                        opp_id = str(uuid.uuid4())
                        logging.info(f"✅ Creating new Opportunity ID '{opp_id}'.")
                        
                        # Find the earliest mention of this opportunity
                        earliest_mention_date = find_earliest_mention(enhanced_opp, relevant_emails)
                        first_mention_date = earliest_mention_date if earliest_mention_date else received_dt
                        
                        logging.info(f"📅 First mention date for opportunity: {first_mention_date[:10]}")
                        
                        contact_email = enhanced_opp.get("contact_email", "").strip()
                        # Note: The 8th column (index 7) in OpportunitiesMaster should be "First Mention Date"
                        # This represents when this opportunity was first mentioned, not when it was created in CRM
                        new_opportunity_rows.append([
                            opp_id, opp.get("contact_name", sender_name), 
                            opp.get("contact_company", "NA"), contact_email,
                            "", opp.get("title", subject), "New Lead", first_mention_date, conv_id, 
                            opp.get("summary", "N/A")
                        ])
                        interaction_rows.append([
                            opp_id, received_dt, "New Lead", "Email", sender_name, 
                            opp.get("summary", "N/A")[:500], opp.get("action_item", "N/A"), ""
                        ])
                        # Add to existing opportunities for subsequent matching in this run
                        existing_opportunities_list.append({
                            "id": opp_id, "summary": opp.get("summary", "N/A"), 
                            "title": opp.get("title", subject), "company": opp.get("contact_company", "NA")
                        })
            else:
                # Check if it's a follow-up to existing opportunity
                logging.info("ℹ️ No new opportunities found. Checking for follow-ups...")
                temp_opp = {
                    "title": subject, 
                    "summary": body_text[:500], 
                    "contact_company": "NA",
                    "contact_email": sender_email,
                    "sender_name": sender_name
                }
                
                opp_id, relevant_emails = find_related_opportunity_with_ai(
                    temp_opp, 
                    existing_opportunities_list, 
                    historical_emails
                )
                
                if opp_id:
                    logging.info(f"🤖 Linked general email to Opportunity ID '{opp_id}'.")
                    interaction_rows.append([
                        opp_id, received_dt, "General Communication", "Email", sender_name, 
                        body_text[:500], "Review", ""
                    ])
                else:
                    # CREATE NEW OPPORTUNITY FOR GENERAL EMAIL
                    opp_id = str(uuid.uuid4())
                    logging.info(f"✅ Creating new Opportunity ID '{opp_id}' for general email.")
                    
                    # Find the earliest mention of this general communication
                    earliest_mention_date = find_earliest_mention(temp_opp, relevant_emails)
                    first_mention_date = earliest_mention_date if earliest_mention_date else received_dt
                    
                    logging.info(f"📅 First mention date for general opportunity: {first_mention_date[:10]}")
                    
                    # Create new opportunity row for general email
                    new_opportunity_rows.append([
                        opp_id, sender_name, "NA", sender_email,
                        "", subject, "General Communication", first_mention_date, conv_id, 
                        body_text[:500]
                    ])
                    interaction_rows.append([
                        opp_id, received_dt, "General Communication", "Email", sender_name, 
                        body_text[:500], "Review", ""
                    ])
                    # Add to existing opportunities for subsequent matching in this run
                    existing_opportunities_list.append({
                        "id": opp_id, "summary": body_text[:500], 
                        "title": subject, "company": "NA"
                    })

            # Mark email as processed
            processed_emails.add(msg_id)

        # Save to Excel
        if new_opportunity_rows or interaction_rows:
            append_rows_to_excel(new_opportunity_rows, "OpportunitiesTable", SHEET_OPPORTUNITIES, excel_file_id, headers)
            append_rows_to_excel(interaction_rows, "InteractionsTable", SHEET_INTERACTIONS, excel_file_id, headers)
        
        # Save processed emails and timestamp
        save_processed_emails(processed_emails)
        write_last_run_timestamp(current_run_timestamp)
        
        logging.info(f"\n--- Cycle Complete ---")
        logging.info(f"✅ Processed {len(new_messages)} emails")
        logging.info(f"✅ Created {len(new_opportunity_rows)} new opportunities")
        logging.info(f"✅ Logged {len(interaction_rows)} interactions")

    except Exception as e:
        logging.error(f"❌ A critical error occurred in the main process: {e}", exc_info=True)
        raise

if __name__ == "__main__":
    main()
