### AI Email Lead Tracker
This script automatically reads new emails from an Outlook inbox, uses Google's AI to identify and analyze potential sales opportunities, and logs them into a structured Excel file on SharePoint.

Features
* **Smart Email Fetching**: Only processes emails that have arrived since the last time it ran.
* **AI-Powered Analysis**: Uses Google Gemini to find sales leads and extract key details like contact info, company, and a summary.
* **Intelligent De-duplication**: Automatically links follow-up emails to existing opportunities, even if they are in a brand new conversation thread.
* **Structured Logging**: Organizes data into two separate, linked tables in Excel: a master list of opportunities and a detailed log of all interactions.
* **Secure Authentication**: Connects to your Microsoft 365 account securely.

Setup Guide
Prerequisites:-
* Python 3.8+
* A Microsoft 365 work or school account
* A Google AI account to get an API key

Step 1: Azure App Registration
You need to give the script permission to read your emails.
1. Go to the **Azure Portal** and sign in.
2. Navigate to **Azure Active Directory** > **App registrations** > **+ New registration**.
3. Give it a name (e.g., `EmailBot`) and click **Register**.
4. On the app's overview page, copy the **Application (client) ID** and the **Directory (tenant) ID**.
5. Go to **API permissions** > **+ Add a permission** > **Microsoft Graph** > **Delegated permissions**.
6. Add these three permissions: `User.Read`, `Mail.Read`, and `Files.ReadWrite.All`.
7. Go to **Authentication** > **+ Add a platform** > **Mobile and desktop applications**.
8. Check the box for the `https://login.microsoftonline.com/common/oauth2/nativeclient` redirect URI.
9. At the bottom of the page, set **"Allow public client flows"** to **Yes** and click **Save**.

Step 2: Get Your Google Gemini API Key
1. Go to **Google AI Studio.
2. Click **"Get API key"** and copy your key.

Step 3: Prepare the Excel File
1. On your SharePoint or OneDrive, create a **new Excel workbook** (e.g., `LeadTracker.xlsx`).
2. Rename the first sheet to exactly `OpportunitiesMaster`.
3. Create a second sheet named exactly `InteractionLog`.
4. **On the **`OpportunitiesMaster` sheet, create these headers in the first row and format them as a table named `OpportunitiesTable`:
   * `Opportunity ID`, `Contact Name`, `Contact Company`, `Contact Email`, `Phone`, `Title`, `Status`, `Date Created`, `Conversation ID`, `Summary`
5. **On the **`InteractionLog` sheet, create these headers and format them as a table named `InteractionsTable`:
   * `Opportunity ID`, `Interaction Date`, `Status`, `Type`, `Sender`, `Summary`, `Action Item`, `Deadline`

Step 4: Configure the Script
Open the `EmailBot2.py` file and paste your credentials directly into the configuration section at the top of the script.

```
# === 2. Configuration ===
# --- PASTE YOUR CREDENTIALS HERE ---
CLIENT_ID = "YOUR_AZURE_CLIENT_ID"
TENANT_ID = "YOUR_AZURE_TENANT_ID"
EXCEL_SHARE_LINK = "YOUR_SHAREPOINT_EXCEL_SHARE_LINK"
GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"

```

---

## 🤖 AUTOMATION SETUP: Run Every Day at 9 AM Automatically

Want your email lead tracker to run automatically every morning without your laptop? Here's how to set it up using GitHub Actions (completely free!):

### Step 5: Prepare for Automation

#### 5.1 Modify Your Script Configuration
**CHANGE THIS** in your `EmailBot2.py` file:

```python
# === 2. Configuration ===
# --- OLD WAY (Manual) ---
CLIENT_ID = "YOUR_AZURE_CLIENT_ID"
TENANT_ID = "YOUR_AZURE_TENANT_ID"
EXCEL_SHARE_LINK = "YOUR_SHAREPOINT_EXCEL_SHARE_LINK"
GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"
```

**TO THIS** (Environment Variables):

```python
# === 2. Configuration ===
# --- NEW WAY (Automation-Ready) ---
CLIENT_ID = os.getenv("CLIENT_ID", "YOUR_AZURE_CLIENT_ID")  # Fallback to manual if needed
TENANT_ID = os.getenv("TENANT_ID", "YOUR_AZURE_TENANT_ID")
EXCEL_SHARE_LINK = os.getenv("EXCEL_SHARE_LINK", "YOUR_SHAREPOINT_EXCEL_SHARE_LINK")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY")
```

#### 5.2 Add Cloud Environment Setup
**ADD THIS** to the top of your script (after the imports):

```python
# === AUTOMATION SETUP ===
def setup_cloud_environment():
    """Setup for GitHub Actions environment"""
    os.makedirs('cache', exist_ok=True)
    os.makedirs('logs', exist_ok=True)
    
    # Enhanced logging for automation
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler('logs/email_processor.log')
        ]
    )

def validate_environment():
    """Ensure all required environment variables are set"""
    required_vars = ['CLIENT_ID', 'TENANT_ID', 'EXCEL_SHARE_LINK', 'GEMINI_API_KEY']
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        logging.error(f"❌ Missing required environment variables: {missing_vars}")
        raise ValueError(f"Missing environment variables: {missing_vars}")
    
    logging.info("✅ All required environment variables are set")
```

#### 5.3 Update File Paths for Cloud
**CHANGE THESE LINES**:

```python
# --- OLD WAY ---
TOKEN_CACHE_FILE = "msal_token_cache.bin"
TIMESTAMP_FILE = "last_run_timestamp.txt"
PROCESSED_EMAILS_FILE = "processed_emails.json"
```

**TO THIS**:

```python
# --- NEW WAY (Cloud-friendly paths) ---
TOKEN_CACHE_FILE = "cache/msal_token_cache.bin"
TIMESTAMP_FILE = "cache/last_run_timestamp.txt"
PROCESSED_EMAILS_FILE = "cache/processed_emails.json"
```

#### 5.4 Update Your Main Function
**ADD THESE LINES** at the beginning of your `main()` function:

```python
def main():
    """Main execution function with automation support"""
    # Setup for automation (only runs in cloud, harmless locally)
    if os.getenv('GITHUB_ACTIONS'):
        setup_cloud_environment()
        validate_environment()
        logging.info("🤖 Running in automation mode")
    
    current_run_timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    
    # ... rest of your existing main() function stays the same
```

### Step 6: GitHub Repository Setup

#### 6.1 Create Repository Structure
Create a new **private** GitHub repository with this structure:

```
your-email-bot/
├── .github/
│   └── workflows/
│       └── daily-email-check.yml
├── src/
│   └── EmailBot2.py
├── requirements.txt
├── README.md
└── .gitignore
```

#### 6.2 Create `requirements.txt`
Create this file with your dependencies:

```txt
requests==2.31.0
msal==1.24.1
html2text==2020.1.16
google-generativeai==0.3.2
python-dotenv==1.0.0
```

#### 6.3 Create `.gitignore`
```txt
# Cache and logs
cache/
logs/
*.log

# Python
__pycache__/
*.pyc
*.pyo

# Environment files (if you use them locally)
.env
```

#### 6.4 Create GitHub Workflow
Create `.github/workflows/daily-email-check.yml`:

```yaml
name: Daily Email Lead Check

on:
  schedule:
    # Runs every day at 9:00 AM UTC (adjust for your timezone)
    - cron: '0 9 * * *'
  
  # Allow manual testing
  workflow_dispatch:

jobs:
  check-emails:
    runs-on: ubuntu-latest
    
    steps:
    - name: Checkout code
      uses: actions/checkout@v4
    
    - name: Set up Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.11'
    
    - name: Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install -r requirements.txt
    
    - name: Create cache directories
      run: |
        mkdir -p cache logs
    
    - name: Restore cache
      uses: actions/cache@v3
      with:
        path: cache/
        key: email-bot-cache-${{ github.run_id }}
        restore-keys: |
          email-bot-cache-
    
    - name: Run Email Bot
      env:
        CLIENT_ID: ${{ secrets.CLIENT_ID }}
        TENANT_ID: ${{ secrets.TENANT_ID }}
        EXCEL_SHARE_LINK: ${{ secrets.EXCEL_SHARE_LINK }}
        GEMINI_API_KEY: ${{ secrets.GEMINI_API_KEY }}
      run: |
        python src/EmailBot2.py
    
    - name: Upload logs
      if: always()
      uses: actions/upload-artifact@v3
      with:
        name: email-bot-logs-${{ github.run_number }}
        path: logs/
        retention-days: 30
```

### Step 7: Configure GitHub Secrets

1. Go to your repository on GitHub
2. Click **Settings** → **Secrets and variables** → **Actions**
3. Add these **Repository Secrets**:
   - `CLIENT_ID` = Your Azure Client ID
   - `TENANT_ID` = Your Azure Tenant ID
   - `EXCEL_SHARE_LINK` = Your SharePoint Excel link
   - `GEMINI_API_KEY` = Your Google Gemini API key


### CRITICAL STEP ###
Update Azure App Registration for Automation
IMPORTANT: You MUST change your Azure app settings for automation to work!
8.1 Update API Permissions (Required!)

Go back to your Azure Portal → App registrations → Your app
Click API permissions
REMOVE all existing Delegated permissions
Click + Add a permission → Microsoft Graph → Application permissions
Add these Application permissions (NOT Delegated):

User.Read.All
Mail.Read (under Application permissions)
Files.ReadWrite.All


Click "Grant admin consent" - This is CRUCIAL!
Wait for the green checkmarks to appear

8.2 Create Client Secret (Required!)

In your Azure app, go to Certificates & secrets
Click + New client secret
Add a description (e.g., "GitHub Automation")
Set expiration (recommend 24 months)
Click Add and COPY THE SECRET VALUE IMMEDIATELY (you can't see it again!)

### Step 8: Handle Authentication (Important!)

Since Microsoft requires interactive login, you need to establish authentication once:

1. **Run your script locally first** (with the new environment variable setup)
2. Complete the device flow authentication
3. **Commit the generated token cache**:
   ```bash
   git add cache/msal_token_cache.bin
   git commit -m "Add authentication token cache"
   git push
   ```

**Security Note**: Only do this with a **private repository** since the token cache contains authentication data.

### Step 9: Test and Deploy

1. **Push all files** to your GitHub repository
2. **Test manually**: Go to Actions → "Daily Email Lead Check" → "Run workflow"
3. **Check logs** in the Actions tab to ensure everything works
4. **Monitor daily runs** - it will automatically run every morning at 9 AM!

### Step 10: Customize Schedule (Optional)

To change the time, edit the cron schedule in your workflow file:

- **9 AM EST**: `cron: '0 14 * * *'` (2 PM UTC)
- **9 AM PST**: `cron: '0 17 * * *'` (5 PM UTC)
- **9 AM IST**: `cron: '30 3 * * *'` (3:30 AM UTC)

### 🎉 That's It!

Your email lead tracker will now run automatically every morning at 9 AM, processing new emails and updating your Excel file - no laptop required!

**Free Usage Limits:**
- GitHub Actions: 2,000 minutes/month (private repos)
- Your script runs in ~3-5 minutes, so ~400-600 runs per month
- Completely covers daily automation needs!
