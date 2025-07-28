
# AI Email Lead Tracker

This script acts as an automated assistant that reads your Outlook inbox, uses AI to identify potential sales opportunities, and organizes them into a structured Excel file on SharePoint.

### What It Does

  * **Finds New Leads**: Automatically processes new emails to find sales opportunities.
  * **Extracts Key Info**: Uses Google's Gemini AI to pull out contact details, company names, and a summary of the lead.
  * **Prevents Duplicates**: Intelligently links follow-up emails to opportunities you're already tracking, even if they're from a different person at the same company.
  * **Keeps You Organized**: Logs everything into two clean, linked tables in an Excel file so you always know what's going on.

-----

## Initial Setup (5-10 Minutes)

First, let's get the bot running on your local machine.

### Step 1: Configure Azure App Registration

You need to give the script permission to access your Microsoft 365 account.

1.  Go to the **Azure Portal** and sign in.
2.  Navigate to **Azure Active Directory** \> **App registrations** \> **+ New registration**.
3.  Give it a name (e.g., `EmailLeadBot`) and click **Register**.
4.  On the app's overview page, copy these two values for later:
      * **Application (client) ID**
      * **Directory (tenant) ID**
5.  Go to **API permissions** \> **+ Add a permission** \> **Microsoft Graph** \> **Delegated permissions**.
6.  Add these three permissions. This allows the script to:
      * `User.Read` (know who you are)
      * `Mail.Read` (read your emails)
      * `Files.ReadWrite.All` (update your Excel file)
7.  Go to **Authentication** \> **+ Add a platform** \> **Mobile and desktop applications**.
8.  Check the box for the `https://login.microsoftonline.com/common/oauth2/nativeclient` option.
9.  At the bottom of the page, find the switch for **"Allow public client flows"**, set it to **Yes**, and click **Save**.

### Step 2: Get Your Google AI Key

1.  Go to **Google AI Studio**.
2.  Click **"Get API key"** and create a new one. Copy the key.

### Step 3: Set Up the Excel File

This is where your leads will be saved.

1.  On your company's SharePoint or OneDrive, create a **new Excel workbook** (e.g., `LeadTracker.xlsx`).
2.  Rename the first sheet to exactly **`OpportunitiesMaster`**.
3.  Create a second sheet named exactly **`InteractionLog`**.
4.  **On the `OpportunitiesMaster` sheet**:
      * Create these headers in the first row: `Opportunity ID`, `Contact Name`, `Contact Company`, `Contact Email`, `Phone`, `Title`, `Status`, `Date Created`, `Conversation ID`, `Summary`.
      * Select the headers, go to the "Home" tab, and click **"Format as Table"**. Name the table `OpportunitiesTable`.
5.  **On the `InteractionLog` sheet**:
      * Create these headers: `Opportunity ID`, `Interaction Date`, `Status`, `Type`, `Sender`, `Summary`, `Action Item`, `Deadline`.
      * Format them as a table named `InteractionsTable`.

### Step 4: Configure the Script

Open the `EmailBot2.py` file and paste the credentials you copied into the configuration section at the top.

```python
# === 2. Configuration ===
CLIENT_ID = "YOUR_AZURE_CLIENT_ID"
TENANT_ID = "YOUR_AZURE_TENANT_ID"
EXCEL_SHARE_LINK = "YOUR_SHAREPOINT_EXCEL_SHARE_LINK" # Get this by clicking Share > Copy Link
GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"
```

You should now be able to run the script successfully on your local machine.

-----

## Full Automation with GitHub Actions

This section explains how to set up the script to run automatically every morning using GitHub Actions.

### Step 5: Prepare Your Script for the Cloud

We need to modify the script to use environment variables for security instead of hardcoding your secret keys.

1.  **Modify Configuration**: Change the configuration block in `EmailBot2.py` **from this:**
    ```python
    CLIENT_ID = "YOUR_AZURE_CLIENT_ID"
    # ...and the rest
    ```
    **To this:**
    ```python
    import os # Make sure this is at the top of your file
    CLIENT_ID = os.getenv("CLIENT_ID")
    TENANT_ID = os.getenv("TENANT_ID")
    EXCEL_SHARE_LINK = os.getenv("EXCEL_SHARE_LINK")
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
    ```
2.  **Update File Paths**: The automation environment requires a cache directory for temporary files. Change the file paths **from this:**
    ```python
    TOKEN_CACHE_FILE = "msal_token_cache.bin"
    # ...and the rest
    ```
    **To this:**
    ```python
    TOKEN_CACHE_FILE = "cache/msal_token_cache.bin"
    TIMESTAMP_FILE = "cache/last_run_timestamp.txt"
    PROCESSED_EMAILS_FILE = "cache/processed_emails.json"
    ```

### Step 6: Set Up Your GitHub Repository

1.  Create a new **private** GitHub repository.
2.  Create the following folder and file structure inside it:
    ```
    your-email-bot/
    ├── .github/workflows/
    │   └── daily-email-check.yml
    ├── src/
    │   └── EmailBot2.py
    └── requirements.txt
    ```
3.  **`requirements.txt`**: Create this file with the required libraries:
    ```txt
    requests
    msal
    html2text
    google-generativeai
    python-dotenv
    ```
4.  **`daily-email-check.yml`**: Create this file inside `.github/workflows/`. It tells GitHub to run your script daily at 9 AM UTC.
    ```yaml
    name: Daily Email Lead Check
    on:
      schedule:
        - cron: '0 9 * * *'
      workflow_dispatch: {}
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
            run: pip install -r requirements.txt
          - name: Restore or Create Cache Dir
            run: mkdir -p cache
          - name: Restore Cached Files
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
            run: python src/EmailBot2.py
    ```

### Step 7: Add Your Keys to GitHub Secrets

This acts as a secure password manager for your script's credentials.

1.  In your GitHub repository, go to **Settings** \> **Secrets and variables** \> **Actions**.
2.  Click **New repository secret** and add the following four secrets:
      * `CLIENT_ID`
      * `TENANT_ID`
      * `EXCEL_SHARE_LINK`
      * `GEMINI_API_KEY`

### Step 8: Critical Azure Update (Do Not Skip)

For automation, the script needs stronger, non-interactive permissions to log in on its own.

1.  Go back to your **Azure Portal** \> **App registrations** \> your app.
2.  Click **API permissions**.
3.  **REMOVE** all three "Delegated" permissions you added in Step 1.
4.  Click **+ Add a permission** \> **Microsoft Graph** \> **Application permissions**.
5.  Add these three **Application** permissions:
      * `User.Read.All`
      * `Mail.Read`
      * `Files.ReadWrite.All`
6.  Click the **"Grant admin consent for..."** button at the top of the permissions list. This is a required step.
7.  Now go to **Certificates & secrets** \> **+ New client secret**.
8.  Give it a description, set the expiration to 24 months, and click **Add**.
9.  **COPY THE SECRET VALUE IMMEDIATELY.** It will disappear after you leave this page.

### Step 9: The One-Time Login

Microsoft security requires one final interactive login to generate an authentication token for the automation to use.

1.  Run your script one last time **on your local machine**.
2.  Complete the device login steps in your browser. This will create the `cache/msal_token_cache.bin` file.
3.  Commit this token file to your **private** GitHub repository:
    ```bash
    git add cache/msal_token_cache.bin
    git commit -m "Add initial authentication token"
    git push
    ```

### You're Done\!

Push all your files to GitHub. You can test the automation manually by going to the "Actions" tab in your repository and clicking "Run workflow." From now on, your bot will run automatically every morning.
