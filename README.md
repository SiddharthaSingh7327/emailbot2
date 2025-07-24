# AI Email Lead Tracker

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

1.  Go to the **[Azure Portal](https://portal.azure.com)** and sign in.
2.  Navigate to **Azure Active Directory** > **App registrations** > **+ New registration**.
3.  Give it a name (e.g., `EmailBot`) and click **Register**.
4.  On the app's overview page, copy the **Application (client) ID** and the **Directory (tenant) ID**.
5.  Go to **API permissions** > **+ Add a permission** > **Microsoft Graph** > **Delegated permissions**.
6.  Add these three permissions: **`User.Read`**, **`Mail.Read`**, and **`Files.ReadWrite.All`**.
7.  Go to **Authentication** > **+ Add a platform** > **Mobile and desktop applications**.
8.  Check the box for the `https://login.microsoftonline.com/common/oauth2/nativeclient` redirect URI.
9.  At the bottom of the page, set **"Allow public client flows"** to **Yes** and click **Save**.

Step 2: Get Your Google Gemini API Key

1.  Go to **[Google AI Studio](https://aistudio.google.com/).
2.  Click **"Get API key"** and copy your key.

Step 3: Prepare the Excel File

1.  On your SharePoint or OneDrive, create a **new Excel workbook** (e.g., `LeadTracker.xlsx`).
2.  Rename the first sheet to exactly `OpportunitiesMaster`.
3.  Create a second sheet named exactly `InteractionLog`.
4.  **On the `OpportunitiesMaster` sheet**, create these headers in the first row and format them as a table named **`OpportunitiesTable`**:
    * `Opportunity ID`, `Contact Name`, `Contact Company`, `Contact Email`, `Phone`, `Title`, `Status`, `Date Created`, `Conversation ID`, `Summary`
5.  **On the `InteractionLog` sheet**, create these headers and format them as a table named **`InteractionsTable`**:
    * `Opportunity ID`, `Interaction Date`, `Status`, `Type`, `Sender`, `Summary`, `Action Item`, `Deadline`

Step 4: Configure the Script

Open the `EmailBot2.py` file and paste your credentials directly into the configuration section at the top of the script.

```python
# === 2. Configuration ===
# --- PASTE YOUR CREDENTIALS HERE ---
CLIENT_ID = "YOUR_AZURE_CLIENT_ID"
TENANT_ID = "YOUR_AZURE_TENANT_ID"
EXCEL_SHARE_LINK = "YOUR_SHAREPOINT_EXCEL_SHARE_LINK"
GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"
