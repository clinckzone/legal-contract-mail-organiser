# Legal Contract Mail Organiser - Setup Instructions

This guide walks you through setting up the automated contract processing system.

## Overview

The system will:
1. Monitor your Gmail inbox every 5 minutes
2. Detect emails with PDF/Word attachments
3. Save attachments to an organized Google Drive folder
4. Log contract details to a Google Sheet for tracking

---

## Prerequisites

- Google Workspace or personal Gmail account
- Access to Google Apps Script (script.google.com)

---

## Setup Steps

### Step 1: Create a New Google Apps Script Project

1. Go to [script.google.com](https://script.google.com)
2. Click **"New Project"**
3. Name the project: `Legal Contract Mail Organiser`

### Step 2: Add the Script Files

Delete the default `Code.gs` content and create the following files:

| File Name | Content |
|-----------|---------|
| `Config.gs` | Copy from `Config.gs` in this repo |
| `Code.gs` | Copy from `Code.gs` in this repo |
| `Setup.gs` | Copy from `Setup.gs` in this repo |

To create a new file in Apps Script:
- Click the **"+"** next to **Files**
- Select **"Script"**
- Name it (without the `.gs` extension)

### Step 3: Update the Manifest

1. In Apps Script, click **Project Settings** (gear icon)
2. Check **"Show 'appsscript.json' manifest file in editor"**
3. Click **Editor** to go back
4. Open `appsscript.json` and replace its contents with the content from `appsscript.json` in this repo

### Step 4: Run Initial Setup

1. In the Apps Script editor, select `Setup.gs` from the file list
2. Select the function `setupDriveFolder` from the dropdown
3. Click **Run**
4. **Grant permissions** when prompted (first time only):
   - Click "Review Permissions"
   - Select your Google account
   - Click "Advanced" → "Go to Legal Contract Mail Organiser"
   - Click "Allow"
5. Check the **Execution Log** (View → Logs) for the Folder ID
6. Copy the Folder ID

### Step 5: Create the Spreadsheet

1. Select the function `setupSpreadsheet` from the dropdown
2. Click **Run**
3. Check the Execution Log for the Spreadsheet ID
4. Copy the Spreadsheet ID

### Step 6: Update Configuration

1. Open `Config.gs`
2. Update these values:
   ```javascript
   SPREADSHEET_ID: 'paste-your-spreadsheet-id-here',
   DRIVE_FOLDER_ID: 'paste-your-folder-id-here',
   TEAM_MEMBERS: [
     'Unassigned',
     'Your Name',
     'Team Member 2',
     // Add your team
   ]
   ```
3. **Save the file** (Ctrl+S / Cmd+S)

### Step 7: Set Up Automatic Trigger

1. Select the function `setupTrigger` from the dropdown
2. Click **Run**
3. The script will now check for new emails every 5 minutes

### Step 8: Test the System

1. Send yourself a test email with a PDF attachment
2. Run `testProcessEmails` manually to process immediately (don't wait 5 min)
3. Check:
   - ✅ Email gets labeled `ContractProcessed`
   - ✅ PDF appears in Google Drive → Legal Contracts folder
   - ✅ New row appears in the Google Sheet

---

## Using the System

### For Legal Counsels

1. **Open the Google Sheet**: Legal Contracts Tracker
2. **View the Dashboard tab** for summary statistics
3. **Go to Contracts tab** to see all contracts
4. **Filter contracts**:
   - Click the filter icon in column headers
   - Or create a Filter View (Data → Create a filter view)
5. **Update a contract**:
   - Change Status dropdown
   - Assign to yourself or team member
   - Add notes in the Notes column
6. **Open the document**: Click the Drive Link

### Filter View Examples

Create these saved filter views for quick access:

| View Name | Filter |
|-----------|--------|
| My Contracts | Assigned To = [Your Name] |
| Needs Assignment | Assigned To = "Unassigned" |
| Under Review | Status = "Under Review" |
| This Week | Received Date = last 7 days |

---

## Sharing with Team

### Share the Google Sheet
1. Open the spreadsheet
2. Click **Share**
3. Add team members with **Editor** access

### Share the Drive Folder
1. Open the Legal Contracts folder in Drive
2. Right-click → **Share**
3. Add team members with **Viewer** or **Editor** access

---

## Troubleshooting

### Emails not being processed
- Check that the trigger is active: Apps Script → Triggers (clock icon)
- Verify the email has a PDF/Word attachment
- Check Execution Log for errors

### Permission errors
- Re-run any setup function and re-authorize if prompted
- Ensure you're using the same Google account for Gmail, Drive, and Sheets

### Duplicate entries
- The system labels processed emails with `ContractProcessed`
- If duplicates appear, check if the label is being applied correctly

---

## Customization

### Change check frequency
Edit `Setup.gs` → `setupTrigger()`:
```javascript
.everyMinutes(10)  // Change from 5 to 10 minutes
```

### Add more file types
Edit `Config.gs`:
```javascript
ALLOWED_EXTENSIONS: ['pdf', 'doc', 'docx', 'xlsx', 'pptx']
```

### Change folder structure
Edit `Code.gs` → `processAttachment()`:
```javascript
// Current: 2026/03-March/
const yearMonth = Utilities.formatDate(receivedDate, Session.getScriptTimeZone(), 'yyyy/MM-MMMM');

// Alternative: Just by year
const yearMonth = Utilities.formatDate(receivedDate, Session.getScriptTimeZone(), 'yyyy');
```

---

## Files Reference

| File | Purpose |
|------|---------|
| `appsscript.json` | Project manifest with permissions |
| `Config.gs` | Configuration values (edit this) |
| `Code.gs` | Main email processing logic |
| `Setup.gs` | One-time setup functions |
