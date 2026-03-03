# Contract Organiser - Gmail Add-on

A Google Workspace Add-on that automatically organizes contract attachments (PDF/Word) from Gmail into Google Drive and tracks them in a spreadsheet.

## Features

- **Automatic Detection**: Monitors Gmail for emails with PDF/Word attachments
- **Drive Organization**: Saves attachments to organized folder structure (Year/Month)
- **Spreadsheet Tracking**: Logs contracts with status tracking
- **Team Collaboration**: Share folder and sheet with team members
- **Simple Setup**: No coding required - just paste shared links

## How It Works

```
Email with PDF arrives → Add-on detects it → Saves to Drive → Logs to Sheet
                                    ↓
                         Legal counsel reviews in Sheet
                                    ↓
                    Updates status: Received → Under Review → Signed → Archived
```

## Installation

### Option 1: Deploy for Your Organization (Recommended)

1. Go to [script.google.com](https://script.google.com)
2. Create a new project: **File → New Project**
3. Name it: `Contract Organiser`
4. Delete the default code and create these files:

| File | Content |
|------|---------|
| `appsscript.json` | Copy from this repo |
| `Code.gs` | Copy from this repo |
| `Cards.gs` | Copy from this repo |
| `Actions.gs` | Copy from this repo |

5. To see `appsscript.json`:
   - Click **Project Settings** (gear icon)
   - Check **"Show 'appsscript.json' manifest file in editor"**

6. Deploy as Add-on:
   - Click **Deploy → Test deployments**
   - Select **Gmail** as the application
   - Click **Install**

### Option 2: Publish to Workspace Marketplace

For organization-wide deployment, follow [Google's publishing guide](https://developers.google.com/workspace/marketplace/how-to-publish).

---

## User Guide

### First User (Creates Shared Resources)

1. **Open Gmail** and find the Contract Organiser in the right sidebar
2. **Create New Folder**: Click "Create New Folder" button
3. **Share the folder**: Click "Open & Share" and share with your team
4. **Create New Spreadsheet**: Click "Create New Spreadsheet" button
5. **Share the spreadsheet**: Click "Open & Share" and share with your team
6. **Enter your name** and click **"Save & Start Monitoring"**
7. **Send the shared links** to your team members

### Other Team Members (Join Existing Setup)

1. **Open Gmail** and find the Contract Organiser in the right sidebar
2. **Paste the folder link** that was shared with you
3. **Paste the spreadsheet link** that was shared with you
4. **Enter your name** and click **"Save & Start Monitoring"**

---

## Dashboard Features

Once configured, the add-on shows:

- **Status**: Whether monitoring is active
- **Statistics**: Total contracts, by status, unassigned count
- **Quick Links**: Open Sheet or Drive folder
- **Actions**:
  - Pause/Start Monitoring
  - Process Emails Now (manual trigger)
  - Reset Setup

---

## Spreadsheet Structure

### Contracts Sheet

| Column | Description |
|--------|-------------|
| Contract ID | Auto-generated (CTR-2026-0001) |
| Subject | Email subject |
| Sender Name | Sender's name |
| Sender Email | Sender's email |
| Received Date | When email arrived |
| Attachment Name | Original filename |
| Drive Link | Link to file in Drive |
| Status | Dropdown: Received, Under Review, Approved, Signed, Archived |
| Assigned To | Dropdown: Team members |
| Notes | Free text |
| Last Updated | Timestamp |

### Dashboard Sheet

Shows summary statistics:
- Total contracts
- Count by status
- Unassigned contracts

---

## Drive Folder Structure

```
📁 Legal Contracts/
├── README.txt
├── 📁 2026/
│   ├── 📁 01-January/
│   │   └── 2026-01-15_vendor_corp_com_Contract.pdf
│   ├── 📁 02-February/
│   └── 📁 03-March/
```

---

## Workflow for Legal Counsels

1. **Check the spreadsheet** for new contracts (Status: Received)
2. **Assign to yourself** using the "Assigned To" dropdown
3. **Click Drive Link** to review the document
4. **Update Status** as you work:
   - `Under Review` - Currently reviewing
   - `Approved` - Approved by legal
   - `Signed` - Fully executed
   - `Archived` - Completed and filed
5. **Add Notes** for any comments

### Using Filters

Create Filter Views in Google Sheets:

| View Name | Filter |
|-----------|--------|
| My Contracts | Assigned To = [Your Name] |
| Needs Assignment | Assigned To = "Unassigned" |
| Under Review | Status = "Under Review" |
| This Week | Received Date = last 7 days |

---

## Troubleshooting

### Add-on not appearing in Gmail
- Make sure you've deployed as a test deployment
- Refresh Gmail
- Check that Gmail is selected in deployment settings

### Emails not being processed
- Check the Dashboard shows "Monitoring Active"
- Verify the email has a .pdf, .doc, or .docx attachment
- Only emails received AFTER setup are processed

### "Cannot access folder/spreadsheet" error
- Ask the owner to share the resource with you
- Make sure you have at least Editor access

### Duplicate entries
- The add-on labels processed emails with "ContractProcessed"
- Check if this label exists in your Gmail

---

## Files Reference

| File | Purpose |
|------|---------|
| `appsscript.json` | Add-on manifest and permissions |
| `Code.gs` | Email processing logic |
| `Cards.gs` | UI card builders |
| `Actions.gs` | Button click handlers |

---

## Privacy & Permissions

The add-on requires these permissions:
- **Gmail**: Read emails, apply labels
- **Drive**: Create folders and files
- **Sheets**: Create and edit spreadsheets

All data stays within your Google Workspace. No external servers are used.
