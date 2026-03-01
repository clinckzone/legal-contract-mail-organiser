/**
 * Configuration for Legal Contract Mail Organiser
 * Update these values after running the setup
 */

const CONFIG = {
  // Google Sheet ID (from the URL: docs.google.com/spreadsheets/d/SHEET_ID/edit)
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',

  // Sheet name where contracts are logged
  SHEET_NAME: 'Contracts',

  // Google Drive folder ID for storing attachments
  // (from the URL: drive.google.com/drive/folders/FOLDER_ID)
  DRIVE_FOLDER_ID: 'YOUR_DRIVE_FOLDER_ID_HERE',

  // Gmail label to mark processed emails
  PROCESSED_LABEL: 'ContractProcessed',

  // File extensions to capture
  ALLOWED_EXTENSIONS: ['pdf', 'doc', 'docx'],

  // Team members for assignment dropdown (update with your team)
  TEAM_MEMBERS: [
    'Unassigned',
    'John Doe',
    'Jane Smith',
    'Alex Johnson'
  ],

  // Contract statuses
  STATUSES: [
    'Received',
    'Under Review',
    'Approved',
    'Signed',
    'Archived'
  ]
};
