/**
 * Setup functions for Legal Contract Mail Organiser
 * Run these once to initialize the system
 */

/**
 * STEP 1: Run this first to create the Drive folder structure
 * Returns the folder ID to add to CONFIG
 */
function setupDriveFolder() {
  const rootFolder = DriveApp.createFolder('Legal Contracts');

  // Create a README file in the folder
  rootFolder.createFile(
    'README.txt',
    'This folder contains contract attachments automatically extracted from Gmail.\n\n' +
    'Structure:\n' +
    '  /YYYY/MM-Month/\n' +
    '    DATE_SENDER_filename.pdf\n\n' +
    'Do not manually rename or move files as they are linked from the tracking sheet.'
  );

  Logger.log('===========================================');
  Logger.log('Drive folder created successfully!');
  Logger.log('Folder Name: Legal Contracts');
  Logger.log('Folder ID: ' + rootFolder.getId());
  Logger.log('Folder URL: ' + rootFolder.getUrl());
  Logger.log('===========================================');
  Logger.log('ACTION: Copy the Folder ID above and paste it into CONFIG.DRIVE_FOLDER_ID in Config.gs');

  return rootFolder.getId();
}

/**
 * STEP 2: Run this to create the Google Sheet with proper structure
 * Returns the spreadsheet ID to add to CONFIG
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.create('Legal Contracts Tracker');
  const sheet = ss.getActiveSheet();
  sheet.setName('Contracts');

  // Set up headers
  const headers = [
    'Contract ID',
    'Subject',
    'Sender Name',
    'Sender Email',
    'Received Date',
    'Attachment Name',
    'Drive Link',
    'Status',
    'Assigned To',
    'Notes',
    'Last Updated'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');

  // Set column widths
  sheet.setColumnWidth(1, 120);  // Contract ID
  sheet.setColumnWidth(2, 250);  // Subject
  sheet.setColumnWidth(3, 150);  // Sender Name
  sheet.setColumnWidth(4, 200);  // Sender Email
  sheet.setColumnWidth(5, 130);  // Received Date
  sheet.setColumnWidth(6, 200);  // Attachment Name
  sheet.setColumnWidth(7, 100);  // Drive Link
  sheet.setColumnWidth(8, 120);  // Status
  sheet.setColumnWidth(9, 120);  // Assigned To
  sheet.setColumnWidth(10, 250); // Notes
  sheet.setColumnWidth(11, 130); // Last Updated

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set up data validation for Status column (column 8)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.STATUSES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('H2:H1000').setDataValidation(statusRule);

  // Set up data validation for Assigned To column (column 9)
  const assigneeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.TEAM_MEMBERS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('I2:I1000').setDataValidation(assigneeRule);

  // Add conditional formatting for status
  addStatusConditionalFormatting(sheet);

  // Create helper sheet for team members (for easy updates)
  const teamSheet = ss.insertSheet('Team');
  teamSheet.getRange(1, 1).setValue('Team Members');
  teamSheet.getRange(1, 1).setFontWeight('bold');
  CONFIG.TEAM_MEMBERS.forEach((member, index) => {
    teamSheet.getRange(index + 2, 1).setValue(member);
  });
  teamSheet.setColumnWidth(1, 200);

  // Create Dashboard sheet with summary
  createDashboardSheet(ss);

  Logger.log('===========================================');
  Logger.log('Spreadsheet created successfully!');
  Logger.log('Spreadsheet Name: Legal Contracts Tracker');
  Logger.log('Spreadsheet ID: ' + ss.getId());
  Logger.log('Spreadsheet URL: ' + ss.getUrl());
  Logger.log('===========================================');
  Logger.log('ACTION: Copy the Spreadsheet ID above and paste it into CONFIG.SPREADSHEET_ID in Config.gs');

  return ss.getId();
}

/**
 * Adds conditional formatting to color-code status values
 */
function addStatusConditionalFormatting(sheet) {
  const statusColumn = sheet.getRange('H2:H1000');

  const rules = [
    { status: 'Received', color: '#e8f5e9' },      // Light green
    { status: 'Under Review', color: '#fff3e0' },  // Light orange
    { status: 'Approved', color: '#e3f2fd' },      // Light blue
    { status: 'Signed', color: '#f3e5f5' },        // Light purple
    { status: 'Archived', color: '#f5f5f5' }       // Light gray
  ];

  const conditionalRules = rules.map(rule => {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(rule.status)
      .setBackground(rule.color)
      .setRanges([statusColumn])
      .build();
  });

  sheet.setConditionalFormatRules(conditionalRules);
}

/**
 * Creates a dashboard sheet with summary statistics
 */
function createDashboardSheet(ss) {
  const dashboard = ss.insertSheet('Dashboard');

  // Move dashboard to first position
  ss.setActiveSheet(dashboard);
  ss.moveActiveSheet(1);

  // Title
  dashboard.getRange('A1').setValue('Legal Contracts Dashboard');
  dashboard.getRange('A1').setFontSize(18).setFontWeight('bold');

  // Summary section
  dashboard.getRange('A3').setValue('Summary Statistics');
  dashboard.getRange('A3').setFontSize(14).setFontWeight('bold');

  // Stats
  const stats = [
    ['Total Contracts', '=COUNTA(Contracts!A:A)-1'],
    ['Received', '=COUNTIF(Contracts!H:H,"Received")'],
    ['Under Review', '=COUNTIF(Contracts!H:H,"Under Review")'],
    ['Approved', '=COUNTIF(Contracts!H:H,"Approved")'],
    ['Signed', '=COUNTIF(Contracts!H:H,"Signed")'],
    ['Archived', '=COUNTIF(Contracts!H:H,"Archived")'],
    ['', ''],
    ['Unassigned', '=COUNTIF(Contracts!I:I,"Unassigned")']
  ];

  dashboard.getRange(4, 1, stats.length, 2).setValues(stats);
  dashboard.getRange('A4:A11').setFontWeight('bold');
  dashboard.getRange('B4:B11').setHorizontalAlignment('center');

  // Recent contracts section
  dashboard.getRange('A13').setValue('Recent Contracts (Last 10)');
  dashboard.getRange('A13').setFontSize(14).setFontWeight('bold');

  dashboard.getRange('A14').setValue('See "Contracts" sheet for full list and filtering');
  dashboard.getRange('A14').setFontStyle('italic');

  // Formatting
  dashboard.setColumnWidth(1, 200);
  dashboard.setColumnWidth(2, 100);
}

/**
 * STEP 3: Run this to set up the time-driven trigger
 */
function setupTrigger() {
  // Remove any existing triggers for this function
  const existingTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of existingTriggers) {
    if (trigger.getHandlerFunction() === 'processIncomingEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Create new trigger to run every 5 minutes
  ScriptApp.newTrigger('processIncomingEmails')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('===========================================');
  Logger.log('Trigger set up successfully!');
  Logger.log('The script will now run every 5 minutes to check for new emails.');
  Logger.log('===========================================');
}

/**
 * Utility: Remove all triggers (for cleanup)
 */
function removeAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
  Logger.log('All triggers removed.');
}

/**
 * MASTER SETUP: Run all setup steps in order
 * Only run this if you want to do everything at once
 */
function runFullSetup() {
  Logger.log('Starting full setup...\n');

  Logger.log('STEP 1: Creating Drive folder...');
  const folderId = setupDriveFolder();

  Logger.log('\nSTEP 2: Creating Spreadsheet...');
  const sheetId = setupSpreadsheet();

  Logger.log('\n===========================================');
  Logger.log('IMPORTANT: Before running Step 3, you must:');
  Logger.log('1. Update CONFIG.DRIVE_FOLDER_ID with: ' + folderId);
  Logger.log('2. Update CONFIG.SPREADSHEET_ID with: ' + sheetId);
  Logger.log('3. Update CONFIG.TEAM_MEMBERS with your team names');
  Logger.log('4. Save the Config.gs file');
  Logger.log('5. Then run setupTrigger() separately');
  Logger.log('===========================================');
}
