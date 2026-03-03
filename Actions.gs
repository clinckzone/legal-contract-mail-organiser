/**
 * Action handlers for the Contract Organiser Add-on
 */

/**
 * Creates a new Drive folder for contracts
 */
function onCreateFolder(e) {
  try {
    const folderName = 'Legal Contracts';
    const folder = DriveApp.createFolder(folderName);

    // Create a README in the folder
    folder.createFile(
      'README.txt',
      'This folder contains contract attachments automatically extracted from Gmail.\n\n' +
      'Structure:\n' +
      '  /YYYY/MM-Month/\n' +
      '    DATE_SENDER_filename.pdf\n\n' +
      'Do not manually rename or move files as they are linked from the tracking sheet.'
    );

    const folderId = folder.getId();
    const folderUrl = folder.getUrl();

    // Temporarily store the folder ID
    const userProps = PropertiesService.getUserProperties();
    userProps.setProperty('TEMP_FOLDER_ID', folderId);

    return buildResourceCreatedCard('Folder', folderName, folderUrl, folderId);

  } catch (error) {
    return buildErrorCard('Error Creating Folder', error.message);
  }
}

/**
 * Creates a new spreadsheet for tracking contracts
 */
function onCreateSheet(e) {
  try {
    const ss = SpreadsheetApp.create('Contract Tracker');
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

    // Set up data validation for Status column
    const statuses = ['Received', 'Under Review', 'Approved', 'Signed', 'Archived'];
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(statuses, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange('H2:H1000').setDataValidation(statusRule);

    // Add conditional formatting for status
    addStatusConditionalFormatting(sheet);

    // Create Dashboard sheet
    createDashboardSheet(ss);

    const sheetId = ss.getId();
    const sheetUrl = ss.getUrl();

    // Temporarily store the sheet ID
    const userProps = PropertiesService.getUserProperties();
    userProps.setProperty('TEMP_SHEET_ID', sheetId);

    return buildResourceCreatedCard('Spreadsheet', 'Contract Tracker', sheetUrl, sheetId);

  } catch (error) {
    return buildErrorCard('Error Creating Spreadsheet', error.message);
  }
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
  dashboard.getRange('A1').setValue('Contract Tracker Dashboard');
  dashboard.getRange('A1').setFontSize(18).setFontWeight('bold');

  // Summary section
  dashboard.getRange('A3').setValue('Summary Statistics');
  dashboard.getRange('A3').setFontSize(14).setFontWeight('bold');

  // Stats formulas
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

  // Formatting
  dashboard.setColumnWidth(1, 200);
  dashboard.setColumnWidth(2, 100);
}

/**
 * Continues setup after resource creation
 */
function onContinueSetup(e) {
  const params = e.commonEventObject.parameters;
  const resourceType = params.resourceType;
  const resourceId = params.resourceId;

  const userProps = PropertiesService.getUserProperties();

  // Store the resource ID permanently if user confirmed
  if (resourceType === 'folder') {
    userProps.setProperty('FOLDER_ID', resourceId);
    userProps.deleteProperty('TEMP_FOLDER_ID');
  } else if (resourceType === 'spreadsheet') {
    userProps.setProperty('SHEET_ID', resourceId);
    userProps.deleteProperty('TEMP_SHEET_ID');
  }

  // Return to setup card
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().popToRoot().updateCard(buildSetupCard()))
    .build();
}

/**
 * Saves the setup configuration and starts monitoring
 */
function onSaveSetup(e) {
  const formInputs = e.commonEventObject.formInputs;

  // Get folder link/ID
  let folderId = PropertiesService.getUserProperties().getProperty('FOLDER_ID');
  if (!folderId && formInputs.folderLink) {
    const folderLink = formInputs.folderLink.stringInputs.value[0];
    folderId = extractIdFromUrl(folderLink, 'folder');
  }

  // Get sheet link/ID
  let sheetId = PropertiesService.getUserProperties().getProperty('SHEET_ID');
  if (!sheetId && formInputs.sheetLink) {
    const sheetLink = formInputs.sheetLink.stringInputs.value[0];
    sheetId = extractIdFromUrl(sheetLink, 'spreadsheet');
  }

  // Get user name
  let userName = '';
  if (formInputs.userName) {
    userName = formInputs.userName.stringInputs.value[0];
  }

  // Validate inputs
  if (!folderId) {
    return buildErrorCard('Missing Folder', 'Please provide a Google Drive folder link or create a new one.');
  }

  if (!sheetId) {
    return buildErrorCard('Missing Spreadsheet', 'Please provide a Google Sheets link or create a new one.');
  }

  if (!userName) {
    return buildErrorCard('Missing Name', 'Please enter your name.');
  }

  // Validate access to resources
  try {
    DriveApp.getFolderById(folderId);
  } catch (error) {
    return buildErrorCard('Cannot Access Folder', 'You don\'t have access to this folder. Ask the owner to share it with you.');
  }

  try {
    SpreadsheetApp.openById(sheetId);
  } catch (error) {
    return buildErrorCard('Cannot Access Spreadsheet', 'You don\'t have access to this spreadsheet. Ask the owner to share it with you.');
  }

  // Save configuration
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty('FOLDER_ID', folderId);
  userProps.setProperty('SHEET_ID', sheetId);
  userProps.setProperty('USER_NAME', userName);

  // Add user to the Team list in the spreadsheet
  addUserToTeamList(sheetId, userName);

  // Set up trigger
  setupUserTrigger();

  return buildSuccessCard(
    'Setup Complete!',
    'Your emails with PDF/Word attachments will now be automatically processed every 5 minutes.\n\n' +
    'Contracts will be saved to the shared Drive folder and logged in the spreadsheet.'
  );
}

/**
 * Extracts Google Drive/Sheets ID from a URL
 */
function extractIdFromUrl(url, type) {
  if (!url) return null;

  url = url.trim();

  // Direct ID (not a URL)
  if (!url.includes('/') && !url.includes('.')) {
    return url;
  }

  if (type === 'folder') {
    // Pattern: /folders/FOLDER_ID or /folders/FOLDER_ID?...
    const folderMatch = url.match(/\/folders\/([a-zA-Z0-9_-]+)/);
    if (folderMatch) return folderMatch[1];
  }

  if (type === 'spreadsheet') {
    // Pattern: /spreadsheets/d/SHEET_ID/... or /spreadsheets/d/SHEET_ID
    const sheetMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (sheetMatch) return sheetMatch[1];
  }

  return null;
}

/**
 * Adds the user to the Team list in the spreadsheet
 */
function addUserToTeamList(sheetId, userName) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    let teamSheet = ss.getSheetByName('Team');

    if (!teamSheet) {
      teamSheet = ss.insertSheet('Team');
      teamSheet.getRange(1, 1).setValue('Team Members');
      teamSheet.getRange(1, 1).setFontWeight('bold');
      teamSheet.getRange(2, 1).setValue('Unassigned');
    }

    // Check if user already exists
    const data = teamSheet.getDataRange().getValues();
    const exists = data.some(row => row[0] === userName);

    if (!exists) {
      const lastRow = teamSheet.getLastRow();
      teamSheet.getRange(lastRow + 1, 1).setValue(userName);

      // Update the Assigned To dropdown in Contracts sheet
      updateAssigneeDropdown(ss);
    }
  } catch (error) {
    Logger.log('Error adding user to team: ' + error.message);
  }
}

/**
 * Updates the Assigned To dropdown with current team members
 */
function updateAssigneeDropdown(ss) {
  try {
    const teamSheet = ss.getSheetByName('Team');
    const contractsSheet = ss.getSheetByName('Contracts');

    if (!teamSheet || !contractsSheet) return;

    const teamData = teamSheet.getRange(2, 1, teamSheet.getLastRow() - 1, 1).getValues();
    const teamMembers = teamData.map(row => row[0]).filter(name => name);

    if (teamMembers.length > 0) {
      const assigneeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(teamMembers, true)
        .setAllowInvalid(false)
        .build();
      contractsSheet.getRange('I2:I1000').setDataValidation(assigneeRule);
    }
  } catch (error) {
    Logger.log('Error updating assignee dropdown: ' + error.message);
  }
}

/**
 * Sets up the time-driven trigger for the current user
 */
function setupUserTrigger() {
  // Remove any existing triggers for this user
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processIncomingEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Set trigger start date
  const userProps = PropertiesService.getUserProperties();
  const now = new Date();
  userProps.setProperty('TRIGGER_START_DATE', now.toISOString());

  // Create new trigger
  ScriptApp.newTrigger('processIncomingEmails')
    .timeBased()
    .everyMinutes(5)
    .create();
}

/**
 * Checks if the monitoring trigger exists
 */
function checkTriggerExists() {
  const triggers = ScriptApp.getProjectTriggers();
  return triggers.some(trigger => trigger.getHandlerFunction() === 'processIncomingEmails');
}

/**
 * Pauses email monitoring
 */
function onPauseMonitoring(e) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processIncomingEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  return buildSuccessCard(
    'Monitoring Paused',
    'Email monitoring has been paused. Your existing contracts are still in the spreadsheet.\n\nClick "Start Monitoring" to resume.'
  );
}

/**
 * Starts/resumes email monitoring
 */
function onStartMonitoring(e) {
  setupUserTrigger();

  return buildSuccessCard(
    'Monitoring Started',
    'Email monitoring is now active. New emails with PDF/Word attachments will be processed every 5 minutes.'
  );
}

/**
 * Processes emails immediately (manual trigger)
 */
function onProcessNow(e) {
  try {
    processIncomingEmails();
    return buildSuccessCard(
      'Processing Complete',
      'Emails have been processed. Check your spreadsheet for any new contracts.'
    );
  } catch (error) {
    return buildErrorCard('Processing Error', error.message);
  }
}

/**
 * Resets the user's setup
 */
function onResetSetup(e) {
  // Remove triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processIncomingEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Clear user properties
  const userProps = PropertiesService.getUserProperties();
  userProps.deleteAllProperties();

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().popToRoot().updateCard(buildSetupCard()))
    .build();
}

/**
 * Gets contract statistics from the spreadsheet
 */
function getContractStats(sheetId) {
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName('Contracts');

  if (!sheet) {
    return { total: 0, received: 0, underReview: 0, unassigned: 0 };
  }

  const data = sheet.getDataRange().getValues();

  let total = Math.max(0, data.length - 1); // Exclude header
  let received = 0;
  let underReview = 0;
  let unassigned = 0;

  for (let i = 1; i < data.length; i++) {
    const status = data[i][7];  // Column H (0-indexed: 7)
    const assignee = data[i][8]; // Column I (0-indexed: 8)

    if (status === 'Received') received++;
    if (status === 'Under Review') underReview++;
    if (!assignee || assignee === 'Unassigned') unassigned++;
  }

  return { total, received, underReview, unassigned };
}
