/**
 * Legal Contract Mail Organiser
 * Main automation script for processing incoming emails with contract attachments
 */

// Constants (not user-configurable)
const PROCESSED_LABEL = 'ContractProcessed';
const ALLOWED_EXTENSIONS = ['pdf', 'doc', 'docx'];
const CONTRACTS_SHEET_NAME = 'Contracts';

/**
 * Main function - processes unread emails with PDF/Word attachments
 * This runs on a time-driven trigger (every 5 minutes)
 */
function processIncomingEmails() {
  // Get user configuration
  const userProps = PropertiesService.getUserProperties();
  const folderId = userProps.getProperty('FOLDER_ID');
  const sheetId = userProps.getProperty('SHEET_ID');
  const triggerStartDateStr = userProps.getProperty('TRIGGER_START_DATE');

  // Validate configuration
  if (!folderId || !sheetId) {
    Logger.log('Setup incomplete. Please configure the add-on first.');
    return;
  }

  if (!triggerStartDateStr) {
    Logger.log('Trigger start date not set.');
    return;
  }

  const triggerStartDate = new Date(triggerStartDateStr);

  // Get resources
  let driveFolder, sheet;
  try {
    driveFolder = DriveApp.getFolderById(folderId);
  } catch (e) {
    Logger.log('Cannot access Drive folder: ' + e.message);
    return;
  }

  try {
    sheet = SpreadsheetApp.openById(sheetId).getSheetByName(CONTRACTS_SHEET_NAME);
    if (!sheet) {
      Logger.log('Contracts sheet not found.');
      return;
    }
  } catch (e) {
    Logger.log('Cannot access spreadsheet: ' + e.message);
    return;
  }

  const processedLabel = getOrCreateLabel(PROCESSED_LABEL);

  // Format date for Gmail search query (YYYY/MM/DD)
  const afterDate = Utilities.formatDate(triggerStartDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  // Search for unread emails that haven't been processed yet, received after trigger was set
  const searchQuery = `is:unread -label:${PROCESSED_LABEL} has:attachment after:${afterDate}`;
  const threads = GmailApp.search(searchQuery, 0, 50);

  let processedCount = 0;

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const message of messages) {
      if (message.isUnread()) {
        const attachments = message.getAttachments();
        const validAttachments = filterValidAttachments(attachments);

        if (validAttachments.length > 0) {
          // Process each valid attachment
          for (const attachment of validAttachments) {
            const result = processAttachment(message, attachment, driveFolder, sheet);
            if (result) {
              processedCount++;
            }
          }
        }
      }
    }

    // Mark the thread as processed
    thread.addLabel(processedLabel);
  }

  if (processedCount > 0) {
    Logger.log(`Processed ${processedCount} contract attachment(s)`);
  }
}

/**
 * Filters attachments to only include PDF and Word documents
 */
function filterValidAttachments(attachments) {
  return attachments.filter(attachment => {
    const fileName = attachment.getName().toLowerCase();
    return ALLOWED_EXTENSIONS.some(ext => fileName.endsWith(`.${ext}`));
  });
}

/**
 * Processes a single attachment - saves to Drive and logs to Sheet
 */
function processAttachment(message, attachment, driveFolder, sheet) {
  try {
    const fileName = attachment.getName();
    const senderEmail = extractEmail(message.getFrom());
    const senderName = extractName(message.getFrom());
    const subject = message.getSubject();
    const receivedDate = message.getDate();

    // Create year/month subfolder structure
    const yearMonth = Utilities.formatDate(receivedDate, Session.getScriptTimeZone(), 'yyyy/MM-MMMM');
    const targetFolder = getOrCreateSubfolder(driveFolder, yearMonth);

    // Generate unique filename: DATE_SENDER_ORIGINALNAME
    const datePrefix = Utilities.formatDate(receivedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const sanitizedSender = senderEmail.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
    const newFileName = `${datePrefix}_${sanitizedSender}_${fileName}`;

    // Save to Drive
    const blob = attachment.copyBlob();
    blob.setName(newFileName);
    const driveFile = targetFolder.createFile(blob);
    const driveLink = driveFile.getUrl();

    // Generate Contract ID
    const contractId = generateContractId(sheet);

    // Log to Sheet
    const now = new Date();
    sheet.appendRow([
      contractId,
      subject,
      senderName,
      senderEmail,
      receivedDate,
      fileName,
      driveLink,
      'Received',  // Initial status
      'Unassigned', // Initial assignment
      '',  // Notes
      now  // Last Updated
    ]);

    Logger.log(`Processed: ${fileName} from ${senderEmail}`);
    return true;

  } catch (error) {
    Logger.log(`Error processing attachment: ${error.message}`);
    return false;
  }
}

/**
 * Gets or creates a Gmail label
 */
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

/**
 * Gets or creates nested subfolders (e.g., "2026/03-March")
 */
function getOrCreateSubfolder(parentFolder, path) {
  const parts = path.split('/');
  let currentFolder = parentFolder;

  for (const part of parts) {
    const folders = currentFolder.getFoldersByName(part);
    if (folders.hasNext()) {
      currentFolder = folders.next();
    } else {
      currentFolder = currentFolder.createFolder(part);
    }
  }

  return currentFolder;
}

/**
 * Extracts email address from "Name <email@example.com>" format
 */
function extractEmail(fromField) {
  const emailMatch = fromField.match(/<([^>]+)>/);
  if (emailMatch) {
    return emailMatch[1];
  }
  // If no angle brackets, assume the whole string is the email
  return fromField.trim();
}

/**
 * Extracts name from "Name <email@example.com>" format
 */
function extractName(fromField) {
  const nameMatch = fromField.match(/^([^<]+)</);
  if (nameMatch) {
    return nameMatch[1].trim().replace(/"/g, '');
  }
  // If no name found, return email
  return extractEmail(fromField);
}

/**
 * Generates a unique contract ID (e.g., CTR-2026-0001)
 */
function generateContractId(sheet) {
  const year = new Date().getFullYear();
  const data = sheet.getDataRange().getValues();

  // Count contracts from this year
  let maxNum = 0;
  const prefix = `CTR-${year}-`;

  for (let i = 1; i < data.length; i++) { // Skip header row
    const id = data[i][0];
    if (id && id.toString().startsWith(prefix)) {
      const num = parseInt(id.toString().replace(prefix, ''), 10);
      if (num > maxNum) {
        maxNum = num;
      }
    }
  }

  const nextNum = String(maxNum + 1).padStart(4, '0');
  return `${prefix}${nextNum}`;
}
