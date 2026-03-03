/**
 * Card UI components for the Contract Organiser Add-on
 */

/**
 * Homepage trigger - shows setup or dashboard based on user state
 */
function onHomepage(e) {
  const userProps = PropertiesService.getUserProperties();
  const folderId = userProps.getProperty('FOLDER_ID');
  const sheetId = userProps.getProperty('SHEET_ID');

  // Check if user has completed setup
  if (folderId && sheetId) {
    return buildDashboardCard();
  } else {
    return buildSetupCard();
  }
}

/**
 * Builds the setup card for new users
 */
function buildSetupCard() {
  const card = CardService.newCardBuilder();

  card.setHeader(
    CardService.newCardHeader()
      .setTitle('Contract Organiser')
      .setSubtitle('Setup')
  );

  // Instructions section
  const instructionSection = CardService.newCardSection()
    .addWidget(
      CardService.newTextParagraph()
        .setText('<b>Welcome!</b>\n\nThis add-on automatically saves PDF/Word attachments from your emails to Google Drive and tracks them in a spreadsheet.')
    );

  card.addSection(instructionSection);

  // Folder input section
  const folderSection = CardService.newCardSection()
    .setHeader('Google Drive Folder');

  folderSection.addWidget(
    CardService.newTextParagraph()
      .setText('Paste a shared folder link, or create a new one:')
  );

  folderSection.addWidget(
    CardService.newTextInput()
      .setFieldName('folderLink')
      .setTitle('Folder Link')
      .setHint('https://drive.google.com/drive/folders/...')
  );

  folderSection.addWidget(
    CardService.newTextButton()
      .setText('Create New Folder')
      .setOnClickAction(
        CardService.newAction().setFunctionName('onCreateFolder')
      )
  );

  card.addSection(folderSection);

  // Sheet input section
  const sheetSection = CardService.newCardSection()
    .setHeader('Tracking Spreadsheet');

  sheetSection.addWidget(
    CardService.newTextParagraph()
      .setText('Paste a shared spreadsheet link, or create a new one:')
  );

  sheetSection.addWidget(
    CardService.newTextInput()
      .setFieldName('sheetLink')
      .setTitle('Spreadsheet Link')
      .setHint('https://docs.google.com/spreadsheets/d/...')
  );

  sheetSection.addWidget(
    CardService.newTextButton()
      .setText('Create New Spreadsheet')
      .setOnClickAction(
        CardService.newAction().setFunctionName('onCreateSheet')
      )
  );

  card.addSection(sheetSection);

  // User name section
  const userSection = CardService.newCardSection()
    .setHeader('Your Details');

  userSection.addWidget(
    CardService.newTextInput()
      .setFieldName('userName')
      .setTitle('Your Name')
      .setHint('How you want to appear in the tracker')
  );

  card.addSection(userSection);

  // Save button
  const actionSection = CardService.newCardSection();

  actionSection.addWidget(
    CardService.newTextButton()
      .setText('Save & Start Monitoring')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(
        CardService.newAction().setFunctionName('onSaveSetup')
      )
  );

  card.addSection(actionSection);

  return card.build();
}

/**
 * Builds the dashboard card for configured users
 */
function buildDashboardCard() {
  const userProps = PropertiesService.getUserProperties();
  const folderId = userProps.getProperty('FOLDER_ID');
  const sheetId = userProps.getProperty('SHEET_ID');
  const userName = userProps.getProperty('USER_NAME') || 'User';
  const triggerStartDate = userProps.getProperty('TRIGGER_START_DATE');

  const card = CardService.newCardBuilder();

  card.setHeader(
    CardService.newCardHeader()
      .setTitle('Contract Organiser')
      .setSubtitle('Dashboard')
  );

  // Status section
  const statusSection = CardService.newCardSection()
    .setHeader('Status');

  const isMonitoring = checkTriggerExists();
  const statusText = isMonitoring ? '✅ Monitoring Active' : '⏸️ Monitoring Paused';

  statusSection.addWidget(
    CardService.newDecoratedText()
      .setText(statusText)
      .setTopLabel('Email Monitoring')
  );

  if (triggerStartDate) {
    const startDate = new Date(triggerStartDate);
    const formattedDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');
    statusSection.addWidget(
      CardService.newDecoratedText()
        .setText(formattedDate)
        .setTopLabel('Monitoring Since')
    );
  }

  statusSection.addWidget(
    CardService.newDecoratedText()
      .setText(userName)
      .setTopLabel('Logged in as')
  );

  card.addSection(statusSection);

  // Stats section
  const statsSection = CardService.newCardSection()
    .setHeader('Statistics');

  try {
    const stats = getContractStats(sheetId);
    statsSection.addWidget(
      CardService.newDecoratedText()
        .setText(stats.total.toString())
        .setTopLabel('Total Contracts')
    );
    statsSection.addWidget(
      CardService.newDecoratedText()
        .setText(stats.received.toString())
        .setTopLabel('Received')
    );
    statsSection.addWidget(
      CardService.newDecoratedText()
        .setText(stats.underReview.toString())
        .setTopLabel('Under Review')
    );
    statsSection.addWidget(
      CardService.newDecoratedText()
        .setText(stats.unassigned.toString())
        .setTopLabel('Unassigned')
    );
  } catch (e) {
    statsSection.addWidget(
      CardService.newTextParagraph()
        .setText('Unable to load stats. Check spreadsheet access.')
    );
  }

  card.addSection(statsSection);

  // Quick links section
  const linksSection = CardService.newCardSection()
    .setHeader('Quick Links');

  linksSection.addWidget(
    CardService.newTextButton()
      .setText('Open Tracking Sheet')
      .setOpenLink(
        CardService.newOpenLink()
          .setUrl('https://docs.google.com/spreadsheets/d/' + sheetId)
          .setOpenAs(CardService.OpenAs.FULL_SIZE)
      )
  );

  linksSection.addWidget(
    CardService.newTextButton()
      .setText('Open Drive Folder')
      .setOpenLink(
        CardService.newOpenLink()
          .setUrl('https://drive.google.com/drive/folders/' + folderId)
          .setOpenAs(CardService.OpenAs.FULL_SIZE)
      )
  );

  card.addSection(linksSection);

  // Actions section
  const actionsSection = CardService.newCardSection()
    .setHeader('Actions');

  if (isMonitoring) {
    actionsSection.addWidget(
      CardService.newTextButton()
        .setText('Pause Monitoring')
        .setOnClickAction(
          CardService.newAction().setFunctionName('onPauseMonitoring')
        )
    );
  } else {
    actionsSection.addWidget(
      CardService.newTextButton()
        .setText('Start Monitoring')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(
          CardService.newAction().setFunctionName('onStartMonitoring')
        )
    );
  }

  actionsSection.addWidget(
    CardService.newTextButton()
      .setText('Process Emails Now')
      .setOnClickAction(
        CardService.newAction().setFunctionName('onProcessNow')
      )
  );

  actionsSection.addWidget(
    CardService.newTextButton()
      .setText('Reset Setup')
      .setOnClickAction(
        CardService.newAction().setFunctionName('onResetSetup')
      )
  );

  card.addSection(actionsSection);

  return card.build();
}

/**
 * Builds a success card after folder/sheet creation
 */
function buildResourceCreatedCard(resourceType, resourceName, resourceUrl, resourceId) {
  const card = CardService.newCardBuilder();

  card.setHeader(
    CardService.newCardHeader()
      .setTitle('✅ ' + resourceType + ' Created!')
  );

  const infoSection = CardService.newCardSection();

  infoSection.addWidget(
    CardService.newDecoratedText()
      .setText(resourceName)
      .setTopLabel(resourceType + ' Name')
  );

  infoSection.addWidget(
    CardService.newTextParagraph()
      .setText('<b>⚠️ Important:</b>\nShare this ' + resourceType.toLowerCase() + ' with your team so they can access it.')
  );

  infoSection.addWidget(
    CardService.newTextButton()
      .setText('Open & Share')
      .setOpenLink(
        CardService.newOpenLink()
          .setUrl(resourceUrl)
          .setOpenAs(CardService.OpenAs.FULL_SIZE)
      )
  );

  card.addSection(infoSection);

  // Copy link section
  const linkSection = CardService.newCardSection()
    .setHeader('Share this link with your team');

  linkSection.addWidget(
    CardService.newTextParagraph()
      .setText(resourceUrl)
  );

  card.addSection(linkSection);

  // Continue button
  const actionSection = CardService.newCardSection();

  actionSection.addWidget(
    CardService.newTextButton()
      .setText('Continue Setup')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(
        CardService.newAction()
          .setFunctionName('onContinueSetup')
          .setParameters({
            'resourceType': resourceType.toLowerCase(),
            'resourceId': resourceId
          })
      )
  );

  card.addSection(actionSection);

  return card.build();
}

/**
 * Builds an error card
 */
function buildErrorCard(title, message) {
  const card = CardService.newCardBuilder();

  card.setHeader(
    CardService.newCardHeader()
      .setTitle('❌ ' + title)
  );

  const section = CardService.newCardSection();

  section.addWidget(
    CardService.newTextParagraph()
      .setText(message)
  );

  section.addWidget(
    CardService.newTextButton()
      .setText('Go Back')
      .setOnClickAction(
        CardService.newAction().setFunctionName('onHomepage')
      )
  );

  card.addSection(section);

  return card.build();
}

/**
 * Builds a success notification card
 */
function buildSuccessCard(title, message) {
  const card = CardService.newCardBuilder();

  card.setHeader(
    CardService.newCardHeader()
      .setTitle('✅ ' + title)
  );

  const section = CardService.newCardSection();

  section.addWidget(
    CardService.newTextParagraph()
      .setText(message)
  );

  section.addWidget(
    CardService.newTextButton()
      .setText('Back to Dashboard')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(
        CardService.newAction().setFunctionName('onHomepage')
      )
  );

  card.addSection(section);

  return card.build();
}
