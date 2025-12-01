/**
 * ============================================
 * Bulk Email Sender
 * ============================================
 * 
 * How to Use:
 * 1. Open Google Sheets
 * 2. Click Extensions → Apps Script
 * 3. Delete existing code and paste this script
 * 4. Click Save (💾)
 * 5. Refresh Google Sheet
 * 6. You will see a new menu: "📧 Email Sender"
 * 7. Click "🔧 Initialize Sheets" to create Recipients and Template tabs
 * 
 * Sheet Structure:
 * - "Recipients" tab: Contains all recipient information
 * - "Template" tab: Contains editable email template and default values
 */

// ============================================
// Configuration
// ============================================

const CONFIG = {
  // Sheet names
  SHEETS: {
    RECIPIENTS: 'Recipients',
    TEMPLATE: 'Template'
  },
  
  // Column names
  COLUMNS: {
    EMAIL: 'email',
    NAME: 'greeting_first_name',
    MESSAGE: 'message',
    STATUS: 'status',
    SENT_DATE: 'sent_date'
  }
};

// Initial default values (only used when creating template)
const INITIAL_DEFAULTS = {
  SUBJECT: 'Hello',
  NAME: 'Sir/Madam',
  MESSAGE: 'This is a test email'
};

// Default email template (only used during initialization)
const DEFAULT_TEMPLATE = `Dear {{greeting_first_name}},

{{message}}

Best regards`;

// ============================================
// Menu Setup
// ============================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 Email Sender')
    .addItem('✉️ Send Emails', 'showSendDialog')
    .addItem('👁️ Preview', 'showPreviewDialog')
    .addSeparator()
    .addItem('🔧 Initialize Sheets', 'initializeSheets')
    .addSeparator()
    .addItem('❓ Help', 'showHelp')
    .addToUi();
}

// ============================================
// Sheet Initialization
// ============================================

/**
 * Initialize all required sheets
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Initialize Sheets',
    'This will create "Recipients" and "Template" tabs.\nExisting tabs with the same name will not be overwritten.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  // Create Recipients sheet
  createRecipientsSheet(ss);
  
  // Create Template sheet
  createTemplateSheet(ss);
  
  ui.alert('✅ Initialization Complete', 
    'Created the following tabs:\n\n' +
    '1. "Recipients" - Add recipient information here\n' +
    '2. "Template" - Edit email template and default values here\n\n' +
    'Please edit the template first, then add recipients.', 
    ui.ButtonSet.OK);
}

/**
 * Create Recipients sheet
 */
function createRecipientsSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.RECIPIENTS);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.RECIPIENTS);
    
    // Add headers
    const headers = [
      CONFIG.COLUMNS.EMAIL,
      CONFIG.COLUMNS.NAME,
      CONFIG.COLUMNS.MESSAGE,
      CONFIG.COLUMNS.STATUS,
      CONFIG.COLUMNS.SENT_DATE
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1a73e8');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // Set column widths
    sheet.setColumnWidth(1, 250); // email
    sheet.setColumnWidth(2, 150); // name
    sheet.setColumnWidth(3, 350); // message
    sheet.setColumnWidth(4, 100); // status
    sheet.setColumnWidth(5, 180); // sent_date
    
    // Add sample data
    sheet.getRange(2, 1, 3, 3).setValues([
      ['example1@email.com', 'John', 'Your recent video really impressed me'],
      ['example2@email.com', 'Jane', 'Your recent project is outstanding'],
      ['example3@email.com', '', '']  // Example using default values
    ]);
    
    // Add notes
    sheet.getRange(2, 4).setNote('Auto-filled after sending');
    sheet.getRange(2, 5).setNote('Auto-filled after sending');
    
    // Freeze header row
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * Create Template sheet
 */
function createTemplateSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEETS.TEMPLATE);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.TEMPLATE);
    
    // Set template structure
    const templateData = [
      ['📧 Email Template Settings', ''],
      ['', ''],
      ['Subject', INITIAL_DEFAULTS.SUBJECT],
      ['', ''],
      ['Body Template', ''],
      [DEFAULT_TEMPLATE, ''],
      ['', ''],
      ['', ''],
      ['', ''],
      ['', ''],
      ['⚙️ Default Values', ''],
      ['', ''],
      ['Default Greeting (used when name is empty)', INITIAL_DEFAULTS.NAME],
      ['Default Message (used when message is empty)', INITIAL_DEFAULTS.MESSAGE],
      ['', ''],
      ['📌 Available Variables', ''],
      ['Variable', 'Description'],
      ['{{greeting_first_name}}', 'Recipient name (uses default greeting if empty)'],
      ['{{message}}', 'Personalized message (uses default message if empty)'],
      ['{{email}}', 'Recipient email address'],
      ['', ''],
      ['💡 Tips', ''],
      ['- Edit content directly in the yellow cells above', ''],
      ['- Variables will be automatically replaced with recipient data', ''],
      ['- If recipient data is empty, default values will be used', '']
    ];
    
    sheet.getRange(1, 1, templateData.length, 2).setValues(templateData);
    
    // Format titles
    sheet.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#1a73e8');
    sheet.getRange('A11').setFontSize(14).setFontWeight('bold').setFontColor('#1a73e8');
    sheet.getRange('A16').setFontSize(14).setFontWeight('bold').setFontColor('#1a73e8');
    sheet.getRange('A22').setFontSize(14).setFontWeight('bold').setFontColor('#1a73e8');
    
    // Format labels
    sheet.getRange('A3').setFontWeight('bold').setBackground('#f3f3f3');
    sheet.getRange('A5').setFontWeight('bold').setBackground('#f3f3f3');
    sheet.getRange('A13').setFontWeight('bold').setBackground('#f3f3f3');
    sheet.getRange('A14').setFontWeight('bold').setBackground('#f3f3f3');
    
    // Format subject input
    sheet.getRange('B3')
      .setBackground('#fff9c4')
      .setBorder(true, true, true, true, false, false, '#fbc02d', SpreadsheetApp.BorderStyle.SOLID);
    
    // Format body input - merge multiple rows for input
    sheet.getRange('A6:B10')
      .merge()
      .setBackground('#fff9c4')
      .setBorder(true, true, true, true, false, false, '#fbc02d', SpreadsheetApp.BorderStyle.SOLID)
      .setVerticalAlignment('top')
      .setWrap(true);
    
    // Format default value inputs
    sheet.getRange('B13')
      .setBackground('#fff9c4')
      .setBorder(true, true, true, true, false, false, '#fbc02d', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange('B14')
      .setBackground('#fff9c4')
      .setBorder(true, true, true, true, false, false, '#fbc02d', SpreadsheetApp.BorderStyle.SOLID);
    
    // Format variable description header
    sheet.getRange('A17:B17')
      .setFontWeight('bold')
      .setBackground('#e8f0fe');
    
    // Format variable list
    sheet.getRange('A18:A20')
      .setFontFamily('Courier New')
      .setBackground('#f8f9fa');
    
    // Set column widths
    sheet.setColumnWidth(1, 300);
    sheet.setColumnWidth(2, 400);
    
    // Set row height (body area)
    sheet.setRowHeight(6, 120);
  }
  
  return sheet;
}

// ============================================
// Data Functions
// ============================================

/**
 * Get template and default values from Template sheet
 */
function getTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TEMPLATE);
  
  if (!sheet) {
    // If template sheet doesn't exist, return initial defaults
    return {
      subject: INITIAL_DEFAULTS.SUBJECT,
      body: DEFAULT_TEMPLATE,
      defaultName: INITIAL_DEFAULTS.NAME,
      defaultMessage: INITIAL_DEFAULTS.MESSAGE
    };
  }
  
  // Read subject (cell B3)
  const subject = sheet.getRange('B3').getValue() || INITIAL_DEFAULTS.SUBJECT;
  
  // Read body (merged cells A6:B10)
  const body = sheet.getRange('A6').getValue() || DEFAULT_TEMPLATE;
  
  // Read default values (cells B13 and B14)
  const defaultName = sheet.getRange('B13').getValue() || INITIAL_DEFAULTS.NAME;
  const defaultMessage = sheet.getRange('B14').getValue() || INITIAL_DEFAULTS.MESSAGE;
  
  return {
    subject: subject.toString(),
    body: body.toString(),
    defaultName: defaultName.toString(),
    defaultMessage: defaultMessage.toString()
  };
}

/**
 * Get all recipients from Recipients sheet
 */
function getRecipients() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RECIPIENTS);
  
  if (!sheet) {
    throw new Error('Cannot find "Recipients" tab. Please click "🔧 Initialize Sheets" first.');
  }
  
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) return [];
  
  const headers = data[0].map(h => h.toString().toLowerCase().trim());
  const emailCol = headers.indexOf(CONFIG.COLUMNS.EMAIL);
  const nameCol = headers.indexOf(CONFIG.COLUMNS.NAME);
  const messageCol = headers.indexOf(CONFIG.COLUMNS.MESSAGE);
  const statusCol = headers.indexOf(CONFIG.COLUMNS.STATUS);
  
  if (emailCol === -1) {
    throw new Error('Cannot find "email" column. Please check the Recipients tab headers.');
  }
  
  const recipients = [];
  
  for (let i = 1; i < data.length; i++) {
    const email = data[i][emailCol]?.toString().trim();
    if (!email || !isValidEmail(email)) continue;
    
    // Skip already sent
    const status = statusCol !== -1 ? data[i][statusCol]?.toString() : '';
    if (status === 'Sent') continue;
    
    recipients.push({
      row: i + 1,
      email: email,
      greeting_first_name: nameCol !== -1 ? data[i][nameCol]?.toString() || '' : '',
      message: messageCol !== -1 ? data[i][messageCol]?.toString() || '' : ''
    });
  }
  
  return recipients;
}

/**
 * Validate email format
 */
function isValidEmail(email) {
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(email);
}

/**
 * Render template with recipient data
 */
function renderTemplate(template, recipient, defaults) {
  let rendered = template;
  
  // Replace {{field}} or {{field|default:'value'}} patterns
  const regex = /\{\{\s*(\w+)(?:\|default:'([^']*)')?\s*\}\}/g;
  
  rendered = rendered.replace(regex, (match, field, defaultValue) => {
    const value = recipient[field];
    if (value && value.toString().trim() !== '') {
      return value;
    }
    if (defaultValue) {
      return defaultValue;
    }
    // Use default values from Template sheet
    if (field === 'greeting_first_name') return defaults.defaultName;
    if (field === 'message') return defaults.defaultMessage;
    return '';
  });
  
  return rendered;
}

// ============================================
// Email Sending
// ============================================

/**
 * Send emails to all pending recipients
 */
function sendEmails(testMode = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RECIPIENTS);
  
  if (!sheet) {
    return { success: false, message: 'Cannot find "Recipients" tab.' };
  }
  
  const recipients = getRecipients();
  const template = getTemplate();
  
  if (recipients.length === 0) {
    return { success: false, message: 'No recipients to send to.' };
  }
  
  // Find status and date columns
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let statusCol = headers.map(h => h.toString().toLowerCase()).indexOf(CONFIG.COLUMNS.STATUS) + 1;
  let dateCol = headers.map(h => h.toString().toLowerCase()).indexOf(CONFIG.COLUMNS.SENT_DATE) + 1;
  
  // Create columns if they don't exist
  if (statusCol === 0) {
    statusCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, statusCol).setValue(CONFIG.COLUMNS.STATUS);
  }
  if (dateCol === 0) {
    dateCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, dateCol).setValue(CONFIG.COLUMNS.SENT_DATE);
  }
  
  let successCount = 0;
  let failCount = 0;
  const errors = [];
  
  const limit = testMode ? 1 : recipients.length;
  
  for (let i = 0; i < limit; i++) {
    const recipient = recipients[i];
    
    try {
      const subject = renderTemplate(template.subject, recipient, template);
      const body = renderTemplate(template.body, recipient, template);
      
      // Send email
      GmailApp.sendEmail(recipient.email, subject, body);
      
      // Update status
      sheet.getRange(recipient.row, statusCol).setValue('Sent');
      sheet.getRange(recipient.row, dateCol).setValue(new Date());
      
      // Highlight row green
      sheet.getRange(recipient.row, 1, 1, sheet.getLastColumn())
        .setBackground('#d4edda');
      
      successCount++;
      
      if (i < limit - 1) {
        Utilities.sleep(500);
      }
      
    } catch (error) {
      sheet.getRange(recipient.row, statusCol).setValue('Error');
      sheet.getRange(recipient.row, 1, 1, sheet.getLastColumn())
        .setBackground('#f8d7da');
      
      errors.push(`${recipient.email}: ${error.message}`);
      failCount++;
    }
  }
  
  return {
    success: true,
    successCount: successCount,
    failCount: failCount,
    errors: errors,
    testMode: testMode
  };
}

/**
 * Send test email to yourself
 */
function sendTestEmail() {
  const template = getTemplate();
  const myEmail = Session.getActiveUser().getEmail();
  
  const testRecipient = {
    email: myEmail,
    greeting_first_name: 'Test User',
    message: 'This is a test email. Please verify the content before actual sending.'
  };
  
  const subject = renderTemplate(template.subject, testRecipient, template) + ' [TEST]';
  const body = renderTemplate(template.body, testRecipient, template);
  
  try {
    GmailApp.sendEmail(myEmail, subject, body);
    return { success: true, email: myEmail };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ============================================
// Dialog Interfaces
// ============================================

/**
 * Show send confirmation dialog
 */
function showSendDialog() {
  let recipients = [];
  let template = { subject: '', body: '', defaultName: '', defaultMessage: '' };
  let errorMessage = '';
  
  try {
    recipients = getRecipients();
    template = getTemplate();
  } catch (error) {
    errorMessage = error.message;
  }
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { 
          font-family: 'Google Sans', Arial, sans-serif; 
          padding: 24px;
          background: #f8f9fa;
          color: #202124;
        }
        .header { margin-bottom: 24px; }
        h2 { font-size: 20px; font-weight: 500; margin-bottom: 8px; }
        .subtitle { color: #5f6368; font-size: 14px; }
        .stats {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 16px;
          margin-bottom: 24px;
        }
        .stat-card {
          background: white;
          padding: 20px;
          border-radius: 12px;
          box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .stat-number {
          font-size: 32px;
          font-weight: 600;
          color: #1a73e8;
        }
        .stat-label {
          font-size: 13px;
          color: #5f6368;
          margin-top: 4px;
        }
        .template-preview {
          background: white;
          padding: 16px;
          border-radius: 12px;
          margin-bottom: 24px;
          box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .template-label {
          font-size: 12px;
          color: #5f6368;
          margin-bottom: 8px;
          text-transform: uppercase;
          letter-spacing: 0.5px;
        }
        .template-subject {
          font-weight: 500;
          margin-bottom: 12px;
          padding-bottom: 12px;
          border-bottom: 1px solid #e0e0e0;
        }
        .template-body {
          font-size: 14px;
          white-space: pre-wrap;
          color: #3c4043;
          line-height: 1.6;
          max-height: 120px;
          overflow-y: auto;
        }
        .template-edit-hint {
          margin-top: 12px;
          padding-top: 12px;
          border-top: 1px solid #e0e0e0;
          font-size: 12px;
          color: #5f6368;
        }
        .template-edit-hint a {
          color: #1a73e8;
          cursor: pointer;
          text-decoration: none;
        }
        .buttons {
          display: flex;
          gap: 12px;
          justify-content: flex-end;
        }
        button {
          padding: 12px 24px;
          border-radius: 8px;
          font-size: 14px;
          font-weight: 500;
          cursor: pointer;
          border: none;
          transition: all 0.2s;
        }
        .btn-secondary {
          background: #f1f3f4;
          color: #3c4043;
        }
        .btn-secondary:hover { background: #e8eaed; }
        .btn-test {
          background: #e8f0fe;
          color: #1a73e8;
        }
        .btn-test:hover { background: #d2e3fc; }
        .btn-primary {
          background: #1a73e8;
          color: white;
        }
        .btn-primary:hover { background: #1557b0; }
        .btn-primary:disabled {
          background: #dadce0;
          cursor: not-allowed;
        }
        .status {
          text-align: center;
          padding: 16px;
          margin-bottom: 16px;
          border-radius: 8px;
          display: none;
        }
        .status.success { background: #e6f4ea; color: #137333; display: block; }
        .status.error { background: #fce8e6; color: #c5221f; display: block; }
        .status.sending { background: #e8f0fe; color: #1a73e8; display: block; }
        .warning {
          background: #fef7e0;
          color: #b06000;
          padding: 12px 16px;
          border-radius: 8px;
          margin-bottom: 16px;
          font-size: 13px;
        }
        .error-box {
          background: #fce8e6;
          color: #c5221f;
          padding: 12px 16px;
          border-radius: 8px;
          margin-bottom: 16px;
          font-size: 13px;
        }
        .progress {
          margin-top: 8px;
          font-size: 13px;
        }
      </style>
    </head>
    <body>
      <div class="header">
        <h2>📧 Send Emails</h2>
        <p class="subtitle">Bulk send with the following content</p>
      </div>
      
      <div id="statusMessage" class="status"></div>
      
      ${errorMessage ? `
        <div class="error-box">
          ❌ ${errorMessage}
        </div>
      ` : ''}
      
      ${!errorMessage && recipients.length === 0 ? `
        <div class="warning">
          ⚠️ No recipients to send to. Please add email addresses in the "Recipients" tab, or clear the "Sent" status.
        </div>
      ` : ''}
      
      <div class="stats">
        <div class="stat-card">
          <div class="stat-number">${recipients.length}</div>
          <div class="stat-label">Pending</div>
        </div>
        <div class="stat-card">
          <div class="stat-number">${Math.ceil(recipients.length / 100)}</div>
          <div class="stat-label">Est. Time (minutes)</div>
        </div>
      </div>
      
      <div class="template-preview">
        <div class="template-label">Current Template</div>
        <div class="template-subject">Subject: ${template.subject}</div>
        <div class="template-body">${template.body}</div>
        <div class="template-edit-hint">
          💡 To edit template, switch to the "Template" tab
        </div>
      </div>
      
      <div class="buttons">
        <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
        <button class="btn-test" onclick="sendTest()" ${errorMessage ? 'disabled' : ''}>🧪 Test Send</button>
        <button class="btn-primary" id="sendBtn" onclick="sendAll()" ${recipients.length === 0 || errorMessage ? 'disabled' : ''}>
          ✉️ Send ${recipients.length} Emails
        </button>
      </div>
      
      <script>
        function sendTest() {
          document.getElementById('statusMessage').className = 'status sending';
          document.getElementById('statusMessage').innerHTML = '⏳ Sending test email...';
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                document.getElementById('statusMessage').className = 'status success';
                document.getElementById('statusMessage').innerHTML = '✅ Test email sent to: ' + result.email;
              } else {
                document.getElementById('statusMessage').className = 'status error';
                document.getElementById('statusMessage').innerHTML = '❌ Error: ' + result.error;
              }
            })
            .withFailureHandler(function(error) {
              document.getElementById('statusMessage').className = 'status error';
              document.getElementById('statusMessage').innerHTML = '❌ Error: ' + error.message;
            })
            .sendTestEmail();
        }
        
        function sendAll() {
          if (!confirm('Are you sure you want to send ${recipients.length} emails?')) return;
          
          document.getElementById('sendBtn').disabled = true;
          document.getElementById('sendBtn').textContent = 'Sending...';
          document.getElementById('statusMessage').className = 'status sending';
          document.getElementById('statusMessage').innerHTML = '⏳ Sending emails, please wait...<div class="progress">Do not close this window</div>';
          
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                document.getElementById('statusMessage').className = 'status success';
                document.getElementById('statusMessage').innerHTML = 
                  '✅ Complete!<br>Success: ' + result.successCount + ' emails' +
                  (result.failCount > 0 ? '<br>Failed: ' + result.failCount + ' emails' : '');
                document.getElementById('sendBtn').textContent = '✓ Done';
              } else {
                document.getElementById('statusMessage').className = 'status error';
                document.getElementById('statusMessage').innerHTML = '❌ ' + result.message;
                document.getElementById('sendBtn').disabled = false;
                document.getElementById('sendBtn').textContent = 'Retry';
              }
            })
            .withFailureHandler(function(error) {
              document.getElementById('statusMessage').className = 'status error';
              document.getElementById('statusMessage').innerHTML = '❌ Error: ' + error.message;
              document.getElementById('sendBtn').disabled = false;
              document.getElementById('sendBtn').textContent = 'Retry';
            })
            .sendEmails(false);
        }
      </script>
    </body>
    </html>
  `)
  .setWidth(480)
  .setHeight(580);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Send Emails');
}

/**
 * Show preview dialog
 */
function showPreviewDialog() {
  let recipients = [];
  let template = { subject: '', body: '', defaultName: '', defaultMessage: '' };
  let errorMessage = '';
  
  try {
    recipients = getRecipients();
    template = getTemplate();
  } catch (error) {
    errorMessage = error.message;
  }
  
  // Preview first recipient or sample
  const previewRecipient = recipients.length > 0 ? recipients[0] : {
    email: 'sample@example.com',
    greeting_first_name: 'John',
    message: 'This is a sample message'
  };
  
  const renderedBody = renderTemplate(template.body, previewRecipient, template);
  const renderedSubject = renderTemplate(template.subject, previewRecipient, template);
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { 
          font-family: 'Google Sans', Arial, sans-serif; 
          padding: 24px;
          background: #f8f9fa;
        }
        h2 { font-size: 18px; margin-bottom: 20px; color: #202124; }
        .error-box {
          background: #fce8e6;
          color: #c5221f;
          padding: 12px 16px;
          border-radius: 8px;
          margin-bottom: 16px;
          font-size: 13px;
        }
        .email-preview {
          background: white;
          border-radius: 12px;
          overflow: hidden;
          box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        .email-header {
          background: #f1f3f4;
          padding: 16px 20px;
          border-bottom: 1px solid #e0e0e0;
        }
        .email-field {
          display: flex;
          margin-bottom: 8px;
          font-size: 14px;
        }
        .email-field:last-child { margin-bottom: 0; }
        .email-label {
          color: #5f6368;
          width: 60px;
          flex-shrink: 0;
        }
        .email-value {
          color: #202124;
        }
        .email-body {
          padding: 24px;
          font-size: 15px;
          line-height: 1.8;
          white-space: pre-wrap;
          color: #202124;
        }
        .recipient-selector {
          margin-bottom: 20px;
        }
        select {
          width: 100%;
          padding: 12px;
          border: 1px solid #dadce0;
          border-radius: 8px;
          font-size: 14px;
          background: white;
        }
        .buttons {
          margin-top: 20px;
          text-align: right;
        }
        button {
          padding: 10px 20px;
          border: none;
          border-radius: 6px;
          font-size: 14px;
          cursor: pointer;
          background: #1a73e8;
          color: white;
        }
        button:hover { background: #1557b0; }
      </style>
    </head>
    <body>
      <h2>👁️ Email Preview</h2>
      
      ${errorMessage ? `<div class="error-box">❌ ${errorMessage}</div>` : ''}
      
      ${recipients.length > 1 ? `
        <div class="recipient-selector">
          <select id="recipientSelect" onchange="updatePreview()">
            ${recipients.map((r, i) => `
              <option value="${i}">${r.greeting_first_name || template.defaultName} (${r.email})</option>
            `).join('')}
          </select>
        </div>
      ` : ''}
      
      <div class="email-preview">
        <div class="email-header">
          <div class="email-field">
            <span class="email-label">To:</span>
            <span class="email-value" id="previewTo">${previewRecipient.email}</span>
          </div>
          <div class="email-field">
            <span class="email-label">Subject:</span>
            <span class="email-value" id="previewSubject">${renderedSubject}</span>
          </div>
        </div>
        <div class="email-body" id="previewBody">${renderedBody}</div>
      </div>
      
      <div class="buttons">
        <button onclick="google.script.host.close()">Close</button>
      </div>
      
      <script>
        const recipients = ${JSON.stringify(recipients)};
        const template = ${JSON.stringify(template)};
        
        function renderTemplate(text, recipient) {
          return text.replace(/\\{\\{\\s*(\\w+)(?:\\|default:'([^']*)')?\\s*\\}\\}/g, 
            function(match, field, defaultValue) {
              const value = recipient[field];
              if (value && value.toString().trim() !== '') return value;
              if (defaultValue) return defaultValue;
              if (field === 'greeting_first_name') return template.defaultName;
              if (field === 'message') return template.defaultMessage;
              return '';
            }
          );
        }
        
        function updatePreview() {
          const select = document.getElementById('recipientSelect');
          const recipient = recipients[select.value];
          
          document.getElementById('previewTo').textContent = recipient.email;
          document.getElementById('previewSubject').textContent = renderTemplate(template.subject, recipient);
          document.getElementById('previewBody').textContent = renderTemplate(template.body, recipient);
        }
      </script>
    </body>
    </html>
  `)
  .setWidth(500)
  .setHeight(480);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Preview');
}

/**
 * Show help dialog
 */
function showHelp() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { 
          font-family: 'Google Sans', Arial, sans-serif; 
          padding: 24px;
          background: #f8f9fa;
          color: #202124;
        }
        h2 { font-size: 20px; margin-bottom: 20px; }
        .step {
          background: white;
          padding: 16px;
          border-radius: 8px;
          margin-bottom: 12px;
          box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        .step-number {
          display: inline-block;
          width: 24px;
          height: 24px;
          background: #1a73e8;
          color: white;
          border-radius: 50%;
          text-align: center;
          line-height: 24px;
          font-size: 13px;
          margin-right: 12px;
        }
        .step p {
          margin-top: 8px;
          font-size: 14px;
          color: #5f6368;
          line-height: 1.6;
        }
        .sheet-info {
          background: #e8f0fe;
          padding: 16px;
          border-radius: 8px;
          margin-bottom: 20px;
        }
        .sheet-info h3 {
          font-size: 14px;
          color: #1a73e8;
          margin-bottom: 12px;
        }
        .sheet-item {
          display: flex;
          align-items: flex-start;
          margin-bottom: 8px;
          font-size: 13px;
        }
        .sheet-item:last-child { margin-bottom: 0; }
        .sheet-icon {
          margin-right: 8px;
        }
        .tip {
          background: #fff8e1;
          padding: 12px 16px;
          border-radius: 8px;
          margin-top: 16px;
          font-size: 13px;
          color: #f57c00;
        }
        button {
          margin-top: 20px;
          padding: 10px 20px;
          background: #1a73e8;
          color: white;
          border: none;
          border-radius: 6px;
          cursor: pointer;
          font-size: 14px;
        }
      </style>
    </head>
    <body>
      <h2>❓ Help</h2>
      
      <div class="sheet-info">
        <h3>📋 Tab Descriptions</h3>
        <div class="sheet-item">
          <span class="sheet-icon">📝</span>
          <span><strong>"Template" tab</strong> - Edit email subject, body, and default values</span>
        </div>
        <div class="sheet-item">
          <span class="sheet-icon">👥</span>
          <span><strong>"Recipients" tab</strong> - Add recipient information</span>
        </div>
      </div>
      
      <div class="step">
        <span class="step-number">1</span>
        <strong>Initialize Sheets</strong>
        <p>Click "🔧 Initialize Sheets" to create required tabs</p>
      </div>
      
      <div class="step">
        <span class="step-number">2</span>
        <strong>Edit Template</strong>
        <p>Switch to "Template" tab and edit in yellow areas:<br>
        - Email subject<br>
        - Email body<br>
        - Default greeting (used when name is empty)<br>
        - Default message (used when message is empty)</p>
      </div>
      
      <div class="step">
        <span class="step-number">3</span>
        <strong>Add Recipients</strong>
        <p>Switch to "Recipients" tab and fill in email, name, and personalized message</p>
      </div>
      
      <div class="step">
        <span class="step-number">4</span>
        <strong>Preview</strong>
        <p>Click "👁️ Preview" to see actual email appearance</p>
      </div>
      
      <div class="step">
        <span class="step-number">5</span>
        <strong>Send</strong>
        <p>Click "✉️ Send Emails" for bulk sending. Recommend test sending first.</p>
      </div>
      
      <div class="tip">
        💡 <strong>Tip:</strong> Gmail daily limit is 500 emails (2000 for Google Workspace)
      </div>
      
      <button onclick="google.script.host.close()">Close</button>
    </body>
    </html>
  `)
  .setWidth(480)
  .setHeight(620);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Help');
}
