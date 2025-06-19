// To learn how to use this script, refer to the documentation:
// https://developers.google.com/apps-script/samples/automations/mail-merge

/*
Copyright 2022 Martin Hawksey

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/
 
/**
 * @OnlyCurrentDoc
*/
 
/**
 * Change these to match the column names you are using for email
 * recipient addresses, email sent column, batch column, and version column.
*/
const RECIPIENT_COL  = "Recipient";
const EMAIL_SENT_COL = "Email Sent";
const BATCH_COL = "Batch";
const VERSION_COL = "Version";

// Rate limiting constants
const EMAIL_RATE_LIMIT = 95; // Send 95 emails per hour (leaving buffer for safety)
const RESUME_DELAY_HOURS = 1; // Wait 1 hour before resuming
const MAX_RESUME_ATTEMPTS = 48; // Maximum 48 hours of retries
const MIN_DAILY_QUOTA_BUFFER = 10; // Minimum daily quota to keep as buffer
 
/**
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Select Batch & Send Emails', 'showBatchSelectionDialog')
      .addSeparator()
      .addItem('Check Job Status', 'showJobStatus')
      .addItem('Cancel Active Job', 'cancelActiveJob')
      .addItem('Check Email Quotas', 'testEmailQuotas')
      .addSeparator()
      .addItem('Email Service Settings', 'showEmailServiceSettings')
      .addItem('Test MailerSend Configuration', 'testMailerSendConfiguration')
      .addSeparator()
      .addItem('Remove Duplicate Email Addresses', 'removeDuplicateEmails')
      .addToUi();
}

/**
 * Shows the batch selection dialog to user
 */
function showBatchSelectionDialog() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const batches = getUniqueBatchValues(sheet);

  if (batches.length === 0) {
    SpreadsheetApp.getUi().alert('No batch values found',
      `No values found in the "${BATCH_COL}" column. Please ensure your spreadsheet has a "${BATCH_COL}" column with data.`,
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const htmlTemplate = HtmlService.createTemplateFromFile('BatchSelection');
  htmlTemplate.batches = batches;

  const html = htmlTemplate.evaluate()
    .setWidth(500)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'Select Batch & Send Emails');
}

/**
 * Shows the email service settings dialog
 */
function showEmailServiceSettings() {
  const htmlTemplate = HtmlService.createTemplateFromFile('EmailServiceSettings');

  const html = htmlTemplate.evaluate()
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Email Service Settings');
}

/**
 * Gets unique batch values from the batch column
 * @param {Sheet} sheet to read data from
 * @return {Array} array of unique batch values
 */
function getUniqueBatchValues(sheet) {
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  const heads = data.shift();

  const batchColIdx = heads.indexOf(BATCH_COL);
  if (batchColIdx === -1) {
    return [];
  }

  const batchValues = data.map(row => row[batchColIdx]).filter(value => value !== '');
  return [...new Set(batchValues)].sort();
}

/**
 * Processes the batch selection from the dialog
 * @param {Object} formData containing selectedBatch, subjectLine, and optional A/B testing fields
 */
function processBatchSelection(formData) {
  try {
    // Log the received form data for debugging
    console.log('Form data received:', formData);

    const selectedBatch = formData.selectedBatch;
    const enableABTesting = formData.enableABTesting;

    if (!selectedBatch) {
      SpreadsheetApp.getUi().alert('Missing Information',
        'Please select a batch.',
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    if (enableABTesting) {
      // A/B Testing mode
      const subjectA = formData.subjectA;
      const subjectB = formData.subjectB;

      if (!subjectA || !subjectB) {
        SpreadsheetApp.getUi().alert('Missing Information',
          'Please enter both subject lines for A/B testing.',
          SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }

      if (subjectA === subjectB) {
        SpreadsheetApp.getUi().alert('Invalid Input',
          'Subject lines A and B must be different for A/B testing.',
          SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }

      console.log('Starting A/B test with subjects:', subjectA, subjectB);
      sendABTestEmails(subjectA, subjectB, SpreadsheetApp.getActiveSheet(), selectedBatch);
    } else {
      // Regular single subject mode
      const subjectLine = formData.subjectLine;

      if (!subjectLine) {
        SpreadsheetApp.getUi().alert('Missing Information',
          'Please enter a subject line.',
          SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }

      console.log('Starting regular email send with subject:', subjectLine);
      sendEmails(subjectLine, SpreadsheetApp.getActiveSheet(), selectedBatch);
    }
  } catch (error) {
    console.error('Error in processBatchSelection:', error);
    SpreadsheetApp.getUi().alert('Error',
      `An error occurred: ${error.message}\n\nPlease check the console logs for more details.`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Removes duplicate email addresses from the spreadsheet
 * Keeps the first occurrence and removes subsequent duplicates
 */
function removeDuplicateEmails() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // Get all data from the sheet
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();

  if (data.length <= 1) {
    ui.alert('No Data Found',
      'The spreadsheet appears to be empty or only contains headers.',
      ui.ButtonSet.OK);
    return;
  }

  // Get headers and find recipient column
  const headers = data[0];
  const recipientColIdx = headers.indexOf(RECIPIENT_COL);

  if (recipientColIdx === -1) {
    ui.alert('Column Not Found',
      `The "${RECIPIENT_COL}" column was not found in the spreadsheet. Please ensure your spreadsheet has a column named "${RECIPIENT_COL}".`,
      ui.ButtonSet.OK);
    return;
  }

  // Find duplicate email addresses
  const duplicateInfo = findDuplicateEmails(data, recipientColIdx);

  if (duplicateInfo.duplicateRows.length === 0) {
    ui.alert('No Duplicates Found',
      'No duplicate email addresses were found in the spreadsheet.',
      ui.ButtonSet.OK);
    return;
  }

  // Show confirmation dialog
  const duplicateCount = duplicateInfo.duplicateRows.length;
  const uniqueEmailCount = duplicateInfo.duplicateEmails.size;

  const confirmMessage = `Found ${duplicateCount} duplicate email address${duplicateCount > 1 ? 'es' : ''} ` +
    `across ${uniqueEmailCount} unique email${uniqueEmailCount > 1 ? 's' : ''}.\n\n` +
    `This will remove ${duplicateCount} row${duplicateCount > 1 ? 's' : ''} from the spreadsheet, ` +
    `keeping only the first occurrence of each email address.\n\n` +
    `Do you want to proceed?`;

  const response = ui.alert('Confirm Duplicate Removal', confirmMessage, ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return;
  }

  // Remove duplicate rows (in reverse order to maintain row indices)
  const rowsToDelete = duplicateInfo.duplicateRows.sort((a, b) => b - a);

  for (const rowIndex of rowsToDelete) {
    sheet.deleteRow(rowIndex + 1); // +1 because sheet rows are 1-indexed
  }

  // Show completion message
  ui.alert('Duplicates Removed',
    `Successfully removed ${duplicateCount} duplicate row${duplicateCount > 1 ? 's' : ''}. ` +
    `The spreadsheet now contains only unique email addresses.`,
    ui.ButtonSet.OK);
}

/**
 * Finds duplicate email addresses in the data
 * @param {Array} data - 2D array of spreadsheet data
 * @param {number} recipientColIdx - Index of the recipient column
 * @return {Object} Object containing duplicateRows array and duplicateEmails set
 */
function findDuplicateEmails(data, recipientColIdx) {
  const emailsSeen = new Map(); // email -> first row index
  const duplicateRows = [];
  const duplicateEmails = new Set();

  // Skip header row (index 0)
  for (let i = 1; i < data.length; i++) {
    const email = data[i][recipientColIdx];

    // Skip empty or invalid email cells
    if (!email || typeof email !== 'string' || email.trim() === '') {
      continue;
    }

    const normalizedEmail = email.trim().toLowerCase();

    if (emailsSeen.has(normalizedEmail)) {
      // This is a duplicate
      duplicateRows.push(i);
      duplicateEmails.add(normalizedEmail);
    } else {
      // First occurrence of this email
      emailsSeen.set(normalizedEmail, i);
    }
  }

  return {
    duplicateRows: duplicateRows,
    duplicateEmails: duplicateEmails
  };
}

/**
 * Sends emails from sheet data with rate limiting and automatic resumption.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
 * @param {string} selectedBatch (optional) to filter emails by batch
 * @param {Object} resumeState (optional) state for resuming interrupted jobs
*/
function sendEmails(subjectLine, sheet=SpreadsheetApp.getActiveSheet(), selectedBatch=null, resumeState=null) {
  try {
    // Check if there's already an active job
    const existingJob = getActiveJobState();
    if (existingJob && !resumeState) {
      SpreadsheetApp.getUi().alert('Job Already Active',
        `There is already an active email job for batch "${existingJob.batchId}". ` +
        `Please wait for it to complete or cancel it first.`,
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // option to skip browser prompt if you want to use this code in other projects
    if (!subjectLine){
      subjectLine = Browser.inputBox("Mail Merge",
                                        "Type or copy/paste the subject line of the Gmail " +
                                        "draft message you would like to mail merge with:",
                                        Browser.Buttons.OK_CANCEL);

      if (subjectLine === "cancel" || subjectLine == ""){
      // If no subject line, finishes up
      return;
      }
    }

    // Gets the draft Gmail message to use as a template
    // Note: Even when using MailerSend, we still use Gmail drafts for the email template
    const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  
    // Initialize or resume job state
    let jobState;
    if (resumeState) {
      jobState = resumeState;
      jobState.resumeCount++;
    } else {
      // Count total emails to send
      const totalEmails = countPendingEmails(sheet, selectedBatch);

      if (totalEmails === 0) {
        SpreadsheetApp.getUi().alert('No Emails to Send',
          `No unsent emails found in batch "${selectedBatch}".`,
          SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }

      // Create new job state
      jobState = {
        batchId: selectedBatch,
        subjectLine: subjectLine,
        totalEmails: totalEmails,
        emailsSent: 0,
        startTime: new Date().toISOString(),
        resumeCount: 0,
        isABTest: false
      };
    }

    // Save job state
    saveJobState(jobState);

    // Gets the data from the passed sheet
    const dataRange = sheet.getDataRange();
    const data = dataRange.getDisplayValues();
    const heads = data.shift();
    const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

    // Convert to object array for easier processing
    const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

    // Track emails sent in this session
    let emailsSentThisSession = 0;
    const out = [];

    // Process emails with rate limiting
    for (let i = 0; i < obj.length; i++) {
      const row = obj[i];
      const batchMatches = !selectedBatch || row[BATCH_COL] === selectedBatch;

      if (row[EMAIL_SENT_COL] == '' && batchMatches) {
        // Check quotas (only for Gmail, MailerSend has much higher limits)
        const emailConfig = getEmailServiceConfig();
        const quotaCheck = checkEmailQuotas(emailsSentThisSession, emailConfig);

        if (!quotaCheck.canSend) {
          console.log(`Quota limit reached. Reason: ${quotaCheck.reason}`);

          // Update job state for resume
          jobState.emailsSent += emailsSentThisSession;

          let alertTitle, alertMessage, shouldScheduleResume = true;

          if (quotaCheck.reason === 'daily_quota_exhausted') {
            // Daily quota exhausted - schedule resume for tomorrow
            const tomorrow = new Date();
            tomorrow.setDate(tomorrow.getDate() + 1);
            tomorrow.setHours(0, 30, 0, 0); // Resume at 12:30 AM tomorrow

            jobState.nextResumeTime = tomorrow.toISOString();
            alertTitle = 'Daily Email Quota Exhausted';
            alertMessage = `Sent ${emailsSentThisSession} emails this session.\n` +
              `Daily quota remaining: ${quotaCheck.dailyQuotaRemaining}\n` +
              `Total progress: ${jobState.emailsSent + emailsSentThisSession} of ${jobState.totalEmails} emails.\n\n` +
              `The job will automatically resume tomorrow at 12:30 AM.`;
          } else {
            // Hourly rate limit - schedule resume in 1 hour
            jobState.nextResumeTime = new Date(Date.now() + RESUME_DELAY_HOURS * 60 * 60 * 1000).toISOString();
            alertTitle = 'Hourly Rate Limit Reached';
            alertMessage = `Sent ${emailsSentThisSession} emails this session.\n` +
              `Daily quota remaining: ${quotaCheck.dailyQuotaRemaining}\n` +
              `Total progress: ${jobState.emailsSent + emailsSentThisSession} of ${jobState.totalEmails} emails.\n\n` +
              `The job will automatically resume in ${RESUME_DELAY_HOURS} hour(s).`;
          }

          if (shouldScheduleResume) {
            // Create resume trigger
            const triggerId = createResumeTrigger();
            jobState.triggerId = triggerId;
          }

          saveJobState(jobState);

          // Update spreadsheet with current progress
          if (out.length > 0) {
            sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
          }

          // Show progress message
          const remaining = jobState.totalEmails - jobState.emailsSent - emailsSentThisSession;
          SpreadsheetApp.getUi().alert(alertTitle,
            alertMessage + `\nRemaining: ${remaining} emails.`,
            SpreadsheetApp.getUi().ButtonSet.OK);

          return;
        }

        try {
          // Get email service configuration
          const emailConfig = getEmailServiceConfig();

          if (emailConfig.service === 'mailersend') {
            // Use MailerSend
            const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

            const result = sendEmailWithMailerSend(
              row[RECIPIENT_COL],
              row['Name'] || row[RECIPIENT_COL], // Use Name column if available, fallback to email
              msgObj.subject,
              msgObj.text,
              msgObj.html,
              emailConfig
            );

            if (result.success) {
              out.push([new Date()]);
              emailsSentThisSession++;
            } else {
              throw new Error(`MailerSend error: ${result.error}`);
            }
          } else {
            // Use Gmail (default)
            const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

            GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
              htmlBody: msgObj.html,
              attachments: emailTemplate.attachments,
              inlineImages: emailTemplate.inlineImages
            });

            out.push([new Date()]);
            emailsSentThisSession++;
          }

        } catch(e) {
          console.error('Error sending email:', e.message);
          out.push([e.message]);

          // Stop the batch on error as requested
          clearJobState();

          // Update spreadsheet with current progress
          if (out.length > 0) {
            sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
          }

          SpreadsheetApp.getUi().alert('Email Send Error',
            `An error occurred while sending emails: ${e.message}\n\n` +
            `The batch has been stopped. ${emailsSentThisSession} emails were sent successfully.`,
            SpreadsheetApp.getUi().ButtonSet.OK);
          return;
        }
      } else {
        out.push([row[EMAIL_SENT_COL]]);
      }
    }

    // Update the sheet with results
    sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);

    // Job completed successfully
    clearJobState();

    SpreadsheetApp.getUi().alert('Emails Sent Successfully',
      `All ${emailsSentThisSession} emails have been sent successfully for batch "${selectedBatch}".`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('Error in sendEmails:', error);
    SpreadsheetApp.getUi().alert('Email Send Error',
      `An error occurred while sending emails: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK);
    throw error; // Re-throw so the calling function can handle it
  }
}

/**
 * Get a Gmail draft message by matching the subject line.
 * @param {string} subject_line to search for draft message
 * @return {object} containing the subject, plain and html message body and attachments
 */
function getGmailTemplateFromDrafts_(subject_line) {
  try {
    // get drafts
    const drafts = GmailApp.getDrafts();
    // filter the drafts that match subject line
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    // get the message object
    const msg = draft.getMessage();

    // Handles inline images and attachments so they can be included in the merge
    // Based on https://stackoverflow.com/a/65813881/1027723
    // Gets all attachments and inline image attachments
    const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
    const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
    const htmlBody = msg.getBody();

    // Creates an inline image object with the image name as key
    // (can't rely on image index as array based on insert order)
    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

    //Regexp searches for all img string positions with cid
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];

    //Initiates the allInlineImages object
    const inlineImagesObj = {};
    // built an inlineImagesObj from inline image matches
    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

    return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody},
            attachments: attachments, inlineImages: inlineImagesObj };
  } catch(e) {
    throw new Error("Oops - can't find Gmail draft");
  }

  /**
   * Filter draft objects with the matching subject linemessage by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} GmailDraft object
   */
  function subjectFilter_(subject_line) {
    return function(element) {
      if (element.getMessage().getSubject() === subject_line) {
        return element;
      }
    }
  }
}

/**
 * Fill template string with data object
 * @see https://stackoverflow.com/a/378000/1027723
 * @param {string} template string containing {{}} markers which are replaced with data
 * @param {object} data object used to replace {{}} markers
 * @return {object} message replaced with data
 */
function fillInTemplateFromObject_(template, data) {
  // We have two templates one for plain text and the html body
  // Stringifing the object means we can do a global replace
  let template_string = JSON.stringify(template);

  // Token replacement
  template_string = template_string.replace(/{{[^{}]+}}/g, key => {
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });
  return  JSON.parse(template_string);
}

/**
 * Escape cell data to make JSON safe
 * @see https://stackoverflow.com/a/9204218/1027723
 * @param {string} str to escape JSON special characters from
 * @return {string} escaped string
 */
function escapeData_(str) {
  return str
    .replace(/[\\]/g, '\\\\')
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
}

/**
 * Assigns A/B test versions to emails in a batch using deterministic random assignment
 * @param {Array} emailData - Array of email objects for the batch
 * @param {string} batchId - The batch identifier for consistent seeding
 * @return {Array} Array of email objects with version assignments
 */
function assignABVersions(emailData, batchId) {
  // Create a simple hash function for deterministic randomness
  function simpleHash(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
      const char = str.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash; // Convert to 32-bit integer
    }
    return Math.abs(hash);
  }

  // Assign versions based on hash of batch + email
  return emailData.map((emailRow, index) => {
    const email = emailRow[RECIPIENT_COL] || '';
    const seed = `${batchId}_${email}_${index}`;
    const hash = simpleHash(seed);
    const version = hash % 2 === 0 ? 'A' : 'B';

    return {
      ...emailRow,
      [VERSION_COL]: version
    };
  });
}

/**
 * Sends A/B test emails from sheet data with rate limiting and automatic resumption
 * @param {string} subjectA - Subject line for version A
 * @param {string} subjectB - Subject line for version B
 * @param {Sheet} sheet - Sheet to read data from
 * @param {string} selectedBatch - Batch to filter emails by
 * @param {Object} resumeState - Optional resume state for interrupted jobs
 */
function sendABTestEmails(subjectA, subjectB, sheet = SpreadsheetApp.getActiveSheet(), selectedBatch, resumeState = null) {
  try {
    // Check if there's already an active job
    const existingJob = getActiveJobState();
    if (existingJob && !resumeState) {
      SpreadsheetApp.getUi().alert('Job Already Active',
        `There is already an active email job for batch "${existingJob.batchId}". ` +
        `Please wait for it to complete or cancel it first.`,
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Get email templates for both versions
    // Note: Even when using MailerSend, we still use Gmail drafts for the email templates
    const emailTemplateA = getGmailTemplateFromDrafts_(subjectA);
    const emailTemplateB = getGmailTemplateFromDrafts_(subjectB);

    // Initialize or resume job state
    let jobState;
    if (resumeState) {
      jobState = resumeState;
      jobState.resumeCount++;
    } else {
      // Count total emails to send
      const totalEmails = countPendingEmails(sheet, selectedBatch);

      if (totalEmails === 0) {
        SpreadsheetApp.getUi().alert('No Emails to Send',
          `No unsent emails found in batch "${selectedBatch}".`,
          SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }

      // Confirm with user for new jobs
      const confirmMessage = `Ready to send A/B test emails to batch "${selectedBatch}":\n\n` +
        `Subject A: "${subjectA}"\n` +
        `Subject B: "${subjectB}"\n\n` +
        `Total: ${totalEmails} emails\n\n` +
        `Do you want to proceed?`;

      const response = SpreadsheetApp.getUi().alert('Confirm A/B Test Send',
        confirmMessage, SpreadsheetApp.getUi().ButtonSet.YES_NO);

      if (response !== SpreadsheetApp.getUi().Button.YES) {
        return;
      }

      // Create new job state
      jobState = {
        batchId: selectedBatch,
        subjectA: subjectA,
        subjectB: subjectB,
        totalEmails: totalEmails,
        emailsSent: 0,
        startTime: new Date().toISOString(),
        resumeCount: 0,
        isABTest: true
      };
    }

    // Save job state
    saveJobState(jobState);

    // Gets the data from the passed sheet
    const dataRange = sheet.getDataRange();
    const data = dataRange.getDisplayValues();
    const heads = data.shift();

    // Get column indices
    const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
    const versionColIdx = heads.indexOf(VERSION_COL);

    // Check if version column exists
    if (versionColIdx === -1) {
      clearJobState();
      SpreadsheetApp.getUi().alert('Version Column Missing',
        `Please add a "${VERSION_COL}" column to your spreadsheet for A/B testing.`,
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Convert to object array
    const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

    // Filter for the selected batch and unsent emails
    const batchEmails = obj.filter((row) => {
      const batchMatches = row[BATCH_COL] === selectedBatch;
      const notSent = row[EMAIL_SENT_COL] === '';
      return batchMatches && notSent;
    });

    // Assign A/B versions if not already assigned
    const emailsWithVersions = assignABVersions(batchEmails, selectedBatch);

    // Track emails sent in this session
    let emailsSentThisSession = 0;
    const emailSentUpdates = [];
    const versionUpdates = [];

    // Process emails with rate limiting
    for (let i = 0; i < obj.length; i++) {
      const row = obj[i];
      const batchMatches = row[BATCH_COL] === selectedBatch;
      const notSent = row[EMAIL_SENT_COL] === '';

      if (batchMatches && notSent) {
        // Check quotas (only for Gmail, MailerSend has much higher limits)
        const emailConfig = getEmailServiceConfig();
        const quotaCheck = checkEmailQuotas(emailsSentThisSession, emailConfig);

        if (!quotaCheck.canSend) {
          console.log(`Quota limit reached. Reason: ${quotaCheck.reason}`);

          // Update job state for resume
          jobState.emailsSent += emailsSentThisSession;

          let alertTitle, alertMessage, shouldScheduleResume = true;

          if (quotaCheck.reason === 'daily_quota_exhausted') {
            // Daily quota exhausted - schedule resume for tomorrow
            const tomorrow = new Date();
            tomorrow.setDate(tomorrow.getDate() + 1);
            tomorrow.setHours(0, 30, 0, 0); // Resume at 12:30 AM tomorrow

            jobState.nextResumeTime = tomorrow.toISOString();
            alertTitle = 'Daily Email Quota Exhausted';
            alertMessage = `Sent ${emailsSentThisSession} emails this session.\n` +
              `Daily quota remaining: ${quotaCheck.dailyQuotaRemaining}\n` +
              `Total progress: ${jobState.emailsSent + emailsSentThisSession} of ${jobState.totalEmails} emails.\n\n` +
              `The A/B test will automatically resume tomorrow at 12:30 AM.`;
          } else {
            // Hourly rate limit - schedule resume in 1 hour
            jobState.nextResumeTime = new Date(Date.now() + RESUME_DELAY_HOURS * 60 * 60 * 1000).toISOString();
            alertTitle = 'Hourly Rate Limit Reached';
            alertMessage = `Sent ${emailsSentThisSession} emails this session.\n` +
              `Daily quota remaining: ${quotaCheck.dailyQuotaRemaining}\n` +
              `Total progress: ${jobState.emailsSent + emailsSentThisSession} of ${jobState.totalEmails} emails.\n\n` +
              `The A/B test will automatically resume in ${RESUME_DELAY_HOURS} hour(s).`;
          }

          if (shouldScheduleResume) {
            // Create resume trigger
            const triggerId = createResumeTrigger();
            jobState.triggerId = triggerId;
          }

          saveJobState(jobState);

          // Update spreadsheet with current progress
          if (emailSentUpdates.length > 0) {
            sheet.getRange(2, emailSentColIdx + 1, emailSentUpdates.length).setValues(emailSentUpdates);
            sheet.getRange(2, versionColIdx + 1, versionUpdates.length).setValues(versionUpdates);
          }

          // Show progress message
          const remaining = jobState.totalEmails - jobState.emailsSent - emailsSentThisSession;
          SpreadsheetApp.getUi().alert(alertTitle,
            alertMessage + `\nRemaining: ${remaining} emails.`,
            SpreadsheetApp.getUi().ButtonSet.OK);

          return;
        }

        // Find the version assignment for this email
        const emailWithVersion = emailsWithVersions.find(e => e[RECIPIENT_COL] === row[RECIPIENT_COL]);
        const version = emailWithVersion ? emailWithVersion[VERSION_COL] : 'A';

        // Choose the appropriate template and subject
        const emailTemplate = version === 'A' ? emailTemplateA : emailTemplateB;
        const currentSubject = version === 'A' ? subjectA : subjectB;

        try {
          // Create custom template with the correct subject
          const customTemplate = {
            ...emailTemplate,
            message: {
              ...emailTemplate.message,
              subject: currentSubject
            }
          };

          const msgObj = fillInTemplateFromObject_(customTemplate.message, row);

          // Get email service configuration
          const emailConfig = getEmailServiceConfig();

          if (emailConfig.service === 'mailersend') {
            // Use MailerSend
            const result = sendEmailWithMailerSend(
              row[RECIPIENT_COL],
              row['Name'] || row[RECIPIENT_COL], // Use Name column if available, fallback to email
              msgObj.subject,
              msgObj.text,
              msgObj.html,
              emailConfig
            );

            if (!result.success) {
              throw new Error(`MailerSend error: ${result.error}`);
            }
          } else {
            // Use Gmail (default)
            GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
              htmlBody: msgObj.html,
              attachments: emailTemplate.attachments,
              inlineImages: emailTemplate.inlineImages
            });
          }

          // Record success
          emailSentUpdates.push([new Date()]);
          versionUpdates.push([version]);
          emailsSentThisSession++;

        } catch(e) {
          console.error('Error sending A/B test email:', e.message);
          emailSentUpdates.push([e.message]);
          versionUpdates.push([version]);

          // Stop the batch on error as requested
          clearJobState();

          // Update spreadsheet with current progress
          if (emailSentUpdates.length > 0) {
            sheet.getRange(2, emailSentColIdx + 1, emailSentUpdates.length).setValues(emailSentUpdates);
            sheet.getRange(2, versionColIdx + 1, versionUpdates.length).setValues(versionUpdates);
          }

          SpreadsheetApp.getUi().alert('A/B Test Error',
            `An error occurred during A/B testing: ${e.message}\n\n` +
            `The batch has been stopped. ${emailsSentThisSession} emails were sent successfully.`,
            SpreadsheetApp.getUi().ButtonSet.OK);
          return;
        }
      } else {
        // Keep existing values for non-matching rows
        emailSentUpdates.push([row[EMAIL_SENT_COL]]);
        versionUpdates.push([row[VERSION_COL]]);
      }
    }

    // Update the sheet with results
    sheet.getRange(2, emailSentColIdx + 1, emailSentUpdates.length).setValues(emailSentUpdates);
    sheet.getRange(2, versionColIdx + 1, versionUpdates.length).setValues(versionUpdates);

    // Job completed successfully
    clearJobState();

    // Count final versions for reporting
    const versionCounts = { A: 0, B: 0 };
    emailsWithVersions.forEach(email => {
      versionCounts[email[VERSION_COL]]++;
    });

    SpreadsheetApp.getUi().alert('A/B Test Complete',
      `Successfully sent ${emailsSentThisSession} emails for batch "${selectedBatch}":\n` +
      `Version A: ${versionCounts.A} emails\n` +
      `Version B: ${versionCounts.B} emails`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('Error in sendABTestEmails:', error);
    SpreadsheetApp.getUi().alert('A/B Test Error',
      `An error occurred during A/B testing: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK);
    throw error; // Re-throw so the calling function can handle it
  }
}

/**
 * Test function to debug form data processing
 * This can be called directly from the Apps Script editor to test
 */
function testProcessBatchSelection() {
  // Test with regular email data - UPDATE THIS SUBJECT LINE TO MATCH YOUR GMAIL DRAFT
  const testFormData = {
    selectedBatch: "test",
    enableABTesting: false,
    subjectLine: "quick question {{First name}}"  // Change this to match your actual Gmail draft subject
  };

  console.log('Testing processBatchSelection with:', testFormData);
  processBatchSelection(testFormData);
}

/**
 * Test function for rate limiting - creates a small test batch
 * This can be used to test the rate limiting functionality without sending many emails
 */
function testRateLimiting() {
  // You can modify EMAIL_RATE_LIMIT temporarily for testing
  // For example: EMAIL_RATE_LIMIT = 2; // Test with only 2 emails per batch

  const testFormData = {
    selectedBatch: "test",
    enableABTesting: false,
    subjectLine: "Test Subject"  // Make sure you have a Gmail draft with this subject
  };

  console.log('Testing rate limiting functionality');
  processBatchSelection(testFormData);
}

/**
 * Utility function to clear any stuck job states during testing
 */
function clearTestJobState() {
  clearJobState();
  console.log('Test job state cleared');
}

/**
 * Test function to check current email quotas
 * This helps understand your current Gmail sending limits
 */
function testEmailQuotas() {
  try {
    const emailConfig = getEmailServiceConfig();
    const dailyQuotaRemaining = emailConfig.service === 'gmail' ? MailApp.getRemainingDailyQuota() : 'unlimited';
    const quotaCheck = checkEmailQuotas(0, emailConfig); // Check with 0 emails sent this session

    console.log('=== Email Quota Status ===');
    console.log(`Email service: ${emailConfig.service.toUpperCase()}`);
    console.log(`Daily quota remaining: ${dailyQuotaRemaining}`);
    console.log(`Can send emails: ${quotaCheck.canSend}`);
    console.log(`Max emails this session: ${quotaCheck.maxEmails}`);
    console.log(`Quota check reason: ${quotaCheck.reason}`);

    // Show in UI as well
    const serviceName = emailConfig.service === 'gmail' ? 'Gmail' : 'MailerSend';
    const message = `Email Service: ${serviceName}\n\n` +
      `Daily Quota Remaining: ${dailyQuotaRemaining}\n` +
      `Can Send Emails: ${quotaCheck.canSend ? 'Yes' : 'No'}\n` +
      `Max Emails This Session: ${quotaCheck.maxEmails}\n` +
      `Status: ${quotaCheck.reason}`;

    SpreadsheetApp.getUi().alert('Email Quota Check', message, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    console.error('Error checking email quotas:', error);
    SpreadsheetApp.getUi().alert('Quota Check Error',
      `Unable to check email quotas: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Helper function to list all Gmail draft subject lines
 * Run this to see what drafts you have available
 */
function listGmailDrafts() {
  try {
    const drafts = GmailApp.getDrafts();
    console.log(`Found ${drafts.length} Gmail drafts:`);

    drafts.forEach((draft, index) => {
      const subject = draft.getMessage().getSubject();
      console.log(`${index + 1}. "${subject}"`);
    });

    if (drafts.length === 0) {
      console.log('No Gmail drafts found. Please create some drafts first.');
    }
  } catch (error) {
    console.error('Error listing drafts:', error);
  }
}

// ============================================================================
// MAILERSEND EMAIL FUNCTIONS
// ============================================================================

/**
 * Sends an email using MailerSend API
 * @param {string} toEmail - Recipient email address
 * @param {string} toName - Recipient name
 * @param {string} subject - Email subject
 * @param {string} textContent - Plain text content
 * @param {string} htmlContent - HTML content
 * @param {Object} config - MailerSend configuration
 * @return {Object} Response object with success status and message ID
 */
function sendEmailWithMailerSend(toEmail, toName, subject, textContent, htmlContent, config) {
  const url = 'https://api.mailersend.com/v1/email';

  const payload = {
    from: {
      email: config.settings.fromEmail,
      name: config.settings.fromName
    },
    to: [
      {
        email: toEmail,
        name: toName || toEmail
      }
    ],
    subject: subject,
    text: textContent,
    html: htmlContent
  };

  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${config.settings.apiToken}`,
      'Content-Type': 'application/json',
      'X-Requested-With': 'XMLHttpRequest'
    },
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 202) {
      // Success - email queued for sending
      const messageId = response.getHeaders()['x-message-id'] || 'unknown';
      return {
        success: true,
        messageId: messageId
      };
    } else {
      // Error response
      const responseText = response.getContentText();
      console.error('MailerSend API error:', responseCode, responseText);

      let errorMessage = `HTTP ${responseCode}`;
      try {
        const errorData = JSON.parse(responseText);
        if (errorData.message) {
          errorMessage = errorData.message;
        }
        if (errorData.errors) {
          errorMessage += ': ' + JSON.stringify(errorData.errors);
        }
      } catch (e) {
        errorMessage += ': ' + responseText;
      }

      return {
        success: false,
        error: errorMessage
      };
    }
  } catch (error) {
    console.error('Error calling MailerSend API:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// ============================================================================
// EMAIL SERVICE CONFIGURATION FUNCTIONS
// ============================================================================

/**
 * Gets the current email service configuration
 * @return {Object} Configuration object with service type and settings
 */
function getEmailServiceConfig() {
  const properties = PropertiesService.getScriptProperties();
  const configJson = properties.getProperty('EMAIL_SERVICE_CONFIG');

  if (!configJson) {
    // Default to Gmail
    return {
      service: 'gmail',
      settings: {}
    };
  }

  try {
    return JSON.parse(configJson);
  } catch (error) {
    console.error('Error parsing email service config:', error);
    return {
      service: 'gmail',
      settings: {}
    };
  }
}

/**
 * Saves the email service configuration
 * @param {Object} config - Configuration object
 */
function saveEmailServiceConfig(config) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('EMAIL_SERVICE_CONFIG', JSON.stringify(config));
}

/**
 * Processes the email service settings form data
 * @param {Object} formData - Form data from the settings dialog
 */
function processEmailServiceSettings(formData) {
  try {
    console.log('Email service settings received:', formData);

    const config = {
      service: formData.emailService,
      settings: {}
    };

    if (formData.emailService === 'mailersend') {
      if (!formData.mailerSendToken) {
        SpreadsheetApp.getUi().alert('Missing Information',
          'Please enter your MailerSend API token.',
          SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }

      if (!formData.fromEmail) {
        SpreadsheetApp.getUi().alert('Missing Information',
          'Please enter your verified sender email address.',
          SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }

      config.settings = {
        apiToken: formData.mailerSendToken,
        fromEmail: formData.fromEmail,
        fromName: formData.fromName || 'Mail Merge'
      };
    }

    saveEmailServiceConfig(config);

    SpreadsheetApp.getUi().alert('Settings Saved',
      `Email service has been set to ${config.service === 'gmail' ? 'Gmail' : 'MailerSend'}.`,
      SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    console.error('Error processing email service settings:', error);
    SpreadsheetApp.getUi().alert('Settings Error',
      `An error occurred while saving settings: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Gets the current email service configuration for the settings dialog
 * @return {Object} Current configuration
 */
function getCurrentEmailServiceConfig() {
  return getEmailServiceConfig();
}

/**
 * Test function to verify MailerSend configuration
 * This sends a test email to verify the MailerSend setup is working
 */
function testMailerSendConfiguration() {
  const config = getEmailServiceConfig();

  if (config.service !== 'mailersend') {
    SpreadsheetApp.getUi().alert('MailerSend Not Configured',
      'MailerSend is not currently selected as the email service. Please configure it in Email Service Settings first.',
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Get user's email for test
  const userEmail = Session.getActiveUser().getEmail();

  if (!userEmail) {
    SpreadsheetApp.getUi().alert('Test Email Error',
      'Could not determine your email address for the test. Please ensure you are logged in.',
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const result = sendEmailWithMailerSend(
    userEmail,
    'Test User',
    'MailerSend Configuration Test',
    'This is a test email to verify your MailerSend configuration is working correctly.',
    '<p>This is a test email to verify your <strong>MailerSend configuration</strong> is working correctly.</p>',
    config
  );

  if (result.success) {
    SpreadsheetApp.getUi().alert('MailerSend Test Successful',
      `Test email sent successfully!\n\nMessage ID: ${result.messageId}\n\nCheck your inbox at ${userEmail} to confirm delivery.`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('MailerSend Test Failed',
      `Test email failed to send.\n\nError: ${result.error}\n\nPlease check your MailerSend configuration and try again.`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ============================================================================
// RATE LIMITING AND JOB MANAGEMENT FUNCTIONS
// ============================================================================

/**
 * Gets the current active job state from script properties
 * @return {Object|null} Job state object or null if no active job
 */
function getActiveJobState() {
  const properties = PropertiesService.getScriptProperties();
  const jobStateJson = properties.getProperty('ACTIVE_EMAIL_JOB');

  if (!jobStateJson) {
    return null;
  }

  try {
    return JSON.parse(jobStateJson);
  } catch (error) {
    console.error('Error parsing job state:', error);
    return null;
  }
}

/**
 * Saves the current job state to script properties
 * @param {Object} jobState - Job state object to save
 */
function saveJobState(jobState) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('ACTIVE_EMAIL_JOB', JSON.stringify(jobState));
}

/**
 * Clears the active job state and any associated triggers
 */
function clearJobState() {
  const properties = PropertiesService.getScriptProperties();
  const jobState = getActiveJobState();

  // Clean up any triggers
  if (jobState && jobState.triggerId) {
    try {
      const triggers = ScriptApp.getProjectTriggers();
      const trigger = triggers.find(t => t.getUniqueId() === jobState.triggerId);
      if (trigger) {
        ScriptApp.deleteTrigger(trigger);
      }
    } catch (error) {
      console.error('Error cleaning up trigger:', error);
    }
  }

  properties.deleteProperty('ACTIVE_EMAIL_JOB');
}

/**
 * Creates a time-based trigger to resume email sending after the rate limit period
 * @return {string} Trigger ID
 */
function createResumeTrigger() {
  const trigger = ScriptApp.newTrigger('resumeEmailSending')
    .timeBased()
    .after(RESUME_DELAY_HOURS * 60 * 60 * 1000) // Convert hours to milliseconds
    .create();

  return trigger.getUniqueId();
}

/**
 * Shows the current job status to the user
 */
function showJobStatus() {
  const jobState = getActiveJobState();
  const ui = SpreadsheetApp.getUi();

  if (!jobState) {
    ui.alert('No Active Job', 'There is no active email sending job.', ui.ButtonSet.OK);
    return;
  }

  const progress = `${jobState.emailsSent} of ${jobState.totalEmails}`;
  const percentage = Math.round((jobState.emailsSent / jobState.totalEmails) * 100);
  const nextResumeTime = jobState.nextResumeTime ? new Date(jobState.nextResumeTime).toLocaleString() : 'Unknown';

  // Get current daily quota
  let dailyQuotaInfo = '';
  try {
    const dailyQuotaRemaining = MailApp.getRemainingDailyQuota();
    dailyQuotaInfo = `\nDaily Quota Remaining: ${dailyQuotaRemaining}`;
  } catch (error) {
    dailyQuotaInfo = '\nDaily Quota: Unable to check';
  }

  const message = `Active Email Job Status:\n\n` +
    `Batch: ${jobState.batchId}\n` +
    `Type: ${jobState.isABTest ? 'A/B Test' : 'Regular Email'}\n` +
    `Progress: ${progress} emails (${percentage}%)\n` +
    `Started: ${new Date(jobState.startTime).toLocaleString()}\n` +
    `Next Resume: ${nextResumeTime}\n` +
    `Resume Attempts: ${jobState.resumeCount}/${MAX_RESUME_ATTEMPTS}` +
    dailyQuotaInfo;

  ui.alert('Job Status', message, ui.ButtonSet.OK);
}

/**
 * Cancels the active email job
 */
function cancelActiveJob() {
  const jobState = getActiveJobState();
  const ui = SpreadsheetApp.getUi();

  if (!jobState) {
    ui.alert('No Active Job', 'There is no active email sending job to cancel.', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert('Cancel Job',
    `Are you sure you want to cancel the active email job for batch "${jobState.batchId}"?\n\n` +
    `Progress: ${jobState.emailsSent} of ${jobState.totalEmails} emails sent.`,
    ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    clearJobState();
    ui.alert('Job Cancelled', 'The active email job has been cancelled.', ui.ButtonSet.OK);
  }
}

/**
 * Main function called by time-based trigger to resume email sending
 */
function resumeEmailSending() {
  try {
    const jobState = getActiveJobState();

    if (!jobState) {
      console.log('No active job found to resume');
      return;
    }

    // Check if we've exceeded maximum resume attempts
    if (jobState.resumeCount >= MAX_RESUME_ATTEMPTS) {
      console.error('Maximum resume attempts exceeded for job:', jobState.batchId);
      clearJobState();
      return;
    }

    console.log(`Resuming email job for batch: ${jobState.batchId}, attempt ${jobState.resumeCount + 1}`);

    // Resume the appropriate type of email sending
    if (jobState.isABTest) {
      resumeABTestEmails(jobState);
    } else {
      resumeRegularEmails(jobState);
    }

  } catch (error) {
    console.error('Error resuming email sending:', error);

    // Try to update job state with error info
    const jobState = getActiveJobState();
    if (jobState) {
      jobState.lastError = error.message;
      jobState.lastErrorTime = new Date().toISOString();
      saveJobState(jobState);
    }
  }
}

/**
 * Counts pending emails for a batch
 * @param {Sheet} sheet - The spreadsheet sheet
 * @param {string} batchId - The batch identifier
 * @return {number} Number of pending emails
 */
function countPendingEmails(sheet, batchId) {
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  const heads = data.shift();

  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  const batchColIdx = heads.indexOf(BATCH_COL);

  if (emailSentColIdx === -1 || batchColIdx === -1) {
    throw new Error('Required columns not found');
  }

  let pendingCount = 0;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const batchMatches = row[batchColIdx] === batchId;
    const notSent = row[emailSentColIdx] === '';

    if (batchMatches && notSent) {
      pendingCount++;
    }
  }

  return pendingCount;
}

/**
 * Checks both daily and hourly email quotas to determine how many emails can be sent
 * @param {number} emailsSentThisSession - Number of emails already sent in current session
 * @param {Object} emailConfig - Email service configuration
 * @return {Object} Object with canSend (boolean), reason (string), and maxEmails (number)
 */
function checkEmailQuotas(emailsSentThisSession, emailConfig = null) {
  // Get email config if not provided
  if (!emailConfig) {
    emailConfig = getEmailServiceConfig();
  }

  // MailerSend has much higher limits, so we can bypass most restrictions
  if (emailConfig.service === 'mailersend') {
    return {
      canSend: true,
      reason: 'mailersend_unlimited',
      maxEmails: 999999, // Effectively unlimited for our purposes
      dailyQuotaRemaining: 'unlimited'
    };
  }

  // Gmail quota checking (existing logic)
  try {
    // Get remaining daily quota from Gmail
    const dailyQuotaRemaining = MailApp.getRemainingDailyQuota();

    console.log(`Daily quota remaining: ${dailyQuotaRemaining}`);
    console.log(`Emails sent this session: ${emailsSentThisSession}`);

    // Check if we have enough daily quota (keeping a buffer)
    if (dailyQuotaRemaining <= MIN_DAILY_QUOTA_BUFFER) {
      return {
        canSend: false,
        reason: 'daily_quota_exhausted',
        maxEmails: 0,
        dailyQuotaRemaining: dailyQuotaRemaining
      };
    }

    // Check hourly rate limit
    if (emailsSentThisSession >= EMAIL_RATE_LIMIT) {
      return {
        canSend: false,
        reason: 'hourly_rate_limit',
        maxEmails: 0,
        dailyQuotaRemaining: dailyQuotaRemaining
      };
    }

    // Calculate how many more emails we can send this session
    const remainingHourlyQuota = EMAIL_RATE_LIMIT - emailsSentThisSession;
    const remainingDailyQuota = dailyQuotaRemaining - MIN_DAILY_QUOTA_BUFFER;
    const maxEmails = Math.min(remainingHourlyQuota, remainingDailyQuota);

    return {
      canSend: maxEmails > 0,
      reason: maxEmails > 0 ? 'can_send' : 'quota_exhausted',
      maxEmails: maxEmails,
      dailyQuotaRemaining: dailyQuotaRemaining
    };

  } catch (error) {
    console.error('Error checking email quotas:', error);
    // Fallback to hourly limit only if daily quota check fails
    const canSend = emailsSentThisSession < EMAIL_RATE_LIMIT;
    return {
      canSend: canSend,
      reason: canSend ? 'can_send_fallback' : 'hourly_rate_limit',
      maxEmails: canSend ? EMAIL_RATE_LIMIT - emailsSentThisSession : 0,
      dailyQuotaRemaining: 'unknown'
    };
  }
}

/**
 * Resumes regular email sending from saved state
 * @param {Object} jobState - Saved job state
 */
function resumeRegularEmails(jobState) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();

    console.log(`Resuming regular emails for batch: ${jobState.batchId}`);

    // Call sendEmails with resume state
    sendEmails(jobState.subjectLine, sheet, jobState.batchId, jobState);

  } catch (error) {
    console.error('Error resuming regular emails:', error);
    throw error;
  }
}

/**
 * Resumes A/B test email sending from saved state
 * @param {Object} jobState - Saved job state
 */
function resumeABTestEmails(jobState) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();

    console.log(`Resuming A/B test emails for batch: ${jobState.batchId}`);

    // Call sendABTestEmails with resume state
    sendABTestEmails(jobState.subjectA, jobState.subjectB, sheet, jobState.batchId, jobState);

  } catch (error) {
    console.error('Error resuming A/B test emails:', error);
    throw error;
  }
}
