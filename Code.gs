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

// MailerSend rate limiting constants
const MAILERSEND_REQUESTS_PER_MINUTE = 8; // Conservative limit (your account allows 10/min)
const MAILERSEND_DELAY_MS = 8000; // 8 seconds between requests (60000ms / 8 requests = 7.5s, rounded up)
 
/**
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('📧 Send Emails', 'showBatchSelectionDialog')
      .addItem('📊 Job Progress & Status', 'showProgressDialog')
      .addSeparator()
      .addSubMenu(ui.createMenu('⚙️ Settings & Tools')
          .addItem('Email Service Settings', 'showEmailServiceSettings')
          .addItem('Test MailerSend Configuration', 'testMailerSendConfiguration')
          .addItem('Check Email Quotas', 'testEmailQuotas')
          .addItem('🔍 Debug Job Status', 'debugJobStatus')
          .addItem('▶️ Manually Resume Job', 'manuallyResumeJob')
          .addItem('🔄 Reset MailerSend Rate Limit', 'resetMailerSendRateLimit')
          .addSeparator()
          .addItem('Remove Duplicate Emails', 'removeDuplicateEmails')
          .addItem('🧹 Cleanup All Triggers', 'cleanupAllTriggers'))
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
 * Shows the progress dialog for active email jobs
 */
function showProgressDialog() {
  const htmlTemplate = HtmlService.createTemplateFromFile('ProgressDialog');

  const html = htmlTemplate.evaluate()
    .setWidth(550)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Email Sending Progress');
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

          if (quotaCheck.reason === 'mailersend_api_rate_limit') {
            // MailerSend API rate limit - wait 1 minute and resume
            alertTitle = 'MailerSend API Rate Limit';
            alertMessage = `MailerSend API rate limit reached (10 requests/minute). ` +
              `The system will automatically resume sending in 1 minute.`;

            // Schedule resume in 1 minute
            const resumeTime = new Date(Date.now() + 60000); // 1 minute from now

            // Clean up any existing triggers first
            if (jobState.triggerId) {
              try {
                const triggers = ScriptApp.getProjectTriggers();
                const oldTrigger = triggers.find(t => t.getUniqueId() === jobState.triggerId);
                if (oldTrigger) {
                  ScriptApp.deleteTrigger(oldTrigger);
                }
              } catch (e) {
                console.error('Error cleaning up old trigger:', e);
              }
            }

            const trigger = ScriptApp.newTrigger('resumeEmailSending')
              .timeBased()
              .at(resumeTime)
              .create();

            jobState.triggerId = trigger.getUniqueId();
            saveJobState(jobState);
          } else if (quotaCheck.reason === 'mailersend_batch_limit') {
            // MailerSend batch limit reached - this shouldn't happen with proper rate limiting
            alertTitle = 'MailerSend Batch Limit';
            alertMessage = `MailerSend batch limit reached. The system will resume in 1 minute.`;
            shouldScheduleResume = true;
          } else if (quotaCheck.reason === 'daily_quota_exhausted') {
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

              // Update job state with current progress (increment by 1 for this email)
              jobState.emailsSent++;
              jobState.lastEmailTime = new Date().toISOString();
              saveJobState(jobState);

              // Add delay between MailerSend requests to respect rate limits
              if (emailsSentThisSession < obj.length) { // Don't delay after the last email
                console.log(`MailerSend: Waiting ${MAILERSEND_DELAY_MS/1000} seconds before next email...`);
                Utilities.sleep(MAILERSEND_DELAY_MS);
              }
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

            // Update job state with current progress (increment by 1 for this email)
            jobState.emailsSent++;
            jobState.lastEmailTime = new Date().toISOString();
            saveJobState(jobState);
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

    // Job completed successfully - update final state and clear
    jobState.emailsSent += emailsSentThisSession;
    jobState.completedTime = new Date().toISOString();
    saveJobState(jobState);

    // Clear job state and triggers
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

          if (quotaCheck.reason === 'mailersend_api_rate_limit') {
            // MailerSend API rate limit - wait 1 minute and resume
            alertTitle = 'MailerSend API Rate Limit';
            alertMessage = `MailerSend API rate limit reached (10 requests/minute). ` +
              `The A/B test will automatically resume in 1 minute.`;

            // Schedule resume in 1 minute
            const resumeTime = new Date(Date.now() + 60000); // 1 minute from now

            // Clean up any existing triggers first
            if (jobState.triggerId) {
              try {
                const triggers = ScriptApp.getProjectTriggers();
                const oldTrigger = triggers.find(t => t.getUniqueId() === jobState.triggerId);
                if (oldTrigger) {
                  ScriptApp.deleteTrigger(oldTrigger);
                }
              } catch (e) {
                console.error('Error cleaning up old trigger:', e);
              }
            }

            const trigger = ScriptApp.newTrigger('resumeEmailSending') // Use same function for consistency
              .timeBased()
              .at(resumeTime)
              .create();

            jobState.triggerId = trigger.getUniqueId();
            saveJobState(jobState);
          } else if (quotaCheck.reason === 'mailersend_batch_limit') {
            // MailerSend batch limit reached
            alertTitle = 'MailerSend Batch Limit';
            alertMessage = `MailerSend batch limit reached. The A/B test will resume in 1 minute.`;
            shouldScheduleResume = true;
          } else if (quotaCheck.reason === 'daily_quota_exhausted') {
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

            // Add delay between MailerSend requests to respect rate limits
            if (emailsSentThisSession < obj.length - 1) { // Don't delay after the last email
              console.log(`MailerSend A/B: Waiting ${MAILERSEND_DELAY_MS/1000} seconds before next email...`);
              Utilities.sleep(MAILERSEND_DELAY_MS);
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

          // Update job state with current progress (increment by 1 for this email)
          jobState.emailsSent++;
          jobState.lastEmailTime = new Date().toISOString();
          saveJobState(jobState);

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

    // Job completed successfully - update final state and clear
    jobState.emailsSent += emailsSentThisSession;
    jobState.completedTime = new Date().toISOString();
    saveJobState(jobState);

    // Clear job state and triggers
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
 * Debug function to check current job status and triggers
 * This helps diagnose why email sending might have stopped
 */
function debugJobStatus() {
  try {
    const jobState = getActiveJobState();
    const triggers = ScriptApp.getProjectTriggers();
    const emailTriggers = triggers.filter(trigger => trigger.getHandlerFunction() === 'resumeEmailSending');

    console.log('=== Job Debug Status ===');
    console.log('Active job state:', jobState);
    console.log(`Total triggers: ${triggers.length}`);
    console.log(`Email resume triggers: ${emailTriggers.length}`);

    if (emailTriggers.length > 0) {
      emailTriggers.forEach((trigger, index) => {
        console.log(`Trigger ${index + 1}: ID=${trigger.getUniqueId()}, Handler=${trigger.getHandlerFunction()}`);
      });
    }

    let message = '';
    if (!jobState) {
      message = 'No active job found. The job may have completed or been cancelled.';
    } else {
      message = `Active Job Found:\n` +
        `Batch: ${jobState.batchId}\n` +
        `Progress: ${jobState.emailsSent} of ${jobState.totalEmails} emails\n` +
        `Resume Count: ${jobState.resumeCount || 0}\n` +
        `Next Resume Time: ${jobState.nextResumeTime || 'Not set'}\n` +
        `Trigger ID: ${jobState.triggerId || 'None'}\n\n` +
        `Active Resume Triggers: ${emailTriggers.length}`;
    }

    SpreadsheetApp.getUi().alert('Job Debug Status', message, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    console.error('Error in debugJobStatus:', error);
    SpreadsheetApp.getUi().alert('Debug Error',
      `Error checking job status: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Manually resume a stuck email job
 * This can be used when the automatic resume isn't working
 */
function manuallyResumeJob() {
  try {
    const jobState = getActiveJobState();

    if (!jobState) {
      SpreadsheetApp.getUi().alert('No Active Job',
        'There is no active email job to resume.',
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Resume Job',
      `Do you want to manually resume the email job for batch "${jobState.batchId}"?\n\n` +
      `Current progress: ${jobState.emailsSent} of ${jobState.totalEmails} emails sent.\n\n` +
      `This will attempt to continue sending emails immediately.`,
      ui.ButtonSet.YES_NO);

    if (response === ui.Button.YES) {
      console.log('Manually resuming email job...');

      // Clear any existing triggers to prevent conflicts
      if (jobState.triggerId) {
        try {
          const triggers = ScriptApp.getProjectTriggers();
          const oldTrigger = triggers.find(t => t.getUniqueId() === jobState.triggerId);
          if (oldTrigger) {
            ScriptApp.deleteTrigger(oldTrigger);
          }
        } catch (e) {
          console.error('Error cleaning up old trigger:', e);
        }
      }

      // Clear the next resume time to allow immediate resumption
      jobState.nextResumeTime = null;
      jobState.triggerId = null;
      saveJobState(jobState);

      // Clear MailerSend rate limiting state to allow immediate sending
      const properties = PropertiesService.getScriptProperties();
      properties.deleteProperty('MAILERSEND_LAST_REQUESTS');

      // Resume the job
      resumeEmailSending();

      ui.alert('Job Resumed',
        'The email job has been manually resumed. Check the progress dialog for updates.',
        ui.ButtonSet.OK);
    }

  } catch (error) {
    console.error('Error in manuallyResumeJob:', error);
    SpreadsheetApp.getUi().alert('Resume Error',
      `Error resuming job: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Reset MailerSend rate limiting state
 * This can help if the rate limiting gets stuck
 */
function resetMailerSendRateLimit() {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('MAILERSEND_LAST_REQUESTS');

    SpreadsheetApp.getUi().alert('Rate Limit Reset',
      'MailerSend rate limiting state has been reset. You can now send emails immediately.',
      SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    console.error('Error resetting rate limit:', error);
    SpreadsheetApp.getUi().alert('Reset Error',
      `Error resetting rate limit: ${error.message}`,
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
 * Records a MailerSend API request timestamp for rate limiting tracking
 * This is a non-blocking version that just records the request
 */
function recordMailerSendRequest() {
  const properties = PropertiesService.getScriptProperties();
  const now = new Date().getTime();

  // Get the last request timestamps (stored as JSON array)
  const lastRequestsJson = properties.getProperty('MAILERSEND_LAST_REQUESTS');
  let lastRequests = [];

  if (lastRequestsJson) {
    try {
      lastRequests = JSON.parse(lastRequestsJson);
    } catch (e) {
      console.error('Error parsing MailerSend request history:', e);
      lastRequests = [];
    }
  }

  // Remove requests older than 1 minute
  const oneMinuteAgo = now - 60000;
  lastRequests = lastRequests.filter(timestamp => timestamp > oneMinuteAgo);

  // Add current request timestamp
  lastRequests.push(now);

  // Save updated timestamps
  properties.setProperty('MAILERSEND_LAST_REQUESTS', JSON.stringify(lastRequests));
}

/**
 * Manages MailerSend API rate limiting by tracking request timestamps
 * @return {boolean} True if we can make a request now, false if we need to wait
 */
function checkMailerSendRateLimit() {
  const properties = PropertiesService.getScriptProperties();
  const now = new Date().getTime();

  // Get the last request timestamps (stored as JSON array)
  const lastRequestsJson = properties.getProperty('MAILERSEND_LAST_REQUESTS');
  let lastRequests = [];

  if (lastRequestsJson) {
    try {
      lastRequests = JSON.parse(lastRequestsJson);
    } catch (e) {
      console.error('Error parsing MailerSend request history:', e);
      lastRequests = [];
    }
  }

  // Remove requests older than 1 minute
  const oneMinuteAgo = now - 60000;
  lastRequests = lastRequests.filter(timestamp => timestamp > oneMinuteAgo);

  // Check if we're under the rate limit
  if (lastRequests.length >= MAILERSEND_REQUESTS_PER_MINUTE) {
    return false; // Rate limit exceeded
  }

  return true; // OK to make request
}

/**
 * Waits for MailerSend rate limit to allow next request
 */
function waitForMailerSendRateLimit() {
  const properties = PropertiesService.getScriptProperties();
  const lastRequestsJson = properties.getProperty('MAILERSEND_LAST_REQUESTS');

  if (!lastRequestsJson) {
    return; // No previous requests, no need to wait
  }

  try {
    const lastRequests = JSON.parse(lastRequestsJson);
    const now = new Date().getTime();
    const oneMinuteAgo = now - 60000;

    // Get requests from the last minute
    const recentRequests = lastRequests.filter(timestamp => timestamp > oneMinuteAgo);

    if (recentRequests.length >= MAILERSEND_REQUESTS_PER_MINUTE) {
      // Calculate how long to wait
      const oldestRecentRequest = Math.min(...recentRequests);
      const waitTime = (oldestRecentRequest + 60000) - now;

      if (waitTime > 0) {
        console.log(`MailerSend rate limit reached. Waiting ${Math.ceil(waitTime/1000)} seconds...`);
        Utilities.sleep(waitTime);
      }
    }
  } catch (e) {
    console.error('Error in waitForMailerSendRateLimit:', e);
    // If there's an error, wait the standard delay to be safe
    Utilities.sleep(MAILERSEND_DELAY_MS);
  }
}

/**
 * Sends an email using MailerSend API with rate limiting
 * @param {string} toEmail - Recipient email address
 * @param {string} toName - Recipient name
 * @param {string} subject - Email subject
 * @param {string} textContent - Plain text content
 * @param {string} htmlContent - HTML content
 * @param {Object} config - MailerSend configuration
 * @return {Object} Response object with success status and message ID
 */
function sendEmailWithMailerSend(toEmail, toName, subject, textContent, htmlContent, config) {
  // For MailerSend, we'll handle rate limiting with delays instead of blocking
  // This prevents the job from getting stuck and needing manual resume
  recordMailerSendRequest(); // Just record the request, don't block
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

  // Clean up any triggers associated with this job
  if (jobState && jobState.triggerId) {
    try {
      const triggers = ScriptApp.getProjectTriggers();
      const trigger = triggers.find(t => t.getUniqueId() === jobState.triggerId);
      if (trigger) {
        ScriptApp.deleteTrigger(trigger);
        console.log('Deleted trigger:', jobState.triggerId);
      }
    } catch (error) {
      console.error('Error cleaning up specific trigger:', error);
    }
  }

  // Also clean up any orphaned email-related triggers as a safety measure
  cleanupOrphanedTriggers();

  properties.deleteProperty('ACTIVE_EMAIL_JOB');
}

/**
 * Cleans up orphaned triggers that may be left over from previous jobs
 */
function cleanupOrphanedTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const emailTriggers = triggers.filter(trigger => {
      const handlerFunction = trigger.getHandlerFunction();
      return handlerFunction === 'resumeEmailSending';
    });

    console.log(`Found ${emailTriggers.length} email-related triggers`);

    // If we have more than 2 email triggers, something's wrong - clean them up
    if (emailTriggers.length > 2) {
      console.log('Cleaning up excess email triggers...');
      emailTriggers.forEach(trigger => {
        try {
          ScriptApp.deleteTrigger(trigger);
          console.log('Deleted orphaned trigger:', trigger.getUniqueId());
        } catch (error) {
          console.error('Error deleting orphaned trigger:', error);
        }
      });
    }
  } catch (error) {
    console.error('Error in cleanupOrphanedTriggers:', error);
  }
}

/**
 * Creates a time-based trigger to resume email sending after the rate limit period
 * @return {string} Trigger ID
 */
function createResumeTrigger() {
  // Clean up any existing triggers first to prevent accumulation
  cleanupOrphanedTriggers();

  const trigger = ScriptApp.newTrigger('resumeEmailSending')
    .timeBased()
    .after(RESUME_DELAY_HOURS * 60 * 60 * 1000) // Convert hours to milliseconds
    .create();

  console.log('Created resume trigger:', trigger.getUniqueId());
  return trigger.getUniqueId();
}

/**
 * Emergency function to clean up ALL triggers (use if you get "too many triggers" error)
 */
function cleanupAllTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    console.log(`Found ${triggers.length} total triggers`);

    let deletedCount = 0;
    triggers.forEach(trigger => {
      try {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      } catch (error) {
        console.error('Error deleting trigger:', error);
      }
    });

    console.log(`Deleted ${deletedCount} triggers`);

    // Also clear any active job state since triggers are gone
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('ACTIVE_EMAIL_JOB');

    SpreadsheetApp.getUi().alert('Triggers Cleaned Up',
      `Deleted ${deletedCount} triggers and cleared job state.\n\n` +
      `You can now start new email jobs without the "too many triggers" error.`,
      SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    console.error('Error in cleanupAllTriggers:', error);
    SpreadsheetApp.getUi().alert('Cleanup Error',
      `Error cleaning up triggers: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Gets detailed job progress data for the progress dialog
 * @return {Object} Job progress data
 */
function getJobProgress() {
  const jobState = getActiveJobState();
  const emailConfig = getEmailServiceConfig();

  if (!jobState) {
    return {
      active: false,
      message: 'No active email job'
    };
  }

  // Calculate if job is complete
  const isComplete = jobState.emailsSent >= jobState.totalEmails;

  // Determine if job is paused (has a next resume time in the future)
  const isPaused = jobState.nextResumeTime && new Date(jobState.nextResumeTime) > new Date();

  // Get pause reason
  let pauseReason = '';
  if (isPaused) {
    const resumeTime = new Date(jobState.nextResumeTime);
    const now = new Date();
    const minutesUntilResume = Math.ceil((resumeTime - now) / 1000 / 60);

    if (minutesUntilResume > 60) {
      const hoursUntilResume = Math.ceil(minutesUntilResume / 60);
      pauseReason = `Rate limit reached. Resuming in ${hoursUntilResume} hour(s) at ${resumeTime.toLocaleTimeString()}.`;
    } else {
      pauseReason = `Rate limit reached. Resuming in ${minutesUntilResume} minute(s) at ${resumeTime.toLocaleTimeString()}.`;
    }
  }

  return {
    active: true,
    batchId: jobState.batchId,
    emailService: emailConfig.service || 'gmail',
    isABTest: jobState.isABTest || false,
    totalEmails: jobState.totalEmails,
    emailsSent: jobState.emailsSent,
    startTime: jobState.startTime,
    nextResumeTime: jobState.nextResumeTime,
    resumeCount: jobState.resumeCount || 0,
    isComplete: isComplete,
    isPaused: isPaused,
    pauseReason: pauseReason,
    lastUpdated: new Date().toISOString()
  };
}

/**
 * Shows the current job status to the user (legacy function, now redirects to progress dialog)
 */
function showJobStatus() {
  const jobState = getActiveJobState();

  if (!jobState) {
    SpreadsheetApp.getUi().alert('No Active Job', 'There is no active email sending job.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Show the progress dialog instead of the old alert
  showProgressDialog();
}

/**
 * Cancels the active email job (called from progress dialog)
 * @return {boolean} True if job was cancelled, false if no job or user cancelled
 */
function cancelActiveJob() {
  const jobState = getActiveJobState();

  if (!jobState) {
    throw new Error('No active email sending job to cancel.');
  }

  clearJobState();
  return true;
}

/**
 * Cancels the active email job with UI confirmation (called from menu)
 */
function cancelActiveJobWithConfirmation() {
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

  // MailerSend has much higher email limits than Gmail
  // We handle rate limiting with delays between requests instead of blocking the entire job
  if (emailConfig.service === 'mailersend') {
    return {
      canSend: true,
      reason: 'mailersend_ready',
      maxEmails: 1000, // Allow large batches since MailerSend handles the rate limiting with delays
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
