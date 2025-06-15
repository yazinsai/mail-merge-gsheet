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
 * recipient addresses, email sent column, and batch column.
*/
const RECIPIENT_COL  = "Recipient";
const EMAIL_SENT_COL = "Email Sent";
const BATCH_COL = "Batch";
 
/**
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Select Batch & Send Emails', 'showBatchSelectionDialog')
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
    .setWidth(400)
    .setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(html, 'Select Batch for Mail Merge');
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
 * @param {Object} formData containing selectedBatch and subjectLine
 */
function processBatchSelection(formData) {
  const selectedBatch = formData.selectedBatch;
  const subjectLine = formData.subjectLine;

  if (!selectedBatch || !subjectLine) {
    SpreadsheetApp.getUi().alert('Missing Information',
      'Please select a batch and enter a subject line.',
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  sendEmails(subjectLine, SpreadsheetApp.getActiveSheet(), selectedBatch);
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
 * Sends emails from sheet data.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
 * @param {string} selectedBatch (optional) to filter emails by batch
*/
function sendEmails(subjectLine, sheet=SpreadsheetApp.getActiveSheet(), selectedBatch=null) {
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
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  
  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // Fetches displayed values for each row in the Range HT Andrew Roberts 
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  // Assumes row 1 contains our column headings
  const heads = data.shift(); 
  
  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  
  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // For a pretty version, see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // Creates an array to record sent emails
  const out = [];

  // Loops through all the rows of data
  obj.forEach(function(row, rowIdx){
    // Only sends emails if email_sent cell is blank and not hidden by a filter
    // Also filter by selected batch if specified
    const batchMatches = !selectedBatch || row[BATCH_COL] === selectedBatch;
    if (row[EMAIL_SENT_COL] == '' && batchMatches){
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        // See https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
        // If you need to send emails with unicode/emoji characters change GmailApp for MailApp
        // Uncomment advanced parameters as needed (see docs for limitations)
        GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          // bcc: 'a.bcc@email.com',
          // cc: 'a.cc@email.com',
          // from: 'an.alias@email.com',
          // name: 'name of the sender',
          // replyTo: 'a.reply@email.com',
          // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        });
        // Edits cell to record email sent date
        out.push([new Date()]);
      } catch(e) {
        // modify cell to record error
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });
  
  // Updates the sheet with new data
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);
  
  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
  */
  function getGmailTemplateFromDrafts_(subject_line){
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
    function subjectFilter_(subject_line){
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
  };
}
