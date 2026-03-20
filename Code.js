// 1. Serve the Web App
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Daily Task Logger')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 2. Receive data from the form and log it to the Sheet
function logTask(formData) {
  try {
    // Securely grab the FULL email address
    const userEmail = Session.getActiveUser().getEmail();

    // Connect to the specific database sheet
    const ss = SpreadsheetApp.openById("1CCgoKVGRKneXcUt1xjoZETft-asvXU7FXVWguDJHdq4");
    const sheet = ss.getSheetByName("TaskLogs");
    
    // Append the row with the full email address and status FALSE
    sheet.appendRow([userEmail, formData.date, formData.day, formData.category, formData.task, false]);
    
    return "Success";
  } catch (error) {
    Logger.log("Error logging task: " + error.toString());
    throw new Error("Failed to save to database.");
  }
}

// ==============================================================================
// PHASE 4: THE AUTOMATION ENGINE (Daily Digest)
// ==============================================================================

function generateDailyDigests() {
  const ss = SpreadsheetApp.openById("1CCgoKVGRKneXcUt1xjoZETft-asvXU7FXVWguDJHdq4");
  const sheet = ss.getSheetByName("TaskLogs");
  const data = sheet.getDataRange().getValues();
  
  // Remove the header row
  const headers = data.shift(); 
  
  const tasksByUser = {};
  const rowsToMarkAsProcessed = [];

  // 1. Group Data by User (Full Email)
  data.forEach((row, index) => {
    const user = row[0]; // Full email address
    const date = row[1];
    const day = row[2];
    const category = row[3];
    const task = row[4];
    const isProcessed = row[5];

    // Only process rows that have a user and are marked FALSE
    if (user !== "" && isProcessed === false) {
      if (!tasksByUser[user]) {
        tasksByUser[user] = {
          date: date,
          day: day,
          tasks: []
        };
      }
      tasksByUser[user].tasks.push({ category: category, description: task });
      
      // Store the exact row number (+2 offset for header and 0-indexing)
      rowsToMarkAsProcessed.push(index + 2); 
    }
  });

  // Stop if no new tasks
  if (Object.keys(tasksByUser).length === 0) {
    Logger.log("No new tasks to compile today.");
    return;
  }

  // 2. Folder Management
  let archiveFolder;
  const folderIterator = DriveApp.getFoldersByName("Daily Task Digests");
  if (folderIterator.hasNext()) {
    archiveFolder = folderIterator.next();
  } else {
    archiveFolder = DriveApp.createFolder("Daily Task Digests");
  }

  // 3. Generate Docs and Send Emails
  for (const user in tasksByUser) {
    const userData = tasksByUser[user];
    
    let displayDate = userData.date;
    if (userData.date instanceof Date) {
      displayDate = Utilities.formatDate(userData.date, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    // Create Doc
    const docName = `Task Digest - ${user} - ${displayDate}`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();

    // Styling
    const title = body.insertParagraph(0, "Daily Task Digest");
    title.setFontSize(16).setBold(true).setForegroundColor('#4F46E5').setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    const subTitleText = `User: ${user} | Date: ${displayDate} | Day: ${userData.day}`;
    const subTitle = body.appendParagraph(subTitleText);
    subTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setForegroundColor('#6B7280').setFontSize(11);

    body.appendHorizontalRule();
    body.appendParagraph("");

    userData.tasks.forEach(t => {
      const listItem = body.appendListItem("");
      listItem.appendText(`${t.category}: `).setBold(true);
      listItem.appendText(t.description).setBold(false);
    });

    doc.saveAndClose();
    const docFile = DriveApp.getFileById(doc.getId());
    docFile.moveTo(archiveFolder);

    // Permissions
    const targetEmail = user; 
    docFile.addEditor(targetEmail, {sendNotificationEmail: false});

    // Extract First Name for greeting
    const firstNameRaw = targetEmail.split('@')[0].split('.')[0];
    const firstName = firstNameRaw.charAt(0).toUpperCase() + firstNameRaw.slice(1);

    // Email Delivery (No-Reply)
    const docUrl = doc.getUrl();
    const emailSubject = `Your Daily Task Digest - ${displayDate}`;
    const emailBody = `Hello ${firstName},\n\nYour tasks for the day have been successfully compiled.\n\nYou can view and edit your digest here:\n${docUrl}\n\nBest regards,\nYour Automated System`;

    MailApp.sendEmail({
      to: targetEmail,
      subject: emailSubject,
      body: emailBody,
      noReply: true
    });
  }

  // 4. Update the Spreadsheet (Cleanup)
  rowsToMarkAsProcessed.forEach(rowIndex => {
    const targetCell = sheet.getRange(rowIndex, 6); // Column F
    targetCell.setValue(true); 
    targetCell.setBackground('#c4f7c4'); // Color only processed cells
  });
}