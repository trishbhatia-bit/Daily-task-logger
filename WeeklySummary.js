function compileWeeklySummaries() {
  const ss = SpreadsheetApp.openById("1CCgoKVGRKneXcUt1xjoZETft-asvXU7FXVWguDJHdq4"); 
  const sheet = ss.getSheetByName("TaskLogs");
  const data = sheet.getDataRange().getValues();
  
  const headers = data.shift(); 
  const tasksByUser = {};
  
  // 1. Calculate the Date Range (Monday to Friday of the current week)
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Start of Saturday
  
  const targetMonday = new Date(today);
  targetMonday.setDate(today.getDate() - 5);
  
  const targetFriday = new Date(today);
  targetFriday.setDate(today.getDate() - 1);
  targetFriday.setHours(23, 59, 59, 999); 

  // 2. Filter and Group Data by User (User column now contains full email)
  data.forEach((row) => {
    const user = row[0];
    const taskDate = new Date(row[1]);
    const day = row[2];
    const category = row[3];
    const task = row[4];

    if (user !== "" && taskDate >= targetMonday && taskDate <= targetFriday) {
      if (!tasksByUser[user]) {
        tasksByUser[user] = { tasks: [] };
      }
      tasksByUser[user].tasks.push(`Category: ${category} | Task: ${task}`);
    }
  });

  if (Object.keys(tasksByUser).length === 0) {
    Logger.log("No tasks found for the Monday-Friday period.");
    return;
  }

  // 3. Folder Management
  let archiveFolder;
  const folderIterator = DriveApp.getFoldersByName("Weekly Task Summaries");
  if (folderIterator.hasNext()) {
    archiveFolder = folderIterator.next();
  } else {
    archiveFolder = DriveApp.createFolder("Weekly Task Summaries");
  }

  const formattedMonday = Utilities.formatDate(targetMonday, Session.getScriptTimeZone(), "MMM dd");
  const formattedFriday = Utilities.formatDate(targetFriday, Session.getScriptTimeZone(), "MMM dd");
  const dateRangeStr = `${formattedMonday} to ${formattedFriday}`;

  // 4. Generate AI Summaries and Docs
  for (const user in tasksByUser) {
    const userData = tasksByUser[user];
    
    // Extract first name from full email (trish.bhatia@highspring.in -> trish -> Trish)
    const firstNameRaw = user.split('@')[0].split('.')[0];
    const firstName = firstNameRaw.charAt(0).toUpperCase() + firstNameRaw.slice(1);

    const rawTaskList = userData.tasks.join("\n");

    // Construct the strict, professional prompt
    const prompt = `You are an AI generating a professional weekly self-report. Below is a list of tasks completed by ${firstName} between ${dateRangeStr}. 
    Write a concise, highly professional, and informative 2-paragraph summary of the week written strictly in the FIRST PERSON ("I"), as if ${firstName} is reporting on their own work. 
    Highlight the main focus areas, key accomplishments, and overall productivity. 
    Tone constraints: Even though it uses "I", do NOT write it like a personal journal or diary entry. Keep the tone objective, analytical, and structured for management review (e.g., "This week, my primary focus was...", "I successfully completed...").
    
    Tasks:
    ${rawTaskList}`;

    const aiSummary = callGeminiAPI(prompt);

    // Create Doc
    const docName = `Weekly Summary - ${user} - ${dateRangeStr}`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();

    const title = body.insertParagraph(0, "Weekly Task Summary");
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setForegroundColor('#4F46E5');

    const subTitleText = `User: ${user} | Period: ${dateRangeStr}`;
    const subTitle = body.appendParagraph(subTitleText);
    subTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setForegroundColor('#6B7280').setFontSize(11);

    body.appendHorizontalRule();
    body.appendParagraph(""); 

    body.appendParagraph("Executive Summary").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(aiSummary);
    
    body.appendParagraph(""); 

    body.appendParagraph("Appendix: Raw Task Log").setHeading(DocumentApp.ParagraphHeading.HEADING3);
    userData.tasks.forEach(t => {
      body.appendListItem(t).setFontSize(10).setForegroundColor('#374151');
    });

    doc.saveAndClose();
    const docFile = DriveApp.getFileById(doc.getId());
    docFile.moveTo(archiveFolder);
    
    const targetEmail = user;
    
    // Use Advanced Drive API to grant edit access silently
    Drive.Permissions.create(
      {
        role: 'writer',
        type: 'user',
        emailAddress: targetEmail
      },
      doc.getId(),
      {
        sendNotificationEmail: false
      }
    );  

    // 5. Send the Email with no-reply
    const docUrl = doc.getUrl();
    const emailSubject = `Your Weekly Task Summary - ${dateRangeStr}`;
    const emailBody = `Hello ${firstName},\n\nYour tasks for the week of ${dateRangeStr} have been successfully compiled and summarized.\n\nYou can review your weekly summary here:\n${docUrl}\n\nBest regards,\nYour Automated System`;

    MailApp.sendEmail({
      to: targetEmail,
      subject: emailSubject,
      body: emailBody,
      noReply: true
    });
  }
}

// Helper function to handle the external API call securely
function callGeminiAPI(promptText) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  
  if (!apiKey) {
    Logger.log("Error: GEMINI_API_KEY is missing from Script Properties.");
    return "Notice: AI configuration error. The API key was not found in Script Properties.";
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  
  const payload = {
    "contents": [{
      "parts": [{"text": promptText}]
    }]
  };
  
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.candidates && json.candidates.length > 0) {
      return json.candidates[0].content.parts[0].text;
    } else {
      return "Notice: The AI system was unable to generate a summary for this data. Please refer to the raw task log below.";
    }
  } catch (e) {
    Logger.log("AI API Error: " + e);
    return "Notice: The AI compilation service is currently unavailable. Please refer to the raw task log below.";
  }
}

// Diagnostic Debug Function
function debugGeminiAPI() {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  Logger.log("1. API Key found in properties: " + (apiKey ? "YES" : "NO"));
  
  if (!apiKey) return;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const payload = {
    "contents": [{
      "parts": [{"text": "Write a one sentence greeting."}]
    }]
  };
  
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true 
  };
  
  Logger.log("2. Sending request to Gemini...");
  const response = UrlFetchApp.fetch(url, options);
  
  Logger.log("3. HTTP Status Code: " + response.getResponseCode());
  Logger.log("4. Raw Response Body: \n" + response.getContentText());
}