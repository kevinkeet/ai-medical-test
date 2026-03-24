/**
 * Google Apps Script Web App for AI Medical Test Session Logging
 *
 * SETUP INSTRUCTIONS:
 * 1. Create a new Google Sheet
 * 2. Create two tabs: "Sessions" and "Messages"
 * 3. Add headers to "Sessions" tab (Row 1):
 *    Timestamp | Session ID | Name | Email | Role | Specialty | Time Taken (sec) | Sections Completed | Total AI Interactions | Session Data (JSON)
 * 4. Add headers to "Messages" tab (Row 1):
 *    Timestamp | Session ID | Participant | Email | Case Number | Case Title | Message Role | Message Content | Message Order
 * 5. Go to Extensions > Apps Script
 * 6. Paste this code into Code.gs
 * 7. Click Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 8. Copy the web app URL and add it to CONFIG.loggingEndpoint in index.html
 */

// Replace with your actual spreadsheet ID after creating it
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const sessionId = Utilities.getUuid();
    const timestamp = new Date().toISOString();

    // Write to Sessions tab
    logSession(ss, data, sessionId, timestamp);

    // Write to Messages tab
    logMessages(ss, data, sessionId, timestamp);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, sessionId: sessionId }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const action = e.parameter.action || 'getSessions';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  if (action === 'getSessions') {
    return ContentService
      .createTextOutput(JSON.stringify(getSessions(ss)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getMessages') {
    const sessionId = e.parameter.sessionId;
    return ContentService
      .createTextOutput(JSON.stringify(getMessages(ss, sessionId)))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getAllData') {
    return ContentService
      .createTextOutput(JSON.stringify({
        sessions: getSessions(ss),
        messages: getAllMessages(ss)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function logSession(ss, data, sessionId, timestamp) {
  const sheet = ss.getSheetByName('Sessions');
  if (!sheet) throw new Error('Sessions tab not found');

  const participant = data.participant || {};
  const timeTaken = data.timeTaken || 0;
  const sectionsCompleted = data.sectionsCompleted || '0/5';
  const totalInteractions = data.totalAiInteractions || 0;

  // Store full session data as JSON for detailed analysis
  const sessionJson = JSON.stringify({
    responses: data.responses,
    cases: data.cases
  });

  sheet.appendRow([
    timestamp,
    sessionId,
    participant.name || '',
    participant.email || '',
    participant.role || '',
    participant.specialty || '',
    timeTaken,
    sectionsCompleted,
    totalInteractions,
    sessionJson
  ]);
}

function logMessages(ss, data, sessionId, timestamp) {
  const sheet = ss.getSheetByName('Messages');
  if (!sheet) throw new Error('Messages tab not found');

  const participant = data.participant || {};
  const responses = data.responses || [];
  const cases = data.cases || [];

  const rows = [];

  responses.forEach(function(response, caseIndex) {
    const chatHistory = response.chatHistory || [];
    const caseTitle = cases[caseIndex] || ('Case ' + (caseIndex + 1));

    chatHistory.forEach(function(msg, msgIndex) {
      rows.push([
        timestamp,
        sessionId,
        participant.name || '',
        participant.email || '',
        caseIndex + 1,
        caseTitle,
        msg.role,
        msg.content,
        msgIndex + 1
      ]);
    });
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function getSessions(ss) {
  const sheet = ss.getSheetByName('Sessions');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(header, i) {
      obj[header] = row[i];
    });
    return obj;
  });
}

function getMessages(ss, sessionId) {
  const sheet = ss.getSheetByName('Messages');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  return data.slice(1)
    .filter(function(row) { return !sessionId || row[1] === sessionId; })
    .map(function(row) {
      var obj = {};
      headers.forEach(function(header, i) {
        obj[header] = row[i];
      });
      return obj;
    });
}

function getAllMessages(ss) {
  const sheet = ss.getSheetByName('Messages');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(header, i) {
      obj[header] = row[i];
    });
    return obj;
  });
}

/**
 * Helper: Initialize sheet headers (run once manually)
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sessions tab
  var sessionsSheet = ss.getSheetByName('Sessions');
  if (!sessionsSheet) {
    sessionsSheet = ss.insertSheet('Sessions');
  }
  sessionsSheet.getRange(1, 1, 1, 10).setValues([[
    'Timestamp', 'Session ID', 'Name', 'Email', 'Role', 'Specialty',
    'Time Taken (sec)', 'Sections Completed', 'Total AI Interactions', 'Session Data (JSON)'
  ]]);
  sessionsSheet.getRange(1, 1, 1, 10).setFontWeight('bold');

  // Messages tab
  var messagesSheet = ss.getSheetByName('Messages');
  if (!messagesSheet) {
    messagesSheet = ss.insertSheet('Messages');
  }
  messagesSheet.getRange(1, 1, 1, 9).setValues([[
    'Timestamp', 'Session ID', 'Participant', 'Email', 'Case Number',
    'Case Title', 'Message Role', 'Message Content', 'Message Order'
  ]]);
  messagesSheet.getRange(1, 1, 1, 9).setFontWeight('bold');

  // Remove default Sheet1 if it exists
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
}
