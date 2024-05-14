function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Extractor')
    .addItem('Extract Emails from URLs', 'runBatchExtract')
    .addItem('Extract Emails from Cell Text', 'extractEmailsFromCell')
    .addToUi();
}

function batchExtractEmails(startRow, numRows) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var urlRange = sheet.getRange(startRow, 1, numRows, 1);  // URLs are in Column A
  var urls = urlRange.getValues();

  var results = [];
  urls.forEach(function(row) {
    var url = row[0];
    if (!url || !url.startsWith("http")) {
      results.push(["Invalid URL"]);
      return;
    }
    try {
      var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      var html = response.getContentText();
      if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
        var emails = extractEmailsFromText(html);
        results.push([emails.length > 0 ? emails.join(", ") : "No emails found"]);
      } else {
        results.push(["HTTP Error " + response.getResponseCode()]);
      }
    } catch (e) {
      results.push(["Error fetching URL: " + e.toString()]);
    }
  });
  var outputRange = sheet.getRange(startRow, 2, results.length, 1);  // Output in Column B
  outputRange.setValues(results);
}

function extractEmailsFromText(text) {
  var emails = [];
  var normalEmailRegex = /[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}/g;
  var obfuscatedEmailRegex = /([a-zA-Z0-9._-]+)\s*\[at\]\s*([a-zA-Z0-9.-]+)(?:\s*\[dot\]\s*([a-zA-Z]{2,6}))?/gi;

  var match;
  while (match = normalEmailRegex.exec(text)) {
    emails.push(match[0]);
  }
  while (match = obfuscatedEmailRegex.exec(text)) {
    var domain = match[2];
    var topLevelDomain = match[3] ? '.' + match[3] : '';  // Add the TLD if it exists
    var formattedEmail = match[1] + '@' + domain + topLevelDomain;
    emails.push(formattedEmail);
  }

  return Array.from(new Set(emails));  // Remove duplicates
}

function extractEmailsFromCell() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Enter cell reference (e.g., "A1") to extract emails from its text:');
  if (result.getSelectedButton() == ui.Button.OK) {
    var cellRef = result.getResponseText();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var text = sheet.getRange(cellRef).getValue();
    var emails = extractEmailsFromText(text);
    ui.alert('Emails found in ' + cellRef + ':\n' + emails.join(', '));
  }
}

function runBatchExtract() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();  // Dynamically find the last row with data
  batchExtractEmails(1, lastRow);  // Run extraction for all rows with URLs
}
