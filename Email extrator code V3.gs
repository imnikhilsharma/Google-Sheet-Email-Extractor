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
  urls.forEach(function(row, index) {
    var url = row[0];
    if (!url || !url.startsWith("http")) {
      results.push(["Invalid URL"]);
      return;
    }

    try {
      Logger.log("Fetching URL (" + (startRow + index) + "): " + url);
      
      var options = {
        muteHttpExceptions: true,
        timeoutMs: 5000  // Set timeout to 5 seconds
      };
      var response = UrlFetchApp.fetch(url, options);
      
      if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
        var html = response.getContentText();
        
        var emails = extractEmailsFromText(html);
        if (emails.length > 0) {
          results.push([emails.join(", ")]);
        } else {
          var facebookUrl = extractFacebookUrlFromHtml(html);
          if (facebookUrl) {
            results.push(["No emails found, Facebook URL: " + facebookUrl]);
          } else {
            results.push(["No emails or Facebook URL found"]);
          }
        }
      } else {
        results.push(["HTTP Error " + response.getResponseCode()]);
      }
    } catch (e) {
      Logger.log("Error fetching URL or timed out: " + e.toString());
      results.push(["Error fetching URL or timed out: " + e.toString()]);
    }
  });

  Logger.log("Writing results to the sheet...");
  var outputRange = sheet.getRange(startRow, 2, results.length, 1);  // Output in Column B
  outputRange.setValues(results);
}

function extractEmailsFromText(text) {
  var emails = [];
  var emailRegex = /(?:mailto:)?[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}/gi; // General email regex

  var match;
  while (match = emailRegex.exec(text)) {
    emails.push(match[0]);
  }

  return Array.from(new Set(emails));  // Remove duplicates
}

function extractFacebookUrlFromHtml(html) {
  var facebookUrlRegex = /href=["'](https?:\/\/(www\.)?facebook\.com\/[^\s"']+)["']/gi;
  var match = facebookUrlRegex.exec(html);
  return match ? match[1] : null;
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
