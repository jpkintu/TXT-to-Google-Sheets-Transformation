function onOpen() {
  // Create custom menu when spreadsheet opens
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Data Import')
    .addItem('Refresh Data Now', 'updateSheetsFromAllFiles')
    .addToUi();
}
function updateSheetsFromAllFiles() {
  try {
    // Configuration
    var folderId = '1366REWX7kxvstO74Hbllg5ZZOgvq-sZb';
    var spreadsheetId = '1YgSJges-r1U72hFwO-x7MwZUAXQhY4vPHGzNeNCslq8';
    var mainSheetName = 'Received Extractions';
    var malformedSheetName = 'Malformed Data';
    var expectedColumns = 15;
    
    // Get folder and spreadsheet
    var folder = DriveApp.getFolderById(folderId);
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // Get or create main sheet
    var mainSheet = spreadsheet.getSheetByName(mainSheetName) || 
                   spreadsheet.getSheets().find(s => s.getSheetId() == 0);
    if (!mainSheet) throw new Error('Main sheet not found');
    
    // Get or create malformed data sheet
    var malformedSheet = spreadsheet.getSheetByName(malformedSheetName);
    if (!malformedSheet) {
      malformedSheet = spreadsheet.insertSheet(malformedSheetName);
      // Set headers for malformed sheet
      malformedSheet.getRange(1, 1, 1, 4)
        .setValues([['Source File', 'Row Number', 'Column Status', 'Original Data']]);
    }
    
    // Clear existing data (keep headers)
    clearSheetData(mainSheet);
    clearSheetData(malformedSheet);
    
    // Process all files
    var files = folder.getFiles();
    var allData = [];
    var badData = [];
    var processedFiles = 0;
    var maxColumns = 0; // Track maximum columns found
    
    while (files.hasNext()) {
      var file = files.next();
      var content = file.getBlob().getDataAsString();
      var rows = content.split('\n');
      
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i].trim();
        if (row !== '') {
          var data = row.split(';');
          allData.push(data);
          maxColumns = Math.max(maxColumns, data.length); // Update max columns
          
          // Only flag in malformed sheet if column count doesn't match
          if (data.length !== expectedColumns) {
            badData.push({
              file: file.getName(),
              row: i+1,
              data: data,
              expected: expectedColumns,
              actual: data.length
            });
          }
        }
      }
      processedFiles++;
    }
    
    // Write all data to main sheet
    if (allData.length > 0) {
      // Create a 2D array with consistent column count
      var formattedData = allData.map(row => {
        var newRow = new Array(maxColumns).fill(''); // Create empty array with max columns
        for (var i = 0; i < row.length; i++) {
          newRow[i] = row[i]; // Copy existing values
        }
        return newRow;
      });
      
      // Write data starting from first column
      mainSheet.getRange(2, 1, formattedData.length, maxColumns)
        .setValues(formattedData);
    }
    
    // Format and write bad data to malformed sheet
    if (badData.length > 0) {
      var formattedBadData = badData.map(item => {
        return [
          item.file,
          item.row,
          item.actual + ' (expected ' + item.expected + ')',
          item.data.join(';') // Original malformed row
        ];
      });
      
      malformedSheet.getRange(2, 1, formattedBadData.length, 4)
        .setValues(formattedBadData);
    }
    
    // Log results
    var message = [
      'Processed ' + processedFiles + ' files',
      'Added ' + allData.length + ' total rows to ' + mainSheetName,
      'Maximum columns found: ' + maxColumns,
      'Flagged ' + badData.length + ' malformed rows in ' + malformedSheetName,
      badData.length > 0 ? 'Review malformed data for corrections' : ''
    ].join('\n');
    
    Logger.log(message);
    SpreadsheetApp.getUi().alert('Data Refresh Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (err) {
    Logger.log('Error: ' + err);
    SpreadsheetApp.getUi().alert('Error', 'An error occurred during data refresh:\n\n' + err, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Helper function to clear sheet data while preserving headers
function clearSheetData(sheet) {
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn())
      .clearContent();
  }
}