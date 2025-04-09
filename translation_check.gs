function justBatch() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TranslationChecking');
  const dataRange = sheet.getRange(4, 1, sheet.getLastRow() - 3, 2);
  const data = dataRange.getValues();
  const GPT_API = "ADD GPT API HERE";
  sheet.getRange('K1').setValue("");
  sheet.getRange('K4').setValue("");


  ///Check if there is language in I4
  var cellValue = sheet.getRange('I4').getValue();
  
  if (!cellValue || cellValue.trim() === '') {
    throw new Error('Enter the translation language in I4');
  }

  ///Check if there is text in C4
  var cellValue = sheet.getRange('C4').getValue();
  
  if (!cellValue || cellValue.trim() === '') {
    throw new Error('Paste the English text into column C');
  }

  ///Check if there is text in D4
  var cellValue = sheet.getRange('D4').getValue();
  
  if (!cellValue || cellValue.trim() === '') {
    throw new Error('Paste the translated text into column D');
  }

  // Get the data in column AG, starting from AG4
  var wdataRange = sheet.getRange(4, 33, sheet.getLastRow() - 1, 1);
  var wdata = wdataRange.getValues();
  
  // Create an array to hold the JSONL lines
  var jsonlLines = [];
  
  // Loop through the data and add non-empty rows to the jsonlLines array
  for (var i = 0; i < wdata.length; i++) {
    if (wdata[i][0]) {
      jsonlLines.push(wdata[i][0]);
    }
  }
  
  // Join the lines with newline characters
  var jsonlContent = jsonlLines.join("\n");
  
  // Get the current date and time
  var now = new Date();
  var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HHmm');

  // Get the sheet file name and replace spaces with underscores
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getName().replace(/\s+/g, '_');

  // Create the file name with sheet name, date, and time
  var fileName = `${sheetName}_${formattedDate}_${formattedTime}.jsonl`;

  // Get the current spreadsheet file and its parent folder
  var spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parentFolder = spreadsheetFile.getParents().next();

  var folderName = 'Created_' + formattedDate;

  // Check if the folder already exists
  var folders = parentFolder.getFoldersByName(folderName);
  var dailyFolder;
  if (folders.hasNext()) {
    // Folder exists
    dailyFolder = folders.next();
  } else {
    // Folder does not exist, create it
    dailyFolder = parentFolder.createFolder(folderName);
    }



  // Create the new file in the same folder as the spreadsheet
  var file = dailyFolder.createFile(fileName, jsonlContent, MimeType.PLAIN_TEXT);

  Logger.log('File created with ID: ' + file.getId());

  // Details - Update File Name
  const FILE_NAME = fileName;

  // Create the payload for the file upload request
  const files = dailyFolder.getFilesByName(FILE_NAME);
  
  // Check if the file exists
  if (!files.hasNext()) {
    Logger.log('File not found');
    return;
  }
  
  const fileBlob = files.next().getBlob();

  // Create the payload for the file upload request
  const uploadPayload = {
    purpose: 'batch',
    file: fileBlob
  };

  // Create the options for the UrlFetchApp file upload request
  const uploadOptions = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + GPT_API
    },
    payload: uploadPayload
  };

  // Make the request to the OpenAI API to upload the file
  const uploadResponse = UrlFetchApp.fetch('https://api.openai.com/v1/files', uploadOptions);
  const uploadResponseData = JSON.parse(uploadResponse.getContentText());
  
  // Log the response and get the file ID
  Logger.log(uploadResponseData);
  const fileId = uploadResponseData.id;
  
  if (fileId) {
    // Create the payload for the batch creation request
    const batchPayload = {
      input_file_id: fileId,
      endpoint: '/v1/chat/completions',
      completion_window: '24h'
    };

    // Create the options for the UrlFetchApp batch creation request
    const batchOptions = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + GPT_API
      },
      payload: JSON.stringify(batchPayload)
    };

    // Make the request to the OpenAI API to create the batch
    const batchResponse = UrlFetchApp.fetch('https://api.openai.com/v1/batches', batchOptions);
    const batchResponseData = JSON.parse(batchResponse.getContentText());
    
    // Log the batch response and get the batch ID
    Logger.log(batchResponseData);
    const batchId = batchResponseData.id;
    sheet.getRange('M4').setValue(batchId);

    const startTime = new Date().getTime();
    const TIMEOUT_LIMIT = 30 * 1000; // 30 seconds in milliseconds

    // Check the batch status and download the output file once it's ready
    while (true) {
      const currentTime = new Date().getTime();
      const elapsedTime = currentTime - startTime;
 
      // Time-out Monitor
      if (elapsedTime > TIMEOUT_LIMIT) {
        throw new Error('GPT needs longer to run, please wait for 30 minutes, then click the "CHECK IF BATCH IS COMPLETE" button');
      }
      // Create the options for the UrlFetchApp request
      const statusOptions = {
        method: 'get',
        headers: {
          'Authorization': 'Bearer ' + GPT_API
        }
      };

      // Make the request to the OpenAI API to check the batch status
      const statusResponse = UrlFetchApp.fetch('https://api.openai.com/v1/batches/' + batchId, statusOptions);
      const statusResponseData = JSON.parse(statusResponse.getContentText());
      
      // Log the batch status
      Logger.log(statusResponseData);



      // PROGRESS TRACKER START Extract the request counts
      var requestCounts = statusResponseData.request_counts || { total: 0, completed: 0, failed: 0 };

      var progressSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TranslationChecking');
      var total = requestCounts.total;
      var completed = requestCounts.completed;

      // Update the progress text in F1
      var progressText = "Number of responses: " + total + "\nNumber completed: " + completed;
      progressSheet.getRange('K1').setValue(progressText);

      // Calculate the percentage of completion
      var progressPercentage = total > 0 ? (completed / total) * 100 : 0;

      // Fill F2 with green blocks based on progress
      var cell = progressSheet.getRange('K4');
      var progressBlocks = Math.round(progressPercentage / 10);
      var progressBar = "▓".repeat(progressBlocks) + "░".repeat(10 - progressBlocks); // 20 blocks total
      cell.setValue(progressBar);

      // Force the spreadsheet to update
      SpreadsheetApp.flush();

      // PROGRESS TRACKER START Extract the request counts

      
      // Check if the batch is processed
      if (statusResponseData.status === 'completed' || statusResponseData.status === 'cancelled') {
        const outputFileId = statusResponseData.output_file_id;
        
        // Create the options for the UrlFetchApp request to download the file content
        const downloadOptions = {
          method: 'get',
          headers: {
            'Authorization': 'Bearer ' + GPT_API
          }
        };

        // Make the request to the OpenAI API to download the file content
        const downloadResponse = UrlFetchApp.fetch('https://api.openai.com/v1/files/' + outputFileId + '/content', downloadOptions);
        const outputBlob = downloadResponse.getBlob();

        // Save the file to Google Drive and store the file name in the global variable
        generatedFileName = FILE_NAME.replace('.jsonl', '_OUTPUT.jsonl');
        dailyFolder.createFile(outputBlob).setName(generatedFileName);
  
        // Log the success message
        Logger.log('Output file downloaded and saved to Google Drive');
        Logger.log('Generated file name: ' + generatedFileName);
        break;
      } else {
        // Wait for 5 seconds before checking again
        Utilities.sleep(5000);
      }
    }
  }

  Logger.log('justBatch completed.');
  // Call the second function at the end of the first function
  useBatchFileToBuildQCodeFrame();
}

function useBatchFileToBuildQCodeFrame() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TranslationChecking');
  
    // Get the current spreadsheet file and its parent folder
  var spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parentFolder = spreadsheetFile.getParents().next();

  var now = new Date();
  var formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HHmm');
  var folderName = 'Created_' + formattedDate;

  // Check if the folder already exists
  var folders = parentFolder.getFoldersByName(folderName);
  var dailyFolder;
  if (folders.hasNext()) {
    // Folder exists
    dailyFolder = folders.next();
  } else {
    // Folder does not exist, create it
    dailyFolder = parentFolder.createFolder(folderName);
    }

  //// ADD THE NEW JSONL FILE NAME HERE

  var files = dailyFolder.getFilesByName(generatedFileName); 
  
  if (!files.hasNext()) {
    Logger.log('File not found.');
    return;
  }

  var jsonlFile = files.next();
  var jsonlContent = jsonlFile.getBlob().getDataAsString().split('\n');

  // Filter out empty lines and parse JSON
  var jsonObjects = jsonlContent
    .filter(line => line.trim() !== '') // Remove empty lines
    .map(line => {
      try {
        return JSON.parse(line);
      } catch (e) {
        Logger.log('Error parsing line: ' + line);
        return null;
      }
    })
    .filter(obj => obj !== null); // Remove null entries from failed parsing

  // Get all custom_id values from column A
  var customIdsRange = sheet.getRange(4, 1, sheet.getLastRow() - 3, 1);
  var customIds = customIdsRange.getValues().flat();

  // Create a map for quick lookup of row index by custom_id
  var customIdToRowIndexMap = {};
  customIds.forEach((id, index) => {
    customIdToRowIndexMap[id] = index + 4; // +4 because sheet rows are 1-indexed and we start from row 4
  });

  jsonObjects.forEach(obj => {
    var customId = parseInt(obj.custom_id);
    var message = obj.response.body.choices[0].message.content;
    var rowIndex = customIdToRowIndexMap[customId];
    
    if (rowIndex) {
      sheet.getRange(rowIndex, 5).setValue(message); // Column E is the 5th column
    }
  });

  Logger.log('process completed.');
}


function translationChecking(){
  justBatch();
}
