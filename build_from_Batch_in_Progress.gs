function useBatchIDToBuildQCodeFrame() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TranslationChecking');
  
    // Get the current spreadsheet file and its parent folder
  var spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getName().replace(/\s+/g, '_');
  var parentFolder = spreadsheetFile.getParents().next();
  const GPT_API = "ADD GPT API KEY HERE";

  //We need to download the file to the right folder 

  ///CHECK IF BATCH IS FINISHED

  // Check if there is text in M4
  var batchId = sheet.getRange('M4').getValue();
  if (!batchId || batchId.trim() === '') {
    SpreadsheetApp.getUi().alert('Error: There is no Batch ID in M4');
    throw new Error('Error: There is no Batch ID in M4');
  }

  // Check the batch status and download the output file once it's ready
  while (true) {
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
   
    // Check if the batch is processed
    if (statusResponseData.status === 'completed'  || statusResponseData.status === 'cancelled') {
      const outputFileId = statusResponseData.output_file_id;
      
      // Create the options for the UrlFetchApp request to download the file content
      const downloadOptions = {
        method: 'get',
        headers: {
          'Authorization': 'Bearer ' + GPT_API
        }
       }
      
      

      // Make the request to the OpenAI API to download the file content
      const downloadResponse = UrlFetchApp.fetch('https://api.openai.com/v1/files/' + outputFileId + '/content', downloadOptions);
      const outputBlob = downloadResponse.getBlob();



      // Define dailyFolder

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
      //////////


      // Save the file to Google Drive and store the file name in the global variable
      generatedFileName = `${sheetName}_${formattedDate}_${formattedTime}_OUTPUT.jsonl`;
      dailyFolder.createFile(outputBlob).setName(generatedFileName);

      // Log the success message
     Logger.log('Output file downloaded and saved to Google Drive');
     Logger.log('Generated file name: ' + generatedFileName);
     break;
    } else {
         SpreadsheetApp.getUi().alert('The batch process is still running\n\nPlease check again later');
         throw new Error('The batch process is still running\n\nPlease check again later');
      }
    
  }

  //// ADD THE NEW JSONL FILE NAME HERE

  var files = dailyFolder.getFilesByName(generatedFileName); 
  
  if (!files.hasNext()) {
    SpreadsheetApp.getUi().alert('Error: File with this name is not saved in todays folder');
    throw new Error('Error: File with this name is not saved in todays folder');
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
}
