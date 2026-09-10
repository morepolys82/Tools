function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    if (data.tasks && data.tasks.length > 0) {
      data.tasks.forEach(function(task) {
        var imageUrl = "";
        
        // 1. Process and upload image if present
        if (task.imageBase64) {
          try {
            var parts = task.imageBase64.split(",");
            var contentType = parts[0].match(/:(.*?);/)[1];
            var base64Data = parts[1];
            var decodedData = Utilities.base64Decode(base64Data);
            var blob = Utilities.newBlob(decodedData, contentType, "task_" + task.id + ".jpg");
            
            // Save file to Google Drive
            var file = DriveApp.createFile(blob);
            
            // Set public read permissions so Google Sheets can load the image
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            
            // Format URL using Google's direct image preview CDN
            imageUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
          } catch (imgError) {
            Logger.log("Image upload failed: " + imgError.toString());
          }
        }

        // 2. Append text fields (Columns A through E)
        var nextRow = sheet.getLastRow() + 1;
        sheet.getRange(nextRow, 1, 1, 5).setValues([[
          task.id,
          task.description,
          task.dateTime,
          task.status,
          new Date().toISOString()
        ]]);
        
        // 3. Insert Cell Image Thumbnail into Column F (Column 6)
        if (imageUrl !== "") {
          var imageCell = SpreadsheetApp.newCellImage()
            .setSourceUrl(imageUrl)
            .build();
          sheet.getRange(nextRow, 6).setValue(imageCell);
          sheet.setRowHeight(nextRow, 60); // Expand row height so the thumbnail fits neatly
        }
      });
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  var tasks = [];
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0]) {
      tasks.push({
        id: rows[i][0],
        description: rows[i][1],
        dateTime: rows[i][2],
        status: rows[i][3]
      });
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", tasks: tasks }))
    .setMimeType(ContentService.MimeType.JSON);
}
