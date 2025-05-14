function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastText = lastRow >= 1 ? sheet.getRange(lastRow, 2).getValue() : "";
  return ContentService.createTextOutput(
    JSON.stringify({
      clipboard: lastText,
      rowCount: lastRow
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var command = e.parameter.command;
  var clipboardText = e.parameter.clipboard;
  var timestamp = new Date();

  if (command === "CLEAR_CLIPBOARD") {
    sheet.clear();  // 清空整張表
    return ContentService.createTextOutput("CLEARED");
  }

  // 寫入剪貼簿內容
  if (clipboardText !== undefined) {
    sheet.appendRow([timestamp, clipboardText]);
    return ContentService.createTextOutput("OK");
  }

  return ContentService.createTextOutput("NO_ACTION");
}
