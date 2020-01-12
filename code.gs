function udonariumXmlDownload() {
  var html = HtmlService.createTemplateFromFile("dialog").evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "Now Downloading..");
}


function getData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  //基本情報
  // 名前
  var name = data[4][1];
  // プレイヤー
  var player = data[5][1];
  // 職業
  
  

  var xml = '<test>' + name + ' / ' + player + '</test>';
  return xml
}

function getFileName() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getActiveSheet();
  var now = new Date();
  var datetime = Utilities.formatDate( now, 'Asia/Tokyo', 'yyyyMMddHHmm');
  return sheet.getName() + '_' + datetime + '.xml';
}
