function udonariumXmlDownload() {
  var html = HtmlService.createTemplateFromFile("dialog").evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "Now Downloading..");
}


function getData() {
  //データ読み込み
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var root = getXml_root();
  var data_character = getXml_character();
  var data_detail = getXml_detail();
  
  
  //offset,end計算
  var common_offset = getOffset_common(data);
  var common_end = getEnd_common(data);
  var info_offset = getOffset_info(data);
  var info_end = getEnd_info(data);
  
  //データ取り込み
  var data_common = getXml_data_common(data, common_offset, common_end);
  var data_info = getXml_data_info(data, info_offset, info_end);
  
  
  //XMLデータ積み込み
  data_character.addContent(data_common);
  data_detail.addContent(data_info);
  data_character.addContent(data_detail);
  root.addContent(data_character);

  //XMLテキスト出力
  var document = XmlService.createDocument(root);
  var xml = XmlService.getPrettyFormat().format(document);
  return xml
}

//-----------------------------------------------------------------------------
//root
function getXml_root(){
  var root = XmlService.createElement('character')
  .setAttribute("location.name", "table")
  .setAttribute("location.x","0")
  .setAttribute("location.y","0")
  .setAttribute("posZ","0")
  .setAttribute("rotate","0")
  .setAttribute("roll","0");
  return root;
}

//character
function getXml_character(){
  var data_character = XmlService.createElement('data')
  .setAttribute("name", "character");
  return data_character;
}

//基本情報
function getXml_data_common(data, offset){
  if(offset < 0){
    return null;
  }
  //基本情報
  // 名前
  var name = data[offset+0][1];
  // プレイヤー
  var player = data[offset+1][1];
  
  var data_common = XmlService.createElement('data')
  .setAttribute("name", "common");
  
  //name
  var data_name = XmlService.createElement('data')
  .setAttribute("name", "name");
  data_name.setText(name + " / " + player);
  data_common.addContent(data_name);

  //size
  var data_size = XmlService.createElement('data')
  .setAttribute("name", "size");
  data_size.setText("1");
  data_common.addContent(data_size);

  return data_common;
}

//character
function getXml_detail(){
  var data_detail = XmlService.createElement('data')
  .setAttribute("name", "detail");
  return data_detail;
}

//getXml
//情報
function getXml_data_common(data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }
  //情報
  // 名前
  var name = data[offset+0][1];
  // プレイヤー
  var player = data[offset+1][1];
  
  var data_common = XmlService.createElement('data')
  .setAttribute("name", "common");
  
  //name
  var data_name = XmlService.createElement('data')
  .setAttribute("name", "name");
  data_name.setText(name + " / " + player);
  data_common.addContent(data_name);

  //size
  var data_size = XmlService.createElement('data')
  .setAttribute("name", "size");
  data_size.setText("1");
  data_common.addContent(data_size);

  return data_common;
}

//情報
function getXml_data_info(data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }
  
  var data_info = XmlService.createElement('data')
  .setAttribute("name", "情報");
  //情報
  for(var i=offset;i<end;i++){
    var name = data[i][0];
    var value = data[i][1];
    if(!(name=="" || value=="")){
      var data_name = XmlService.createElement('data')
  .setAttribute("name", name);
      data_name.setText(value);
      data_info.addContent(data_name);
    }
  }
  
  return data_info;
}

//-----------------------------------------------------------------------------
//offset
//offset_common
function getOffset_common(data){
  for(var i=0;i<data.length;i++){
    var name = data[i][0];
    if(name == "基本情報"){
      return i+1;
    }
  }
  return -1;
}

//offset_info
function getOffset_info(data){
  for(var i=0;i<data.length;i++){
    var name = data[i][0];
    if(name == "基本情報"){
      return i+1;
    }
  }
  return -1;
}


//end
//end_common
function getEnd_common(data){
  for(var i=0;i<data.length;i++){
    var name = data[i][0];
    if(name == "能力値"){
      return i;
    }
  }
  return -1;
}

//end_info
function getEnd_info(data){
  for(var i=0;i<data.length;i++){
    var name = data[i][0];
    if(name == "能力値"){
      return i;
    }
  }
  return -1;
}

//end_info
function getEnd_info(data){
  for(var i=0;i<data.length;i++){
    var name = data[i][0];
    if(name == "能力値"){
      return i;
    }
  }
  return -1;
}

//-----------------------------------------------------------------------------
function getFileName() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getActiveSheet();
  var now = new Date();
  var datetime = Utilities.formatDate( now, 'Asia/Tokyo', 'yyyyMMddHHmm');
  return sheet.getName() + '_' + datetime + '.xml';
}
