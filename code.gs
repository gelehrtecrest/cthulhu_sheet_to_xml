function udonariumXmlDownload() {
  var html = HtmlService.createTemplateFromFile("dialog").evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "Now Downloading..");
}

var chat = "";
var db = "";

function getData() {
  //データ読み込み
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var root = getXml_root();
  var data_character = getXml_character();
  var data_detail = getXml_detail();
  
  //チャットパレットの初期化
  init_chat();
  //offset,end計算
  var common_offset = getOffset(data, '基本情報');
  var common_end = getEnd(data, '能力値');
  var info_offset = getOffset(data, '基本情報');
  var info_end = getEnd(data, '能力値');
  var ability_offset = getOffset(data, '能力値');
  var ability_end = getEnd(data, '技能名');
  var skill_offset = getOffset(data, '技能');
  var skill_end = getEnd(data, '戦闘');
  var battle_offset = getOffset(data, '戦闘');
  var battle_end = getEnd(data, '武器');
  var weapon_offset = getOffset(data, '武器');
  var weapon_end = getEnd(data, 'バックストーリー');
  var bs_offset = getOffset(data, 'バックストーリー');
  var bs_end = getEnd(data, '装備品と所持品');
  var equip_offset = getOffset(data, '装備品と所持品');
  var equip_end = getEnd(data, '収入と財産');
  var money_offset = getOffset(data, '収入と財産');

  //データ取り込み
  var data_image = getXml_data_image();
  var data_common = getXml_data_common(data, common_offset, common_end);
  var data_info = getXml_data_info(data, info_offset, info_end);
  var data_ability = getXml_data_ability(data, ability_offset, ability_end);
  var data_skill = getXml_data_skill(data, skill_offset, skill_end);
  var data_battle = getXml_data_battle(data, battle_offset, battle_end);
  var data_weapon = getXml_data_weapon(data, weapon_offset, weapon_end);
  var data_bs = getXml_data_bs(data, bs_offset, bs_end);
  var data_equip = getXml_data_equip(data, equip_offset, equip_end);
  var data_money = getXml_data_money(data, money_offset);
  var data_chat = getXml_data_chat();

  //XMLデータ積み込み
  data_character.addContent(data_image);
  data_character.addContent(data_common);
  data_detail.addContent(data_info);
  data_detail.addContent(data_ability);
  data_detail.addContent(data_skill);
  data_detail.addContent(data_battle);
  data_detail.addContent(data_weapon);
  data_detail.addContent(data_bs);
  data_detail.addContent(data_equip);
  data_detail.addContent(data_money);
  data_character.addContent(data_detail);
  root.addContent(data_character);
  root.addContent(data_chat);

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

//image
function getXml_data_image(){
  var data_image = XmlService.createElement('data')
  .setAttribute("name", "image");
  var data_name = XmlService.createElement('data')
  .setAttribute("type", "image")
  .setAttribute("name", "imageIdentifier");
  data_image.addContent(data_name);
  return data_image;
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

//能力
function getXml_data_ability(data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }
  
  var data_ability = XmlService.createElement('data')
  .setAttribute("name", "リソース");
  //能力
  for(var i=offset;i<end;i++){
    var name = data[i][0];
    var value = data[i][4];
    if(name!=""){
      var data_name = XmlService.createElement('data')
      .setAttribute("type", "numberResource")
      .setAttribute("currentValue", value)
      .setAttribute("name", name);
      data_name.setText(value);
      data_ability.addContent(data_name);
    }
  }
  
  return data_ability;
}

//技能
function getXml_data_skill(data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }
  
  var data_skill = XmlService.createElement('data')
  .setAttribute("name", "戦闘技能");
  //技能
  for(var i=offset;i<end;i++){
    var name = data[i][0];
    var value = data[i][5];
    if(name!=""){
      var data_name = XmlService.createElement('data')
      .setAttribute("name", name);
      data_name.setText(value);
      data_skill.addContent(data_name);
      chat += '\nCC<={' + name + '}  :' + name;
    }
  }
  
  return data_skill;
}

//戦闘
function getXml_data_battle(data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }
  
  var data_battle = XmlService.createElement('data')
  .setAttribute("name", "戦闘");
  //戦闘
  for(var i=offset;i<end;i++){
    var name = data[i][0];
    var value = data[i][1];
    if(name!="" && name!="回避"){
      var data_name = XmlService.createElement('data')
      .setAttribute("name", name);
      data_name.setText(value);
      data_battle.addContent(data_name);
      if(name="ダメージボーナス"){
        db = value;
      }
    }
  }
  
  return data_battle;
}

//武器
function getXml_data_weapon(data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }
  chat += '\n\n//-----武器威力判定\n'
  var data_weapon = XmlService.createElement('data')
  .setAttribute("name", "武器");
  //武器
  for(var i=offset;i<end;i++){
    var name = data[i][0];
    var value = data[i][4];
    if(name!=""){
      var value_str = value.toString().replace('\+DB', db);
      var data_name = XmlService.createElement('data')
      .setAttribute("name", name);
      data_name.setText(value_str);
      data_weapon.addContent(data_name);
      chat += '\n' + value_str + ' :' + name;
    }
  }
  
  return data_weapon;
}

//バックストーリー
function getXml_data_bs(data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }
  
  var data_bs = XmlService.createElement('data')
  .setAttribute("name", "バックストーリー");
  //バックストーリー
  for(var i=offset;i<end;i++){
    var name = data[i][0];
    if(!(name=="")){
      var data_name = XmlService.createElement('data')
      .setAttribute("type", "note")
      .setAttribute("name", name);
      var value = ""
      for(var j = 1;data[i][j] != "";j++){
        if(value == ""){
          value = data[i][j];
        } else {
          value = value + "\n" + data[i][j];
        }
      }
      data_name.setText(value);
      data_bs.addContent(data_name);
    }
  }
  
  return data_bs;
}

//装備品と所持品
function getXml_data_equip(data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }
  
  var data_equip = XmlService.createElement('data')
  .setAttribute("name", "装備品と所持品");
  //装備品と所持品
  for(var i=offset;i<end;i++){
    var name = data[i][0];
    var value = data[i][1];
    if(!(name=="")){
      var data_name = XmlService.createElement('data')
      .setAttribute("type", "note")
      .setAttribute("name", name);
      data_name.setText(value);
      data_equip.addContent(data_name);
    }
  }
  
  return data_equip;
}

//収入と財産
function getXml_data_money(data, offset){
  if(offset < 0){
    return null;
  }
  
  var data_money = XmlService.createElement('data')
  .setAttribute("name", "収入と財産");
  //収入と財産
  for(var i=offset;i<data.length;i++){
    var name = data[i][0];
    var value_d = data[i][1];
    var value_y = data[i][2];
    if(!(name=="")){
      var data_name_d = XmlService.createElement('data')
      .setAttribute("name", name + ' (ドル)');
      data_name_d.setText(value_d);
      data_money.addContent(data_name_d);
      var data_name_y = XmlService.createElement('data')
      .setAttribute("name", name + ' (円)');
      data_name_y.setText(value_y);
      data_money.addContent(data_name_y);
    }
  }
  
  return data_money;
}

//チャットパレット
function getXml_data_chat(){
  var data_chat = XmlService.createElement('chat-palette')
  .setAttribute("dicebot", "Cthulhu7th");
  data_chat.setText(chat);
  return data_chat;
}

//-----------------------------------------------------------------------------
//offset
function getOffset(data, key){
  for(var i=0;i<data.length;i++){
    var name = data[i][0];
    if(name == key){
      return i+1;
    }
  }
  return -1;
}

//end
function getEnd(data, key){
  for(var i=0;i<data.length;i++){
    var name = data[i][0];
    if(name == key){
      return i;
    }
  }
  return -1;
}
//-----------------------------------------------------------------------------
function init_chat(){
  chat = "CC<={SAN}  :SANチェック \n現在SAN値 {SAN値} \n\n//-----神話技能\nCC<={クトゥルフ神話技能}  :クトゥルフ神話技能\n\n//-----技能\n";
}

//-----------------------------------------------------------------------------
function getFileName() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getActiveSheet();
  var now = new Date();
  var datetime = Utilities.formatDate( now, 'Asia/Tokyo', 'yyyyMMddHHmm');
  return sheet.getName() + '_' + datetime + '.xml';
}
