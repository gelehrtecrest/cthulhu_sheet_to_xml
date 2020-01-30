var style_udonarium = 1;
var style_trpg_studio = 2;

//-----------------------------------------------------------------------------
function udonariumXmlDownload() {
  var html = HtmlService.createTemplateFromFile("udonarium_dialog").evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "Now Downloading a xml file..");
}

function trpgStudioJsonDownload() {
  var html = HtmlService.createTemplateFromFile("trpg_studio_dialog").evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "Now Downloading a json file..");
}

//-----------------------------------------------------------------------------

var chat = "";
var db = "";

function getDataUdonariumXml() {
  //データ読み込み
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var style = style_udonarium;
  return getData(style, data);
}

function getDataTrpgStudioJson() {
  //データ読み込み
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var style = style_trpg_studio;
  return getData(style, data);
}

//-----------------------------------------------------------------------------

function getData(style, data){
  var root = get_root(style);
  var data_character = get_character(style);
  var data_detail = get_detail(style);
  
  //チャットパレットの初期化
  init_chat();
  //offset,end計算
  var common_offset = getOffset(data, '基本情報');
  var common_end = getEnd(data, '能力値');
  var info_offset = getOffset(data, '基本情報');
  var info_end = getEnd(data, '能力値');
  var ability_offset = getOffset(data, '能力値');
  var ability_end = getEnd(data, '技能');
  var skill_offset = getOffset(data, '技能名');
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
  var data_image = get_data_image(style);
  var data_common = get_data_common(style, data, common_offset, common_end);
  var data_info = get_data_info(style, data, info_offset, info_end);
  var data_ability = get_data_ability(style, data, ability_offset, ability_end);
  var data_skill = get_data_skill(style, data, skill_offset, skill_end);
  var data_battle = get_data_battle(style, data, battle_offset, battle_end);
  var data_weapon = get_data_weapon(style, data, weapon_offset, weapon_end);
  var data_bs = get_data_bs(style, data, bs_offset, bs_end);
  var data_equip = get_data_equip(style, data, equip_offset, equip_end);
  var data_money = get_data_money(style, data, money_offset);
  var data_chat = get_data_chat(style);

  if(style == style_udonarium){
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
  } else if(style == style_trpg_studio){
    root["info"] = data_info;
    var array_forms = []
    array_forms.push(data_ability);
    array_forms.push(data_skill);
    array_forms.push(data_battle);
    array_forms.push(data_weapon);
    array_forms.push(data_bs);
    array_forms.push(data_equip);
    array_forms.push(data_money);
    root["array_forms"] = array_forms;
    json = JSON.stringify(root)
    return json;
  } else {
    return "";
  }
}

//-----------------------------------------------------------------------------
//root
function get_root(style){
  switch(style){
    case style_udonarium:
      var root = XmlService.createElement('character')
      .setAttribute("location.name", "table")
      .setAttribute("location.x","0")
      .setAttribute("location.y","0")
      .setAttribute("posZ","0")
      .setAttribute("rotate","0")
      .setAttribute("roll","0");
      return root;
    case style_trpg_studio:
      var root = {};
      return root;
    default:
      return null;
  }
}

//character
function get_character(style){
  var data_character = XmlService.createElement('data')
  .setAttribute("name", "character");
  return data_character;
}

//image
function get_data_image(style){
  var data_image = XmlService.createElement('data')
  .setAttribute("name", "image");
  var data_name = XmlService.createElement('data')
  .setAttribute("type", "image")
  .setAttribute("name", "imageIdentifier");
  data_image.addContent(data_name);
  return data_image;
}

//character
function get_detail(){
  var data_detail = XmlService.createElement('data')
  .setAttribute("name", "detail");
  return data_detail;
}

//getXml
//情報
function get_data_common(style, data, offset, end){
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
function get_data_info(style, data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }
  
  var data_info = null;
  switch(style){
    case style_udonarium:
      data_info = XmlService.createElement('data')
      .setAttribute("name", "情報");
      break;
    case style_trpg_studio:
      data_info = {};
      break;
  }

  //情報
  for(var i=offset;i<end;i++){
    var name = data[i][0];
    var value = data[i][1];
    if(!(name=="" || value=="")){
      switch(style){
        case style_udonarium:
          var data_name = XmlService.createElement('data')
          .setAttribute("name", name);
          data_name.setText(value);
          data_info.addContent(data_name);
          break;
        case style_trpg_studio:
          switch(name){
            case "名前":
              data_info["chara_name"] = value;
              break;
            case "年齢":
              data_info["age"] = value;
              break;
            case "性別":
              data_info["sex"] = value;
              break;
            case "職業":
              data_info["job"] = value;
              break;
            default:
              if(data_info["remarks"] == null){
                data_info["remarks"] = name + ":" + value;
              } else {
                data_info["remarks"] += "\n" + name + ":" + value;
              }
          }
        default:
          break;
      }
    }
  }
  return data_info;
}

//能力
function get_data_ability(style, data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }

  chat += '\n\n//-----能力\n'
  var data_ability = null;
  switch(style){
    case style_udonarium:
      data_ability = XmlService.createElement('data')
      .setAttribute("name", "リソース");
      break;
    case style_trpg_studio:
      data_ability = {};
      data_ability["type"] = "charaSheetInputCloneNumber";
      data_ability["title"] = "リソース";
      data_ability["forms"] = [];
      break;
  }

  //能力
  for(var i=offset;i<end;i++){
    var name = data[i][0];
    var value = data[i][4];
    if(name!=""){
      switch(style){
        case style_udonarium:
          var data_name = XmlService.createElement('data')
          .setAttribute("type", "numberResource")
          .setAttribute("currentValue", value)
          .setAttribute("name", name);
          data_name.setText(value);
          data_ability.addContent(data_name);
          if(!(name == "SAN値" ||
               name == "耐久力" ||
               name == "マジック・ポイント" ||
               name == "移動率")){
            if(name == "INT"){
              chat += '\nCC<={' + name + '}  :' + name + ' / アイデア';
            } else {
              chat += '\nCC<={' + name + '}  :' + name;
            }
          }
          break;
        case style_trpg_studio:
          var data_name = {};
          data_name["text"] = name;
          data_name["panel"] = false;
          data_name["number"] = value;
          data_ability["forms"].push(data_name);
          break;
      }
    }
  }
  return data_ability;
}

//技能
function get_data_skill(style, data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }

  chat += '\n\n//-----技能\n'
  var data_skill = null;
  switch(style){
    case style_udonarium:
      data_skill = XmlService.createElement('data')
      .setAttribute("name", "技能");
      break;
    case style_trpg_studio:
      data_skill = {};
      data_skill["type"] = "charaSheetInputCloneNumber";
      data_skill["title"] = "技能";
      data_skill["forms"] = [];
      break;
  }
  //技能
  for(var i=offset;i<end;i++){
    var name = data[i][0];
    var value = data[i][5];
    if(name!=""){
      switch(style){
        case style_udonarium:
          var data_name = XmlService.createElement('data')
          .setAttribute("name", name);
          data_name.setText(value);
          data_skill.addContent(data_name);
          chat += '\nCC<={' + name + '}  :' + name;
          break;
        case style_trpg_studio:
          var data_name = {};
          data_name["text"] = name;
          data_name["panel"] = false;
          data_name["number"] = value;
          data_skill["forms"].push(data_name);
          break;
      }
    }
  }
  
  return data_skill;
}

//戦闘
function get_data_battle(style, data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }
  
  var data_battle = null;
  switch(style){
    case style_udonarium:
      data_battle = XmlService.createElement('data')
      .setAttribute("name", "戦闘");
      break;
    case style_trpg_studio:
      data_battle = {};
      data_battle["type"] = "charaSheetInputCloneText";
      data_battle["title"] = "戦闘";
      data_battle["forms"] = [];
      break;
  }
  //戦闘
  for(var i=offset;i<end;i++){
    var name = data[i][0];
    var value = data[i][1];
    if(name!="" && name!="回避"){
      switch(style){
        case style_udonarium:
          var data_name = XmlService.createElement('data')
          .setAttribute("name", name);
          data_name.setText(value);
          data_battle.addContent(data_name);
          if(name="ダメージボーナス"){
            db = value;
          }
          break;
        case style_trpg_studio:
          var data_name = {};
          data_name["text1"] = name;
          data_name["panel"] = false;
          data_name["text2"] = value;
          data_battle["forms"].push(data_name);
          break;
      }
    }
  }
  
  return data_battle;
}

//武器
function get_data_weapon(style, data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }
  chat += '\n\n//-----武器威力判定\n'
  var data_weapon = null;
  switch(style){
    case style_udonarium:
      data_weapon = XmlService.createElement('data')
      .setAttribute("name", "武器");
      break;
    case style_trpg_studio:
      data_weapon = {};
      data_weapon["type"] = "charaSheetInputCloneText";
      data_weapon["title"] = "武器";
      data_weapon["forms"] = [];
      break;
  }
  //武器
  for(var i=offset;i<end;i++){
    var name = data[i][0];
    var value = data[i][4];
    if(name!=""){
      switch(style){
        case style_udonarium:
          var value_str = value.toString().replace('\+DB', db);
          var data_name = XmlService.createElement('data')
          .setAttribute("name", name);
          data_name.setText(value_str);
          data_weapon.addContent(data_name);
          chat += '\n' + value_str + ' :' + name;
          break;
        case style_trpg_studio:
          var data_name = {};
          data_name["text1"] = name;
          data_name["panel"] = false;
          data_name["text2"] = value;
          data_weapon["forms"].push(data_name);
          break;
      }
    }
  }
  
  return data_weapon;
}

//バックストーリー
function get_data_bs(style, data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }
  
  var data_bs = null;
  switch(style){
    case style_udonarium:
      data_bs = XmlService.createElement('data')
      .setAttribute("name", "バックストーリー");
      break;
    case style_trpg_studio:
      data_bs = {};
      data_bs["type"] = "charaSheetInputCloneText";
      data_bs["title"] = "バックストーリー";
      data_bs["forms"] = [];
      break;
  }
  //バックストーリー
  for(var i=offset;i<end;i++){
    var name = data[i][0];
    if(!(name=="")){
      var value = ""
      for(var j = 1;data[i][j] != "";j++){
        if(value == ""){
          value = data[i][j];
        } else {
          value = value + "\n" + data[i][j];
        }
      }
      switch(style){
        case style_udonarium:
          var data_name = XmlService.createElement('data')
          .setAttribute("type", "note")
          .setAttribute("name", name);
          data_name.setText(value);
          data_bs.addContent(data_name);
          break;
      case style_trpg_studio:
          var data_name = {};
          data_name["text1"] = name;
          data_name["panel"] = false;
          data_name["text2"] = value;
          data_bs["forms"].push(data_name);
          break;
      }
    }
  }
  
  return data_bs;
}

//装備品と所持品
function get_data_equip(style, data, offset, end){
  if(offset < 0 || end < 0){
    return null;
  }
  
  var data_equip = null;
  switch(style){
    case style_udonarium:
      data_equip = XmlService.createElement('data')
      .setAttribute("name", "装備品と所持品");
      break;
    case style_trpg_studio:
      data_equip = {};
      data_equip["type"] = "charaSheetInputCloneText";
      data_equip["title"] = "装備品と所持品";
      data_equip["forms"] = [];
      break;
  }
  //装備品と所持品
  for(var i=offset;i<end;i++){
    var name = data[i][0];
    var value = data[i][1];
    if(!(name=="")){
      switch(style){
        case style_udonarium:
          var data_name = XmlService.createElement('data')
          .setAttribute("type", "note")
          .setAttribute("name", name);
          data_name.setText(value);
          data_equip.addContent(data_name);
          break;
       case style_trpg_studio:
          var data_name = {};
          data_name["text1"] = name;
          data_name["panel"] = false;
          data_name["text2"] = value;
          data_equip["forms"].push(data_name);
          break;
      }
    }
  }
  
  return data_equip;
}

//収入と財産
function get_data_money(style, data, offset){
  if(offset < 0){
    return null;
  }
  
  var data_money = null;
  switch(style){
    case style_udonarium:
      data_money = XmlService.createElement('data')
      .setAttribute("name", "収入と財産");
      break;
    case style_trpg_studio:
      data_money = {};
      data_money["type"] = "charaSheetInputCloneText";
      data_money["title"] = "収入と財産";
      data_money["forms"] = [];
      break;
  }
  //収入と財産
  for(var i=offset;i<data.length;i++){
    var name = data[i][0];
    var value_d = data[i][1];
    var value_y = data[i][2];
    if(!(name=="")){
      switch(style){
        case style_udonarium:
          var data_name_d = XmlService.createElement('data')
          .setAttribute("name", name + ' (ドル)');
          data_name_d.setText(value_d);
          data_money.addContent(data_name_d);
          var data_name_y = XmlService.createElement('data')
          .setAttribute("name", name + ' (円)');
          data_name_y.setText(value_y);
          data_money.addContent(data_name_y);
          break;
        case style_trpg_studio:
          var data_name = {};
          data_name["text1"] = name + ' (ドル)';
          data_name["panel"] = false;
          data_name["text2"] = value_d;
          data_money["forms"].push(data_name);
          data_name["text1"] = name + ' (円)';
          data_name["panel"] = false;
          data_name["text2"] = value_y;
          data_money["forms"].push(data_name);
          break;
      }
    }
  }
  
  return data_money;
}

//チャットパレット
function get_data_chat(style){
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
  chat = "CC<={SAN}  :SANチェック \n現在SAN値 {SAN値} \n\n//-----神話技能\nCC<={クトゥルフ神話技能}  :クトゥルフ神話技能\n\n";
}

//-----------------------------------------------------------------------------
function getXmlFileName() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getActiveSheet();
  var now = new Date();
  var datetime = Utilities.formatDate( now, 'Asia/Tokyo', 'yyyyMMddHHmm');
  return sheet.getName() + '_' + datetime + '.xml';
}

function getJsonFileName() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getActiveSheet();
  var now = new Date();
  var datetime = Utilities.formatDate( now, 'Asia/Tokyo', 'yyyyMMddHHmm');
  return sheet.getName() + '_' + datetime + '.json';
}
//-----------------------------------------------------------------------------
