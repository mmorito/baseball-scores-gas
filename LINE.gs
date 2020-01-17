//CHANNEL_ACCESS_TOKENを設定
//LINE developerで登録をした、CHANNEL_ACCESS_TOKENを入力する
var CHANNEL_ACCESS_TOKEN = "<CHANNEL_ACCESS_TOKEN>"; 
var line_endpoint = "https://api.line.me/v2/bot/message/reply";
var SPREAD_ID = "1rXt5ZpL7rW0YtBgOsr0I1iuRA1R5BB8Pk6cUJDzPK6Q"

//SpreadSheetの取得
var SS = SpreadsheetApp.openById(SPREAD_ID);
var sheet = SS.getSheetByName("サマリ"); //Spreadsheetのシート名（タブ名）
var members = [
  {key: ['ほげ太郎1', 'ほげ1'], value: 'ほげ太郎1'},
  {key: ['ほげ太郎2', 'ほげ2'], value: 'ほげ太郎2'},
  {key: ['ほげ太郎3', 'ほげ3'], value: 'ほげ太郎3'},
  {key: ['ほげ太郎4', 'ほげ4'], value: 'ほげ太郎4'},
  {key: ['ほげ太郎5', 'ほげ5'], value: 'ほげ太郎5'},
  {key: ['ほげ太郎6', 'ほげ6'], value: 'ほげ太郎6'},
  {key: ['ほげ太郎7', 'ほげ7'], value: 'ほげ太郎7'},
  {key: ['ほげ太郎8', 'ほげ8'], value: 'ほげ太郎8'},
  {key: ['ほげ太郎9', 'ほげ9'], value: 'ほげ太郎9'},
  {key: ['ほげ太郎10', 'ほげ10'], value: 'ほげ太郎10'},
  {key: ['ほげ太郎11', 'ほげ11'], value: 'ほげ太郎11'},
  {key: ['ほげ太郎12', 'ほげ12'], value: 'ほげ太郎12'},
  {key: ['ほげ太郎13', 'ほげ13'], value: 'ほげ太郎13'},
  {key: ['ほげ太郎14', 'ほげ14'], value: 'ほげ太郎14'},
  {key: ['ほげ太郎15', 'ほげ15'], value: 'ほげ太郎15'},
  {key: ['ほげ太郎16', 'ほげ16'], value: 'ほげ太郎16'},
  {key: ['ほげ太郎17', 'ほげ17'], value: 'ほげ太郎17'},
  {key: ['ほげ太郎18', 'ほげ18'], value: 'ほげ太郎18'}
];

// メイン処理
// POSTデータ取得、JSONをパースする
function doPost(e) {
  // 受け取ったメッセージを解析する
  var res = receiveMessage(e);
  
  var user_message = parseUserMessage(res.user_message);
  
  // messageのキーワードを拾って返信するメッセージを決める
  var reply_messages = chooseResponseMessage(user_message);
  
  // messageを返信する
  sendMessage(res.reply_token, [reply_messages]);
}

// messageのキーワードを拾って返信するメッセージを決める
function chooseResponseMessage(user_message) {
  var reply_messages;
  if (user_message.indexOf('成績') >= 0) {
    for (var i = 0; i < members.length; i++) {
      for (var j = 0; j < members[i].key.length; j++) {
        if (user_message.indexOf(members[i].key[j]) >= 0) {
          reply_messages = getIndivisualScore(members[i].value);
          return reply_messages;
        }
      }
    }
  } else if (user_message.indexOf('試合数') >= 0) {
    return '2019年の試合数は16です';
  } else if (user_message === 'うるさい') {
    return 'ごめんなさい';
  } else if (user_message.indexOf('らびっつ') >= 0) {
    return 'https://teams.one/teams/h-reddogs';
  }
}

function getIndivisualScore(name) {
  // [[[],[],[]...],[[],[],[]...]]
  var val = sheet.getRange('B8:V37').getValues();
  var title1 = sheet.getRange('Y8:AA17').getValues();
  var title2 = sheet.getRange('Y41:AA45').getValues();
  for (var i = 0; i < val.length; i++) {
    if (val[i][0] === name) {
      var res = '2019年 ' + name + 'さんの成績.\n';
      if (!isEmpty(val[i][1])) { res += '打席:　 ' + empty(val[i][1]) + '\n'; }
      if (!isEmpty(val[i][3])) { res += '得点: 　' + empty(val[i][3]) + '\n'; }
      if (!isEmpty(val[i][2])) { res += '打数: 　' + empty(val[i][2]) + '\n'; }
      if (!isEmpty(val[i][5])) { res += '二塁打: ' + empty(val[i][5]) + '\n'; }
      if (!isEmpty(val[i][4])) { res += '安打: 　' + empty(val[i][4]) + '\n'; }
      if (!isEmpty(val[i][7])) { res += '本塁打: ' + empty(val[i][7]) + '\n'; }
      if (!isEmpty(val[i][6])) { res += '三塁打: ' + empty(val[i][6]) + '\n'; }
      if (!isEmpty(val[i][9])) { res += '盗塁: 　' + empty(val[i][9]) + '\n'; }
      if (!isEmpty(val[i][8])) { res += '打点: 　' + empty(val[i][8]) + '\n'; }
      if (!isEmpty(val[i][11])) { res += '犠飛: 　' + empty(val[i][11]) + '\n'; }
      if (!isEmpty(val[i][10])) { res += '犠打: 　' + empty(val[i][10]) + '\n'; }
      if (!isEmpty(val[i][13])) { res += '死球: 　' + empty(val[i][13]) + '\n'; }
      if (!isEmpty(val[i][12])) { res += '四球: 　' + empty(val[i][12]) + '\n'; }
      if (!isEmpty(val[i][15])) { res += '失策: 　' + empty(val[i][15]) + '\n'; }
      if (!isEmpty(val[i][14])) { res += '三振: 　' + empty(val[i][14]) + '\n'; }
      if (!isEmpty(val[i][17])) { res += '打率: 　' + Math.round(val[i][17] * 1000) / 1000 + '\n'; }
      if (!isEmpty(val[i][16])) { res += '塁打数: ' + Math.round(val[i][16] * 1000) / 1000 + '\n'; }
      if (!isEmpty(val[i][19])) { res += '出塁率: ' + Math.round(val[i][19] * 1000) / 1000 + '\n'; }
      if (!isEmpty(val[i][18])) { res += '長打率: ' + Math.round(val[i][18] * 1000) / 1000 + '\n'; }

      res += '\n取得タイトル(打者)\n';
      var title1Result = '';
      for (var j = 0; j < title1.length; j++) {
        if (title1[j][1] === name) {
          title1Result += title1[j][0] + '\n';
        }
      }
      if (!title1Result) { title1Result = 'なし\n'; }
      res += title1Result;
      
      res += '\n取得タイトル(投手)\n';
      var title2Result = '';
      for (var j = 0; j < title2.length; j++) {
        if (title2[j][1] === name) {
          title2Result += title2[j][0] + '\n';
        }
      }
      if (!title2Result) { title2Result = 'なし\n'; }
      res += title2Result;
      
      return res;
    }
  }
}

function empty(val) {
  return val || 0;
}

function isEmpty(val) {
  return !val || val === 0;
}

function parseUserMessage(val) {
  if (!val) return val;
  return val.replace(/　/g," ");
}

// 受け取ったメッセージを解析する
function receiveMessage(data) {
  var json = JSON.parse(data.postData.contents);

  //返信するためのトークン取得
  var reply_token= json.events[0].replyToken;
  if (typeof reply_token === 'undefined') {
    return;
  }

  //送られたLINEメッセージを取得
  var user_message = json.events[0].message.text;
  
  return {
    reply_token: reply_token,
    user_message: user_message
  };
}

// messageを返信する
function sendMessage(reply_token, reply_messages) {
  if (!reply_messages || reply_messages.length === 0 || !reply_messages[0]) return;
  // メッセージを返信
  var messages = reply_messages.map(function (v) {
    return {'type': 'text', 'text': v};    
  });    
  UrlFetchApp.fetch(line_endpoint, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': reply_token,
      'messages': messages,
    }),
  });
  return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
}

// 動作確認用
function test() {
  var name = '成績 せっきん';
  var aaa = parseUserMessage(name);
  var t = chooseResponseMessage(aaa);
  Logger.log(t);
}