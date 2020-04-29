var NUMBER_OF_MATCHS = 5; // 試合数

function onOpen(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var myMenuEntries = [];
  myMenuEntries.push({name: "成績集計", functionName: "calcPersonalScore"});
  ss.addMenu("コマンド", myMenuEntries);
}

function calcPersonalScore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sum = ss.getSheetByName('サマリ');
  clearSummary(sum);
  var cnt = {'fielder': 0, 'pitcher': 0};
  var members = getTeamMembers(ss);
  // 正規メンバーをループ
  for (var j = 0; j < members.length; j++) {
    var name = members[j].name;
    var fielderScore = [[]];
    var pitcherScore = [[]];
    Logger.log(name);
    // 全試合ループ
    for (var i = 1; i <= NUMBER_OF_MATCHS; i++) {
      var sn = i + "試合目";
      var s = ss.getSheetByName(sn);
      // 1試合分のスコア取得
      var fieldersScore = s.getRange(12, 2, 20, 16).getValues();
      var pitchersScore = s.getRange(42, 2, 10, 20).getValues();
      // 野手の成績を集計
      for (var k = 0; k < fieldersScore.length; k++) {
        if (name === fieldersScore[k][0]) {
          for (var l = 0; l <= 14; l++) {
            fielderScore[0][l] = fielderScore[0][l] || 0;
            if (fieldersScore[k][l + 1]) { fielderScore[0][l] += fieldersScore[k][l + 1]; }
            if (fielderScore[0][l] === 0) { fielderScore[0][l] = ''; }
          }
          break;
        }
      }
      // 投手の成績を集計
      for (var k = 0; k < pitchersScore.length; k++) {
        if (name === pitchersScore[k][0]) {
          for (var l = 0; l <= 18; l++) {
            pitcherScore[0][l] = pitcherScore[0][l] || 0;
            if (pitchersScore[k][l + 1]) { pitcherScore[0][l] += pitchersScore[k][l + 1]; }
            if (pitcherScore[0][l] === 0) { pitcherScore[0][l] = ''; }
          }
          break;
        }
      }
    }
    // 画面描画
    if (fielderScore[0].length > 0) {
      sum.getRange(8 + cnt.fielder, 2).setValue(name);
      sum.getRange(8 + cnt.fielder, 3, 1, fielderScore[0].length).setValues(fielderScore);
      cnt.fielder++;
    } 
    if (pitcherScore[0].length > 0) {
      sum.getRange(42 + cnt.pitcher, 2).setValue(name);
      sum.getRange(42 + cnt.pitcher, 3, 1, pitcherScore[0].length).setValues(pitcherScore);
      cnt.pitcher++;
    }
  }
}

function clearSummary(sum) {
  sum.getRange(8, 2, 30, 16).clear();
  sum.getRange(42, 2, 30, 20).clear();
}

function getTeamMembers(ss) {
  const dev = ss.getSheetByName("work").getDataRange().getValues();
  var members = [];
  for (var i = 0; i < dev.length; i++) {
    members.push({'name': dev[i][0], 'fielder': [], 'pitcher': []});
  }
  Logger.log(members);
  return members;
}


/**
 * OPSを自動計算するよ！
 * @param {float}   strokeRate  長打率
 * @param {float}   birthRate   出塁率
 * @return {string}             OPS(https://calculator.jp/science/ops)
**/
function calcOPS(strokeRate, birthRate) {
  var ops = strokeRate + birthRate
  var res = "G"
  if (ops >= 0.9000) {
    res = "A";
  } else if (ops >= 0.8334 && ops <= 0.8999) {
    res = "B";
  } else if (ops >= 0.7667 && ops <= 0.8333) {
    res = "C";
  } else if (ops >= 0.7000 && ops <= 0.7666) {
    res = "D";
  } else if (ops >= 0.6334 && ops <= 0.6999) {
    res = "E";
  } else if (ops >= 0.5667 && ops <= 0.6333) {
    res = "F";
  }
  return res;
}
