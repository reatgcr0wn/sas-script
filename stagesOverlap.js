/*
記載されているステージが2つ以上存在していないかチェックします
mainScript3()
で実行します。

カラムが、
バンド名,出演日,出演ステージ,部,開始時刻,終了時刻,出演可能時間1日目,出演可能時間2日目,大学,団体,バンド読み,代表者,代表者カナ,代表者重複バンド,
メンバー1名前,メンバー1重複バンド,メンバー2名前,メンバー2重複バンド,....,メンバー9名前,メンバー9重複バンド,
電話番号,メール1,メール1確認,メール2,メール確認,曲名1,URL1,曲名2,URL2,申請時備考欄,バンド点,コメント
順で並んでおり、1行目に項目のヘッダーがあるスプレッドシートを操作します。
開始時刻,終了時刻のフォーマットは、Date型で扱える文字列です

setRowColor()は重いのでエラーの行だけに適用しよう
*/
function mainScript3(){
  checkStageOverlap();
}

//本番の日付を設定
var real = {};
real.year = 2016 //年
real.month = 11 // 1月= 0,　12月　= 11
real.day1 = 3　//1日目
real.day2 = 4 //2日目

var idletime　= 45//何分未満を警告対象とするか

var mainobj = {};
//操作対象スプレッドシート
mainobj.url = 'https://docs.google.com/spreadsheets/d/11F0_XZ5CB2zBg6U9iC8SJNK4rN0s-CeN-0k59icTjyo/edit#gid=13218088096=347892'
//取得するシート名
mainobj.sheetname = 'エントリー2016' 

mainobj.spreadsheet = SpreadsheetApp.openByUrl(mainobj.url);
mainobj.sheets = mainobj.spreadsheet.getSheets();
mainobj.sheet = mainobj.spreadsheet.getSheetByName(mainobj.sheetname)

function getSheetData3(sheet){
  
  //シートの最終行番号、最終列番号を取得
  var startrow = 1;
  var startcol = 1;
  var lastrow = sheet.getLastRow();
  var lastcol = sheet.getLastColumn();
  
  //全部配列にぶち込む
  var range = sheet.getRange(startrow, startcol, lastrow, lastcol);
  var values = range.getValues();
  
  //例えばA1セルは[0][0]、B3セルは[2][1]
  //Logger.log(values[2][1]);
  return values
}

//ステージ重複チェック
function checkStageOverlap(){
  var values = getSheetData3(mainobj.sheet);
  
  //for (var i = 1; i < values.length; i++ ){
  for (var i = 1; i < 443; i++ ){  
    var day = values[i][1];//出演日
    var place = values[i][2];
    var part = values[i][3];//出演部
    var start = values[i][4];//ステージ開始時刻
    var end = values[i][5];//ステージ終了時刻
    
    if(day&&part&&start&&end){
      for(var j = i+1;j<443;j++){
        var another_row = values[j];
        if(day.toString()==another_row[1].toString()&&place.toString()==another_row[2].toString()&&start.toString()==another_row[4].toString()&&end.toString()==another_row[5].toString()&&j!=i){
          Logger.log((i+1)+':'+values[i][0]+'と'+(j+1)+':'+another_row[0]+'のステージが一致');
        }
      }
    }
      
  }

}
