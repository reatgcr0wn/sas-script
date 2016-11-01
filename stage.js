/*
バンド名,出演日,出演ステージ,部,開始時刻,終了時刻,出演可能時間1日目,出演可能時間2日目,大学,団体,バンド読み,代表者,代表者カナ,代表者重複バンド,
メンバー1名前,メンバー1重複バンド,メンバー2名前,メンバー2重複バンド,....,メンバー9名前,メンバー9重複バンド,
電話番号,メール1,メール1確認,メール2,メール確認,曲名1,URL1,曲名2,URL2,申請時備考欄,バンド点,コメント

順で並んでおり、1行目に項目のヘッダーがあるスプレッドシートを操作します。
開始時刻,終了時刻のフォーマットは、Date型で扱える文字列です

setRowColor()は重いのでエラーの行だけに適用しよう
*/

function mainScript(){
  checkStageTime();
}

//本番の日付を設定
var real = {};
real.year = 2016 //年
real.month = 11 // 1月= 0,　12月　= 11
real.day1 = 3　//1日目
real.day2 = 4 //2日目

var mainobj = {};
//操作対象スプレッドシート
mainobj.url = 'https://docs.google.com/spreadsheets/d/19yegjkmq6B2nPubV6Z_PPA9sOLxIr-pB503oE1FNxN8/edit#gid=347892'
//取得するシート名
mainobj.sheetname = 'エントリー2016' 

mainobj.spreadsheet = SpreadsheetApp.openByUrl(mainobj.url);
mainobj.sheets = mainobj.spreadsheet.getSheets();
mainobj.sheet = mainobj.spreadsheet.getSheetByName(mainobj.sheetname)

function getSheetData(){
  
  //シートの最終行番号、最終列番号を取得
  var startrow = 1;
  var startcol = 1;
  var lastrow = mainobj.sheet.getLastRow();
  var lastcol = mainobj.sheet.getLastColumn();
  
  //全部配列にぶち込む
  var range = mainobj.sheet.getRange(startrow, startcol, lastrow, lastcol);
  var values = range.getValues();
  
  //例えばA1セルは[0][0]、B3セルは[2][1]
  //Logger.log(values[2][1]);
  return values
}

//ステージ時間が申請時間帯と一致するか確認する
function checkStageTime(){
  values = getSheetData();
  
  for (var i = 1; i < values.length; i++ ){
    
    var day = values[i][1];//出演日
    var part = values[i][3];//出演部
    var start = values[i][4];//ステージ開始時刻
    var end = values[i][5];//ステージ終了時刻
    var reqday1 = values[i][6];//1日目申請時間帯
    var reqday2 = values[i][7];//2日目申請時間帯
    
    if(day&&part&&start&&end){
      if(!isSuitToReqDay(day,part,start,end,reqday1,reqday2)){
        setRowColor(i+1,'#FA6367');
        Logger.log((i+1)+':'+values[i][0]+' 出演時間帯が申請と一致していません')
      }else{
        //setRowColor(i+1,'#FFFFFF');
      }
    }else{
      setRowColor(i+1,'#CFDDDE');
      Logger.log((i+1)+':'+values[i][0]+' 出演ステージデータ欄が入力されていないか、無効な値です。')
    }
  }

  Browser.msgBox("チェック終了");
}

function isSuitsToReqDay(day,part,start,end,reqday1,reqday2){
  start = new Date(start); 
  end = new Date(end);
  
  start.setFullYear(real.year,real.month);
  end.setFullYear(real.year,real.month);
  
  //1日目
  if (day == 1){
    start.setDate(real.day1);
    end.setDate(real.day1);
    
    switch (part){
      case '1部':
        if(reqday1.match(/1部/)) return true;
        break;
      case '2部':
        if(reqday1.match(/2部/)) return true;
        break;
      case '3部':
        if(reqday1.match(/3部/)) return true;
        break;
      case 'Final':
        if(reqday1.match(/3部/)) return true;
        break;
      case 'SSS':
        if(reqday1.match(/3部/)) return true;
        break;
      case '天望デッキ':
        if(end < new Date(real.year,real.month,real.day1,16,30)){//16:30より前は2部
          if(reqday1.match(/2部/)) return true;
        }else if(start >= new Date(real.year,real.month,real.day1,16,30)){
          if(reqday1.match(/3部/)) return true;
        }
      break;  
    }
  }else if (day==2){
    start.setDate(real.day2);
    end.setDate(real.day2);
    switch (part){
      case '1部':
        if(reqday2.match(/1部/)) return true;
        break;
      case '2部':
        if(reqday2.match(/2部/)) return true;
        break;
      case '3部':
        if(reqday2.match(/3部/)) return true;
        break;
      case 'Final':
        if(reqday2.match(/3部/)) return true;
        break;
      case 'SSS':
        if(reqday2.match(/3部/)) return true;
        break; 
      case '天望デッキ':
        if(end < new Date(real.year,real.month,real.day2,16,30)){//16:30より前は2部
          if(reqday2.match(/2部/)) return true;
        }else if(start >= new Date(real.year,real.month,real.day2,16,30)){
          if(reqday2.match(/3部/)) return true;
        }
      break; 
    }
  }
  
  return false;
}

//色変える
function setRowColor(row,color){
 mainobj.sheet.getRange(row, 1, 1,mainobj.sheet.getLastColumn()).setBackgroundColor(color);
}