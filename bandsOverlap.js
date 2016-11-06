/*
バンドメンバーが、一定時間以内に他のバンドでの出演がないかをチェックします。
mainScript2()
で実行します。

カラムが、
バンド名,出演日,出演ステージ,部,開始時刻,終了時刻,出演可能時間1日目,出演可能時間2日目,大学,団体,バンド読み,代表者,代表者カナ,代表者重複バンド,
メンバー1名前,メンバー1重複バンド,メンバー2名前,メンバー2重複バンド,....,メンバー9名前,メンバー9重複バンド,
電話番号,メール1,メール1確認,メール2,メール確認,曲名1,URL1,曲名2,URL2,申請時備考欄,バンド点,コメント
順で並んでおり、1行目に項目のヘッダーがあるスプレッドシートを操作します。
開始時刻,終了時刻のフォーマットは、Date型で扱える文字列です

setRowColor()は重いのでエラーの行だけに適用しよう
*/
function mainScript2(){
  check();
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
mainobj.url = 'https://docs.google.com/spreadsheets/d/11F0_XZ5CB2zBg6U9iC8SJNK4rN0s-CeN-0k59icTjyo/edit#gid=1321808809d=347892'
//取得するシート名
mainobj.sheetname = 'エントリー2016' 

mainobj.spreadsheet = SpreadsheetApp.openByUrl(mainobj.url);
mainobj.sheets = mainobj.spreadsheet.getSheets();
mainobj.sheet = mainobj.spreadsheet.getSheetByName(mainobj.sheetname)

function getSheetData2(sheet){
  
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

//メンバーが同時間帯に出演していないかチェックする
function check(){
  var values = getSheetData2(mainobj.sheet);
  
  //for (var i = 1; i < values.length; i++ ){
  for (var i = 1; i < 443; i++ ){  
    var day = values[i][1];//出演日
    var part = values[i][3];//出演部
    var start = values[i][4];//ステージ開始時刻
    var end = values[i][5];//ステージ終了時刻
    var reqday1 = values[i][6];//1日目申請時間帯
    var reqday2 = values[i][7];//2日目申請時間帯
    
    if(day&&part&&start&&end){
      if(getOverlappedMembers(i,day,part,start,end,values).length>0){
        setRowColor(i+1,'#FFE1AE');
        //Logger.log((i+1)+':'+values[i][0]+' 30分以内に出演時間帯が重複するメンバーがいます')
      }else{
        //setRowColor(i+1,'#FFFFFF');
      }
    }else{
      //setRowColor(i+1,'#CFDDDE');
      Logger.log((i+1)+':'+values[i][0]+' 出演ステージデータ欄が入力されていないか、無効な値です。')
    }
  }

}

function getOverlappedMembers(bandindex,day,part,start,end,values){
var overlap = [];
var members = [];
var another_members = [];
//  for (var i = 1; i < values.length; i++ ){
  for (var i = 1; i < 443; i++ ){
    if(i!=bandindex){
      if(values[i][0]&&values[i][1]&&values[i][3]&&values[i][4]){
        if(values[i][1] == values[bandindex][1]){//dayをチェック
          
          //dayが同じ→時間だけ比べればよいので、
          //年月日は無視してDateオブジェクトを生成
          var band_start = +new Date(start); 
          var another_start = +new Date(values[i][4]);
          var diff_min = (another_start - band_start)/1000/60;
          
          if( diff_min >= 0&&diff_min < idletime){//時間が0分以上30分未満
            
            //メンバーチェック
            var band_row = values[bandindex];//何回も使うのでキャッシュ
            var another_row = values[i];
            //メンバー追加
            members.length = 0;
            members.push(trimText(band_row[11]),trimText(band_row[14]),trimText(band_row[16]),trimText(band_row[18]),trimText(band_row[20]),trimText(band_row[22]),trimText(band_row[24]));
            members = members.filter(function(e){return e !== "";});//空を削除
            
            another_members.length = 0;
            another_members.push(trimText(another_row[11]),trimText(another_row[14]),trimText(another_row[16]),trimText(another_row[18]),trimText(another_row[20]),trimText(another_row[22]),trimText(another_row[24]));
            another_members = another_members.filter(function(e){return e !== "";});

            for (var j in another_members){
              if(members.indexOf(another_members[j]) !== -1){//メンバー被ってたら
                if(overlap.indexOf(another_members[j]) === -1){//被ってる人リストになければ追加
                  Logger.log((bandindex+1)+':'+band_row[0]+' の '+another_members[j]+' さんが '+diff_min+'分後のバンド '+(i+1)+':'+another_row[0]+' と重複しています。');
                  overlap.push(another_members[j]);
                }
              }
            }
          }
        }
       }
    }
  }
  return overlap;
}

//空白を取り除くよ
function　trimText(txt){
txt = txt.toString();
 if(txt!=''){
  var nbsp = String.fromCharCode(160);//&nbsp;
  txt = txt.split(nbsp).join('　');
  txt = txt.replace(' ','');
  txt = txt.replace('　','');
  txt = txt.replace(/\r/g,'');
  txt = txt.replace(/\n/g,'');
  txt = txt.replace(/^\t/g,'');
  txt = txt.replace(/\t$/g,'');
  txt = txt.replace(/^ /g,'');
  txt = txt.replace(/ $/g,'');
  }
  return txt;
};
