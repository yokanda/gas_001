function main(){

// スプレッドシートIDとシート名を入力
  var ss = SpreadsheetApp.openById('1Hd5yZ0_0y2HjKhXftYIuILhAEKBSDffft5OROzYFsHQ');
  var ss2 = SpreadsheetApp.openById('1Hd5yZ0_0y2HjKhXftYIuILhAEKBSDffft5OROzYFsHQ');
   
  var sh1 = ss.getSheetByName('check');
  var sh2 = ss.getSheetByName('check 印刷用')
  var sh3 = ss2.getSheetByName('一覧表')


// 利用者、契約日、営業担当、取引を格納する空の変数を作成
// 最終行の番号を取得する
  var riyosha = sh1.getRange(6,3).getValue();
  var kdate = sh1.getRange(4,3).getValue();
  var kyotaku = sh1.getRange(8,3).getValue();
  var caremane = sh1.getRange(10,3).getValue();
  var eigyo   = sh1.getRange(12,3).getValue();
  var tori    = sh1.getRange(14,3).getValue();
  var kingaku    = sh1.getRange(16,3).getValue();  
  var tekiyo    = sh1.getRange(18,3).getValue();  
  var sakusei    = sh1.getRange(20,3).getValue(); 

//印刷シート転記処理
  sh2.getRange(4,3).setValue(riyosha);
  sh2.getRange(6,3).setValue(kdate);
  sh2.getRange(8,3).setValue(eigyo);
  sh2.getRange(10,3).setValue(tori);
  sh2.getRange(12,3).setValue(kingaku); 
  sh2.getRange(14,3).setValue(tekiyo); 
  sh2.getRange(16,3).setValue(sakusei);
    
// 必須入力チェック
   if (riyosha == "") {
        Browser.msgBox("エラー：利用者は必須入力項目です！");
        var range=sh1.getRange(6,3) ;
        range.setBackground("yellow");
        return;
   　} else   {
        var range=sh1.getRange(6,3) ;
        range.setBackground("white");
    }         

  if (kdate == "") {
        Browser.msgBox("エラー：契約日は必須入力項目です！");
        var range=sh1.getRange(4,3) ;
        range.setBackground("yellow");
        return;
   　} else   {
        var range=sh1.getRange(4,3) ;
        range.setBackground("white");
    }         

  if (eigyo == "") {
        Browser.msgBox("エラー：営業担当は必須入力項目です！");
        var range=sh1.getRange(12,3) ;
        range.setBackground("yellow");
        return;
   　} else   {
        var range=sh1.getRange(12,3) ;
        range.setBackground("white");
    }         

  if (tori == "") {
        Browser.msgBox("エラー：取引は必須入力項目です！");
        var range=sh1.getRange(14,3) ;
        range.setBackground("yellow");
        return;
   　} else   {
        var range=sh1.getRange(14,3) ;
        range.setBackground("white");
    }         

  if (tori == "介護保険：販売"　&& kingaku == 0 ) {
        Browser.msgBox("エラー：金額は必須入力項目です！");
        var range=sh1.getRange(16,3) ;
        range.setBackground("yellow");
        return;
   　} else   {
        var range=sh1.getRange(16,3) ;
        range.setBackground("white");
    }         

  if (tori == "介護保険：住宅改修"　&& kingaku == 0 ) {
        Browser.msgBox("エラー：金額は必須入力項目です！");
        var range=sh1.getRange(16,3) ;
        range.setBackground("yellow");
        return;
   　} else   {
        var range=sh1.getRange(16,3) ;
        range.setBackground("white");
    }         

if (tori == "介護保険：販売"){
} else if  (tori == "介護保険：住宅改修") {
} else if  (kingaku != 0) {
Browser.msgBox("エラー：金額は入力できません！");
        var range=sh1.getRange(16,3) ;
        range.setBackground("yellow");
        return;
} else {
var range=sh1.getRange(16,3) ;
        range.setBackground("white");
}


  if (sakusei == "") {
        Browser.msgBox("エラー：取引は必須入力項目です！");
        var range=sh1.getRange(20,3) ;
        range.setBackground("yellow");
        return;
   　} else   {
        var range=sh1.getRange(20,3) ;
        range.setBackground("white");
    }         

// メッセージ　BOX
   
    var box = '【契約】　' + tori + '\\n'
             + '【営業担当】　' + eigyo + '\\n'
             + '【契約日】　' + Utilities.formatDate(kdate,'JST','yyyyMMdd') + '\\n'
             + '【利用者】　' + riyosha + ' さま'+ '\\n' 
             + '【金　額】　' + kingaku + '\\n' 
             + '　処理しますか？' ;
        
// メッセージ表示
   var hyoji = Browser.msgBox(box, Browser.Buttons.OK_CANCEL);
   if (hyoji == 'cancel') {
        Browser.msgBox("処理はキャンセルされました！");
        return;
    }  
      

// 書類名称セット
  const p1 = "・ケアプラン" ;
  const p2 = "・担当者会議録";
  const p3 = "・サービス計画書（選定提案）" ;
  const p4 = "・サービス計画書（利用計画）";
  const p5 = "・完了報告書" ;
  const p6 = "・フェイスシート";
  const p7 = "・例外給付" ;
  const p8 = "・契約書";
  const p9 = "・重要事項説明書" ;
  const p10 = "・介護保険証";
  const p11 = "・負担割合証" ;
  const p12 = "・領収書";
  const p13 = "・ニコス申込書" ;
  const p14 = " ";
  const p15 = " ";
  const p16 = " ";
 
// 書類名称クリア処理

 for(let i = 17;i < 31;i++){
       sh2.getRange(i,2).setValue(""); 
}

// Browser.msgbox(tori、Browser.Buttons.OK_CANCEL);
   if (tori == "レンタル：新規") {
       sh2.getRange(17,2).setValue(p1); 
       sh2.getRange(18,2).setValue(p2);
       sh2.getRange(19,2).setValue(p3);
       sh2.getRange(20,2).setValue(p4);
       sh2.getRange(21,2).setValue(p5);
       sh2.getRange(22,2).setValue(p6);
       sh2.getRange(23,2).setValue(p7); 
       sh2.getRange(24,2).setValue(p8);
       sh2.getRange(25,2).setValue(p9);
       sh2.getRange(26,2).setValue(p10);
       sh2.getRange(27,2).setValue(p11);
       sh2.getRange(28,2).setValue(p12);
//       sh1.getRange(29,2).setValue(p13);
//     sh1.getRange(30,2).setValue(p14);

　} else if (tori == "レンタル：更新")  {
       sh2.getRange(17,2).setValue(p1); 
       sh2.getRange(18,2).setValue(p2);
       sh2.getRange(19,2).setValue(p4);
       sh2.getRange(20,2).setValue(p6);
       sh2.getRange(21,2).setValue(p7);
       sh2.getRange(22,2).setValue(p10);
       
　} else if (tori == "レンタル：追加")  {
       sh2.getRange(17,2).setValue(p1); 
       sh2.getRange(18,2).setValue(p2);
       sh2.getRange(19,2).setValue(p3);
       sh2.getRange(20,2).setValue(p4);
       sh2.getRange(21,2).setValue(p5);
       sh2.getRange(22,2).setValue(p6);
       sh2.getRange(23,2).setValue(p7); 
       sh2.getRange(24,2).setValue(p8);
       
　} else if (tori == "レンタル：機種変更")  {
       sh2.getRange(17,2).setValue(p1); 
       sh2.getRange(18,2).setValue(p2);
       sh2.getRange(19,2).setValue(p3);
       sh2.getRange(20,2).setValue(p4);
       sh2.getRange(21,2).setValue(p5);
       sh2.getRange(22,2).setValue(p7);
       sh2.getRange(23,2).setValue(p8); 
       
　} else if (tori == "レンタル：一部解約")  {
       sh2.getRange(17,2).setValue(p4); 
       sh2.getRange(18,2).setValue(p5);
        
　} else if (tori == "レンタル：退院")  {
       sh2.getRange(17,2).setValue(p1);  //ケアプラン
       sh2.getRange(18,2).setValue(p2);　//担当者会議録
        
　} else if (tori == "レンタル：区変")  {
       sh2.getRange(17,2).setValue(p1);  //ケアプラン
       sh2.getRange(18,2).setValue(p2);　//担当者会議録
        
　} else if (tori == "レンタル：見直し")  {
       sh2.getRange(17,2).setValue(p1);  //ケアプラン
       sh2.getRange(18,2).setValue(p2);　//担当者会議録
        
　} else if (tori == "レンタル：居宅変更")  {
       sh2.getRange(17,2).setValue(p1);  //ケアプラン
       sh2.getRange(18,2).setValue(p2);　//担当者会議録
        
　} else if (tori == "レンタル：ケアマネ変更")  {
       sh2.getRange(17,2).setValue(p1);  //ケアプラン
       sh2.getRange(18,2).setValue(p2);　//担当者会議録
        
　} else if (tori == "レンタル：一般")  {
       sh2.getRange(17,2).setValue(p8);  //契約書
       sh2.getRange(18,2).setValue(p13);　//ニコス申込書
                             
　} else if (tori == "介護保険：販売")  {
       sh2.getRange(17,2).setValue(p1);  //ケアプラン
       sh2.getRange(18,2).setValue(p2);　//担当者会議録
       sh2.getRange(19,2).setValue(p4);  //利用計画
       sh2.getRange(20,2).setValue(p6);  //フェイスシート
       sh2.getRange(21,2).setValue(p9);  //重要事項説明書
       sh2.getRange(22,2).setValue(p10);  //介護保険証
       sh2.getRange(23,2).setValue(p11);  //負担割合証 
       sh2.getRange(24,2).setValue(p12);  //領収書

　} else if (tori == "介護保険：住宅改修")  {
       sh2.getRange(17,2).setValue(p1);  //ケアプラン
       sh2.getRange(18,2).setValue(p2);　//担当者会議録
       sh2.getRange(19,2).setValue(p6);  //フェイスシート
       sh2.getRange(20,2).setValue(p10);  //介護保険証
       sh2.getRange(21,2).setValue(p11);  //負担割合証
       sh2.getRange(22,2).setValue(p12);  //領収書

}


// チェック一覧表へ転記処理
   var lastRow = sh3.getLastRow()+1;
   sh3.getRange(lastRow,2).setValue(kdate); 
   sh3.getRange(lastRow,3).setValue(riyosha); 
   sh3.getRange(lastRow,4).setValue(eigyo); 
   sh3.getRange(lastRow,5).setValue(tori); 
   sh3.getRange(lastRow,6).setValue(kingaku); 
   sh3.getRange(lastRow,7).setValue(tekiyo); 
   sh3.getRange(lastRow,10).setValue(sakusei);    

// PDFの保存先となるフォルダID 確認方法は後述
  var folderid = "17WguLww5B55RZpg2SeeyQ48VD2oi7urR";
  
  // マイドライブ直下に保存したい場合は以下
  // var root= DriveApp.getRootFolder();
  // var folderid = root.getId();
  
  /////////////////////////////////////////////  
  // 現在開いているスプレッドシートをPDF化したい場合//
  ////////////////////////////////////////////
  // 現在開いているスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 現在開いているスプレッドシートのIDを取得
  var ssid = ss.getId();
  
// アクティブシートをcheck印刷用に変更
  var num = 1
  ss.setActiveSheet(ss.getSheets()[num]); 
  
// 現在開いているスプレッドシートのシートIDを取得
  var sheetid = ss.getActiveSheet().getSheetId();
  
//  var sheetid = ss.getSheetByName('check印刷用');
  // getActiveSheetの後の()を忘れると、TypeError: オブジェクト function getActiveSheet() {/* */} で関数 getSheetId が見つかりません。

  // ファイル名に使用する名前を取得
  var customer_name = sh2.getRange("C4").getValue();
  var keiyaku_date =  sh2.getRange("C6").getValue();
  var eigyo_tanto =  sh2.getRange("C8").getValue();
   var tori_kbn =  sh2.getRange("C10").getValue();
  
  Logger.log(customer_name);
  
  // ここで例として使用しているスプレッドシートのC15に顧客の名前が入っているため、それをファイル名用に取得しているだけです。

  // ファイル名に使用するタイムスタンプを取得
  var timestamp = getTimestamp();
  var tstamp = Moment.moment();
  sh3.getRange(lastRow,9).setValue(tstamp.format('YYYY年M月D日H時m分'));  
//  sh3.getRange(lastRow,9).setValue(timestamp);

  // フォルダ作成
  var folder_name = Utilities.formatDate(kdate,'JST','yyyyMMdd')  + "_" + customer_name + "_" + tori_kbn ; 
  DriveApp.createFolder(folder_name) ;

  Logger.log(folder_name);
 
  // PDF作成関数
  createPDF( folder_name , ssid, sheetid, Utilities.formatDate(kdate,'JST','yyyyMMdd')  + "_" + customer_name + "_" + tori_kbn  );
  
// 通知メッセージ　chat送信 レンタル新規、介護保険・販売、住宅改修のみ通知

if (tori == "介護保険：販売"){
} else if  (tori == "介護保険：住宅改修") {
} else if  (tori == "レンタル：新規") {
} else {
return;
}
    Logger.log(tori);
    var text = '【契約】　' + tori + '\n'
             + '【営業担当】　' + eigyo + '\n'
             + '【契約日】　' + Utilities.formatDate(kdate,'JST','yyyyMMdd') + '\n'
             + '【利用者】　' + riyosha + ' さま'+ '\n' 
             + '【金　額】　' + kingaku + '\n' 
             + '　以上、ご報告いたします◎' ;
        
// ペイロード        
    var payload = {
  　'text' : text
  　 }
   
// エンコード
    var json = JSON.stringify(payload);
  
// WebhookURL
    var url = 'https://chat.googleapis.com/v1/spaces/AAAAR3QgIn4/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=QuRRmJV3LYRdg9HQSIp5EZcvhZf_SKKUl9l4wTeJcos%3D';
   
// ポストするためにヘッダーとかボディをまとめて入力
    var options = {
    'method' : 'POST',
    'contentType' : 'application/json; charset=utf-8',
    'payload' : json
    }
   
// 送信！ 
    var response = UrlFetchApp.fetch(url, options); 
    

//
//初期状態　入力項目クリア処理
   
//    var box = '　入力項目をクリアしますか？' ;
        
// メッセージ表示
//   var hyoji = Browser.msgBox(box, Browser.Buttons.OK_CANCEL);
//    if (hyoji == 'ok') {
//    sh1.getRange(4,3).setValue("");
//    sh1.getRange(6,3).setValue("");
//      sh1.getRange(8,3).setValue("");
//      sh1.getRange(10,3).setValue("");
//      sh1.getRange(12,3).setValue(""); 
//     sh1.getRange(14,3).setValue(""); 
//    }     
  
//  sh1.getRange(16,3).setValue("");
  
 }