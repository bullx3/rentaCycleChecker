const SHEET_MAIL_LIST = 'gmail';
const SHEET_NO_FINISH = 'no_finish';
const SHEET_HISTORY = 'history';
const SHEET_CONFIG = 'config';

const GMAIL_LABEL_CHECKED = 'cycle_checked';
const GMAIL_LABEL_ARCHIVE_MY_NOTIFY = 'archive_my_notificaton'
const SPREAD_SHEET_URL = getThisSpreadSheetUrl();
const SPREAD_SHEET_WEB_URL = getThisSpreadSheetWebUrl();
const SEND_MAIL_ADDRESS = getSendMailAddress();
const SEND_MAIL_NAME = "レンタサイクル自動監視";
const GMAIL_LABEL_ARCHIVE = "archive_cycle_share";


function executeCycleCheck(){
  console.log("Start executeCycleCheck");
  // gmailを検索してあたらしい通知を取得 -> gmailシート更新
  getGmailNotification();
  // gmailシートとno_finishシートから対になっていない情報を取得 -> no_finishシート更新
  checkCycleReturn();

  // 前回までに送信した通知メールを既読 ＆ アーカイブ
  // この処理はメール送信の前にすること  
  archiveMyNotification();
  
  // no_finishシートを元にメールを送信
  checkAndSendNotify();
  
  backupMailHistory();

  // 施錠通知を既読 ＆ アーカイブ  
  archiveLockNotification();

  console.log("Finish executeCycleCheck");

}

/*
  gmailから貸し出しと返却メールを抽出しシートに出力する。
  送信日時でソートする（昇順）
  戻り値は取得件数（返却と抽出で２件計算）
*/
function getGmailNotification(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIL_LIST);
  sheet.clear();
  let rent_subject = "【CycleShare】貸出手続き完了";
  let return_subject = "【CycleShare】返却完了";
  let search_subjects = "subject:"+ rent_subject + " OR " + "subject:" + return_subject;
  var now = new Date();
  var sirial_time_before = now.getTime();
  sirial_time_before -= now.getHours() * 60 * 60 * 1000;
  sirial_time_before -= now.getMinutes() * 60 * 1000;
  sirial_time_before -= now.getSeconds() * 1000;
  sirial_time_before -= now.getMilliseconds();  
  sirial_time_before -= 1 * 24 * 60 * 60 * 1000; // 一日前
  var search_date_only = new Date();
  search_date_only.setTime(sirial_time_before);
//  console.log(search_date_only);
  var search_date = search_date_only.getFullYear() + "/" + (search_date_only.getMonth() + 1) + "/" + search_date_only.getDate();

  var search = search_subjects + " after:" + search_date + " label:inbox";
  console.log(search);

  var threads = GmailApp.search(search);
  var msg_count = 0;
  for(var t in threads){
    var thread = threads[t];
    var msgs = thread.getMessages();
    for(var m in msgs){
      
      var msg = msgs[m];
      
      if(!msg.isInInbox()){
        // 検索時にlabel:inboxをつけていてもスレッドで取得すればinboxのないメールも含まれてしまう為、メール毎にチェック
        continue;
      }
      
      var bike_no;
      var type;
      if(msg.getSubject() == rent_subject){
        type = "rent";
        // フォーマット
        // 自転車番号：TYO13308
        bike_no = getBikeNo(msg.getPlainBody(), /自転車番号：([A-Z]*?[0-9]*?)\r\n/);
      }else{
        type = "return";
        // フォーマット
        // 【自転車番号】TYO13308
        bike_no = getBikeNo(msg.getPlainBody(), /【自転車番号】([A-Z]*?[0-9]*?)\r\n/);
      }
//      sheet.appendRow([type, bike_no, msg.getDate(), msg.getSubject(), msg.getPlainBody()]);
      sheet.appendRow([type, bike_no, msg.getDate(), t, m]);
      console.log("Gmail thread: " + t + " msg:" + m + " " + type + " " + bike_no + " " + msg.getDate());

      msg_count++;
    }
  }

  if(msg_count == 0){
    return 0;
  }

  // 昇順でソート
  var range = sheet.getRange("A:Z");
  range.sort({column: 3, ascending: true});

  // スレッドにラベルを追加
  let label = GmailApp.getUserLabelByName(GMAIL_LABEL_CHECKED);
  label.addToThreads(threads);

  // アーカイブして受信トレイを外す
  GmailApp.moveThreadsToArchive(threads);
  
  // 既読にする
  GmailApp.markThreadsRead(threads);
  
  return msg_count;
}

/*
  返却されているかどうかのチェック
*/
function checkCycleReturn(){
  
  var sheet_no_finish = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NO_FINISH);
  let no_finish_count = sheet_no_finish.getLastRow();
  var no_finish_values = [];
  if(no_finish_count > 0){
    no_finish_values = sheet_no_finish.getRange(1, 1, no_finish_count, 3).getValues();
  }

  var sheet_mail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIL_LIST);
  let mail_count = sheet_mail.getLastRow();
  var mail_values = [];
  if(mail_count > 0){
    mail_values = sheet_mail.getRange(1, 1, mail_count, 3).getValues();
  }
  
  
  /*
   no_finish_values(前回までに残っている結果)
   mail_values(今回gmailから取得した結果)
   を結合し、対になっていない項目を検出する。
   対になっていない項目はsheet_no_finishに出力する(前回までのは全削除)
  */
  
  var all_values = no_finish_values.slice(); // 結果的に配列コピー
  for(var i = 0; i < mail_count; i++){
    all_values.push(mail_values[i]);
  }

  var rent_list = [];

  for(var row = 0; row < all_values.length; row++){
    
    var type = all_values[row][0];
    var bike_no = all_values[row][1];
    var date = all_values[row][2];
    
    let len = rent_list.length;
    var search_flg = false;
    for(var list_index = 0; list_index < len; list_index++){
      // 借り、返却で対になる内容を探す
      // メールが前後する関係も考えて順番は考慮しないが、基本的に古い履歴から埋めていく
      var rent_info = rent_list[list_index];

      if(rent_info["is_finish"] || bike_no != rent_info["bike_no"]){
        // すでに対が見つかっている
        // or  自転車番号が一致してない
        continue;
      }
            
      // 対になるリストが見つかっった
      if(type == "rent" && rent_info["is_return"]){
        rent_info["is_rent"] = true;
        rent_info["is_finish"] = true;
        rent_info["start_date"] = date;
        search_flg = true;
        break;
      }else if(type == "return" && rent_info["is_rent"]){
        rent_info["is_return"] = true;
        rent_info["is_finish"] = true;
        rent_info["finish_date"] = date;
        search_flg = true;
        break;
      }
    }

    if(!search_flg){
      // 対になる内容が見つかっていないのでrent_listに追加
      
      if(type == "rent"){
        rent_list.push({is_rent:true, is_return:false, bike_no:bike_no, start_date:date, is_finish: false});
      }else{
        rent_list.push({is_rent:false, is_return:true, bike_no:bike_no, finish_date:date, is_finish: false});        
      }      
    }
  }  // all_valuesのfor文end
  
  // 対になっていないリストno_finishのシートに出力する
  sheet_no_finish.clear();
  rent_list.forEach(function(rent_value){
    if(!rent_value["is_finish"]){
      // 対になっていないため出力
      if(rent_value["is_rent"]){
        sheet_no_finish.appendRow(["rent", rent_value["bike_no"], rent_value["start_date"]]);
      }else{
        sheet_no_finish.appendRow(["return", rent_value["bike_no"], rent_value["finish_date"]]);        
      }
    }
  });

}

/*
  no_finishのシートを参照してのこり時間に合わせてメールを送信する
*/
function checkAndSendNotify(){
  var sheet_no_finish = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NO_FINISH);
  let no_finish_count = sheet_no_finish.getLastRow();
  if(no_finish_count == 0){
    return;
  }

    // 降順でソート
  sheet_no_finish.sort(3,false);

  var no_finish_values = sheet_no_finish.getRange(1, 1, no_finish_count, 3).getValues();

  
  // 最後にrentで受信したメールを未返却のメールとし、それ以外をエラーとする
  
  var subject = "";
  var error_subject = "";
  var error_message = "";
  
  if(no_finish_values[0][0] == "rent"){
    var bike_no = no_finish_values[0][1];
    var mail_date = no_finish_values[0][2];
    if(no_finish_count > 1){
      error_subject = "【error】"
      error_message = "\n\nエラー履歴\n";
      for(var i = 1; i < no_finish_count; i++){
        var error_type = no_finish_values[i][0];
        var error_bike_no = no_finish_values[i][1];
        var error_date = no_finish_values[i][2];
        error_message += error_type + " 車両番号:" + error_bike_no + " 受信時刻: " + dateFormat(error_date) + "\n";
      }
    }

    // のこり時間をチェックして送信
    const TRIGER_MIN = 10; // 10分毎に確認の場合
    var now_date = new Date();
    var interval_min = Math.ceil((now_date.getTime() -  mail_date.getTime()) / (1000*60));
    
    // 返却リミットは4Hとしてチェックのインターバルを考慮した時間でチェックする
    if(interval_min < TRIGER_MIN ){ // 開始直後
      // 監視開始を送信
      subject = "【通知】"+ error_subject + bike_no + " の監視をはじめました";
    }else if(interval_min <= (2*60) && interval_min > (2*60 - TRIGER_MIN) ){ // 1h50m〜2h
      subject = "【通知】"+ error_subject + "残り約２時間";
      
    }else if(interval_min <= (3*60) && interval_min > (3*60 - TRIGER_MIN) ){ // 2h50m〜3h
      subject = "【通知】"+ error_subject + "残り約1時間";

    }else if(interval_min < (4*60) && interval_min > (3*60 + 30 - TRIGER_MIN)){ // 3h20m 〜 4h
      // 10分毎に送信
      var remind_min = 4 * 60 - interval_min;
      subject = "【警告】" + error_subject + "残り" + remind_min +"分！至急返却求む!";
      
    }else if(interval_min >= (4*60)){ // 4H以上
      var min = interval_min % 60;
      if((min >= 20 && min < 30) || (min >= 50 && min < 60)){ // 30分おき
          subject = "【！！警告！！】" + error_subject + "すでに4Hを超過しています!至急返却せよ!"         
      }      
    }
    
　  // メール送信
    if(subject != ""){
      var toAddr = SEND_MAIL_ADDRESS;
      var name = SEND_MAIL_NAME;
      var body = "車両No: " + bike_no + "\n開始時刻:" + dateFormat(mail_date) + error_message + '\n\n' + SPREAD_SHEET_URL + '\n' + SPREAD_SHEET_WEB_URL + '\n';
      var message = {to:toAddr, subject:subject, name:name, body:body};
      MailApp.sendEmail(message);
    }

  }else{
    // エラーのみ（１H毎に通知）
    var now_min = (new Date()).getMinutes();
//    now_min = 51;
    
    if((now_min >= 50 && now_min < 60)){ // １時間おき
      subject = "【error】整合性の合わないデータがあります";
      for(var i = 0; i < no_finish_count; i++){
        var type_str =  no_finish_values[i][0];
        var error_bike_no = no_finish_values[i][1];
        var error_date = no_finish_values[i][2];
        error_message += type_str + " 車両番号:" + error_bike_no + " 受信時刻: " + dateFormat(error_date) + "\n";
      }
    }
    
　  // メール送信
    if(subject != ""){
      var toAddr = SEND_MAIL_ADDRESS;
      var name = SEND_MAIL_NAME;
      var body = error_message + "\n\n" + SPREAD_SHEET_URL;
      var message = {to:toAddr, subject:subject, name:name, body:body};
      MailApp.sendEmail(message);
    }


  }  
}


/*
  

*/
function backupMailHistory(){
  // gmail内容をhistoryに追記
  var sheet_mail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIL_LIST);
  let mail_count = sheet_mail.getLastRow();
  if(mail_count == 0){
    return;
  }

  var mail_values = sheet_mail.getRange(1, 1, mail_count, 5).getValues();

  var sheet_history = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_HISTORY);
  
  var now = new Date();  
  
  mail_values.forEach(function(value){
    value.push(now);
    sheet_history.appendRow(value);
  });
}



/*
  引き数のDateをHH:mm:ss 形式の文字列で返却
*/
function dateFormat(date){
  var str = date.getHours() + ":" + date.getMinutes() + ":" + date.getSeconds();
  return str;
}


/*
  施錠確認通知をアーカイブ ＆ 既読にする
*/
function archiveLockNotification(){
  let SEARCH_SUBJECT = "【CycleShare】施錠確認通知";
  let search_subjects = "subject:"+ SEARCH_SUBJECT;
  var search = search_subjects + " label:inbox";
  
  var threads = GmailApp.search(search);
/*
  var msg_count = 0;
  for(var t in threads){
    var thread = threads[t];
    var msgs = thread.getMessages();
    for(var m in msgs){
      var msg = msgs[m];
      console.log("thread:" + t + " msg:" + m + " date:" + msg.getDate());
    }
  }
*/

  // ラベル追加
  let label_append = GmailApp.getUserLabelByName(GMAIL_LABEL_ARCHIVE);
  label_append.addToThreads(threads);
    
  // 受信トレイ(inbox)のラベル削除はできないのでアーカイブする
  GmailApp.moveThreadsToArchive(threads);
  
  // 既読にする
  GmailApp.markThreadsRead(threads);
  
}

/*
  自分で送った通知メールを既読 ＆ アーカイブにする
  （送ったばかりのメールを既読にされないように送信前に実行する）
*/
function archiveMyNotification(){
  var search = "from:" + SEND_MAIL_NAME + " label:inbox";
  var threads = GmailApp.search(search);

  // ラベル追加
  let label_append = GmailApp.getUserLabelByName(GMAIL_LABEL_ARCHIVE_MY_NOTIFY);
  label_append.addToThreads(threads);
  
  
  // 受信トレイからけすためにアーカイブする
  GmailApp.moveThreadsToArchive(threads);
  
  // 既読にする
  GmailApp.markThreadsRead(threads);  

}

function getSendMailAddress(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  let addr = sheet.getRange('B1').getValue();
  return addr;
}

function getThisSpreadSheetUrl(){
  let url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  return url;
}

function getThisSpreadSheetWebUrl(){
  let url = ScriptApp.getService().getUrl();
  let last_index = url.lastIndexOf('/');
  let exec_url = url.slice(0,last_index) + "/exec";
  
  return exec_url;  
}


