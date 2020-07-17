function exceptInboxAndAppendLabel(){
  let SEARCH_SUBJECT = "【CycleShare】施錠確認通知";
  let search_subjects = "subject:"+ SEARCH_SUBJECT;
  var search = search_subjects + " label:inbox";
  
  var threads = GmailApp.search(search);
  var msg_count = 0;
  for(var t in threads){
    var thread = threads[t];
    var msgs = thread.getMessages();
    for(var m in msgs){
      var msg = msgs[m];
      console.log("thread:" + t + " msg:" + m + " date:" + msg.getDate());
    }
  }

  // ラベル追加
  let label_append = GmailApp.getUserLabelByName("test1");
  label_append.addToThreads(threads);
  
  // 受信トレイ削除
//  let label_remove = GmailApp.getUserLabelByName("受信トレイ");
//  label_remove.removeFromThreads(threads);
  
  // 受信トレイ(inboxも)のラベル削除はできなかったのでアーカイブにしてみた
  GmailApp.moveThreadsToArchive(threads);
  
}


function getLabels(){
  var labels = GmailApp.getUserLabels();
  labels.forEach(function(label){
    console.log(label.getName());
  });
}

function test() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('active');
  sheet.appendRow(['日付', new Date()]);

  // 行追加
  sheet.insertRowAfter(1);
  
  // 行削除
  sheet.deleteRow(4);

  // 値取得
  var range = sheet.getRange(1, 2);
  var value = range.getValue();
  console.log(value);
  
  // 値書き込み
  var range1 = sheet.getRange(7, 1);
  range1.setValue('set');
  var range2 = sheet.getRange("A8:B8");
  range2.setValues([['set', new Date()]]);

  

}





function regTest(){
  var str = "自転車番号：TYO13308\n";
//  var reg = /自転車番号：(.*?)\n/;
  var reg = /自転車番号：([A-Z]*?[0-9]*?)\n/;
  var no = str.match(reg);
  console.log(no);
  console.log(no[0]);
  console.log(no[1]);
  console.log(no[2]);
}

function dateTest(){
  var now = new Date();
  var old = new Date(2020, 3, 10, 0, 0, 0);
  console.log(now);
  console.log(old);
  var interval = now.getTime() - old.getTime();
  console.log("interval:" + interval);
  var interval_date = new Date(interval);
  console.log(interval_date);
  // 経過時間はマイクロ秒
  var interval_days = Math.floor(interval_date / (1000*60*60*24));
  var interval_hours = Math.floor(interval_date / (1000*60*60));
  var interval_minitus = Math.floor(interval_date / (1000*60));
  var interval_secconds = Math.floor(interval_date / (1000));
  
  console.log(interval_days);
  console.log(interval_hours);
  console.log(interval_minitus);
  console.log(interval_secconds);
  
}

function sendGmailTest(){
  var toAddr = "ryuko2010@gmail.com";
  var subject = "テスト送信";
  var name = "自動送信"
  var body = "これは\n送信\nテスト\nです";
  var message = {to:toAddr, subject:subject, name:name, body:body};
  MailApp.sendEmail(message);
}


function testDeleteRow(){
  var sheet_no_finish = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NO_FINISH);
  // 行削除。引数は行番号（１行目の場合は1）
  sheet_no_finish.deleteRow(2);
}

function test2daysBefore(){
  var now = new Date("2020/1/1 11:00:24");
  console.log(now);

  // 日付までを抽出する(年月日で作成パターン)
  var now_date_only_str = now.getFullYear() + "/" + (now.getMonth() + 1) + "/" + now.getDate();
  console.log(now_date_only_str);
  var now_date_only = new Date(now_date_only_str);
  console.log(now_date_only);
  

  // 日付までを抽出する(時分秒を削除)
  console.log(now.getTime());
  var now_date_only_sirial = now.getTime() - (now.getHours() * 60 * 60 * 1000) - (now.getMinutes() * 60 * 1000) - (now.getSeconds() * 1000) - now.getMilliseconds();
  console.log(now_date_only_sirial);
  var now_date_only2 = new Date();
  now_date_only2.setTime(now_date_only_sirial);
  console.log(now_date_only2);
  var now_date_only2_before = new Date();
  //一日前
  now_date_only2_before.setTime(now_date_only_sirial - 1 * 24 * 60 * 60 * 1000);
  console.log(now_date_only2_before);

  
  
  
}






function myFunction(){
  
  let msg_count = getGmailNotification();

  
  // 返却してない内容に対して時間に応じてメールを送る
  let limit_interval = 4 * 60;
  let notify_intaval_1 = 2 * 60;
  let notify_intaval_2 = 3 * 60 + 20;
  let notify_intaval_3 = 4 * 60;
  

  var notify_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('notify');
  var notify_send_time = notify_sheet.getRange("A2").getValue();
  var notify_span_time = notify_sheet.getRange("D2").getValue();
  var notify_send_count = notify_sheet.getRange("E2").getValue();
  
  var send_count = 0;
  var subject = "";
  var body = "";
  var error_body = "";
  
  for(var list_index = rent_list.length - 1; list_index > 0 ; list_index--){
    let rent_info = rent_list[list_index];
//    console.log(rent_info);

    if(!rent_info["is_finish"]){
      // 対がない(終わってない)
      if(!is_return){
        
        if(send_count > 0){
          error_body += "\nエラー! 複数の未返却があります\n 開始時刻:" + rent_info["start_date"] + "\n 自転車番号:" + rent_info["bike_no"] + "\n";
          continue;
        }
        // 普通はこちら。返却が行われていない
        var now_date = new Date();
        var interval_min = (now_date.getTime() -  rent_info["start_date"].getTime()) / (1000*60);
        
        if(interval_min > notify_intaval_3){
          // 4Hを超えているのでまずい！
          subject = "【レンタサイクル】【警告】すでに4Hを超過しています!至急返却せよ!"
          body = createBody(rent_info["bike_no"], rent_info["start_date"]);
          send_count++;
        }else if(interval_min > notify_intaval_2){
          // 3.5H経過
          // 10分毎に実行する予定なので毎回送信
          var remind_min = limit_interval - interval_min;
          subject = "【レンタサイクル】【警告】残り" + remind_min +"分！至急返却求む!"
          send_count++;
          
        }else if(interval_min > notify_intaval_1){
          // 2H経過
          // 1回だけ送信
          subject = "【レンタサイクル】【通知】残り約２H";
          send_count++;
        }else{
          // 2H以下
          // タイマーカウントを始めたことを一度だけ送信
          subject = "【レンタサイクル】【通知】カウント開始（のこり約4H）";
          send_count++;
        }
      }else{
        // 貸しがなくて返却があるので通常はありえない。
        // メールがうまく送られていないか、バグ
        error_body += "\nエラー! 返却していますが、貸し出し記録がありません"
      }
    }
  }
  
  if(send_count > 0){
    console.log();
    if(error_body.length > 0){
      // 何らかのエラーがあった
      subject += "(エラーあり)";
      body += error_body;
    }
    sendMail(subject , body);
  }
}

function getBikeNo(gmailBody, reg){
  var bike_no = gmailBody.match(reg);
//  console.log("bike_no");
//  console.log(bike_no);
  if(bike_no != null){
    return bike_no[1];
  }
  return "null";  
}

function createBody(bike_no, start_date){
  return "自転車番号:" + bike_no + " 開始時刻:" + start_date + "\n";
}

function sendMail(subject, body){
  var toAddr = SEND_MAIL_ADDRESS;
  var name = "【自動送信】レンタサイクルカウント";
  var message = {to:toAddr, subject:subject, name:name, body:body};
  MailApp.sendEmail(message);  
}


function setLastNotify(bike_no, start_date, span, send_count){
  var notify_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('notify');
  notify_sheet.getRange("A2").setValue(new Date());
  notify_sheet.getRange("B2").setValue(start_date);
  notify_sheet.getRange("C2").setValue(bike_no);
  notify_sheet.getRange("D2").setValue(span);
  notify_sheet.getRange("D2").setValue(send_count);  
}



