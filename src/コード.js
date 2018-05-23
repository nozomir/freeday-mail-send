///////////////////////////////////////
// フリーデー 
///////////////////////////////////////

// リーダー陣にリマインダー送る
function sendFreedayForReminder() {
  var toMails = [
    'rosshibonber@gmail.com',
    'masa7898@gmail.com',
    'sakameg@gmail.com',
    'q.hamamura@gmail.com',
    'kuni.eguchi@gmail.com',
    'amybirkin61@gmail.com',
    'minami.shining100@gmail.com',
    'jun.beck.1oo7@gmail.com',
    'jam.1.over.drive@gmail.com'
  ];
  var eventsText = "リーダーの皆様☆\n\n"
                       + "いつもフリーデーのお時間をいただきありがとうございます！！(≧∇≦)☆\n"
                       + "フリーデーについて、以下の内容でよいかどうか、ご確認いただけますと幸いです♪\n\n"
                       + "※15日の23時過ぎにメーリスに流れるため、それまでに修正お願いいたします☆\n\n"
                       + "よろしくお願いいたします☆\n\n"
                       + "---------------------------\n\n";
  getFreedayEventsAndEmailSend(eventsText, toMails, "【本日中にご確認をお願いいたします！！】フリーデーのスケジュールについて☆");
}

// メーリングリストにスケジュールを送る
function sendFreedayToMailingList() {
  getFreedayEventsAndEmailSend("", ['Rosshi-hyper@team-site.jp'], "");
}

// フリーデースケジュールを送信する
function getFreedayEventsAndEmailSend(eventsText, toMails, title) {
  var address = 'masameg.schedule@gmail.com';//カレンダーのアドレス
  // 来月の月初の日付を取得
  var beginningOfMonth = beginningOfMonthFunc();
  
  // 来月の月末の日付を取得
  var endOfMonth = endOfMonthFunc();
  
  // メールの件名
  if(!title) title = Utilities.formatDate(beginningOfMonth, "Asia/Tokyo", "M月のフリーデー☆");
  
  // 対象のカレンダーを取得
  // フリーデー
  var freedayCal = CalendarApp.getCalendarById(address);
  // 翌月のイベントをeventsに格納
  var events = freedayCal.getEvents(beginningOfMonth, endOfMonth);
  // 本文TOP
  eventsText += (beginningOfMonth.getMonth() + 1) + "月のフリーデーです！！(≧∇≦)\n"
                    + "手帳にメモメモー！！φ(..)\n\n"
                    + "------------------------------\n\n";
  
  // 曜日配列
  const myTbl = new Array("日","月","火","水","木","金","土","日");
  
  // 予定をeventsTextにまとめる
  for(var i = 0; i < events.length; i++){
    // 予定日
    var date = Utilities.formatDate(events[i].getStartTime(), "Asia/Tokyo", "M/dd");
    // 曜日
    var myDay = Utilities.formatDate(events[i].getStartTime(), "Asia/Tokyo" , "u");
    // 本文
    eventsText += date + "(" + myTbl[myDay] + ") " + events[i].getTitle() + "\n";
  }
  
  eventsText += "\n\n";
  
  var name = Utilities.formatDate(beginningOfMonth, "Asia/Tokyo", "M月のフリーデー☆");
  //メール送信
  for (var i = 0, max = toMails.length; i < max; i++) {
    MailApp.sendEmail({
      to: toMails[i],
      subject: title,
      body: eventsText,
      name: name
    });
  }
}

///////////////////////////////////////
// フリーデーここまで
///////////////////////////////////////

///////////////////////////////////////
// セミナースケジュール
///////////////////////////////////////

// セミナースケジュールを送信する
function getSeminarEventsAndEmailSend(eventsText, toMails, title) {
  // 来月の月初の日付を取得
  var beginningOfMonth = beginningOfMonthFunc();
  
  // 来月の月末の日付を取得
  var endOfMonth = endOfMonthFunc();
  
  // メールの件名
  if(!title) title = Utilities.formatDate(beginningOfMonth, "Asia/Tokyo", "M月のセミナースケジュール☆");
  
  // 対象のカレンダーを取得
  // セミナースケジュール
  var cal = CalendarApp.getCalendarById('ju9ep2f0cadt7srd9vj416g04g@group.calendar.google.com');
  // 翌月のイベントをeventsに格納
  var events = cal.getEvents(beginningOfMonth, endOfMonth);
  // 本文TOP
  eventsText += (beginningOfMonth.getMonth() + 1) + "月のセミナースケジュールです！！(≧∇≦)" + "\n\n"
                    + "手帳にメモメモ～φ(^▽^)メモメモ～♪♪\n\n\n"
                    + "【" + (beginningOfMonth.getMonth() + 1) + "月】" + "\n\n";
  
  // 予定をeventsTextにまとめる
  for (var i = 0; i < events.length; i++) {
    var startDateTime = events[i].getStartTime();
    var endDateTime = events[i].getEndTime();
    if (startDateTime.getDate() === endDateTime.getDate()) {
      // 予定日
      var date = Utilities.formatDate(startDateTime, "Asia/Tokyo", "d日");
      // 曜日
      var dow = isHoliday(startDateTime) ? ("(" + dayOfWeek(startDateTime) + "・祝)" ) : ("(" + dayOfWeek(startDateTime) + ")" );
      // 開始時間
      var startTime = Utilities.formatDate(startDateTime, "Asia/Tokyo", "HH:mm");
      // 終了時間
      var endTime = Utilities.formatDate(endDateTime, "Asia/Tokyo", "HH:mm");
      // 本文
      eventsText += "◯" + date + dow + "\n"
                  + startTime + "～" + endTime + "\n"
                  + events[i].getTitle() + "\n\n\n";
    } else {
      // 予定開始日
      var startDate = Utilities.formatDate(startDateTime, "Asia/Tokyo", "d日");
      var startDayOfWeek = isHoliday(startDateTime) ? ("(" + dayOfWeek(startDateTime) + "・祝)" ) : ("(" + dayOfWeek(startDateTime) + ")" );
      // 予定終了日
      endDateTime.setDate(endDateTime.getDate() - 1);
      var endDate = Utilities.formatDate(endDateTime, "Asia/Tokyo", "d日");
      var endDayOfWeek = isHoliday(endDateTime) ? ("(" + dayOfWeek(endDateTime) + "・祝)" ) : ("(" + dayOfWeek(endDateTime) + ")" );
      // 本文
      eventsText += "◯" + startDate + startDayOfWeek + "～" + endDate + endDayOfWeek + "\n"
                  + events[i].getTitle() + "\n\n\n";
    }
  }
  
  //メール送信
  for (var i = 0, max = toMails.length; i < max; i++) {
    MailApp.sendEmail({
      to: toMails[i],
      subject: title,
      body: eventsText,
      name: "セミナースケジュール☆"
    });
  }
}

// 事業家セミナースケジュールをメーリスに送る
function sendSeminarScheduleToMailngList() {
  getSeminarEventsAndEmailSend("", ['YRSG-P-all_pr@team-site.jp'], "");
}

// 事業家セミナースケジュールをリマインダーする
function sendSeminarScheduleForReminder() {
  var eventsText = "まささん☆\nめぐさん☆\nえぐさん☆\n\n"
                   + "いつもありがとうございます！！\n\n"
                   + "セミナースケジュールについて、以下の内容でよいかどうか、ご確認お願いいたします☆\n\n"
                   + "※15日の23時過ぎにメーリスに流れるため、それまでに修正お願いいたします☆\n\n"
                   + "-----------------------------------------------------------------------\n\n";
  var title = '【本日中に確認をお願いいたします！！】セミナースケジュールについて☆';
  var toMails = [
    'masa7898@gmail.com',
    'sakameg@gmail.com',
    'kuni.eguchi@gmail.com',
    'jam.1.over.drive@gmail.com'
  ]
  getSeminarEventsAndEmailSend(eventsText, toMails, title);
}

///////////////////////////////////////
// セミナースケジュールここまで
///////////////////////////////////////

function beginningOfMonthFunc() {
  var beginningOfMonth = new Date();
  beginningOfMonth.setDate(1);
  beginningOfMonth.setMonth(beginningOfMonth.getMonth() + 1);
  return beginningOfMonth;
}

function endOfMonthFunc() {
  var endOfMonth = new Date();
  endOfMonth.setDate(1);
  endOfMonth.setMonth(endOfMonth.getMonth() + 2);
  endOfMonth.setDate(0);
  return endOfMonth;
}

/*
 * 与えられた日付が祝日か判定する
 */
function isHoliday(date){
  //祝日か判定
  var calendarId = "ja.japanese#holiday@group.v.calendar.google.com";
  var calendar = CalendarApp.getCalendarById(calendarId);
  var todayEvents = calendar.getEventsForDay(date);
  if(todayEvents.length > 0){
    return true;
  }

  return false;
}

/*
 * 与えられた数字の曜日を日本語で返す
 */

function dayOfWeek(date) {
  // 曜日配列
  const myTbl = new Array("日","月","火","水","木","金","土","日");
  
  var myDay = Utilities.formatDate(date, "Asia/Tokyo" , "u");
  return myTbl[myDay];
}