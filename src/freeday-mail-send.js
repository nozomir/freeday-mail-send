///////////////////////////////////////
// フリーデー 
///////////////////////////////////////

// リーダー陣にリマインダー送る
function sendFreedayForReminder() {
  var toMails = [
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
  var beginningOfMonth = beginningOfNextMonthFunc();

  // 来月の月末の日付を取得
  var endOfMonth = endOfNextMonthFunc();

  // メールの件名
  if (!title) title = Utilities.formatDate(beginningOfMonth, "Asia/Tokyo", "M月のフリーデー☆");

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
  const myTbl = new Array("日", "月", "火", "水", "木", "金", "土", "日");

  // 予定をeventsTextにまとめる
  for (var i = 0; i < events.length; i++) {
    // 予定日
    var date = Utilities.formatDate(events[i].getStartTime(), "Asia/Tokyo", "M/dd");
    // 曜日
    var myDay = Utilities.formatDate(events[i].getStartTime(), "Asia/Tokyo", "u");
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
