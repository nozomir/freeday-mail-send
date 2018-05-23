///////////////////////////////////////
// セミナースケジュール
///////////////////////////////////////

// セミナースケジュールを3ヶ月分送信する
function seminarScheduleSend(eventsText, toMails, title, isProd) {
    // セミナースケジュールを送った就きを管理しているスプレッドシートを取得
    var sheet = SpreadsheetApp
        .openByUrl('https://docs.google.com/spreadsheets/d/1RwE1be-huRpBmpFiOxvKuXILqiRa2BYqJwPNLu-csJg/edit#gid=0')
        .getSheetByName('セミナースケジュール');

    for (i = 1; i <= 3; i++) {
        var nowDate = new Date();
        nowDate.setDate(1);

        // 既にメール送信済みかどうかを調べる
        nowDate.setMonth(nowDate.getMonth() + i);
        var yearMonth = Utilities.formatDate(nowDate, "Asia/Tokyo", "Y/M");

        // 既にメール送信済みなら、次の月を調べる
        if (findRow(sheet, yearMonth, 1) !== 0) {
            continue;
        }
        getSeminarEventsAndEmailSend(yearMonth, eventsText, toMails, title);

        // 本番メール送信なら、送信済みとしてスプレッドシートに値を書き込む
        if (isProd) {
            sheet.getRange(sheet.getLastRow() + 1, 1).setValue(yearMonth);
        }
    }
}

// セミナースケジュールを送信する
function getSeminarEventsAndEmailSend(yearMonth, eventsText, toMails, title) {
    // 月初の日付を取得
    var beginningOfMonth = new Date(yearMonth + '/1');

    // 月末の日付を取得
    var endOfMonth = new Date(yearMonth + '/1');
    endOfMonth.setMonth(endOfMonth.getMonth() + 1);
    endOfMonth.setDate(0);

    // メールの件名
    if (!title) title = Utilities.formatDate(beginningOfMonth, "Asia/Tokyo", "M月のセミナースケジュール☆");

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
            var dow = isHoliday(startDateTime) ? ("(" + dayOfWeek(startDateTime) + "・祝)") : ("(" + dayOfWeek(startDateTime) + ")");
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
            var startDayOfWeek = isHoliday(startDateTime) ? ("(" + dayOfWeek(startDateTime) + "・祝)") : ("(" + dayOfWeek(startDateTime) + ")");
            // 予定終了日
            endDateTime.setDate(endDateTime.getDate() - 1);
            var endDate = Utilities.formatDate(endDateTime, "Asia/Tokyo", "d日");
            var endDayOfWeek = isHoliday(endDateTime) ? ("(" + dayOfWeek(endDateTime) + "・祝)") : ("(" + dayOfWeek(endDateTime) + ")");
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
    seminarScheduleSend("", ['YRSG-P-all_pr@team-site.jp'], "", true);
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
    ];
    seminarScheduleSend(eventsText, toMails, title, false);
}