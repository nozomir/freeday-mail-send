function beginningOfNextMonthFunc() {
    var beginningOfMonth = new Date();
    beginningOfMonth.setDate(1);
    beginningOfMonth.setMonth(beginningOfMonth.getMonth() + 1);
    return beginningOfMonth;
}

function endOfNextMonthFunc() {
    var endOfMonth = new Date();
    endOfMonth.setDate(1);
    endOfMonth.setMonth(endOfMonth.getMonth() + 2);
    endOfMonth.setDate(0);
    return endOfMonth;
}

/*
 * 与えられた日付が祝日か判定する
 */
function isHoliday(date) {
    //祝日か判定
    var calendarId = "ja.japanese#holiday@group.v.calendar.google.com";
    var calendar = CalendarApp.getCalendarById(calendarId);
    var todayEvents = calendar.getEventsForDay(date);
    if (todayEvents.length > 0) {
        return true;
    }

    return false;
}

/*
 * 与えられた数字の曜日を日本語で返す
 */

function dayOfWeek(date) {
    // 曜日配列
    const myTbl = new Array("日", "月", "火", "水", "木", "金", "土", "日");

    var myDay = Utilities.formatDate(date, "Asia/Tokyo", "u");
    return myTbl[myDay];
}

function findRow(sheet, val, col) {

    var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得

    for (var i = 0; i < dat.length; i++) {
        if (dat[i][col - 1] === val) {
            return i + 1;
        }
    }
    return 0;
}
