// ====== 夕方メール（17:00 / 18:00）======
// 内容: 残業申請状況 + 休日出勤申請 + 休暇未承認通知

/**
 * 夕方1回目の統合メール（17:00）
 * 残業申請状況 + 休暇1次未承認通知
 */
function sendEveningMail1() {
  var to = getMailTo_();
  if (!to) { Logger.log('送信先メールが未設定'); return; }

  var now = new Date();
  var dateLabel = fmtDate_(now, 'yyyy/MM/dd');

  var lines = [];
  lines.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('【統合業務メール】夕方の報告①　' + dateLabel);
  lines.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('');

  // ── 1. 残業・休日出勤 申請状況 ──
  try {
    buildOtSection_(lines, now, dateLabel);
  } catch (e) {
    lines.push('（残業データ取得でエラー: ' + e.message + '）');
    lines.push('');
  }

  lines.push('───────────────────────────');
  lines.push('');

  // ── 2. 休暇届 1次未承認通知 ──
  try {
    var pending1 = readPendingApproval1_();
    lines.push('【休暇届 1次承認未完了通知】');
    lines.push('');
    if (pending1.length === 0) {
      lines.push('1次承認待ちの休暇届はありません。');
    } else {
      lines.push('以下の休暇届が申請されていますが、まだ1次承認されていません。');
      lines.push('');
      var groups = {};
      for (var i = 0; i < pending1.length; i++) {
        var p = pending1[i];
        if (!groups[p.deptName]) groups[p.deptName] = [];
        groups[p.deptName].push(p);
      }
      var depts = Object.keys(groups).sort();
      for (var i = 0; i < depts.length; i++) {
        var arr = groups[depts[i]];
        lines.push('■ ' + depts[i] + '(' + arr.length + '件)');
        for (var j = 0; j < arr.length; j++) {
          lines.push('  - ' + arr[j].workerName + '　' + arr[j].leaveDate + '　' + arr[j].leaveType + '(' + arr[j].leaveKubun + ')');
        }
        lines.push('');
      }
      lines.push('合計: ' + pending1.length + '件');
    }
    var leaveUrl = getLeaveAppUrl_();
    if (leaveUrl) {
      lines.push('');
      lines.push('承認はこちらから：' + leaveUrl + '?page=top');
    }
  } catch (e) {
    lines.push('（休暇未承認データ取得でエラー: ' + e.message + '）');
  }
  lines.push('');

  lines.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('※本メールは統合業務メール通知システムより自動送信されています。');

  var subject = '【統合業務メール】夕方の報告① ' + dateLabel;
  GmailApp.sendEmail(to, subject, lines.join('\n'));
  Logger.log('夕方メール①送信完了: ' + to);
}

/**
 * 夕方2回目の統合メール（18:00）
 * 残業申請状況（更新版）+ 休暇2次未承認通知
 */
function sendEveningMail2() {
  var to = getMailTo_();
  if (!to) { Logger.log('送信先メールが未設定'); return; }

  var now = new Date();
  var dateLabel = fmtDate_(now, 'yyyy/MM/dd');

  var lines = [];
  lines.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('【統合業務メール】夕方の報告②　' + dateLabel);
  lines.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('');

  // ── 1. 残業・休日出勤 申請状況（更新版） ──
  try {
    buildOtSection_(lines, now, dateLabel);
  } catch (e) {
    lines.push('（残業データ取得でエラー: ' + e.message + '）');
    lines.push('');
  }

  lines.push('───────────────────────────');
  lines.push('');

  // ── 2. 休暇届 2次承認待ち通知 ──
  try {
    var pending2 = readPendingApproval2_();
    lines.push('【休暇届 2次承認未完了通知】');
    lines.push('');
    if (pending2.length === 0) {
      lines.push('2次承認待ちの休暇届はありません。');
    } else {
      lines.push('以下の休暇届が1次承認済みですが、2次承認がまだ完了していません。');
      lines.push('');
      var groups = {};
      for (var i = 0; i < pending2.length; i++) {
        var p = pending2[i];
        if (!groups[p.deptName]) groups[p.deptName] = [];
        groups[p.deptName].push(p);
      }
      var depts = Object.keys(groups).sort();
      for (var i = 0; i < depts.length; i++) {
        var arr = groups[depts[i]];
        lines.push('■ ' + depts[i] + '(' + arr.length + '件)');
        for (var j = 0; j < arr.length; j++) {
          lines.push('  - ' + arr[j].workerName + '　' + arr[j].leaveDate + '　' + arr[j].leaveType + '(' + arr[j].leaveKubun + ')　1次承認: ' + arr[j].approvedAt);
        }
        lines.push('');
      }
      lines.push('合計: ' + pending2.length + '件');
    }
    var leaveUrl = getLeaveAppUrl_();
    if (leaveUrl) {
      lines.push('');
      lines.push('2次承認はこちらから：' + leaveUrl + '?page=somuAdmin');
    }
  } catch (e) {
    lines.push('（休暇2次承認データ取得でエラー: ' + e.message + '）');
  }
  lines.push('');

  lines.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('※本メールは統合業務メール通知システムより自動送信されています。');

  var subject = '【統合業務メール】夕方の報告② ' + dateLabel;
  GmailApp.sendEmail(to, subject, lines.join('\n'));
  Logger.log('夕方メール②送信完了: ' + to);
}

// ====== 共通: 残業・休日出勤セクション生成 ======
function buildOtSection_(lines, dateObj, dateLabel) {
  var overtimeItems = readOtOvertimeRequests_(dateObj);
  var holidayItems = readOtHolidayRequests_(dateObj);
  var dateParam = fmtDate_(dateObj, 'yyyy-MM-dd');

  lines.push('【残業・休日出勤 申請状況報告】' + dateLabel);
  lines.push('');

  if (overtimeItems.length === 0 && holidayItems.length === 0) {
    lines.push('本日分の残業申請および今週末の休日出勤申請はありません。');
    lines.push('');
    return;
  }

  // 残業セクション
  if (overtimeItems.length > 0) {
    lines.push('【本日の残業申請】');
    lines.push('');
    var otGroups = {};
    for (var i = 0; i < overtimeItems.length; i++) {
      var it = overtimeItems[i];
      if (!otGroups[it.dept]) otGroups[it.dept] = [];
      otGroups[it.dept].push(it);
    }
    var otDepts = Object.keys(otGroups).sort();
    for (var i = 0; i < otDepts.length; i++) {
      lines.push('■ ' + otDepts[i]);
      var arr = otGroups[otDepts[i]];
      for (var j = 0; j < arr.length; j++) {
        lines.push('- ' + arr[j].workerName + '：残業 ' + fmtMinutesJa_(arr[j].approvedMinutes) + '（' + arr[j].statusLabel + '）');
      }
      lines.push('');
    }
  }

  // 休日出勤セクション
  if (holidayItems.length > 0) {
    lines.push('【今週末の休日出勤申請】');
    lines.push('');
    var hdGroups = {};
    for (var i = 0; i < holidayItems.length; i++) {
      var it = holidayItems[i];
      if (!hdGroups[it.dept]) hdGroups[it.dept] = [];
      hdGroups[it.dept].push(it);
    }
    var hdDepts = Object.keys(hdGroups).sort();
    for (var i = 0; i < hdDepts.length; i++) {
      lines.push('■ ' + hdDepts[i]);
      var arr = hdGroups[hdDepts[i]];
      for (var j = 0; j < arr.length; j++) {
        lines.push('- ' + arr[j].workerName + '：休日出勤 ' + arr[j].targetDateLabel + ' ' + fmtMinutesJa_(arr[j].approvedMinutes) + '（' + arr[j].statusLabel + '）');
      }
      lines.push('');
    }
  }

  var otUrl = getOtAppUrl_();
  if (otUrl) {
    lines.push('残業アプリ：' + otUrl);
    lines.push('二次承認ページ：' + otUrl + '?page=approve2&date=' + dateParam);
    lines.push('');
  }
}
