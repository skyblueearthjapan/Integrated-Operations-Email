// ====== 朝メール（07:00）======
// 内容: 本日休暇者 + 前日残業実績 + 有給警告（該当日のみ） + 作業完了未記録

/**
 * 朝の統合メール送信（トリガーから呼ばれる）
 */
function sendMorningMail() {
  var to = getMailTo_();
  if (!to) { Logger.log('送信先メールが未設定'); return; }

  var now = new Date();
  var dateLabel = fmtDate_(now, 'yyyy/MM/dd');
  var yesterday = new Date(now);
  yesterday.setDate(yesterday.getDate() - 1);

  var lines = [];
  lines.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('【統合業務メール】朝の報告 ' + dateLabel);
  lines.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('');

  // ── 1. 作業完了未記録アラート ──
  try {
    var incompleteList = readIncompleteRequests_();
    if (incompleteList.length > 0) {
      lines.push('【⚠ 作業完了未記録のお知らせ】');
      lines.push('以下の作業員は作業完了ボタンが押されておらず、実働時間が記録されていません。');
      lines.push('');

      var incByDate = {};
      for (var k = 0; k < incompleteList.length; k++) {
        var it = incompleteList[k];
        if (!incByDate[it.targetDate]) incByDate[it.targetDate] = [];
        incByDate[it.targetDate].push(it);
      }

      var dates = Object.keys(incByDate).sort();
      for (var di = 0; di < dates.length; di++) {
        var targetDate = dates[di];
        var dateItems = incByDate[targetDate];
        var deptMap = {};
        for (var j = 0; j < dateItems.length; j++) {
          var item = dateItems[j];
          if (!deptMap[item.dept]) deptMap[item.dept] = [];
          deptMap[item.dept].push(item);
        }
        lines.push('▼ ' + targetDate.replace(/-/g, '/') + ' の未記録');
        var depts = Object.keys(deptMap).sort();
        for (var dj = 0; dj < depts.length; dj++) {
          lines.push('  ■ ' + depts[dj]);
          var arr = deptMap[depts[dj]];
          for (var aj = 0; aj < arr.length; aj++) {
            var typeJa = arr[aj].requestType === 'overtime' ? '残業' : '休日出勤';
            lines.push('    - ' + arr[aj].workerName + '：' + typeJa + ' 承認' + fmtMinutesJa_(arr[aj].approvedMinutes));
          }
        }
        var otUrl = getOtAppUrl_();
        if (otUrl) {
          lines.push('  → 作業員TOP画面：' + otUrl + '?page=top');
          lines.push('  → 総務部手動入力：' + otUrl + '?page=admin&date=' + targetDate);
        }
        lines.push('');
      }
      lines.push('───────────────────────────');
      lines.push('');
    }
  } catch (e) {
    lines.push('（作業完了未記録の取得でエラー: ' + e.message + '）');
    lines.push('');
  }

  // ── 2. 本日の休暇者 ──
  try {
    var todayLeaves = readTodayLeaves_();
    lines.push('【本日の休暇者】' + dateLabel);
    lines.push('');
    if (todayLeaves.length === 0) {
      lines.push('本日の休暇者はいません。');
    } else {
      var leaveGroups = {};
      for (var i = 0; i < todayLeaves.length; i++) {
        var lv = todayLeaves[i];
        if (!leaveGroups[lv.deptName]) leaveGroups[lv.deptName] = [];
        leaveGroups[lv.deptName].push(lv);
      }
      var lvDepts = Object.keys(leaveGroups).sort();
      for (var i = 0; i < lvDepts.length; i++) {
        lines.push('■ ' + lvDepts[i]);
        var lvArr = leaveGroups[lvDepts[i]];
        for (var j = 0; j < lvArr.length; j++) {
          var lv = lvArr[j];
          var detail = lv.leaveType + '　' + lv.leaveKubun;
          if (lv.halfType) detail += '（' + lv.halfType + '）';
          lines.push('  - ' + lv.workerName + '　' + detail);
        }
      }
      lines.push('');
      lines.push('合計: ' + todayLeaves.length + '名');
    }
    var leaveUrl = getLeaveAppUrl_();
    if (leaveUrl) {
      lines.push('');
      lines.push('休暇届け管理画面：' + leaveUrl + '?page=somuAdmin');
    }
  } catch (e) {
    lines.push('（休暇者データ取得でエラー: ' + e.message + '）');
  }
  lines.push('');
  lines.push('───────────────────────────');
  lines.push('');

  // ── 3. 前日の残業・休日出勤 実績 ──
  try {
    var report = readOtMorningReport_(yesterday);
    var yesterdayLabel = fmtDate_(yesterday, 'yyyy/MM/dd');
    lines.push('【残業・休日出勤 実績報告】' + yesterdayLabel);
    lines.push('');

    if (report.items.length === 0) {
      lines.push('前日の承認済み申請はありません。');
    } else {
      var otGroups = {};
      for (var i = 0; i < report.items.length; i++) {
        var it = report.items[i];
        if (!otGroups[it.dept]) otGroups[it.dept] = [];
        otGroups[it.dept].push(it);
      }
      var otDepts = Object.keys(otGroups).sort();
      for (var i = 0; i < otDepts.length; i++) {
        lines.push('■ ' + otDepts[i]);
        var arr = otGroups[otDepts[i]];
        for (var j = 0; j < arr.length; j++) {
          var it = arr[j];
          var typeJa = it.requestType === 'overtime' ? '残業' : '休日出勤';
          lines.push('- ' + it.workerName + '：' + typeJa
            + ' 承認' + fmtMinutesJa_(it.approvedMinutes)
            + ' / 実働' + fmtMinutesJa_(it.actualMinutes)
            + ' / 休憩' + fmtMinutesJa_(it.breakMinutes)
            + ' / 実残業' + fmtMinutesJa_(it.netMinutes)
            + ' [' + it.pdfStatus + ']');
        }
        lines.push('');
      }
      lines.push('PDF作成状況：残業 ' + report.counts.overtime + '件 / 休日 ' + report.counts.holiday + '件 / 合計 ' + report.counts.total + '件');
    }

    var otSettings = getOtSettings_();
    var pdfFolderUrl = normalize_(otSettings['PDF_DRIVE_FOLDER_URL'] || '');
    var accSsId = normalize_(otSettings['ACCUMULATION_SS_ID'] || '');
    var otUrl = getOtAppUrl_();
    lines.push('');
    if (pdfFolderUrl) lines.push('PDFフォルダ：' + pdfFolderUrl);
    if (accSsId) lines.push('データ蓄積：https://docs.google.com/spreadsheets/d/' + accSsId + '/edit');
    if (otUrl) lines.push('残業アプリ：' + otUrl);
  } catch (e) {
    lines.push('（残業実績データ取得でエラー: ' + e.message + '）');
  }
  lines.push('');

  // ── 4. 有給取得警告（該当日のみ） ──
  try {
    var WARN_MAIL_DATES = [
      { month: 7, day: 1 },
      { month: 9, day: 1 },
      { month: 1, day: 1 },
      { month: 2, day: 16 },
    ];
    var m = now.getMonth() + 1;
    var d = now.getDate();
    var isWarnDay = false;
    for (var i = 0; i < WARN_MAIL_DATES.length; i++) {
      if (WARN_MAIL_DATES[i].month === m && WARN_MAIL_DATES[i].day === d) {
        isWarnDay = true;
        break;
      }
    }

    if (isWarnDay) {
      lines.push('───────────────────────────');
      lines.push('');
      var warnData = readWarnWorkers_();
      lines.push('【有給取得警告】有給取得義務 未達者一覧');
      lines.push('');
      lines.push('年度: ' + warnData.fiscalYear + '年度（' + warnData.fiscalYear + '/3/16 - ' + (warnData.fiscalYear + 1) + '/3/15）');
      lines.push('現在の必要有給回数: ' + warnData.currentRequired + '回（' + warnData.currentLabel + '基準）');
      lines.push('');

      if (warnData.workers.length === 0) {
        lines.push('未達者はいません。');
      } else {
        var warnGroups = {};
        for (var i = 0; i < warnData.workers.length; i++) {
          var w = warnData.workers[i];
          if (!warnGroups[w.deptName]) warnGroups[w.deptName] = [];
          warnGroups[w.deptName].push(w);
        }
        var warnDepts = Object.keys(warnGroups).sort();
        for (var i = 0; i < warnDepts.length; i++) {
          lines.push('■ ' + warnDepts[i]);
          var arr = warnGroups[warnDepts[i]];
          for (var j = 0; j < arr.length; j++) {
            var w = arr[j];
            var mark = w.warnLevel === '危険' ? '!!!' : (w.warnLevel === '警告' ? '!!' : '!');
            lines.push('  ' + mark + ' ' + w.workerName + '  有給: ' + w.paidCount + '回  （' + w.warnLevel + '）  最終: ' + (w.lastPaidDate || '-'));
          }
        }
        lines.push('');
        lines.push('対象者数: ' + warnData.workers.length + '名');
      }
      lines.push('');
    }
  } catch (e) {
    lines.push('（有給警告データ取得でエラー: ' + e.message + '）');
  }

  lines.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('※本メールは統合業務メール通知システムより自動送信されています。');

  var subject = '【統合業務メール】朝の報告 ' + dateLabel;
  GmailApp.sendEmail(to, subject, lines.join('\n'));
  Logger.log('朝メール送信完了: ' + to);
}
