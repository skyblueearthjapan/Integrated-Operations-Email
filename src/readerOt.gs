// ====== 残業・休日出勤申請app データ読み取り ======

/**
 * 当日の残業申請一覧（submitted + approved）
 * 追加: 本日提出の過去日遡及申請（targetDate < 本日 かつ targetDate >= 年度開始3/16）もOR条件で含める
 */
function readOtOvertimeRequests_(dateObj) {
  var ss = getOtSS_();
  var info = getSheetHeaderIndex_(ss, 'Requests', 1);
  var sh = info.sh;
  var idx = info.idx;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var ymd = fmtDate_(dateObj, 'yyyy-MM-dd');
  var fyStartYmd = fmtDate_(getFiscalYearStartDate_(dateObj), 'yyyy-MM-dd');
  var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var status = normalize_(row[idx['status(submitted/approved/canceled)']]);
    if (!status || status === 'canceled') continue;

    var requestType = normalize_(row[idx['requestType(overtime/holiday)']]);
    if (requestType !== 'overtime') continue;

    var targetDateVal = row[idx['targetDate']];
    var targetYmd;
    try {
      targetYmd = targetDateVal instanceof Date
        ? fmtDate_(targetDateVal, 'yyyy-MM-dd')
        : fmtDate_(new Date(targetDateVal), 'yyyy-MM-dd');
    } catch (e) { continue; }

    // submittedAt を取得
    var submittedAtVal = idx['submittedAt'] !== undefined ? row[idx['submittedAt']] : null;
    var submittedYmd = '';
    if (submittedAtVal) {
      try {
        submittedYmd = submittedAtVal instanceof Date
          ? fmtDate_(submittedAtVal, 'yyyy-MM-dd')
          : fmtDate_(new Date(submittedAtVal), 'yyyy-MM-dd');
      } catch (e) { submittedYmd = ''; }
    }

    // 条件A: 既存 → 対象日が本日
    var matchToday = (targetYmd === ymd);
    // 条件B: 追加 → 本日提出かつ対象日が過去（年度開始以降）
    var matchRetro = (submittedYmd === ymd) && (targetYmd < ymd) && (targetYmd >= fyStartYmd);
    if (!matchToday && !matchRetro) continue;

    var approvedBy2 = idx['approvedBy2'] !== undefined ? normalize_(row[idx['approvedBy2']]) : '';
    var statusLabel = '未承認';
    if (approvedBy2) statusLabel = '二次承認済';
    else if (status === 'approved') statusLabel = '承認済';

    var isRetroactive = computeIsRetroactive_(submittedAtVal, targetYmd);
    var daysAgo = isRetroactive ? computeDaysAgo_(targetYmd, ymd) : 0;

    out.push({
      dept: normalize_(row[idx['dept']]),
      workerName: normalize_(row[idx['workerName']]),
      approvedMinutes: Number(row[idx['approvedMinutes']] || 0),
      statusLabel: statusLabel,
      targetDate: targetYmd,
      isRetroactive: isRetroactive,
      daysAgo: daysAgo,
    });
  }

  out.sort(function(a, b) { return (a.dept + a.workerName).localeCompare(b.dept + b.workerName, 'ja'); });
  return out;
}

/**
 * 今週末〜の休日出勤申請一覧
 * 追加: 本日提出の過去日遡及休日出勤申請（年度開始以降〜本日未満）もOR条件で含める
 */
function readOtHolidayRequests_(dateObj) {
  var ss = getOtSS_();
  var info = getSheetHeaderIndex_(ss, 'Requests', 1);
  var sh = info.sh;
  var idx = info.idx;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var now = dateObj;
  var todayYmd = fmtDate_(now, 'yyyy-MM-dd');
  var fyStartYmd = fmtDate_(getFiscalYearStartDate_(now), 'yyyy-MM-dd');
  var dow = now.getDay();
  var monday = new Date(now);
  monday.setDate(now.getDate() - (dow === 0 ? 6 : dow - 1));
  monday.setHours(0, 0, 0, 0);
  var saturday = new Date(monday);
  saturday.setDate(monday.getDate() + 5);
  var nextMonday = new Date(monday);
  nextMonday.setDate(monday.getDate() + 7);

  var weekendStart = fmtDate_(saturday, 'yyyy-MM-dd');
  var weekendEnd = fmtDate_(nextMonday, 'yyyy-MM-dd');

  var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  var dayNames = ['日','月','火','水','木','金','土'];
  var out = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var status = normalize_(row[idx['status(submitted/approved/canceled)']]);
    if (!status || status === 'canceled') continue;

    var requestType = normalize_(row[idx['requestType(overtime/holiday)']]);
    if (requestType !== 'holiday') continue;

    var targetDateVal = row[idx['targetDate']];
    var targetYmd;
    try {
      targetYmd = targetDateVal instanceof Date
        ? fmtDate_(targetDateVal, 'yyyy-MM-dd')
        : fmtDate_(new Date(targetDateVal), 'yyyy-MM-dd');
    } catch (e) { continue; }

    // submittedAt 取得
    var submittedAtVal = idx['submittedAt'] !== undefined ? row[idx['submittedAt']] : null;
    var submittedYmd = '';
    if (submittedAtVal) {
      try {
        submittedYmd = submittedAtVal instanceof Date
          ? fmtDate_(submittedAtVal, 'yyyy-MM-dd')
          : fmtDate_(new Date(submittedAtVal), 'yyyy-MM-dd');
      } catch (e) { submittedYmd = ''; }
    }

    // 条件A: 既存 → 今週末範囲
    var matchWeekend = (targetYmd >= weekendStart && targetYmd <= weekendEnd);
    // 条件B: 追加 → 本日提出の過去日休日出勤
    var matchRetro = (submittedYmd === todayYmd) && (targetYmd < todayYmd) && (targetYmd >= fyStartYmd);
    if (!matchWeekend && !matchRetro) continue;

    var d = new Date(targetYmd + 'T00:00:00');
    var dateLabel = (d.getMonth()+1) + '/' + d.getDate() + '(' + dayNames[d.getDay()] + ')';

    var approvedBy2 = idx['approvedBy2'] !== undefined ? normalize_(row[idx['approvedBy2']]) : '';
    var statusLabel = '未承認';
    if (approvedBy2) statusLabel = '二次承認済';
    else if (status === 'approved') statusLabel = '承認済';

    var isRetroactive = computeIsRetroactive_(submittedAtVal, targetYmd);
    var daysAgo = isRetroactive ? computeDaysAgo_(targetYmd, todayYmd) : 0;

    out.push({
      dept: normalize_(row[idx['dept']]),
      workerName: normalize_(row[idx['workerName']]),
      approvedMinutes: Number(row[idx['approvedMinutes']] || 0),
      statusLabel: statusLabel,
      targetDate: targetYmd,
      targetDateLabel: dateLabel,
      isRetroactive: isRetroactive,
      daysAgo: daysAgo,
    });
  }

  out.sort(function(a, b) {
    if (a.targetDate !== b.targetDate) return a.targetDate < b.targetDate ? -1 : 1;
    return (a.dept + a.workerName).localeCompare(b.dept + b.workerName, 'ja');
  });
  return out;
}

/**
 * 前日の承認済み残業・休日出勤の実績（朝メール用）
 */
function readOtMorningReport_(dateObj) {
  var ss = getOtSS_();
  var info = getSheetHeaderIndex_(ss, 'Requests', 1);
  var sh = info.sh;
  var idx = info.idx;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { items: [], counts: { overtime: 0, holiday: 0, total: 0 } };

  var ymd = fmtDate_(dateObj, 'yyyy-MM-dd');
  var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  // WorkLogs読み取り
  var workMap = readWorkLogsMap_(ss);

  var items = [];
  var pdfOt = 0, pdfHd = 0;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var status = normalize_(row[idx['status(submitted/approved/canceled)']]);
    if (status !== 'approved') continue;

    var targetDateVal = row[idx['targetDate']];
    var targetYmd;
    try {
      targetYmd = targetDateVal instanceof Date
        ? fmtDate_(targetDateVal, 'yyyy-MM-dd')
        : fmtDate_(new Date(targetDateVal), 'yyyy-MM-dd');
    } catch (e) { continue; }
    if (targetYmd !== ymd) continue;

    var requestId = normalize_(row[idx['requestId']]);
    var requestType = normalize_(row[idx['requestType(overtime/holiday)']]);
    var wl = workMap[requestId] || {};
    var hasPdf = !!(row[idx['pdfFileId']] && normalize_(row[idx['pdfFileId']]));

    if (hasPdf) {
      if (requestType === 'overtime') pdfOt++;
      else pdfHd++;
    }

    items.push({
      dept: normalize_(row[idx['dept']]),
      workerName: normalize_(row[idx['workerName']]),
      requestType: requestType,
      approvedMinutes: Number(row[idx['approvedMinutes']] || 0),
      actualMinutes: wl.actualMinutes || 0,
      breakMinutes: wl.breakMinutes || 0,
      netMinutes: wl.netMinutes || 0,
      pdfStatus: hasPdf ? 'PDF作成済' : 'PDF未作成',
    });
  }

  items.sort(function(a, b) { return (a.dept + a.workerName).localeCompare(b.dept + b.workerName, 'ja'); });

  return {
    items: items,
    counts: { overtime: pdfOt, holiday: pdfHd, total: pdfOt + pdfHd },
  };
}

/**
 * 作業完了未記録の申請一覧（過去7日以内）
 */
function readIncompleteRequests_() {
  var ss = getOtSS_();
  var info = getSheetHeaderIndex_(ss, 'Requests', 1);
  var sh = info.sh;
  var idx = info.idx;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var now = new Date();
  var todayYmd = fmtDate_(now, 'yyyy-MM-dd');
  var sevenDaysAgo = new Date(now);
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
  var limitYmd = fmtDate_(sevenDaysAgo, 'yyyy-MM-dd');

  var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  var workMap = readWorkLogsMap_(ss);

  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var status = normalize_(row[idx['status(submitted/approved/canceled)']]);
    if (status !== 'approved') continue;

    var targetDateVal = row[idx['targetDate']];
    var targetYmd;
    try {
      targetYmd = targetDateVal instanceof Date
        ? fmtDate_(targetDateVal, 'yyyy-MM-dd')
        : fmtDate_(new Date(targetDateVal), 'yyyy-MM-dd');
    } catch (e) { continue; }

    if (targetYmd >= todayYmd || targetYmd < limitYmd) continue;

    var requestId = normalize_(row[idx['requestId']]);
    var wl = workMap[requestId] || {};
    if (wl.actualEndAt) continue;  // 完了済みはスキップ

    out.push({
      requestType: normalize_(row[idx['requestType(overtime/holiday)']]),
      dept: normalize_(row[idx['dept']]),
      workerName: normalize_(row[idx['workerName']]),
      targetDate: targetYmd,
      approvedMinutes: Number(row[idx['approvedMinutes']] || 0),
    });
  }

  out.sort(function(a, b) {
    if (a.dept !== b.dept) return a.dept.localeCompare(b.dept, 'ja');
    return a.workerName.localeCompare(b.workerName, 'ja');
  });
  return out;
}

/**
 * WorkLogsをrequestIdのMapとして読み取り
 */
function readWorkLogsMap_(ss) {
  var sh = ss.getSheetByName('WorkLogs');
  if (!sh) return {};
  var lastRow = sh.getLastRow();
  if (lastRow < 3) return {};

  var header = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h) { return normalize_(h); });
  var idx = buildHeaderIndex_(header);

  var ridCol = idx['requestId'];
  if (ridCol === undefined) return {};

  var values = sh.getRange(3, 1, lastRow - 2, sh.getLastColumn()).getValues();
  var map = {};

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rid = normalize_(row[ridCol]);
    if (!rid) continue;
    map[rid] = {
      actualEndAt: row[idx['actualEndAt']] ? String(row[idx['actualEndAt']]) : '',
      actualMinutes: Number(row[idx['actualMinutes']] || 0),
      breakMinutes: Number(row[idx['breakMinutes']] || 0),
      netMinutes: Number(row[idx['netMinutes']] || 0),
    };
  }
  return map;
}
