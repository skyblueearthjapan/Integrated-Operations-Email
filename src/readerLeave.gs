// ====== 休暇届けapp データ読み取り ======

// ヘッダエイリアス（英語 → 日本語）
var HEADER_ALIAS_ = {
  'FY_START_YEAR': '年度',
  'STATUS': '承認状態',
  'APPLIED_AT': '申請日時',
  'DEPT_ID': '部署ID',
  'DEPT_NAME': '部署名',
  'WORKER_ID': '作業員ID',
  'WORKER_NAME': '作業員名',
  'LEAVE_DATE': '休暇日',
  'DAY_TYPE': '休暇区分',
  'HALF_TYPE': '半日区分',
  'LEAVE_KIND': '休暇種類',
  'SUBSTITUTE_FOR_WORK_DATE': '振替元出勤日',
  'SPECIAL_REASON_REF': '特別理由参照',
  'SPECIAL_REASON_TEXT': '特別理由',
  'APPROVED_AT': '承認日時',
  'APPROVED_AT2': '2次承認日時',
};

function buildLeaveHeaderIndex_(header) {
  var idx = {};
  var reverseAlias = {};
  for (var key in HEADER_ALIAS_) {
    if (HEADER_ALIAS_.hasOwnProperty(key)) {
      reverseAlias[HEADER_ALIAS_[key]] = key;
    }
  }
  for (var i = 0; i < header.length; i++) {
    var h = header[i];
    idx[h] = i;
    if (HEADER_ALIAS_[h] && idx[HEADER_ALIAS_[h]] === undefined) {
      idx[HEADER_ALIAS_[h]] = i;
    }
    if (reverseAlias[h] && idx[reverseAlias[h]] === undefined) {
      idx[reverseAlias[h]] = i;
    }
  }
  return idx;
}

/**
 * 本日の休暇者一覧を取得（承認済み、休暇日＝今日）
 * 追加: 本日提出された過去日休暇申請（承認済み）も拾う
 */
function readTodayLeaves_() {
  var ss = getLeaveSS_();
  var sh = ss.getSheetByName('T_LEAVE_REQUEST');
  if (!sh) return [];

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var rawHeader = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var header = rawHeader.map(function(h) { return normalize_(h); });
  var idx = buildLeaveHeaderIndex_(header);

  var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  var now = new Date();
  var todayYmd = fmtDate_(now, 'yyyy-MM-dd');
  var fyStartYmd = fmtDate_(getFiscalYearStartDate_(now), 'yyyy-MM-dd');
  var todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
  var todayEnd = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);

  var appliedAtCol = idx['申請日時'] !== undefined ? idx['申請日時'] : idx['APPLIED_AT'];

  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var status = normalize_(row[idx['承認状態']]);
    if (status !== '承認済') continue;

    var leaveDate = row[idx['休暇日']];
    if (!(leaveDate instanceof Date)) continue;

    var leaveYmd = fmtDate_(leaveDate, 'yyyy-MM-dd');
    var appliedAt = appliedAtCol !== undefined ? row[appliedAtCol] : null;
    var appliedYmd = '';
    if (appliedAt) {
      try {
        appliedYmd = appliedAt instanceof Date
          ? fmtDate_(appliedAt, 'yyyy-MM-dd')
          : fmtDate_(new Date(appliedAt), 'yyyy-MM-dd');
      } catch (e) { appliedYmd = ''; }
    }

    // 条件A: 既存 → 休暇日＝本日
    var matchToday = (leaveDate >= todayStart && leaveDate <= todayEnd);
    // 条件B: 追加 → 本日提出かつ過去日休暇（年度開始以降）
    var matchRetro = (appliedYmd === todayYmd) && (leaveYmd < todayYmd) && (leaveYmd >= fyStartYmd);
    if (!matchToday && !matchRetro) continue;

    var isRetroactive = computeIsRetroactive_(appliedAt, leaveYmd);
    var daysAgo = isRetroactive ? computeDaysAgo_(leaveYmd, todayYmd) : 0;

    out.push({
      deptName: normalize_(row[idx['部署名']]),
      workerName: normalize_(row[idx['作業員名']]),
      leaveType: normalize_(row[idx['休暇種類']]),
      leaveKubun: normalize_(row[idx['休暇区分']]),
      halfType: normalize_(row[idx['半日区分']]),
      leaveDate: fmtDate_(leaveDate, 'yyyy/MM/dd'),
      targetDate: leaveYmd,
      isRetroactive: isRetroactive,
      daysAgo: daysAgo,
    });
  }

  out.sort(function(a, b) {
    return (a.deptName + a.workerName).localeCompare(b.deptName + b.workerName, 'ja');
  });
  return out;
}

/**
 * 1次未承認の休暇申請を取得（申請中 + 休暇日が今日以降）
 * 追加: 本日提出された過去日休暇のうち未承認のものも拾う
 */
function readPendingApproval1_() {
  var ss = getLeaveSS_();
  var sh = ss.getSheetByName('T_LEAVE_REQUEST');
  if (!sh) return [];

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var rawHeader = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var header = rawHeader.map(function(h) { return normalize_(h); });
  var idx = buildLeaveHeaderIndex_(header);

  var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  var now = new Date();
  var todayYmd = fmtDate_(now, 'yyyy-MM-dd');
  var fyStartYmd = fmtDate_(getFiscalYearStartDate_(now), 'yyyy-MM-dd');
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var appliedAtCol = idx['申請日時'] !== undefined ? idx['申請日時'] : idx['APPLIED_AT'];

  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var status = normalize_(row[idx['承認状態']]);
    if (status !== '申請中') continue;

    var leaveDate = row[idx['休暇日']];
    if (!(leaveDate instanceof Date)) continue;
    var ld = new Date(leaveDate);
    ld.setHours(0, 0, 0, 0);
    var leaveYmd = fmtDate_(leaveDate, 'yyyy-MM-dd');

    var appliedAt = appliedAtCol !== undefined ? row[appliedAtCol] : null;
    var appliedYmd = '';
    if (appliedAt) {
      try {
        appliedYmd = appliedAt instanceof Date
          ? fmtDate_(appliedAt, 'yyyy-MM-dd')
          : fmtDate_(new Date(appliedAt), 'yyyy-MM-dd');
      } catch (e) { appliedYmd = ''; }
    }

    // 条件A: 既存 → 休暇日が今日以降
    var matchFuture = (ld >= today);
    // 条件B: 追加 → 本日提出かつ過去日休暇（年度開始以降）
    var matchRetro = (appliedYmd === todayYmd) && (leaveYmd < todayYmd) && (leaveYmd >= fyStartYmd);
    if (!matchFuture && !matchRetro) continue;

    var isRetroactive = computeIsRetroactive_(appliedAt, leaveYmd);
    var daysAgo = isRetroactive ? computeDaysAgo_(leaveYmd, todayYmd) : 0;

    out.push({
      deptName: normalize_(row[idx['部署名']]),
      workerName: normalize_(row[idx['作業員名']]),
      leaveDate: fmtDate_(leaveDate, 'yyyy/MM/dd'),
      leaveType: normalize_(row[idx['休暇種類']]),
      leaveKubun: normalize_(row[idx['休暇区分']]),
      targetDate: leaveYmd,
      isRetroactive: isRetroactive,
      daysAgo: daysAgo,
    });
  }

  out.sort(function(a, b) {
    return (a.deptName + a.workerName).localeCompare(b.deptName + b.workerName, 'ja');
  });
  return out;
}

/**
 * 2次承認待ちの休暇申請を取得（承認済 + 2次承認日時が空）
 */
function readPendingApproval2_() {
  var ss = getLeaveSS_();
  var sh = ss.getSheetByName('T_LEAVE_REQUEST');
  if (!sh) return [];

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var rawHeader = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var header = rawHeader.map(function(h) { return normalize_(h); });
  var idx = buildLeaveHeaderIndex_(header);

  var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  var at2Col = idx['2次承認日時'] !== undefined ? idx['2次承認日時'] : idx['APPROVED_AT2'];
  var approvedAtCol = idx['承認日時'] !== undefined ? idx['承認日時'] : idx['APPROVED_AT'];
  var appliedAtCol = idx['申請日時'] !== undefined ? idx['申請日時'] : idx['APPLIED_AT'];

  var now = new Date();
  var todayYmd = fmtDate_(now, 'yyyy-MM-dd');

  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var status = normalize_(row[idx['承認状態']]);
    if (status !== '承認済') continue;

    // 2次承認日時が空 → 未完了
    if (at2Col !== undefined && row[at2Col] && String(row[at2Col]).trim() !== '') continue;

    var leaveDate = row[idx['休暇日']];
    if (!(leaveDate instanceof Date)) continue;

    var leaveYmd = fmtDate_(leaveDate, 'yyyy-MM-dd');
    var appliedAt = appliedAtCol !== undefined ? row[appliedAtCol] : null;
    var isRetroactive = computeIsRetroactive_(appliedAt, leaveYmd);
    var daysAgo = isRetroactive ? computeDaysAgo_(leaveYmd, todayYmd) : 0;

    var approvedAt = '';
    if (approvedAtCol !== undefined && row[approvedAtCol]) {
      approvedAt = row[approvedAtCol] instanceof Date
        ? fmtDate_(row[approvedAtCol], 'yyyy/MM/dd HH:mm')
        : String(row[approvedAtCol]);
    }

    out.push({
      deptName: normalize_(row[idx['部署名']]),
      workerName: normalize_(row[idx['作業員名']]),
      leaveDate: fmtDate_(leaveDate, 'yyyy/MM/dd'),
      leaveType: normalize_(row[idx['休暇種類']]),
      leaveKubun: normalize_(row[idx['休暇区分']]),
      approvedAt: approvedAt,
      targetDate: leaveYmd,
      isRetroactive: isRetroactive,
      daysAgo: daysAgo,
    });
  }

  out.sort(function(a, b) {
    return (a.deptName + a.workerName).localeCompare(b.deptName + b.workerName, 'ja');
  });
  return out;
}

/**
 * 有給取得警告の未達者一覧を取得
 */
function readWarnWorkers_() {
  var ss = getLeaveSS_();

  // V_LEAVE_SUMMARY から取得
  var sh = ss.getSheetByName('V_LEAVE_SUMMARY');
  if (!sh) return { workers: [], currentRequired: 0, currentLabel: '' };

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { workers: [], currentRequired: 0, currentLabel: '' };

  var rawHeader = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var header = rawHeader.map(function(h) { return normalize_(h); });
  var idx = buildHeaderIndex_(header);

  var fy = computeFiscalYear_(new Date());
  var values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (Number(row[idx['年度']]) !== fy) continue;
    var warnLevel = normalize_(row[idx['警告レベル']]);
    if (!warnLevel) continue;

    out.push({
      deptName: normalize_(row[idx['部署名']]),
      workerName: normalize_(row[idx['作業員名']]),
      paidCount: Number(row[idx['有給回数']] || 0),
      warnLevel: warnLevel,
      lastPaidDate: normalize_(row[idx['最終有給日']]),
    });
  }

  // 現在の必要回数を算出
  var PAID_LEAVE_WARN = [
    { deadline: { month: 6, day: 30 }, required: 1, label: '6月末' },
    { deadline: { month: 8, day: 31 }, required: 2, label: '8月末' },
    { deadline: { month: 12, day: 31 }, required: 3, label: '12月末' },
    { deadline: { month: 2, day: 15 }, required: 4, label: '2/15' },
    { deadline: { month: 3, day: 15 }, required: 5, label: '年度末' },
  ];
  var now = new Date();
  var currentRequired = 0;
  var currentLabel = '';
  for (var w = 0; w < PAID_LEAVE_WARN.length; w++) {
    var pw = PAID_LEAVE_WARN[w];
    var dm = pw.deadline.month;
    var dd = pw.deadline.day;
    var dy = (dm > 3 || (dm === 3 && dd >= 16)) ? fy : fy + 1;
    var deadlineDate = new Date(dy, dm - 1, dd, 23, 59, 59);
    if (now > deadlineDate) {
      currentRequired = pw.required;
      currentLabel = pw.label;
    }
  }

  return {
    workers: out,
    currentRequired: currentRequired,
    currentLabel: currentLabel,
    fiscalYear: fy,
  };
}
