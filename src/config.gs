// ====== 統合メール通知 CONFIG ======
var TZ = 'Asia/Tokyo';

// ★★★ 初回セットアップ時に設定が必要 ★★★
// 両アプリのコンテナバインドスプレッドシートID
var LEAVE_SS_ID = '16v1tQNCFkDTk2mnAZ4E2eEdHeUQ2Bepdcsn5vBAmllw';   // 休暇届けapp
var OT_SS_ID = '1Knx_kaQMZZams65J1oeSDaBeWUt8XXanNe94XSAHKFQ';   // 残業・休日出勤申請app

// メール送信先（両アプリ共通）
// ※ 各アプリのSettingsシートからも読み取り可能
var MAIL_TO = '';       // 空の場合はSettingsシートから取得

// ====== ヘルパー ======
function fmtDate_(d, pattern) {
  pattern = pattern || 'yyyy-MM-dd';
  return Utilities.formatDate(d, TZ, pattern);
}

function normalize_(s) {
  return String(s == null ? '' : s).trim()
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(c) {
      return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
    });
}

function fmtMinutesJa_(mins) {
  var m = Math.max(0, Number(mins || 0));
  var h = Math.floor(m / 60);
  var r = m % 60;
  if (h === 0) return r + '分';
  if (r === 0) return h + '時間';
  return h + '時間' + r + '分';
}

// ====== スプレッドシート接続 ======
var _leaveSS = null;
var _otSS = null;

function getLeaveSS_() {
  if (_leaveSS) return _leaveSS;
  if (!LEAVE_SS_ID) throw new Error('LEAVE_SS_ID が未設定です。config.gs を編集してください。');
  _leaveSS = SpreadsheetApp.openById(LEAVE_SS_ID);
  return _leaveSS;
}

function getOtSS_() {
  if (_otSS) return _otSS;
  if (!OT_SS_ID) throw new Error('OT_SS_ID が未設定です。config.gs を編集してください。');
  _otSS = SpreadsheetApp.openById(OT_SS_ID);
  return _otSS;
}

// ====== 設定値取得 ======
var _leaveSettings = null;
var _otSettings = null;

function getLeaveSettings_() {
  if (_leaveSettings) return _leaveSettings;
  var sh = getLeaveSS_().getSheetByName('M_SYSTEM_SETTING');
  if (!sh) throw new Error('休暇届けSSに M_SYSTEM_SETTING シートがありません');
  var values = sh.getDataRange().getValues();
  var map = {};
  for (var r = 1; r < values.length; r++) {
    var key = String(values[r][0] || '').trim();
    var val = values[r][1];
    if (key) map[key] = val;
  }
  _leaveSettings = map;
  return map;
}

function getOtSettings_() {
  if (_otSettings) return _otSettings;
  var sh = getOtSS_().getSheetByName('Settings');
  if (!sh) throw new Error('残業SSに Settings シートがありません');
  var values = sh.getDataRange().getValues();
  var map = {};
  for (var r = 2; r < values.length; r++) {  // 残業app: 3行目以降がデータ
    var key = String(values[r][0] || '').trim();
    var val = values[r][1];
    if (key) map[key] = val;
  }
  _otSettings = map;
  return map;
}

// 送信先メールアドレス取得
function getMailTo_() {
  if (MAIL_TO) return MAIL_TO;
  // 残業appのHR_MAIL_TOを優先（総務向け統合メール）
  var otSettings = getOtSettings_();
  return normalize_(otSettings['HR_MAIL_TO'] || '');
}

function getSomuEmails_() {
  var leaveSettings = getLeaveSettings_();
  return normalize_(leaveSettings['SOMU_EMAILS'] || '');
}

function getLeaveAppUrl_() {
  return normalize_(getLeaveSettings_()['APP_URL'] || '');
}

function getOtAppUrl_() {
  return normalize_(getOtSettings_()['APP_URL'] || '');
}

// ====== シートヘッダインデックス構築 ======
function buildHeaderIndex_(header) {
  var idx = {};
  for (var i = 0; i < header.length; i++) {
    idx[header[i]] = i;
  }
  return idx;
}

function getSheetHeaderIndex_(ss, sheetName, headerRowNo) {
  headerRowNo = headerRowNo || 1;
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet not found: ' + sheetName);
  var maxCol = sh.getMaxColumns();
  if (maxCol === 0) return { sh: sh, header: [], idx: {} };
  var raw = sh.getRange(headerRowNo, 1, 1, maxCol).getValues()[0];
  var lastH = 0;
  for (var c = 0; c < raw.length; c++) {
    if (String(raw[c] || '').trim() !== '') lastH = c + 1;
  }
  if (lastH === 0) return { sh: sh, header: [], idx: {} };
  var header = [];
  for (var c = 0; c < lastH; c++) {
    header.push(normalize_(raw[c]));
  }
  return { sh: sh, header: header, idx: buildHeaderIndex_(header) };
}

// 年度計算（3/16始まり）
function computeFiscalYear_(dateObj) {
  var d = dateObj || new Date();
  var year = d.getFullYear();
  var month = d.getMonth() + 1;
  var day = d.getDate();
  if (month < 3 || (month === 3 && day < 16)) {
    return year - 1;
  }
  return year;
}
