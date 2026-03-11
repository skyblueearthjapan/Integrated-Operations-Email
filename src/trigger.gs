// ====== トリガー管理 ======

/**
 * 全トリガーを初期設定する（GASエディタから手動実行）
 */
function setupAllTriggers() {
  // 既存トリガーを全削除
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  // 朝メール: 毎日 07:00
  ScriptApp.newTrigger('sendMorningMail')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .nearMinute(0)
    .create();

  // 夕方メール①: 毎日 17:00
  ScriptApp.newTrigger('sendEveningMail1')
    .timeBased()
    .everyDays(1)
    .atHour(17)
    .nearMinute(0)
    .create();

  // 夕方メール②: 毎日 18:00
  ScriptApp.newTrigger('sendEveningMail2')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .nearMinute(0)
    .create();

  Logger.log('統合メール トリガー設定完了: 朝07:00、夕方17:00、夕方18:00');
}

/**
 * 全トリガーを削除する
 */
function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log('全トリガー削除完了');
}

// ====== 手動テスト用 ======

/**
 * 朝メールのテスト送信
 */
function testMorningMail() {
  sendMorningMail();
}

/**
 * 夕方メール①のテスト送信
 */
function testEveningMail1() {
  sendEveningMail1();
}

/**
 * 夕方メール②のテスト送信
 */
function testEveningMail2() {
  sendEveningMail2();
}
