/**
 * Triggers.gs - 時間ベーストリガーの設定・管理
 * setupTriggers() を一度だけGASエディタから手動実行してトリガーを登録する
 */

/**
 * トリガーを初期設定（既存は全削除してから登録）
 */
function setupTriggers() {
  deleteAllTriggers();

  // 毎日8時台: マスタ同期
  ScriptApp.newTrigger('syncAll')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  // 毎日8:00〜8:30台: 日次レポート（集計 + メール送信）。当日締切8:00直後に本日分＋翌出勤土曜分（暫定）を送信
  // ※ nearMinute(15) で 8:15±15分（=8:00〜8:30）の枠に発火を限定。お弁当屋さんの電話時刻と被らないよう8:30までに送信
  ScriptApp.newTrigger('dailyReportJob')
    .timeBased()
    .atHour(8)
    .nearMinute(15)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  // 毎週金曜15時台: 翌日が出勤土曜なら最終確定メールを送信（変更の有無に関わらず必ず送信）
  ScriptApp.newTrigger('saturdayDeadlineMailJob')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(15)
    .inTimezone('Asia/Tokyo')
    .create();

  // 4時間ごと: 管理シート更新（予約変更の反映用、負荷軽減のため4時間間隔）
  // ※ everyHours(N) の有効値は 1/2/4/6/8/12 のみ。3は無効で毎時実行にフォールバックするため4を使用
  // ※ everyHours(N) はタイムゾーンに依存しない（appsscript.jsonのAsia/Tokyo設定で実行）
  ScriptApp.newTrigger('updateAdminSheetsForToday')
    .timeBased()
    .everyHours(4)
    .create();

  Logger.log('Triggers set up successfully');
  listTriggers();
}

/**
 * 既存トリガーを全削除
 */
function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });
  Logger.log('Deleted ' + triggers.length + ' triggers');
}

/**
 * 現在のトリガー一覧をログ出力
 */
function listTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('Current triggers (' + triggers.length + '):');
  triggers.forEach(function(t) {
    Logger.log('  - ' + t.getHandlerFunction() + ' (' + t.getEventType() + ')');
  });
}

/**
 * 初回セットアップ（GASエディタから手動実行）
 * シート初期化 + マスタ同期 + トリガー登録
 */
function firstTimeSetup() {
  Logger.log('=== First time setup started ===');
  if (PORTAL_URL.indexOf('__REPLACE_ME__') >= 0) {
    Logger.log('⚠ PORTAL_URL がプレースホルダのままです。Config.gs で本番URLに差し替えてください。');
  }
  SheetService.initSheets();
  Logger.log('Sheets initialized');
  MasterSync.syncAll();
  Logger.log('Master data synced');
  setupTriggers();
  Logger.log('=== First time setup completed ===');
}
