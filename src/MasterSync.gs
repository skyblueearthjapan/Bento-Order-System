/**
 * MasterSync.gs - 外部マスタシートからのデータ同期
 * 毎日8時台のトリガーで自動実行
 */

var MasterSync = (function() {

  /**
   * 外部マスタスプレッドシート内の指定gidのシートを取得
   */
  function _getSourceSheetByGid(gid) {
    var ss = SpreadsheetApp.openById(EXTERNAL_MASTER_SPREADSHEET_ID);
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === gid) return sheets[i];
    }
    return null;
  }

  /**
   * 作業員マスタの同期
   */
  function syncWorkers() {
    var src = _getSourceSheetByGid(MASTER_WORKER_GID);
    if (!src) {
      Logger.log('syncWorkers: source sheet not found (gid=' + MASTER_WORKER_GID + ')');
      return { count: 0, error: 'source sheet not found' };
    }
    var lastRow = src.getLastRow();
    var lastCol = Math.max(src.getLastColumn(), 6);
    if (lastRow < 2) {
      Logger.log('syncWorkers: no data');
      return { count: 0 };
    }
    var values = src.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var rows = values
      .filter(function(r) { return r[0]; })
      .map(function(r) {
        return [
          String(r[0]).trim(),  // 作業員コード
          String(r[1] || '').trim(),  // 氏名
          String(r[2] || '').trim(),  // 部署
          String(r[3] || '').trim(),  // 担当業務
          String(r[4] || '').trim(),  // 拠点
          String(r[5] || '').trim()   // スタッフ種類
        ];
      });

    var cacheSheet = getOrCreateSheet(SHEET_WORKERS, WORKER_HEADERS);
    var cacheLastRow = cacheSheet.getLastRow();
    if (cacheLastRow >= 2) {
      cacheSheet.getRange(2, 1, cacheLastRow - 1, 6).clearContent();
    }
    if (rows.length > 0) {
      cacheSheet.getRange(2, 1, rows.length, 6).setValues(rows);
    }
    Logger.log('syncWorkers: ' + rows.length + ' rows synced');
    return { count: rows.length };
  }

  /**
   * カレンダーマスタの同期
   */
  function syncCalendar() {
    var src = _getSourceSheetByGid(MASTER_CALENDAR_GID);
    if (!src) {
      Logger.log('syncCalendar: source sheet not found (gid=' + MASTER_CALENDAR_GID + ')');
      return { count: 0, error: 'source sheet not found' };
    }
    var lastRow = src.getLastRow();
    var lastCol = Math.max(src.getLastColumn(), 4);
    if (lastRow < 2) {
      Logger.log('syncCalendar: no data');
      return { count: 0 };
    }
    var values = src.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var rows = values
      .filter(function(r) { return r[0]; })
      .map(function(r) {
        var dateStr = r[0] instanceof Date ? formatDateYmd(r[0]) : String(r[0]).trim();
        return [
          dateStr,
          String(r[1] || '').trim(),
          String(r[2] || '').trim(),
          String(r[3] || '').trim()
        ];
      });

    var cacheSheet = getOrCreateSheet(SHEET_CALENDAR, CALENDAR_HEADERS);
    var cacheLastRow = cacheSheet.getLastRow();
    if (cacheLastRow >= 2) {
      cacheSheet.getRange(2, 1, cacheLastRow - 1, 4).clearContent();
    }
    if (rows.length > 0) {
      cacheSheet.getRange(2, 1, rows.length, 4).setValues(rows);
      cacheSheet.getRange(2, 1, rows.length, 1).setNumberFormat('@');
    }
    Logger.log('syncCalendar: ' + rows.length + ' rows synced');
    return { count: rows.length };
  }

  /**
   * 全マスタ同期（トリガーから呼ばれる）
   */
  function syncAll() {
    try {
      SheetService.initSheets();
      var w = syncWorkers();
      var c = syncCalendar();
      Logger.log('syncAll completed: workers=' + w.count + ', calendar=' + c.count);
      return { workers: w, calendar: c };
    } catch (err) {
      Logger.log('syncAll error: ' + err.message);
      throw err;
    }
  }

  return {
    syncWorkers: syncWorkers,
    syncCalendar: syncCalendar,
    syncAll: syncAll
  };
})();

/**
 * トリガーから呼び出されるラッパー関数
 */
function syncAll() {
  return MasterSync.syncAll();
}
