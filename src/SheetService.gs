/**
 * SheetService.gs - スプレッドシートI/O層
 * 全ての読み書きはこのモジュールを経由する
 */

var SheetService = (function() {

  // ===== シート初期化 =====

  function initSheets() {
    getOrCreateSheet(SHEET_WORKERS, WORKER_HEADERS);
    getOrCreateSheet(SHEET_CALENDAR, CALENDAR_HEADERS);
    getOrCreateSheet(SHEET_RESERVATIONS, RESERVATION_HEADERS);
    getOrCreateSheet(SHEET_EMAIL_RECIPIENTS, ['名前', 'メールアドレス', '有効フラグ']);
    getOrCreateSheet(SHEET_NOTICES, ['種別', '内容', '有効期限', '有効フラグ']);
    getOrCreateSheet(SHEET_OPERATION_LOG, ['日時', 'ユーザー', 'アクション', '作業員コード', '日付', '注文状態', '配達拠点']);
    getOrCreateSheet(SHEET_CONFIG, ['キー', '値']);
  }

  // ===== 作業員マスタ =====

  function getWorkers() {
    var sheet = getOrCreateSheet(SHEET_WORKERS, WORKER_HEADERS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    return values
      .filter(function(r) { return r[0]; })
      .map(function(r) {
        return {
          code: String(r[0]).trim(),
          name: String(r[1]).trim(),
          dept: String(r[2]).trim(),
          location: String(r[4]).trim(),
          staffType: String(r[5]).trim()
        };
      });
  }

  function getWorkerByCode(code) {
    var all = getWorkers();
    for (var i = 0; i < all.length; i++) {
      if (all[i].code === String(code)) return all[i];
    }
    return null;
  }

  function getWorkersMap() {
    var all = getWorkers();
    var map = {};
    all.forEach(function(w) { map[w.code] = w; });
    return map;
  }

  // ===== カレンダーマスタ =====

  function getCalendar() {
    var sheet = getOrCreateSheet(SHEET_CALENDAR, CALENDAR_HEADERS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    return values
      .filter(function(r) { return r[0]; })
      .map(function(r) {
        var dateStr = r[0] instanceof Date ? formatDateYmd(r[0]) : String(r[0]).trim();
        return {
          date: dateStr,
          type: String(r[1] || '').trim(),
          weekday: String(r[2] || '').trim(),
          note: String(r[3] || '').trim()
        };
      });
  }

  function isHoliday(dateStr) {
    if (!dateStr) return false;
    var cal = getCalendar();
    var entry = cal.filter(function(c) { return c.date === dateStr; })[0];
    if (entry) {
      if (entry.type === CALENDAR_TYPE.HOLIDAY || entry.type === CALENDAR_TYPE.NATIONAL) return true;
      if (entry.type === CALENDAR_TYPE.WORK_SAT) return false;
    }
    var d = new Date(dateStr + 'T00:00:00+09:00');
    var dow = d.getDay();
    if (dow === 0) return true;  // 日曜
    if (dow === 6) return true;  // 土曜（出勤土曜マスタ該当なし）
    return false;
  }

  function isWorkSaturday(dateStr) {
    var cal = getCalendar();
    var entry = cal.filter(function(c) { return c.date === dateStr; })[0];
    return !!(entry && entry.type === CALENDAR_TYPE.WORK_SAT);
  }

  // ===== 予約データ =====

  function _readAllReservations() {
    var sheet = getOrCreateSheet(SHEET_RESERVATIONS, RESERVATION_HEADERS);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { sheet: sheet, rows: [] };
    var values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    var rows = values.map(function(r, i) {
      return {
        _row: i + 2,
        workerCode: String(r[0] || '').trim(),
        workerName: String(r[1] || '').trim(),
        date: r[2] instanceof Date ? formatDateYmd(r[2]) : String(r[2] || '').trim(),
        orderState: String(r[3] || '').trim(),
        location: String(r[4] || '').trim(),
        updatedAt: r[5] instanceof Date ? formatDateTime(r[5]) : String(r[5] || ''),
        updatedBy: String(r[6] || '').trim()
      };
    }).filter(function(r) { return r.workerCode && r.date; });
    return { sheet: sheet, rows: rows };
  }

  function getReservationsInRange(fromDate, toDate) {
    var data = _readAllReservations();
    return data.rows.filter(function(r) {
      return r.date >= fromDate && r.date <= toDate;
    }).map(function(r) {
      return {
        workerCode: r.workerCode,
        workerName: r.workerName,
        date: r.date,
        orderState: r.orderState,
        location: r.location,
        updatedAt: r.updatedAt,
        updatedBy: r.updatedBy
      };
    });
  }

  function getUserReservations(workerCode, fromDate, toDate) {
    var data = _readAllReservations();
    return data.rows.filter(function(r) {
      return r.workerCode === String(workerCode) && r.date >= fromDate && r.date <= toDate;
    });
  }

  function upsertReservation(change) {
    var data = _readAllReservations();
    var sheet = data.sheet;
    var existingRow = null;
    for (var i = 0; i < data.rows.length; i++) {
      if (data.rows[i].workerCode === String(change.workerCode) && data.rows[i].date === change.date) {
        existingRow = data.rows[i];
        break;
      }
    }

    var email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
    var now = formatDateTime(new Date());

    // 注文状態が空ならキャンセル扱い: 既存行があれば削除してnullを返す
    if (!change.orderState || change.orderState === '') {
      if (existingRow) {
        sheet.deleteRow(existingRow._row);
      }
      return null;  // キャンセルはnullを返しクライアント側でstateから削除
    }

    if (existingRow) {
      // 更新
      var row = existingRow._row;
      sheet.getRange(row, RESERVATION_COL.ORDER_STATE).setValue(change.orderState);
      sheet.getRange(row, RESERVATION_COL.LOCATION).setValue(change.location || '');
      sheet.getRange(row, RESERVATION_COL.UPDATED_AT).setValue(now);
      sheet.getRange(row, RESERVATION_COL.UPDATED_BY).setValue(email);
      if (change.workerName) {
        sheet.getRange(row, RESERVATION_COL.WORKER_NAME).setValue(change.workerName);
      }
    } else {
      // 新規追加
      var dateCell = change.date;
      var newRow = [
        change.workerCode,
        change.workerName || '',
        dateCell,
        change.orderState,
        change.location || '',
        now,
        email
      ];
      sheet.appendRow(newRow);
      // 日付セルを文字列フォーマット固定
      var appendedRowNum = sheet.getLastRow();
      sheet.getRange(appendedRowNum, RESERVATION_COL.DATE).setNumberFormat('@');
      sheet.getRange(appendedRowNum, RESERVATION_COL.DATE).setValue(dateCell);
    }

    return {
      workerCode: change.workerCode,
      workerName: change.workerName || '',
      date: change.date,
      orderState: change.orderState,
      location: change.location || '',
      updatedAt: now,
      updatedBy: email
    };
  }

  // ===== お知らせ =====

  function getNotice() {
    var sheet = getOrCreateSheet(SHEET_NOTICES, ['種別', '内容', '有効期限', '有効フラグ']);
    var lastRow = sheet.getLastRow();
    var result = { notice: '', warning: '' };
    if (lastRow < 2) return result;
    var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    var today = formatDateYmd(new Date());
    values.forEach(function(r) {
      var type = String(r[0] || '').trim();
      var content = String(r[1] || '').trim();
      var expire = r[2] instanceof Date ? formatDateYmd(r[2]) : String(r[2] || '').trim();
      var enabled = String(r[3] || '').trim().toUpperCase();
      if (enabled !== 'ON') return;
      if (expire && expire < today) return;
      if (type === 'notice') result.notice = content;
      if (type === 'warning') result.warning = content;
    });
    return result;
  }

  // ===== メール送信先 =====

  function getMailRecipients() {
    var sheet = getOrCreateSheet(SHEET_EMAIL_RECIPIENTS, ['名前', 'メールアドレス', '有効フラグ']);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    return values
      .filter(function(r) { return r[1] && String(r[2] || '').trim().toUpperCase() === 'ON'; })
      .map(function(r) { return String(r[1]).trim(); });
  }

  return {
    initSheets: initSheets,
    getWorkers: getWorkers,
    getWorkerByCode: getWorkerByCode,
    getWorkersMap: getWorkersMap,
    getCalendar: getCalendar,
    isHoliday: isHoliday,
    isWorkSaturday: isWorkSaturday,
    getReservationsInRange: getReservationsInRange,
    getUserReservations: getUserReservations,
    upsertReservation: upsertReservation,
    getNotice: getNotice,
    getMailRecipients: getMailRecipients
  };
})();
