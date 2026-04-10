/**
 * Code.gs - GASエントリポイント & 公開API
 * お弁当予約アプリのサーバー側エントリ
 */

// ===== Webアプリ エントリポイント =====

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  var todayStr = formatDateYmd(new Date());
  var currentUser = '';
  try {
    currentUser = Session.getActiveUser().getEmail() || '';
  } catch (err) {
    currentUser = '';
  }
  template.appTitle = APP_TITLE;
  template.portalUrl = PORTAL_URL;
  template.todayStr = todayStr;
  template.currentUserEmail = currentUser;

  return template.evaluate()
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * HTMLインクルードヘルパー
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===== 公開API =====

/**
 * 初回ロード時の一括データ取得
 */
function api_bootstrap() {
  var lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(5000)) {
      throw new Error('他のユーザーがデータを更新中です。数秒後に再試行してください。');
    }

    // 初回のみシート初期化（二回目以降はシートが既存なのでgetOrCreateSheetは軽い）
    SheetService.initSheets();

    var today = new Date();
    var fromDate = new Date(today);
    fromDate.setDate(fromDate.getDate() - 45);  // 過去45日
    var toDate = new Date(today);
    toDate.setDate(toDate.getDate() + 45);      // 未来45日

    var workers = SheetService.getWorkers();
    var calendar = SheetService.getCalendar();
    var reservations = SheetService.getReservationsInRange(formatDateYmd(fromDate), formatDateYmd(toDate));
    var noticeData = SheetService.getNotice();
    var recipients = [];
    try {
      recipients = SheetService.getMailRecipients();
    } catch (e) {
      recipients = [];
    }

    var currentUser = { email: '', name: '' };
    try {
      var u = Session.getActiveUser();
      currentUser.email = u.getEmail() || '';
    } catch (e) {}

    return {
      workers: workers,
      calendar: calendar,
      reservations: reservations,
      notice: noticeData.notice || DEFAULT_NOTICE,
      warning: noticeData.warning || DEFAULT_WARNING,
      rules: {
        deadlineHour: RULES.DEADLINE_HOUR,
        satDeadlineHour: RULES.SATURDAY_DEADLINE_HOUR,
        reservationWindowDays: RULES.RESERVATION_WINDOW_DAYS,
        monthStartDay: RULES.MONTH_START_DAY
      },
      currentUser: currentUser,
      today: formatDateYmd(today)
    };
  } catch (err) {
    Logger.log('api_bootstrap error: ' + err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 予約変更の一括保存
 * @param {Array} changes - [{workerCode, workerName, date, orderState, location}, ...]
 * @return {Array} [{success, workerCode, date, row, error}, ...]
 */
function api_applyPatch(changes) {
  var lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(15000)) {
      throw new Error('他のユーザーが更新中です。数秒後に再試行してください。');
    }
    if (!changes || !Array.isArray(changes) || changes.length === 0) {
      return [];
    }

    Logger.log('api_applyPatch: ' + changes.length + ' changes');

    var results = [];
    changes.forEach(function(change) {
      try {
        var validation = Validation.validateChange(change);
        if (!validation.valid) {
          results.push({
            success: false,
            workerCode: change.workerCode,
            date: change.date,
            row: null,
            error: validation.message
          });
          return;
        }
        var row = SheetService.upsertReservation(change);
        logOperation_('upsert', change);
        results.push({
          success: true,
          workerCode: change.workerCode,
          date: change.date,
          row: row,
          error: null
        });
      } catch (err) {
        Logger.log('applyPatch row error: ' + err.message);
        results.push({
          success: false,
          workerCode: change.workerCode,
          date: change.date,
          row: null,
          error: err.message
        });
      }
    });

    return results;
  } catch (err) {
    Logger.log('api_applyPatch error: ' + err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 期間指定で予約データを再取得（手動リフレッシュ用）
 */
function api_getReservationsInRange(fromDate, toDate) {
  var lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(5000)) {
      throw new Error('データ取得中です。数秒後に再試行してください。');
    }
    return SheetService.getReservationsInRange(fromDate, toDate);
  } finally {
    lock.releaseLock();
  }
}

// ===== 操作ログ =====

function logOperation_(action, data) {
  try {
    var sheet = getOrCreateSheet(SHEET_OPERATION_LOG, ['日時', 'ユーザー', 'アクション', '作業員コード', '日付', '注文状態', '配達拠点']);
    var email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
    sheet.appendRow([
      formatDateTime(new Date()),
      email,
      action,
      data.workerCode || '',
      data.date || '',
      data.orderState || '',
      data.location || ''
    ]);
  } catch (err) {
    Logger.log('logOperation_ error: ' + err.message);
  }
}
