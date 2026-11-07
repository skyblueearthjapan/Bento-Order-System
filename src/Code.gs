/**
 * Code.gs - GASエントリポイント & 公開API
 * お弁当予約アプリのサーバー側エントリ
 */

// ===== Webアプリ エントリポイント =====

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || '';
  var currentUser = MenuService.currentEmail();
  var appUrl = '';
  try { appUrl = ScriptApp.getService().getUrl() || ''; } catch (err) { appUrl = ''; }

  // 総務設定画面（?page=admin）。登録した人だけが開ける
  if (page === 'admin') {
    var admin = false;
    try { admin = MenuService.isAdmin(currentUser); } catch (err) { admin = false; }
    var at = HtmlService.createTemplateFromFile(admin ? 'admin' : 'denied');
    at.appTitle = APP_TITLE;
    at.appUrl = appUrl;
    at.currentUserEmail = currentUser;
    at.todayStr = formatDateYmd(new Date());
    // 総務設定はほかのサイトに埋め込ませない（予約画面からは target="_top" で開く）
    return at.evaluate()
      .setTitle(APP_TITLE + ' 総務設定')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  var template = HtmlService.createTemplateFromFile('index');
  template.appTitle = APP_TITLE;
  template.portalUrl = PORTAL_URL;
  template.todayStr = formatDateYmd(new Date());
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
    var todayStr = formatDateYmd(today);
    var fromDate = new Date(today);
    fromDate.setDate(fromDate.getDate() - 45);  // 過去45日
    var toDate = new Date(today);
    toDate.setDate(toDate.getDate() + 45);      // 未来45日

    var workers = SheetService.getWorkers();
    var calendar = SheetService.getCalendar();
    var reservations = SheetService.getReservationsInRange(formatDateYmd(fromDate), formatDateYmd(toDate));
    var noticeData = SheetService.getNotice();

    var menuCtx = MenuService.loadContext();
    var monthStart = formatDateYmd(new Date(today.getFullYear(), today.getMonth(), 1));
    var email = MenuService.currentEmail();
    var appUrl = '';
    try { appUrl = ScriptApp.getService().getUrl() || ''; } catch (e) {}

    return {
      workers: workers,
      calendar: calendar,
      reservations: reservations,
      notice: noticeData.notice || DEFAULT_NOTICE,
      warning: noticeData.warning || DEFAULT_WARNING,
      rules: {
        deadlineHour: RULES.DEADLINE_HOUR,
        satDeadlineHour: RULES.SATURDAY_DEADLINE_HOUR,
        earlyDeadlineHour: RULES.EARLY_DEADLINE_HOUR,
        reservationWindowDays: RULES.RESERVATION_WINDOW_DAYS,
        monthStartDay: RULES.MONTH_START_DAY
      },
      menus: MENU_DEFS,
      menuPeriods: MenuService.publicPeriods(menuCtx, monthStart),
      tokus: MenuService.publicTokus(menuCtx, monthStart),
      isAdmin: MenuService.isAdmin(email),
      adminUrl: appUrl ? appUrl + '?page=admin' : '',
      currentUser: { email: email, name: '' },
      today: todayStr
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
 * @return {Array} [{success, workerCode, date, row, error, permanent, current}, ...]
 *   失敗時の current はサーバーに残っている予約（なければ null）。画面はこれで表示を戻す。
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

    // ルール判定用のデータは1回だけ読む
    var ctx = MenuService.loadContext();
    ctx.existing = _reservationMapForChanges_(changes);

    var results = [];
    changes.forEach(function(change) {
      var key = change.workerCode + '_' + change.date;
      try {
        var validation = Validation.validateChange(change, ctx);
        if (!validation.valid) {
          results.push({
            success: false,
            workerCode: change.workerCode,
            date: change.date,
            row: null,
            error: validation.message,
            permanent: validation.permanent,
            current: ctx.existing[key] || null
          });
          return;
        }
        var row = SheetService.upsertReservation(change);
        ctx.existing[key] = row;
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
          error: err.message,
          permanent: false,
          current: ctx.existing[key] || null
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
 * 変更対象の日付範囲の予約を {code_date: 予約} で返す
 */
function _reservationMapForChanges_(changes) {
  var dates = changes.map(function(c) { return String(c.date || ''); }).filter(function(d) { return d; }).sort();
  var map = {};
  if (dates.length === 0) return map;
  SheetService.getReservationsInRange(dates[0], dates[dates.length - 1]).forEach(function(r) {
    map[r.workerCode + '_' + r.date] = r;
  });
  return map;
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
