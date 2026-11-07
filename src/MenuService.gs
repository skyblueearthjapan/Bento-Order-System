/**
 * MenuService.gs - 麺・丼の受付期間 / 特弁 / 締切ルール / 総務設定の権限 / チラシ保存
 *
 * 締切ルール
 * - 弁当・おかずのみ: 当日 8:00（出勤土曜は前日金曜 15:00）… 従来どおり（Validation.gs）
 * - 麺・丼・特弁   : 前の平日の 8:00
 *     「前の平日」= 月〜金で、社内カレンダーの休日・祝日ではない日。出勤土曜は数えない。
 *     例) 木曜分 → 水曜 8:00 / 月曜分・出勤土曜分 → 金曜 8:00 / 祝日の翌日 → さらに前の平日 8:00
 *   締切の直後（8:00〜8:30）の日次メールで、その日が締切の分をまとめて総務へ送る。
 */

var MenuService = (function() {

  // ===== 日付ヘルパー（JST の YYYY-MM-DD 文字列で扱う） =====

  function _addDays(dateStr, n) {
    var d = new Date(dateStr + 'T12:00:00+09:00');
    d.setTime(d.getTime() + n * 24 * 60 * 60 * 1000);
    return formatDateYmd(d);
  }

  function _dow(dateStr) {
    return new Date(dateStr + 'T12:00:00+09:00').getUTCDay();  // 12:00 JST = 03:00 UTC なので UTC の曜日 = JST の曜日
  }

  function _toYmd(v) {
    if (v instanceof Date) return formatDateYmd(v);
    var s = String(v || '').trim().replace(/\//g, '-');
    var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (!m) return s;
    return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  }

  function _flag(v) {
    return String(v || '').trim().toUpperCase() === 'ON';
  }

  // ===== 社内カレンダー =====

  function getCalendarMap() {
    var map = {};
    SheetService.getCalendar().forEach(function(c) { map[c.date] = c; });
    return map;
  }

  function isWorkSat(dateStr, calMap) {
    var c = calMap[dateStr];
    return !!(c && c.type === CALENDAR_TYPE.WORK_SAT);
  }

  function isOffMaster(dateStr, calMap) {
    var c = calMap[dateStr];
    return !!(c && (c.type === CALENDAR_TYPE.HOLIDAY || c.type === CALENDAR_TYPE.NATIONAL));
  }

  /** 休日（注文できない日）か: マスタの休日・祝日、日曜、出勤土曜以外の土曜 */
  function isHoliday(dateStr, calMap) {
    var w = _dow(dateStr);
    if (isOffMaster(dateStr, calMap)) return true;
    if (w === 0) return true;
    if (w === 6 && !isWorkSat(dateStr, calMap)) return true;
    return false;
  }

  /** 平日の営業日か（月〜金で休日・祝日でない）。出勤土曜は含めない */
  function isWeekdayBiz(dateStr, calMap) {
    var w = _dow(dateStr);
    return w >= 1 && w <= 5 && !isOffMaster(dateStr, calMap);
  }

  /** 前の平日（YYYY-MM-DD） */
  function prevWeekdayBiz(dateStr, calMap) {
    var d = _addDays(dateStr, -1);
    for (var i = 0; i < 30; i++) {
      if (isWeekdayBiz(d, calMap)) return d;
      d = _addDays(d, -1);
    }
    return d;
  }

  /** 麺・丼・特弁の締切（Date） */
  function earlyDeadline(dateStr, calMap) {
    var pd = prevWeekdayBiz(dateStr, calMap);
    return new Date(pd + 'T' + ('0' + RULES.EARLY_DEADLINE_HOUR).slice(-2) + ':00:00+09:00');
  }

  /** 本日 8:00 が締切になる日（= 本日の朝メールで注文する日）の一覧 */
  function daysDueOn(todayStr, calMap) {
    if (!isWeekdayBiz(todayStr, calMap)) return [];
    var out = [];
    for (var i = 1; i <= 20; i++) {
      var d = _addDays(todayStr, i);
      if (isHoliday(d, calMap)) continue;
      if (prevWeekdayBiz(d, calMap) === todayStr) out.push(d);
      else if (prevWeekdayBiz(d, calMap) > todayStr) break;
    }
    return out;
  }

  // ===== 麺・丼の受付期間 =====

  function getPeriods() {
    var sheet = getOrCreateSheet(SHEET_MENU_PERIODS, MENU_PERIOD_HEADERS);
    var last = sheet.getLastRow();
    if (last < 2) return [];
    return sheet.getRange(2, 1, last - 1, MENU_PERIOD_HEADERS.length).getValues()
      .map(function(r, i) {
        return {
          _row: i + 2,
          id: String(r[0] || '').trim(),
          menu: String(r[1] || '').trim(),
          name: String(r[2] || '').trim(),
          from: _toYmd(r[3]),
          to: _toYmd(r[4]),
          active: _flag(r[5]),
          flyerId: String(r[6] || '').trim(),
          flyerName: String(r[7] || '').trim()
        };
      })
      .filter(function(p) { return p.id && (p.menu === 'men' || p.menu === 'don') && p.from && p.to; });
  }

  // ===== 特弁 =====

  function getTokus() {
    var sheet = getOrCreateSheet(SHEET_TOKUBEN, TOKUBEN_HEADERS);
    var last = sheet.getLastRow();
    if (last < 2) return [];
    return sheet.getRange(2, 1, last - 1, TOKUBEN_HEADERS.length).getValues()
      .map(function(r, i) {
        return {
          _row: i + 2,
          id: String(r[0] || '').trim(),
          date: _toYmd(r[1]),
          name: String(r[2] || '').trim(),
          price: Number(r[3]) || 0,
          active: _flag(r[4]),
          flyerId: String(r[5] || '').trim(),
          flyerName: String(r[6] || '').trim()
        };
      })
      .filter(function(t) { return t.id && t.date; });
  }

  /** ルール判定に必要なデータを1回で読む */
  function loadContext() {
    return { calMap: getCalendarMap(), periods: getPeriods(), tokus: getTokus() };
  }

  function periodOf(key, dateStr, ctx) {
    for (var i = 0; i < ctx.periods.length; i++) {
      var p = ctx.periods[i];
      if (p.active && p.menu === key && p.from <= dateStr && dateStr <= p.to) return p;
    }
    return null;
  }

  function tokuOn(dateStr, ctx) {
    for (var i = 0; i < ctx.tokus.length; i++) {
      var t = ctx.tokus[i];
      if (t.active && t.date === dateStr) return t;
    }
    return null;
  }

  function isEarly(key) {
    return EARLY_MENU_KEYS.indexOf(key) !== -1;
  }

  /** そのメニューをその日に受け付けているか（締切は見ない） */
  function isOffered(key, dateStr, ctx) {
    var def = getMenuDef(key);
    if (!def) return false;
    if (def.kind === 'always') return true;
    if (def.kind === 'day') return !!tokuOn(dateStr, ctx);
    return !!periodOf(key, dateStr, ctx);
  }

  /** その日に受付がある麺・丼・特弁があるか */
  function anyEarlyOffered(dateStr, ctx) {
    return EARLY_MENU_KEYS.some(function(k) { return isOffered(k, dateStr, ctx); });
  }

  /** 画面・メールで「麺・丼・特弁の案内」を出すか（今日以降に有効な設定が1つでもある） */
  function hasUpcoming(todayStr, ctx) {
    return ctx.periods.some(function(p) { return p.active && p.to >= todayStr; }) ||
           ctx.tokus.some(function(t) { return t.active && t.date >= todayStr; });
  }

  /** 画面へ渡す形 */
  function publicPeriods(ctx, fromStr) {
    return ctx.periods.filter(function(p) { return p.active && p.to >= fromStr; }).map(function(p) {
      return { id: p.id, menu: p.menu, name: p.name, from: p.from, to: p.to, flyerId: p.flyerId };
    });
  }
  function publicTokus(ctx, fromStr) {
    return ctx.tokus.filter(function(t) { return t.active && t.date >= fromStr; }).map(function(t) {
      return { id: t.id, date: t.date, name: t.name, price: t.price, flyerId: t.flyerId };
    });
  }

  // ===== 設定できる人 =====

  function _ensureAdminSheet() {
    var ss = getSourceSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_ADMIN_USERS);
    if (!sheet) {
      sheet = getOrCreateSheet(SHEET_ADMIN_USERS, ADMIN_USER_HEADERS);
      if (INITIAL_ADMIN_USERS.length > 0) {
        sheet.getRange(2, 1, INITIAL_ADMIN_USERS.length, ADMIN_USER_HEADERS.length).setValues(INITIAL_ADMIN_USERS);
      }
    }
    return sheet;
  }

  function getAdminEmails() {
    var sheet = _ensureAdminSheet();
    var last = sheet.getLastRow();
    if (last < 2) return [];
    return sheet.getRange(2, 1, last - 1, 3).getValues()
      .filter(function(r) { return r[0] && _flag(r[2]); })
      .map(function(r) { return String(r[0]).trim().toLowerCase(); });
  }

  function currentEmail() {
    try { return (Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) { return ''; }
  }

  function isAdmin(email) {
    email = (email === undefined ? currentEmail() : String(email || '').toLowerCase());
    if (!email) return false;
    return getAdminEmails().indexOf(email) !== -1;
  }

  function assertAdmin() {
    if (!isAdmin()) throw new Error('総務設定を使う権限がありません。');
  }

  // ===== 変更履歴 =====

  function log(text) {
    try {
      var sheet = getOrCreateSheet(SHEET_ADMIN_LOG, ADMIN_LOG_HEADERS);
      sheet.appendRow([formatDateTime(new Date()), currentEmail(), text]);
    } catch (e) {
      Logger.log('admin log error: ' + e.message);
    }
  }

  function getLog(limit) {
    var sheet = getOrCreateSheet(SHEET_ADMIN_LOG, ADMIN_LOG_HEADERS);
    var last = sheet.getLastRow();
    if (last < 2) return [];
    var n = Math.min(limit || 50, last - 1);
    var values = sheet.getRange(last - n + 1, 1, n, 3).getValues();
    return values.reverse().map(function(r) {
      return { at: r[0] instanceof Date ? formatDateTime(r[0]) : String(r[0]), who: String(r[1]), text: String(r[2]) };
    });
  }

  // ===== チラシ（PDF・画像）の保存 =====

  function _flyerFolder() {
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('flyer_folder_id');
    if (id) {
      try { return DriveApp.getFolderById(id); } catch (e) { /* 消された場合は作り直す */ }
    }
    var parent = DriveApp.getFolderById(ADMIN_OUTPUT_FOLDER_ID);
    var it = parent.getFoldersByName(FLYER_FOLDER_NAME);
    var folder = it.hasNext() ? it.next() : parent.createFolder(FLYER_FOLDER_NAME);
    props.setProperty('flyer_folder_id', folder.getId());
    return folder;
  }

  /**
   * チラシを保存し、社内の人が見られるように共有する
   * @param {Object} f {name, mimeType, base64}
   */
  function saveFlyer(f) {
    if (!f || !f.base64) throw new Error('ファイルが空です。');
    var mime = String(f.mimeType || '');
    if (!/^application\/pdf$|^image\//.test(mime)) throw new Error('PDF か画像のファイルを選んでください。');
    var bytes = Utilities.base64Decode(f.base64);
    if (bytes.length > FLYER_MAX_BYTES) throw new Error('ファイルが大きすぎます（10MBまで）。');
    var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    var blob = Utilities.newBlob(bytes, mime, stamp + '_' + String(f.name || 'flyer'));
    var file = _flyerFolder().createFile(blob);
    var shared = true;
    try {
      file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
    } catch (e) {
      shared = false;
      Logger.log('flyer setSharing failed: ' + e.message);
    }
    return { fileId: file.getId(), fileName: String(f.name || ''), shared: shared };
  }

  return {
    addDays: _addDays,
    toYmd: _toYmd,
    getCalendarMap: getCalendarMap,
    isWorkSat: isWorkSat,
    isHoliday: isHoliday,
    isWeekdayBiz: isWeekdayBiz,
    prevWeekdayBiz: prevWeekdayBiz,
    earlyDeadline: earlyDeadline,
    daysDueOn: daysDueOn,
    getPeriods: getPeriods,
    getTokus: getTokus,
    loadContext: loadContext,
    periodOf: periodOf,
    tokuOn: tokuOn,
    isEarly: isEarly,
    isOffered: isOffered,
    anyEarlyOffered: anyEarlyOffered,
    hasUpcoming: hasUpcoming,
    publicPeriods: publicPeriods,
    publicTokus: publicTokus,
    ensureAdminSheet: _ensureAdminSheet,
    isAdmin: isAdmin,
    assertAdmin: assertAdmin,
    currentEmail: currentEmail,
    log: log,
    getLog: getLog,
    saveFlyer: saveFlyer
  };
})();
