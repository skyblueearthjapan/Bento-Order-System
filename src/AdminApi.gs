/**
 * AdminApi.gs - 総務設定画面（?page=admin）から呼ばれるAPI
 * すべて「設定できる人」（admin_users シート）だけが使える
 */

/**
 * 設定画面の初期データ
 */
function api_admin_bootstrap() {
  MenuService.assertAdmin();
  return _adminData_();
}

function _adminData_() {
  SheetService.initSheets();
  var ctx = MenuService.loadContext();
  var all = SheetService.getReservationsInRange('0000-00-00', '9999-99-99');
  var counts = {};  // key: menu|date → 件数
  all.forEach(function(r) {
    var k = r.orderState + '|' + r.date;
    counts[k] = (counts[k] || 0) + 1;
  });
  function countRange(menu, from, to) {
    var n = 0;
    Object.keys(counts).forEach(function(k) {
      var p = k.split('|');
      if (p[0] === menu && p[1] >= from && p[1] <= to) n += counts[k];
    });
    return n;
  }
  return {
    me: MenuService.currentEmail(),
    today: formatDateYmd(new Date()),
    nowIso: new Date().toISOString(),
    menus: MENU_DEFS,
    periods: ctx.periods.map(function(p) {
      return { id: p.id, menu: p.menu, name: p.name, from: p.from, to: p.to, active: p.active,
               flyerId: p.flyerId, flyerName: p.flyerName, count: countRange(p.menu, p.from, p.to) };
    }),
    tokus: ctx.tokus.map(function(t) {
      return { id: t.id, date: t.date, name: t.name, price: t.price, active: t.active,
               flyerId: t.flyerId, flyerName: t.flyerName, count: countRange('toku', t.date, t.date) };
    }),
    workers: SheetService.getWorkers().map(function(w) { return { code: w.code, name: w.name, dept: w.dept, location: w.location }; }),
    calendar: SheetService.getCalendar(),
    log: MenuService.getLog(50)
  };
}

/**
 * チラシ（PDF・画像）を保存する
 * @param {Object} f {name, mimeType, base64}
 */
function api_admin_uploadFlyer(f) {
  MenuService.assertAdmin();
  var res = MenuService.saveFlyer(f);
  MenuService.log('チラシを保存：' + res.fileName + (res.shared ? '' : '（社内共有の設定に失敗。ドライブで共有設定をご確認ください）'));
  return res;
}

// ===== 麺・丼の受付期間 =====

/**
 * 期間を追加・更新
 * @param {Object} p {id?, menu, name, from, to, active, flyerId, flyerName}
 * @param {String} impact 期間から外れる予約の扱い（'keep' | 'bento' | 'cancel'）。未指定で外れる予約があれば確認を返す
 */
function api_admin_savePeriod(p, impact) {
  MenuService.assertAdmin();
  if (!p || (p.menu !== 'men' && p.menu !== 'don')) throw new Error('メニューを選んでください。');
  var from = MenuService.toYmd(p.from), to = MenuService.toYmd(p.to);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) throw new Error('開始日と終了日を入れてください。');
  if (to < from) throw new Error('終了日が開始日より前になっています。');
  var label = getMenuDef(p.menu).label;
  var text = (p.id ? '期間を編集：' : '期間を追加：') + label + (p.name ? '「' + p.name + '」' : '') + ' ' + from + '〜' + to + (p.active ? '' : '（停止中）');
  return _withLockAndImpact_(impact, text, function() {
    var sheet = getOrCreateSheet(SHEET_MENU_PERIODS, MENU_PERIOD_HEADERS);
    var rowVals = [p.id || _newId_('P'), p.menu, String(p.name || ''), from, to, p.active ? 'ON' : 'OFF',
                   String(p.flyerId || ''), String(p.flyerName || ''), formatDateTime(new Date()), MenuService.currentEmail()];
    var target = p.id ? _findRowById_(sheet, p.id, MENU_PERIOD_HEADERS.length) : 0;
    if (p.id && !target) throw new Error('この期間は、ほかの人が削除したようです。画面を開き直してください。');
    _writeRow_(sheet, target, rowVals, [4, 5]);
  });
}

function api_admin_setPeriodActive(id, active, impact) {
  MenuService.assertAdmin();
  return _withLockAndImpact_(impact, (active ? '期間を有効化：' : '期間を一時停止：') + id, function() {
    var sheet = getOrCreateSheet(SHEET_MENU_PERIODS, MENU_PERIOD_HEADERS);
    var row = _findRowById_(sheet, id, MENU_PERIOD_HEADERS.length);
    if (!row) throw new Error('期間が見つかりません。');
    sheet.getRange(row, 6).setValue(active ? 'ON' : 'OFF');
    sheet.getRange(row, 9, 1, 2).setValues([[formatDateTime(new Date()), MenuService.currentEmail()]]);
  });
}

function api_admin_deletePeriod(id, impact) {
  MenuService.assertAdmin();
  return _withLockAndImpact_(impact, '期間を削除：' + id, function() {
    var sheet = getOrCreateSheet(SHEET_MENU_PERIODS, MENU_PERIOD_HEADERS);
    var row = _findRowById_(sheet, id, MENU_PERIOD_HEADERS.length);
    if (row) sheet.deleteRow(row);
  });
}

// ===== 特弁 =====

/**
 * 特弁を追加・更新
 * @param {Object} t {id?, date, name, price, flyerId, flyerName}
 */
function api_admin_saveToku(t, impact) {
  MenuService.assertAdmin();
  if (!t) throw new Error('入力が空です。');
  var date = MenuService.toYmd(t.date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('日付を入れてください。');
  if (!String(t.name || '').trim()) throw new Error('名前を入れてください。');
  var cal = MenuService.getCalendarMap();
  if (MenuService.isHoliday(date, cal)) throw new Error(date + ' は休日です。営業日を選んでください。');
  var price = Math.max(0, Math.round(Number(t.price) || 0));
  var dup = MenuService.getTokus().filter(function(x) { return x.active && x.date === date && x.id !== String(t.id || ''); })[0];
  if (dup) throw new Error(date + ' にはすでに特弁「' + dup.name + '」があります。1日に1つまでです。');
  var text = (t.id ? '特弁を編集：' : '特弁を追加：') + date + ' ' + t.name + ' ¥' + price;
  return _withLockAndImpact_(impact, text, function() {
    var sheet = getOrCreateSheet(SHEET_TOKUBEN, TOKUBEN_HEADERS);
    var rowVals = [t.id || _newId_('T'), date, String(t.name).trim(), price, 'ON',
                   String(t.flyerId || ''), String(t.flyerName || ''), formatDateTime(new Date()), MenuService.currentEmail()];
    var target = t.id ? _findRowById_(sheet, t.id, TOKUBEN_HEADERS.length) : 0;
    if (t.id && !target) throw new Error('この特弁は、ほかの人が削除したようです。画面を開き直してください。');
    _writeRow_(sheet, target, rowVals, [2]);
  });
}

function api_admin_deleteToku(id, impact) {
  MenuService.assertAdmin();
  return _withLockAndImpact_(impact, '特弁を削除：' + id, function() {
    var sheet = getOrCreateSheet(SHEET_TOKUBEN, TOKUBEN_HEADERS);
    var row = _findRowById_(sheet, id, TOKUBEN_HEADERS.length);
    if (row) sheet.deleteRow(row);
  });
}

// ===== 締切後の予約の修正 =====

/**
 * 予約を1件取得（修正欄の「いまの予約」表示用）
 */
function api_admin_getReservation(workerCode, date) {
  MenuService.assertAdmin();
  var r = SheetService.getReservationsInRange(date, date).filter(function(x) { return x.workerCode === String(workerCode); })[0];
  var res = { orderState: r ? r.orderState : '', location: r ? r.location : '', ordered: false };
  if (r && r.orderState) {
    var cal = MenuService.getCalendarMap();
    res.ordered = _isAlreadyOrdered_(r.orderState, date, cal);
  }
  return res;
}

/**
 * 総務が締切後の予約を直す
 * @param {Object} fix {workerCode, date, orderState, reason}
 */
function api_admin_fixReservation(fix) {
  MenuService.assertAdmin();
  if (!fix || !String(fix.reason || '').trim()) throw new Error('理由を入れてください（履歴に残ります）。');
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error('他の人が更新中です。数秒後にもう一度お試しください。');
  try {
    var ctx = MenuService.loadContext();
    var worker = SheetService.getWorkerByCode(fix.workerCode);
    if (!worker) throw new Error('作業員が見つかりません。');
    var v = Validation.validateAdminFix({ workerCode: fix.workerCode, date: fix.date, orderState: fix.orderState || '' }, ctx);
    if (!v.valid) throw new Error(v.message);
    var cur = SheetService.getReservationsInRange(fix.date, fix.date).filter(function(x) { return x.workerCode === String(fix.workerCode); })[0];
    var from = cur ? cur.orderState : '';
    var to = fix.orderState || '';
    if (from === to) throw new Error('いまの予約と同じです。');
    SheetService.upsertReservation({
      workerCode: worker.code,
      workerName: worker.name,
      date: fix.date,
      orderState: to,
      location: (cur && cur.location) || worker.location
    });
    logOperation_('admin_fix', { workerCode: worker.code, date: fix.date, orderState: to, location: (cur && cur.location) || worker.location });
    var lbl = function(k) { return k ? getMenuDef(k).label : 'なし（取消）'; };
    MenuService.log('締切後の修正：' + fix.date + ' ' + worker.name + ' ' + lbl(from) + ' → ' + lbl(to) + '（理由：' + String(fix.reason).trim() + '）');
    var ordered = from ? _isAlreadyOrdered_(from, fix.date, ctx.calMap) : false;
    return { ok: true, ordered: ordered, fromLabel: from ? lbl(from) : '', data: _adminData_() };
  } finally {
    lock.releaseLock();
  }
}

// ===== 共通 =====

/** そのメニューが、すでに西原屋さんへ注文済みの時刻を過ぎているか */
function _isAlreadyOrdered_(menuKey, date, calMap) {
  var now = new Date();
  if (MenuService.isEarly(menuKey)) return now >= MenuService.earlyDeadline(date, calMap);
  if (MenuService.isWorkSat(date, calMap)) return now >= new Date(MenuService.addDays(date, -1) + 'T15:00:00+09:00');
  return now >= new Date(date + 'T08:00:00+09:00');
}

/**
 * 設定を変えたときに、受付から外れる予約を確かめてから反映する
 * - impact 未指定で外れる予約があれば、反映せずに {needConfirm:true, affected:[...]} を返す
 * - impact: 'keep' 残す / 'bento' 弁当に切替 / 'cancel' 取消
 */
function _withLockAndImpact_(impactArg, logText, applyFn) {
  // impactArg: null（まず確かめる）または {how:'keep'|'bento'|'cancel', keys:[確認画面で見せた予約のキー]}
  var impact = null, confirmedKeys = null;
  if (impactArg) {
    impact = typeof impactArg === 'string' ? impactArg : impactArg.how;
    confirmedKeys = typeof impactArg === 'string' ? null : (impactArg.keys || []);
    if (['keep', 'bento', 'cancel'].indexOf(impact) === -1) throw new Error('外れる予約の扱いが正しくありません。');
  }
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error('他の人が更新中です。数秒後にもう一度お試しください。');
  try {
    var today = formatDateYmd(new Date());
    var before = MenuService.loadContext();
    var reservations = SheetService.getReservationsInRange(today, '9999-99-99')
      .filter(function(r) { return MenuService.isEarly(r.orderState); });
    var offeredBefore = {};
    reservations.forEach(function(r) { offeredBefore[r.workerCode + '_' + r.date] = MenuService.isOffered(r.orderState, r.date, before); });

    // 反映してから、外れる予約を数える（確認が必要なら元に戻せるよう、シートの中身を控えておく）
    var backup = _backupSettingSheets_();
    applyFn();
    SpreadsheetApp.flush();
    var after = MenuService.loadContext();
    var affected = reservations.filter(function(r) {
      return offeredBefore[r.workerCode + '_' + r.date] && !MenuService.isOffered(r.orderState, r.date, after);
    });

    // 確認画面を見ている間に対象の予約が変わっていたら、もう一度確かめてもらう
    var keysNow = affected.map(function(r) { return r.workerCode + '_' + r.date; }).sort();
    var changedSinceConfirm = confirmedKeys !== null &&
      JSON.stringify(keysNow) !== JSON.stringify(confirmedKeys.slice().sort());
    if (affected.length > 0 && (!impact || changedSinceConfirm)) {
      _restoreSettingSheets_(backup);
      var wm = SheetService.getWorkersMap();
      return {
        ok: false,
        needConfirm: true,
        affected: affected.map(function(r) {
          var name = r.workerName || (wm[r.workerCode] ? wm[r.workerCode].name : r.workerCode);
          return { key: r.workerCode + '_' + r.date, date: r.date, name: name, menu: r.orderState, ordered: _isAlreadyOrdered_(r.orderState, r.date, after.calMap) };
        }),
        changed: changedSinceConfirm
      };
    }

    if (affected.length > 0 && impact !== 'keep') {
      affected.forEach(function(r) {
        SheetService.upsertReservation({
          workerCode: r.workerCode, workerName: r.workerName, date: r.date,
          orderState: impact === 'bento' ? ORDER_STATE.BENTO : '',
          location: r.location
        });
        logOperation_('admin_period_' + impact, { workerCode: r.workerCode, date: r.date, orderState: impact === 'bento' ? ORDER_STATE.BENTO : '', location: r.location });
      });
    }
    var how = { keep: '残す', bento: '弁当に切替', cancel: '取消' }[impact] || '';
    MenuService.log(logText + (affected.length > 0 ? '（外れた予約' + affected.length + '件を' + how + '）' : ''));
    return { ok: true, data: _adminData_() };
  } finally {
    lock.releaseLock();
  }
}

function _backupSettingSheets_() {
  return [SHEET_MENU_PERIODS, SHEET_TOKUBEN].map(function(name) {
    var sh = getOrCreateSheet(name, name === SHEET_MENU_PERIODS ? MENU_PERIOD_HEADERS : TOKUBEN_HEADERS);
    var last = sh.getLastRow(), cols = sh.getLastColumn();
    return { name: name, values: last >= 1 && cols >= 1 ? sh.getRange(1, 1, last, cols).getValues() : [] };
  });
}

function _restoreSettingSheets_(backup) {
  backup.forEach(function(b) {
    var sh = getSourceSpreadsheet().getSheetByName(b.name);
    if (!sh) return;
    sh.clearContents();
    if (b.values.length > 0) sh.getRange(1, 1, b.values.length, b.values[0].length).setValues(b.values);
  });
  SpreadsheetApp.flush();
}

function _newId_(prefix) {
  return prefix + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + Math.floor(Math.random() * 1000);
}

function _findRowById_(sheet, id, cols) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id)) return i + 2;
  }
  return 0;
}

/** 行を書き込む（row=0 なら末尾に追加）。dateCols の列は文字列のまま保存する */
function _writeRow_(sheet, row, values, dateCols) {
  var r = row || sheet.getLastRow() + 1;
  (dateCols || []).forEach(function(c) { sheet.getRange(r, c).setNumberFormat('@'); });
  sheet.getRange(r, 1, 1, values.length).setValues([values]);
}
