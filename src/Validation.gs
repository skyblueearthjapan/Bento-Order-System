/**
 * Validation.gs - 入力検証・正規化
 */

var Validation = (function() {

  /**
   * 予約変更の検証
   * @param {Object} change - {workerCode, workerName, date, orderState, location}
   * @param {Object} ctx    - MenuService.loadContext() の結果に existing（{code_date: 予約}）を足したもの（省略可）
   * @return {Object} {valid: boolean, message: string, permanent: boolean}
   *   permanent=true は「何度送っても通らない」変更（締切後・期間外など）。画面はキューと下書きから外す。
   */
  function validateChange(change, ctx) {
    if (!change) {
      return _ng('変更データが空です', true);
    }
    if (!change.workerCode) {
      return _ng('作業員コードが必要です', true);
    }
    if (!change.date || !/^\d{4}-\d{2}-\d{2}$/.test(String(change.date))) {
      return _ng('日付形式が不正です (YYYY-MM-DD)', true);
    }
    var next = change.orderState || '';
    if (next !== '' && MENU_KEYS.indexOf(next) === -1) {
      return _ng('注文状態が不正です: ' + next, true);
    }
    if (change.location && change.location !== LOCATION.SHIN && change.location !== LOCATION.HONSHA) {
      return _ng('配達拠点が不正です: ' + change.location, true);
    }

    ctx = ctx || MenuService.loadContext();
    var now = new Date();
    var today = formatDateYmd(now);
    var date = change.date;

    // 過去日（今日より前）への変更は拒否
    if (date < today) {
      return _ng('過去の日付は変更できません', true);
    }
    // 予約できるのは本日から2週間まで
    if (date > MenuService.addDays(today, RULES.RESERVATION_WINDOW_DAYS)) {
      return _ng('予約できるのは2週間先までです', true);
    }
    // 休日に注文は入れられない（取消は通す）
    if (next !== '' && MenuService.isHoliday(date, ctx.calMap)) {
      return _ng('休日は注文できません', true);
    }

    var key = change.workerCode + '_' + date;
    var existing;
    if (ctx.existing) {
      existing = ctx.existing[key] || null;
    } else {
      // 呼び出し元が渡していなくても、締切後の変更制限が効くように自分で読む
      existing = SheetService.getReservationsInRange(date, date).filter(function(r) {
        return r.workerCode === String(change.workerCode);
      })[0] || null;
    }
    var prev = existing ? (existing.orderState || '') : '';
    var locationChanged = false;
    if (existing && change.location && MenuService.isEarly(prev)) {
      // 保存済みの拠点が空なら、作業員マスタの拠点と比べる
      var curLoc = existing.location;
      if (!curLoc) {
        var w = SheetService.getWorkerByCode(change.workerCode);
        curLoc = w ? w.location : '';
      }
      locationChanged = !!curLoc && change.location !== curLoc;
    }

    // 当日分は8時以降は変更不可（弁当・おかずのみの締切）
    if (date === today && now.getHours() >= RULES.DEADLINE_HOUR) {
      return _ng('当日分の締切(AM8:00)を過ぎています', true);
    }
    // 出勤土曜分は前日金曜15時以降は変更不可
    if (MenuService.isWorkSat(date, ctx.calMap)) {
      var friStr = MenuService.addDays(date, -1);
      if (today > friStr || (today === friStr && now.getHours() >= RULES.SATURDAY_DEADLINE_HOUR)) {
        return _ng('出勤土曜分の締切(前日金曜15:00)を過ぎています', true);
      }
    }

    // 麺・丼・特弁: 前の平日の朝8:00が締切。締切後は追加も変更・取消もできない
    var deadline = MenuService.earlyDeadline(date, ctx.calMap);
    var label = function(k) { var d = getMenuDef(k); return d ? d.label : k; };
    if (prev && MenuService.isEarly(prev) && (prev !== next || locationChanged) && now >= deadline) {
      return _ng(label(prev) + 'は西原屋さんへ注文済みのため変更できません（締切: 前の平日の朝8:00）', true);
    }
    // 同じメニューのまま（拠点だけの変更など）は受付の確認をしない。期間変更で残した予約も通す
    if (next && MenuService.isEarly(next) && prev !== next) {
      if (!MenuService.isOffered(next, date, ctx)) {
        return _ng(label(next) + 'はこの日は受付していません', true);
      }
      if (prev !== next && now >= deadline) {
        return _ng(label(next) + 'の締切（前の平日の朝8:00）を過ぎています', true);
      }
    }
    return { valid: true, message: '', permanent: false };
  }

  /**
   * 総務による締切後の修正の検証（締切は見ない。日付・メニューの受付だけ確かめる）
   */
  function validateAdminFix(change, ctx) {
    if (!change || !change.workerCode) return _ng('名前を選んでください', true);
    if (!change.date || !/^\d{4}-\d{2}-\d{2}$/.test(String(change.date))) return _ng('日付形式が不正です', true);
    var next = change.orderState || '';
    if (next !== '' && MENU_KEYS.indexOf(next) === -1) return _ng('注文状態が不正です: ' + next, true);
    ctx = ctx || MenuService.loadContext();
    if (next !== '' && MenuService.isHoliday(change.date, ctx.calMap)) return _ng('休日は注文できません', true);
    if (next !== '' && !MenuService.isOffered(next, change.date, ctx)) {
      return _ng(getMenuDef(next).label + 'はこの日は受付していません', true);
    }
    return { valid: true, message: '', permanent: false };
  }

  function _ng(message, permanent) {
    return { valid: false, message: message, permanent: !!permanent };
  }

  /**
   * 日付文字列の正規化
   */
  function normalizeDate(dateStr) {
    if (!dateStr) return '';
    if (dateStr instanceof Date) return formatDateYmd(dateStr);
    return String(dateStr).trim();
  }

  return {
    validateChange: validateChange,
    validateAdminFix: validateAdminFix,
    normalizeDate: normalizeDate
  };
})();
