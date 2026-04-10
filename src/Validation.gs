/**
 * Validation.gs - 入力検証・正規化
 */

var Validation = (function() {

  /**
   * 予約変更の検証
   * @param {Object} change - {workerCode, workerName, date, orderState, location}
   * @return {Object} {valid: boolean, message: string}
   */
  function validateChange(change) {
    if (!change) {
      return { valid: false, message: '変更データが空です' };
    }
    if (!change.workerCode) {
      return { valid: false, message: '作業員コードが必要です' };
    }
    if (!change.date || !/^\d{4}-\d{2}-\d{2}$/.test(String(change.date))) {
      return { valid: false, message: '日付形式が不正です (YYYY-MM-DD)' };
    }
    var validStates = ['', ORDER_STATE.BENTO, ORDER_STATE.OKAZU];
    if (validStates.indexOf(change.orderState || '') === -1) {
      return { valid: false, message: '注文状態が不正です: ' + change.orderState };
    }
    if (change.location && change.location !== LOCATION.SHIN && change.location !== LOCATION.HONSHA) {
      return { valid: false, message: '配達拠点が不正です: ' + change.location };
    }
    // 過去日（今日より前）への変更は拒否
    var now = new Date();
    var today = formatDateYmd(now);
    if (change.date < today) {
      return { valid: false, message: '過去の日付は変更できません' };
    }
    // 当日分は9時以降は変更不可
    if (change.date === today && now.getHours() >= RULES.DEADLINE_HOUR) {
      return { valid: false, message: '当日分の締切(AM9:00)を過ぎています' };
    }
    // 出勤土曜分は前日金曜15時以降は変更不可
    try {
      if (SheetService.isWorkSaturday(change.date)) {
        var targetD = new Date(change.date + 'T00:00:00+09:00');
        var friD = new Date(targetD);
        friD.setDate(friD.getDate() - 1);
        var friStr = formatDateYmd(friD);
        if (today > friStr) {
          return { valid: false, message: '出勤土曜分の締切(金曜15:00)を過ぎています' };
        }
        if (today === friStr && now.getHours() >= RULES.SATURDAY_DEADLINE_HOUR) {
          return { valid: false, message: '出勤土曜分の締切(金曜15:00)を過ぎています' };
        }
      }
    } catch (e) {
      // SheetServiceがまだロードされていない場合はスキップ
    }
    return { valid: true, message: '' };
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
    normalizeDate: normalizeDate
  };
})();
