/**
 * MailService.gs - 日次レポートメール送信
 *
 * 毎朝 8:00〜8:30 の日次メールの構成
 * 1. 【西原屋さん連絡用】本日のお弁当手配（弁当・おかずのみ）
 *    + 本日の麺・丼・特弁（前の平日の朝に注文済み・確認用）
 * 2. 【手配者名一覧】本日分（メニューごと）
 * 3. 【明日（出勤土曜）分・暫定】（翌日が出勤土曜のときだけ）
 * 4. 【ご注文分】麺・丼・特弁（本日8:00 確定）… 本日8:00が締切の日の分（金曜は土曜出勤分と月曜分）
 * 5. 【受付期間・特弁の予定】
 * 麺・丼・特弁の設定が1つもない期間は、3〜5 の麺・丼・特弁の部分は出さない（従来と同じ見た目）。
 */

var MailService = (function() {

  var LINE = '━━━━━━━━━━━━━━━━━━━━';
  var WD = ['日', '月', '火', '水', '木', '金', '土'];

  /**
   * スプレッドシートをExcel形式でエクスポート
   */
  function exportSpreadsheetAsExcel(ssId) {
    var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=xlsx';
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('Export failed: HTTP ' + response.getResponseCode());
    }
    var dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
    return response.getBlob().setName('お弁当注文一覧_' + dateStr + '.xlsx');
  }

  // ===== 書式ヘルパー =====

  function _d(dateStr) { return new Date(dateStr + 'T12:00:00+09:00'); }
  function _jp(dateStr) { return Utilities.formatDate(_d(dateStr), 'Asia/Tokyo', 'yyyy年MM月dd日'); }
  function _md(dateStr) {
    var d = _d(dateStr);
    return (d.getUTCMonth() + 1) + '月' + d.getUTCDate() + '日（' + WD[d.getUTCDay()] + '）';
  }
  function _mdShort(dateStr) {
    var d = _d(dateStr);
    return (d.getUTCMonth() + 1) + '/' + d.getUTCDate() + '（' + WD[d.getUTCDay()] + '）';
  }
  function _label(key) { var m = getMenuDef(key); return m ? m.label : key; }
  function _unit(key) { return key === 'toku' ? '個' : '杯'; }

  function _group(summary, g) {
    return (summary.byGroup && summary.byGroup[g]) || ReportService.emptyGroup();
  }
  function _num(g, key) { return Number(g[key]) || 0; }

  /** 弁当・おかずのみの1行（従来どおりの形） */
  function _bentoLine(label, g) {
    var b = _num(g, 'bento'), o = _num(g, 'okazu');
    return '▼ ' + label + '：' + (b + o) + '名（弁当' + b + ' / おかずのみ' + o + '）';
  }
  function _bentoGrandLine(summary) {
    var t = summary.grandTotal, b = _num(t, 'bento'), o = _num(t, 'okazu');
    return '▼ 全体合計：' + (b + o) + '名（弁当' + b + ' + おかずのみ' + o + '）';
  }

  /** その日に載せる麺・丼・特弁（受付がある、または予約がある） */
  function _earlyKeys(summary, dateStr, ctx) {
    return EARLY_MENU_KEYS.filter(function(k) {
      if (_num(summary.grandTotal, k) > 0) return true;
      return ctx ? MenuService.isOffered(k, dateStr, ctx) : false;
    });
  }

  /** 氏名をメニューごとに並べる */
  function _namesLines(g, keys, indent) {
    indent = indent || '　';
    var out = [];
    keys.forEach(function(k) {
      var list = (g.namesBy && g.namesBy[k]) || [];
      if (list.length > 0) out.push(indent + _label(k) + '（' + list.length + '）：' + list.join('、'));
    });
    if (out.length === 0) out.push(indent + '（なし）');
    return out;
  }

  /** メニュー名＋特弁の名前・価格（例: 特弁「すき焼き弁当」¥980） */
  function _menuTitle(key, dateStr, ctx) {
    if (key === 'toku' && ctx) {
      var t = MenuService.tokuOn(dateStr, ctx);
      if (t) return '特弁「' + t.name + '」¥' + t.price;
    }
    return _label(key);
  }

  // ===== 本文 =====

  /**
   * 日次メール本文を生成
   * @param summary 当日分の集計
   * @param nextDaySatSummary 翌日が出勤土曜の場合の集計（任意）
   * @param opt {ctx, dueSummaries:[summary...]}（任意。麺・丼・特弁の欄に使う）
   */
  function buildMailBody(summary, nextDaySatSummary, opt) {
    opt = opt || {};
    var ctx = opt.ctx || null;
    var dateLabel = _jp(summary.date);
    var timeStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm');
    var groups = summary.groupOrder;

    var lines = [];
    lines.push('総務ご担当者様');
    lines.push('');
    lines.push('お疲れ様です。');
    lines.push('本日分のお弁当注文一覧を送付いたします。');
    lines.push('添付のExcelファイルをご確認ください。');
    lines.push('');
    lines.push('■ 対象日: ' + dateLabel);
    lines.push('■ 注文集計時刻: ' + timeStr + '時点');
    lines.push('');
    lines.push(LINE);
    lines.push('【西原屋さん連絡用】本日のお弁当手配');
    lines.push(LINE);
    lines.push('');
    groups.forEach(function(g) { lines.push(_bentoLine(g, _group(summary, g))); });
    lines.push('');
    lines.push(_bentoGrandLine(summary));

    // 本日の麺・丼・特弁（前の平日の朝にすでに注文済み）
    var todayEarly = _earlyKeys(summary, summary.date, ctx);
    if (todayEarly.length > 0) {
      lines.push('');
      lines.push('▼ 本日の麺・丼・特弁（前の平日の朝に注文済み・確認用）');
      groups.forEach(function(g) {
        var gg = _group(summary, g);
        lines.push('　' + g + '：' + todayEarly.map(function(k) { return _label(k) + _num(gg, k) + _unit(k); }).join(' / '));
      });
      lines.push('　合計：' + todayEarly.map(function(k) { return _label(k) + _num(summary.grandTotal, k) + _unit(k); }).join('・') + ' 手配済みです');
    }

    lines.push('');
    lines.push(LINE);
    lines.push('【手配者名一覧】本日分');
    lines.push(LINE);
    groups.forEach(function(g) {
      lines.push('■ ' + g + ':');
      _namesLines(_group(summary, g), ['bento', 'okazu'].concat(todayEarly)).forEach(function(l) { lines.push(l); });
    });

    // ===== 翌日が出勤土曜の場合: 弁当・おかずのみの暫定集計を併記 =====
    if (nextDaySatSummary) {
      var satLabel = Utilities.formatDate(_d(nextDaySatSummary.date), 'Asia/Tokyo', 'MM月dd日');
      lines.push('');
      lines.push(LINE);
      lines.push('【明日（出勤土曜 ' + satLabel + '）分・暫定】 締切: 本日15:00');
      lines.push(LINE);
      nextDaySatSummary.groupOrder.forEach(function(g) { lines.push(_bentoLine(g, _group(nextDaySatSummary, g))); });
      lines.push('');
      lines.push(_bentoGrandLine(nextDaySatSummary));
      lines.push('');
      lines.push('【明日分・手配者名一覧】');
      nextDaySatSummary.groupOrder.forEach(function(g) {
        lines.push('■ ' + g + ':');
        _namesLines(_group(nextDaySatSummary, g), ['bento', 'okazu']).forEach(function(l) { lines.push(l); });
      });
      lines.push('');
      lines.push('※ 上記は本日朝時点の暫定値です。本日15時の締切後に変更があれば、別途「追加変更」メールでお知らせします。');
      if ((opt.dueSummaries || []).length > 0) lines.push('※ 明日分の麺・丼・特弁は、下の【ご注文分】に載せています。');
    }

    // ===== 本日8:00が締切の麺・丼・特弁（ご注文分） =====
    var due = opt.dueSummaries || [];
    if (due.length > 0) {
      lines.push('');
      lines.push(LINE);
      lines.push('【ご注文分】麺・丼・特弁（本日8:00 確定）');
      lines.push(LINE);
      lines.push('本日の朝のご連絡で、以下もあわせて西原屋さんへご注文ください。');
      if (due.length > 1) {
        lines.push('（' + due.map(function(s) { return _mdShort(s.date); }).join('・') + ' の分をまとめてお送りします）');
      }
      due.forEach(function(ds) {
        var keys = _earlyKeys(ds, ds.date, ctx);
        lines.push('');
        var isSat = ctx ? MenuService.isWorkSat(ds.date, ctx.calMap) : false;
        var titles = keys.filter(function(k) { return k === 'toku'; }).map(function(k) { return _menuTitle(k, ds.date, ctx); });
        lines.push('◆ ' + _md(ds.date) + (isSat ? '・出勤土曜' : '') + '分' + (titles.length ? '　' + titles.join('　') : ''));
        ds.groupOrder.forEach(function(g) {
          var gg = _group(ds, g);
          lines.push('　▼ ' + g + '：' + keys.map(function(k) { return _label(k) + _num(gg, k); }).join(' / '));
        });
        lines.push('　▼ 合計：' + keys.map(function(k) { return _label(k) + _num(ds.grandTotal, k); }).join(' + '));
        ds.groupOrder.forEach(function(g) {
          lines.push('　■ ' + g + ':');
          _namesLines(_group(ds, g), keys, '　　').forEach(function(l) { lines.push(l); });
        });
      });
    }

    // ===== 受付期間・特弁の予定 =====
    if (ctx) {
      var today = summary.date;
      var ps = ctx.periods.filter(function(p) { return p.active && p.to >= today; })
        .sort(function(a, b) { return a.from < b.from ? -1 : 1; });
      var ts = ctx.tokus.filter(function(t) { return t.active && t.date > today; })
        .sort(function(a, b) { return a.date < b.date ? -1 : 1; });
      if (ps.length > 0 || ts.length > 0) {
        lines.push('');
        lines.push(LINE);
        lines.push('【受付期間・特弁の予定】');
        lines.push(LINE);
        ps.forEach(function(p) {
          lines.push('・' + _label(p.menu) + '：' + _md(p.from) + '〜' + _md(p.to) + (p.name ? '　' + p.name : '') + (p.from <= today ? '（受付中）' : '（受付予定）'));
        });
        ts.forEach(function(t) {
          var dl = MenuService.earlyDeadline(t.date, ctx.calMap);
          lines.push('・特弁：' + _md(t.date) + '　' + t.name + ' ¥' + t.price + '（締切 ' + _mdShort(formatDateYmd(dl)) + ' 8:00）');
        });
      }
    }

    lines.push('');
    lines.push('※ 詳細は添付Excelをご参照ください。');
    lines.push('※ このメールはお弁当予約アプリより自動送信されています。');
    lines.push('※ 内容にご不明点がございましたら、システム管理者までお問い合わせください。');
    return lines.join('\n');
  }

  /**
   * 翌日が出勤土曜なら、その日付（YYYY-MM-DD）を返す。それ以外は null。
   */
  function _nextDayIfWorkSat(dateStr) {
    var nextStr = MenuService.addDays(dateStr, 1);
    try {
      return SheetService.isWorkSaturday(nextStr) ? nextStr : null;
    } catch (e) {
      return null;
    }
  }

  /**
   * 出勤土曜分のスナップショット保存・取得・削除
   */
  function _saveSatSnapshot(dateStr, summary) {
    var minimal = {
      date: summary.date,
      byGroup: {},
      grandTotal: summary.grandTotal,
      groupOrder: summary.groupOrder
    };
    summary.groupOrder.forEach(function(g) {
      var d = summary.byGroup[g];
      var m = { total: d.total, names: d.names.slice() };
      MENU_KEYS.forEach(function(k) { m[k] = d[k] || 0; });
      minimal.byGroup[g] = m;
    });
    PropertiesService.getScriptProperties().setProperty('sat_snapshot_' + dateStr, JSON.stringify(minimal));
  }
  function _getSatSnapshot(dateStr) {
    var s = PropertiesService.getScriptProperties().getProperty('sat_snapshot_' + dateStr);
    return s ? JSON.parse(s) : null;
  }
  function _clearSatSnapshot(dateStr) {
    PropertiesService.getScriptProperties().deleteProperty('sat_snapshot_' + dateStr);
  }

  /**
   * 名前リスト「氏名（弁当）」「氏名（丼）」などを分解
   */
  function _parseNameType(nt) {
    var m = String(nt).match(/^(.+?)（(.+?)）$/);
    return m ? { name: m[1], type: m[2] } : { name: String(nt), type: '' };
  }

  /**
   * スナップショットと現在の集計を比較し、差分を計算
   * 戻り値: { added:[name+type], removed:[name+type], changed:[{name, from, to}] }
   */
  function _diffSatSummary(snapshot, current) {
    var snap = {};   // name → type
    var curr = {};
    if (snapshot) {
      snapshot.groupOrder.forEach(function(g) {
        (snapshot.byGroup[g].names || []).forEach(function(nt) {
          var p = _parseNameType(nt); snap[p.name] = p.type;
        });
      });
    }
    current.groupOrder.forEach(function(g) {
      (current.byGroup[g].names || []).forEach(function(nt) {
        var p = _parseNameType(nt); curr[p.name] = p.type;
      });
    });
    var added = [], removed = [], changed = [];
    Object.keys(curr).forEach(function(name) {
      if (!(name in snap)) {
        added.push(name + '（' + curr[name] + '）');
      } else if (snap[name] !== curr[name]) {
        changed.push({ name: name, from: snap[name], to: curr[name] });
      }
    });
    Object.keys(snap).forEach(function(name) {
      if (!(name in curr)) removed.push(name + '（' + snap[name] + '）');
    });
    return { added: added, removed: removed, changed: changed };
  }

  /**
   * 出勤土曜分の追加変更メール本文
   */
  function _buildSatSupplementalBody(satDateStr, current, diff) {
    var satLabel = _jp(satDateStr);
    var timeStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm');
    var hasChanges = !(diff.added.length === 0 && diff.removed.length === 0 && diff.changed.length === 0);
    var lines = [];
    lines.push('総務ご担当者様');
    lines.push('');
    lines.push('お疲れ様です。');
    lines.push('本日朝のメールでお知らせした明日（' + satLabel + '・出勤土曜）分の');
    if (hasChanges) {
      lines.push('お弁当注文に変更がありました。最終確定状況をご連絡いたします。');
    } else {
      lines.push('お弁当注文に変更はありませんでした。最終確定状況をご連絡いたします。');
    }
    lines.push('');
    lines.push('■ 対象日: ' + satLabel);
    lines.push('■ 締切時刻: 本日15:00');
    lines.push('■ 確定時刻: ' + timeStr);
    lines.push('');
    lines.push(LINE);
    lines.push('【朝メールからの変更点】');
    lines.push(LINE);
    lines.push('▼ 追加: ' + diff.added.length + '件');
    diff.added.forEach(function(s) { lines.push('　・' + s); });
    lines.push('▼ 取消: ' + diff.removed.length + '件');
    diff.removed.forEach(function(s) { lines.push('　・' + s); });
    lines.push('▼ 種別変更: ' + diff.changed.length + '件');
    diff.changed.forEach(function(c) { lines.push('　・' + c.name + '：' + c.from + ' → ' + c.to); });
    lines.push('');
    lines.push(LINE);
    lines.push('【最終確定後の集計】');
    lines.push(LINE);
    current.groupOrder.forEach(function(g) { lines.push(_bentoLine(g, _group(current, g))); });
    lines.push('');
    lines.push(_bentoGrandLine(current));
    var early = EARLY_MENU_KEYS.filter(function(k) { return _num(current.grandTotal, k) > 0; });
    if (early.length > 0) {
      lines.push('▼ 麺・丼・特弁（金曜の朝に注文済み）：' + early.map(function(k) { return _label(k) + _num(current.grandTotal, k); }).join(' / '));
    }
    lines.push('');
    lines.push(LINE);
    lines.push('【最終 手配者名一覧】');
    lines.push(LINE);
    current.groupOrder.forEach(function(g) {
      lines.push('■ ' + g + ':');
      _namesLines(_group(current, g), ['bento', 'okazu'].concat(early)).forEach(function(l) { lines.push(l); });
    });
    lines.push('');
    lines.push('※ 本メールは出勤土曜の前日15時の締切時点で、変更の有無に関わらず送信されます。');
    if (early.length > 0) lines.push('※ 麺・丼・特弁は金曜の朝8時に確定しているため、このメールでは変わりません。');
    lines.push('※ このメールはお弁当予約アプリより自動送信されています。');
    return lines.join('\n');
  }

  /**
   * 出勤土曜の前日15時に呼ばれる最終確定メール送信
   * - 翌日が出勤土曜でなければ何もしない
   * - スナップショットと現在を比較し、変更の有無に関わらず必ず送信（件名・本文で有無を明示）
   */
  function sendSaturdaySupplementalMail() {
    try {
      var today = formatDateYmd(new Date());
      var satDate = _nextDayIfWorkSat(today);
      if (!satDate) {
        Logger.log('Saturday supplemental: tomorrow is not a work Saturday, skip');
        return;
      }
      var snapshot = _getSatSnapshot(satDate);
      var current = ReportService.generateReportSummary(satDate);
      var diff = _diffSatSummary(snapshot, current);
      var hasChanges = !(diff.added.length === 0 && diff.removed.length === 0 && diff.changed.length === 0);
      // 変更の有無に関わらず必ず送信する（変更があれば総務がお弁当屋へ追加電話するため、有無を必ず通知）
      Logger.log('Saturday supplemental: ' + satDate + ' hasChanges=' + hasChanges);
      var recipients = SheetService.getMailRecipients();
      if (!recipients || recipients.length === 0) {
        Logger.log('Saturday supplemental: no recipients');
        return;
      }
      var subject = '【お弁当注文表・' + (hasChanges ? '追加変更' : '変更なし') + '】明日（出勤土曜 ' + Utilities.formatDate(_d(satDate), 'Asia/Tokyo', 'MM月dd日') + '）分の最終確定';
      var body = _buildSatSupplementalBody(satDate, current, diff);

      // 添付: 最終確定後の土曜分Excel
      var options = {};
      var tempSsId = null;
      try {
        var tempSs = _buildOneDaySpreadsheetFromSummary(current, satDate, 'お弁当予約_土曜最終');
        tempSsId = tempSs.getId();
        SpreadsheetApp.flush();
        options.attachments = [exportSpreadsheetAsExcel(tempSsId)];
      } catch (e) {
        Logger.log('Saturday supplemental: attachment generation failed: ' + e.message);
      }

      try {
        recipients.forEach(function(to) {
          try {
            MailApp.sendEmail(to, subject, body, options);
            Logger.log('Saturday supplemental sent to: ' + to + ' (attachment=' + (options.attachments ? 'yes' : 'no') + ')');
          } catch (e) {
            Logger.log('Saturday supplemental send failed to ' + to + ': ' + e.message);
          }
        });
      } finally {
        if (tempSsId) {
          try { DriveApp.getFileById(tempSsId).setTrashed(true); } catch (e) {
            Logger.log('Trash temp ss failed: ' + e.message);
          }
        }
      }
      _clearSatSnapshot(satDate);
    } catch (err) {
      Logger.log('sendSaturdaySupplementalMail error: ' + err.message);
      throw err;
    }
  }

  /**
   * 本日8:00が締切の麺・丼・特弁の集計（受付か予約がある日だけ）
   */
  function _dueSummaries(today, ctx) {
    return MenuService.daysDueOn(today, ctx.calMap).map(function(d) {
      return ReportService.generateReportSummary(d);
    }).filter(function(s) {
      return _earlyKeys(s, s.date, ctx).length > 0;
    });
  }

  /**
   * 日次レポートメール送信
   * - 翌日が出勤土曜の場合は、翌日分の暫定集計を本文に併記しスナップショット保存
   * - 本日8:00が締切の麺・丼・特弁を【ご注文分】として載せる
   */
  function sendDailyReport() {
    try {
      var today = formatDateYmd(new Date());
      var recipients = SheetService.getMailRecipients();
      if (!recipients || recipients.length === 0) {
        Logger.log('sendDailyReport: no recipients');
        return;
      }

      // 集計は予約の保存（api_applyPatch）と同じロックの中で読む。
      // 締切直前に始まった保存が書き終わってから数えるため（麺・丼の注文漏れを防ぐ）
      var summary, satSummary = null, ctx, due;
      var satDate = _nextDayIfWorkSat(today);
      var lock = LockService.getDocumentLock();
      lock.waitLock(30000);
      try {
        summary = ReportService.generateReportSummary(today);
        // 翌日が出勤土曜なら、翌日分も生成して本文に併記＋スナップショット保存
        if (satDate) {
          satSummary = ReportService.generateReportSummary(satDate);
          _saveSatSnapshot(satDate, satSummary);
          Logger.log('Saturday snapshot saved for ' + satDate);
        }
        ctx = MenuService.loadContext();
        due = _dueSummaries(today, ctx);
      } finally {
        lock.releaseLock();
      }

      var ss = ReportService.getOrCreateAdminSpreadsheet();
      var blob = null;
      try {
        blob = exportSpreadsheetAsExcel(ss.getId());
      } catch (e) {
        Logger.log('Excel export failed: ' + e.message);
      }

      var subject = '【お弁当注文表】本日分のお弁当注文一覧（' + _jp(today) + '）'
                  + (satSummary ? ' ＋翌出勤土曜分（暫定）' : '')
                  + (due.length > 0 ? ' ＋麺・丼のご注文分' : '');
      var body = buildMailBody(summary, satSummary, { ctx: ctx, dueSummaries: due });

      var options = {};
      if (blob) options.attachments = [blob];

      recipients.forEach(function(to) {
        try {
          MailApp.sendEmail(to, subject, body, options);
          Logger.log('Mail sent to: ' + to);
        } catch (e) {
          Logger.log('Mail send failed to ' + to + ': ' + e.message);
        }
      });
    } catch (err) {
      Logger.log('sendDailyReport error: ' + err.message);
      throw err;
    }
  }

  return {
    sendDailyReport: sendDailyReport,
    sendSaturdaySupplementalMail: sendSaturdaySupplementalMail,
    exportSpreadsheetAsExcel: exportSpreadsheetAsExcel,
    buildMailBody: buildMailBody,
    dueSummaries: _dueSummaries,
    diffSatSummary: _diffSatSummary,
    buildSatSupplementalBody: _buildSatSupplementalBody
  };
})();

/**
 * トリガーから呼ばれる日次ジョブ
 */
function dailyReportJob() {
  try {
    ReportService.updateAdminSheetsForToday();
    MailService.sendDailyReport();
  } catch (err) {
    Logger.log('dailyReportJob error: ' + err.message);
  }
}

/**
 * トリガーから呼ばれる出勤土曜の前日15:00ジョブ
 * - 翌日が出勤土曜のときだけ動作
 * - 朝メール時のスナップショットと比較し、変更の有無に関わらず最終確定メールを送信
 */
function saturdayDeadlineMailJob() {
  try {
    MailService.sendSaturdaySupplementalMail();
  } catch (err) {
    Logger.log('saturdayDeadlineMailJob error: ' + err.message);
  }
}

/**
 * 本日の日次メールの本文を、送らずに実行ログへ出す（確認用。GASエディタから手動実行）
 */
function previewDailyMail() {
  var today = formatDateYmd(new Date());
  var ctx = MenuService.loadContext();
  var summary = ReportService.generateReportSummary(today);
  var satDate = MenuService.addDays(today, 1);
  var satSummary = SheetService.isWorkSaturday(satDate) ? ReportService.generateReportSummary(satDate) : null;
  var body = MailService.buildMailBody(summary, satSummary, { ctx: ctx, dueSummaries: MailService.dueSummaries(today, ctx) });
  Logger.log(body);
  return body;
}

/**
 * サンプル: 出勤土曜の前日15時の追加変更メール（架空の差分データで送信）
 * - 朝メール時のスナップショット（17名）と、15時時点の最終確定（20名）を架空生成
 * - 追加4件 / 取消1件 / 種別変更1件 を含む差分を生成し、SAMPLE_MAIL_TO 全員に送信
 */
function sendSampleSaturdaySupplementalMail() {
  var today = formatDateYmd(new Date());
  var satDate = MenuService.addDays(today, 1);
  var d = new Date(satDate + 'T12:00:00+09:00');

  // 朝時点のスナップショット
  var snapshot = _buildSatSampleSummary(satDate,
    ['田村 修二', '佐藤 健', '鈴木 一郎', '高橋 美咲', '伊藤 大輔', '渡辺 真理', '山本 剛', '中村 結衣'],
    ['橋本 一夫', '石川 幸子'],
    ['後藤 雄一', '長谷川 涼', '村上 理恵', '近藤 雅人', '坂本 葵', '遠藤 健一'],
    ['竹内 大'],
    { '本社工場': { don: ['青木 恵'] } }
  );

  // 15時時点の最終確定（朝から変更あり）
  // - 追加: 小林 光（弁当）, 加藤 智子（弁当）, 前田 翔太（おかず）, 金子 由佳（おかず）
  // - 取消: 遠藤 健一（弁当）
  // - 種別変更: 山本 剛 弁当→おかず
  var current = _buildSatSampleSummary(satDate,
    ['田村 修二', '佐藤 健', '鈴木 一郎', '高橋 美咲', '伊藤 大輔', '渡辺 真理', '中村 結衣', '小林 光', '加藤 智子'],
    ['橋本 一夫', '石川 幸子', '前田 翔太', '山本 剛'],
    ['後藤 雄一', '長谷川 涼', '村上 理恵', '近藤 雅人', '坂本 葵'],
    ['竹内 大', '金子 由佳'],
    { '本社工場': { don: ['青木 恵'] } }
  );

  var diff = MailService.diffSatSummary(snapshot, current);
  var satLabel = Utilities.formatDate(d, 'Asia/Tokyo', 'MM月dd日');
  var subject = '[サンプル]【お弁当注文表・追加変更】明日（出勤土曜 ' + satLabel + '）分の最終確定';
  var body = '※ これはサンプルメールです。架空の差分データ（追加' + diff.added.length
           + '件・取消' + diff.removed.length + '件・種別変更' + diff.changed.length + '件）で生成しています。\n\n'
           + MailService.buildSatSupplementalBody(satDate, current, diff);

  _sendSampleWithSheet_(subject, body, function() {
    return _buildOneDaySpreadsheetFromSummary(current, satDate, 'お弁当予約_サンプル土曜最終');
  });
}

/**
 * サンプル: 金曜朝メール（本日(金)＋翌日(出勤土曜)分・暫定 ＋ 麺・丼のご注文分 を併記）
 * SAMPLE_MAIL_TO 全員に送信
 */
function sendSampleFridayMail() {
  var today = formatDateYmd(new Date());
  var friDate = today;
  var satDate = MenuService.addDays(today, 1);
  var monDate = MenuService.addDays(today, 3);

  // 金曜の本日分（丼は前日の朝に注文済み）
  var friSummary = _buildSatSampleSummary(friDate,
    ['田村 修二', '佐藤 健', '鈴木 一郎', '高橋 美咲', '伊藤 大輔', '渡辺 真理', '山本 剛', '中村 結衣', '小林 光', '加藤 智子', '吉田 隆', '山田 花子', '松本 拓也', '井上 由美', '木村 慎一'],
    ['橋本 一夫', '石川 幸子', '前田 翔太'],
    ['後藤 雄一', '長谷川 涼', '村上 理恵', '近藤 雅人', '坂本 葵', '遠藤 健一', '青木 恵', '福田 武', '太田 美香', '西村 龍', '藤井 七海'],
    ['竹内 大', '金子 由佳', '和田 進', '中川 美鈴'],
    { '新工場': { don: ['林 香織'] }, '本社工場': { don: ['岡本 真'], men: ['三浦 美咲'] } }
  );
  // 翌日の出勤土曜分・暫定
  var satSummary = _buildSatSampleSummary(satDate,
    ['田村 修二', '佐藤 健', '鈴木 一郎', '高橋 美咲', '伊藤 大輔', '渡辺 真理', '山本 剛', '中村 結衣'],
    ['橋本 一夫', '石川 幸子'],
    ['後藤 雄一', '長谷川 涼', '村上 理恵', '近藤 雅人', '坂本 葵', '遠藤 健一'],
    ['竹内 大'],
    { '本社工場': { don: ['青木 恵'] } }
  );
  // 月曜分（本日8:00締切の麺・丼）
  var monSummary = _buildSatSampleSummary(monDate, [], [], [], [],
    { '新工場': { men: ['山田 花子', '松本 拓也'], don: ['井上 由美'] }, '本社工場': { men: ['福田 武'] } }
  );

  var friLabel = Utilities.formatDate(new Date(friDate + 'T12:00:00+09:00'), 'Asia/Tokyo', 'yyyy年MM月dd日');
  var subject = '[サンプル]【お弁当注文表】本日分のお弁当注文一覧（' + friLabel + '）＋翌出勤土曜分（暫定）＋麺・丼のご注文分';
  var body = '※ これはサンプルメールです。架空の予約データで「金曜朝メール（翌日出勤土曜あり・麺丼のご注文分あり）」を再現しています。\n\n'
           + MailService.buildMailBody(friSummary, satSummary, { dueSummaries: [satSummary, monSummary] });

  _sendSampleWithSheet_(subject, body, function() {
    var friColLabel = Utilities.formatDate(new Date(friDate + 'T12:00:00+09:00'), 'Asia/Tokyo', 'MM月dd日') + '(金)';
    var satColLabel = Utilities.formatDate(new Date(satDate + 'T12:00:00+09:00'), 'Asia/Tokyo', 'MM月dd日') + '(土・暫定)';
    return _buildFridaySampleSpreadsheet(friSummary, satSummary, friColLabel, satColLabel);
  });
}

/**
 * サンプル送信の共通処理（添付を作って送り、一時ファイルを捨てる）
 */
function _sendSampleWithSheet_(subject, body, buildSheet) {
  var options = {};
  var tempSsId = null;
  try {
    var tempSs = buildSheet();
    tempSsId = tempSs.getId();
    SpreadsheetApp.flush();
    options.attachments = [MailService.exportSpreadsheetAsExcel(tempSsId)];
  } catch (e) {
    Logger.log('Sample: attachment generation failed: ' + e.message);
    body += '\n\n[添付生成エラー] ' + e.message;
  }
  try {
    SAMPLE_MAIL_TO.forEach(function(to) {
      try {
        MailApp.sendEmail(to, subject, body, options);
        Logger.log('Sample mail sent to: ' + to + ' (attachment=' + (options.attachments ? 'yes' : 'no') + ')');
      } catch (e) {
        Logger.log('Sample send failed to ' + to + ': ' + e.message);
      }
    });
  } finally {
    if (tempSsId) {
      try { DriveApp.getFileById(tempSsId).setTrashed(true); } catch (e) {
        Logger.log('Trash temp ss failed: ' + e.message);
      }
    }
  }
}

/**
 * 拠点ごとのメニュー別氏名リストから、集計データを作る
 * lists = { '新工場': {bento:[], okazu:[], men:[], don:[], toku:[]}, '本社工場': {...} }
 */
function _summaryFromLists(dateStr, lists) {
  var groups = ['新工場', '本社工場'];
  var byGroup = {};
  var grand = { total: 0 };
  MENU_KEYS.forEach(function(k) { grand[k] = 0; });
  groups.forEach(function(g) {
    var src = lists[g] || {};
    var b = ReportService.emptyGroup();
    MENU_DEFS.forEach(function(m) {
      var arr = src[m.key] || [];
      b[m.key] = arr.length;
      b.total += arr.length;
      b.namesBy[m.key] = arr.slice();
      arr.forEach(function(n) { b.names.push(n + '（' + m.label + '）'); });
      grand[m.key] += arr.length;
      grand.total += arr.length;
    });
    byGroup[g] = b;
  });
  return { date: dateStr, byGroup: byGroup, grandTotal: grand, groupOrder: groups };
}

/**
 * サンプル集計データを生成（弁当・おかずのみ＋任意で麺・丼・特弁）
 * extra = { '新工場': {men:[], don:[], toku:[]}, '本社工場': {...} }
 */
function _buildSatSampleSummary(dateStr, shinB, shinO, honB, honO, extra) {
  extra = extra || {};
  function merge(g, b, o) {
    var e = extra[g] || {};
    return { bento: b, okazu: o, men: e.men || [], don: e.don || [], toku: e.toku || [] };
  }
  return _summaryFromLists(dateStr, {
    '新工場': merge('新工場', shinB, shinO),
    '本社工場': merge('本社工場', honB, honO)
  });
}

/**
 * メール送信周りの診断ヘルパー
 * GASエディタから実行 → 実行ログ画面で確認
 * 1) mail_recipients シートの生値
 * 2) getMailRecipients() の結果
 * 3) 登録済みトリガー一覧
 * 4) MailApp の残り送信可能数
 * 5) 実際に1通テスト送信
 */
function debugMailSetup() {
  Logger.log('===== Mail diagnostics =====');

  // 1) 生のシート内容
  try {
    var sheet = getOrCreateSheet(SHEET_EMAIL_RECIPIENTS, ['名前', 'メールアドレス', '有効フラグ']);
    var lastRow = sheet.getLastRow();
    Logger.log('[1] mail_recipients sheet: lastRow=' + lastRow);
    if (lastRow >= 2) {
      var raw = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
      raw.forEach(function(r, i) {
        Logger.log('  row ' + (i + 2) + ': name="' + r[0] + '", email="' + r[1] + '", flag="' + r[2] + '" (flagType=' + typeof r[2] + ', flagLen=' + String(r[2] || '').length + ')');
      });
    }
  } catch (e) {
    Logger.log('[1] sheet read error: ' + e.message);
  }

  // 2) getMailRecipients() の結果
  try {
    var recipients = SheetService.getMailRecipients();
    Logger.log('[2] getMailRecipients() returned ' + recipients.length + ' addresses:');
    recipients.forEach(function(r) { Logger.log('  - ' + r); });
  } catch (e) {
    Logger.log('[2] getMailRecipients error: ' + e.message);
  }

  // 3) トリガー一覧
  try {
    var triggers = ScriptApp.getProjectTriggers();
    Logger.log('[3] Project triggers: ' + triggers.length);
    triggers.forEach(function(t) {
      Logger.log('  - ' + t.getHandlerFunction() + ' / ' + t.getEventType());
    });
  } catch (e) {
    Logger.log('[3] triggers error: ' + e.message);
  }

  // 4) MailApp 残量
  try {
    var remaining = MailApp.getRemainingDailyQuota();
    Logger.log('[4] MailApp remaining quota today: ' + remaining);
  } catch (e) {
    Logger.log('[4] quota error: ' + e.message);
  }

  // 5) テスト送信（getMailRecipients() の最初のアドレスへ）
  try {
    var rcp = SheetService.getMailRecipients();
    if (rcp.length === 0) {
      Logger.log('[5] No recipients to test with. Add "ON" flag in mail_recipients sheet.');
    } else {
      var testTo = rcp[0];
      MailApp.sendEmail(
        testTo,
        '[診断テスト] お弁当予約アプリ - メール送信テスト',
        'これは debugMailSetup() からのテスト送信です。\n受信できれば mail_recipients シート → MailApp の経路は正常です。\n\n送信時刻: ' + new Date()
      );
      Logger.log('[5] Test mail sent to: ' + testTo);
    }
  } catch (e) {
    Logger.log('[5] test send error: ' + e.message);
  }

  Logger.log('===== End diagnostics =====');
}

/**
 * サンプルメール送信（架空データ・指定アドレスへ送信）
 * GASエディタから sendSampleMail を選んで手動実行する用途
 */
var SAMPLE_MAIL_TO = [
  'imaizumi@lineworks.co.jp',
  'c-matusita@lineworks.co.jp'
];

function sendSampleMail() {
  var today = formatDateYmd(new Date());
  var dLabel = Utilities.formatDate(new Date(today + 'T12:00:00+09:00'), 'Asia/Tokyo', 'yyyy年MM月dd日');
  var summary = _buildSampleSummary(today);
  var subject = '[サンプル]【お弁当注文表】本日分のお弁当注文一覧（' + dLabel + '）';
  var body = '※ これはサンプルメールです。実データではなく架空の予約データで生成しています。\n\n'
           + MailService.buildMailBody(summary);

  _sendSampleWithSheet_(subject, body, function() {
    return _buildSampleAdminSpreadsheet(summary, today);
  });
}

/**
 * サンプル用の架空集計データを構築（新工場25名 / 本社工場20名 = 計45名）
 */
function _buildSampleSummary(dateStr) {
  var shinBento = ['田村 修二', '佐藤 健', '鈴木 一郎', '高橋 美咲', '伊藤 大輔', '渡辺 真理', '山本 剛', '中村 結衣', '小林 光', '加藤 智子', '吉田 隆', '山田 花子', '松本 拓也', '井上 由美', '木村 慎一', '林 香織', '清水 浩二', '山口 美穂', '森 賢治', '池田 裕子'];
  var shinOkazu = ['橋本 一夫', '石川 幸子', '前田 翔太', '藤田 麻衣', '岡田 純一'];
  var honshaBento = ['後藤 雄一', '長谷川 涼', '村上 理恵', '近藤 雅人', '坂本 葵', '遠藤 健一', '青木 恵', '福田 武', '太田 美香', '西村 龍', '藤井 七海', '岡本 真', '三浦 美咲', '中島 修', '原 さつき'];
  var honshaOkazu = ['竹内 大', '金子 由佳', '和田 進', '中川 美鈴', '石井 研'];
  return _buildSatSampleSummary(dateStr, shinBento, shinOkazu, honshaBento, honshaOkazu);
}

/**
 * "氏名（弁当）" / "氏名（丼）" などを分解
 */
function _parseNTGlobal(nt) {
  var m = String(nt).match(/^(.+?)（(.+?)）$/);
  return m ? { name: m[1], type: m[2] } : { name: String(nt), type: '' };
}

/**
 * サマリから拠点ごと・メニューごとの氏名リストを取り出す
 * 戻り値: { '新工場': {bento:[], okazu:[], men:[], don:[], toku:[]}, '本社工場': {...} }
 */
function _peopleListsFromSummary(summary) {
  var labelToKey = {};
  MENU_DEFS.forEach(function(m) { labelToKey[m.label] = m.key; });
  var result = {};
  ['新工場', '本社工場'].forEach(function(g) {
    var lists = {};
    MENU_KEYS.forEach(function(k) { lists[k] = []; });
    var src = summary.byGroup[g] || {};
    if (src.namesBy) {
      MENU_KEYS.forEach(function(k) { lists[k] = (src.namesBy[k] || []).slice(); });
    } else {
      (src.names || []).forEach(function(nt) {
        var p = _parseNTGlobal(nt);
        var k = labelToKey[p.type];
        if (k) lists[k].push(p.name);
      });
    }
    result[g] = lists;
  });
  return result;
}

/** 添付シートに載せるメニュー（弁当・おかずのみ＋人がいる麺・丼・特弁） */
function _menusForLists(listsArray) {
  return MENU_DEFS.filter(function(m) {
    if (m.kind === 'always') return true;
    return listsArray.some(function(l) { return l && l[m.key] && l[m.key].length > 0; });
  });
}

/** 集計行（メニューごと＋合計）を書く */
function _writeSummaryRows_(sheet, startRow, width, menus, countFn, colCount, bg) {
  var rows = menus.map(function(m) {
    var r = [m.label, ''];
    for (var i = 0; i < colCount; i++) r.push(countFn(m.key, i));
    return r;
  });
  var total = ['合計', ''];
  for (var i = 0; i < colCount; i++) {
    total.push(menus.reduce(function(a, m) { return a + countFn(m.key, i); }, 0));
  }
  rows.push(total);
  sheet.getRange(startRow, 1, rows.length, width).setValues(rows);
  sheet.getRange(startRow, 1, rows.length, 1).setFontWeight('bold');
  sheet.getRange(startRow, 1, rows.length, width).setBackground(bg);
}

/**
 * 拠点別シート1日分（氏名×1日付）を書き込み
 * @param lists {bento:[], okazu:[], men:[], don:[], toku:[]}
 */
function _writeOneDayLocationSheet(sheet, dateLabel, lists, deptLabel) {
  var header = ['氏名', '部署', dateLabel];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f0e8d8');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  var menus = _menusForLists([lists]);
  var rows = [];
  menus.forEach(function(m) {
    (lists[m.key] || []).forEach(function(n) { rows.push([n, deptLabel || '', m.sym]); });
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  _writeSummaryRows_(sheet, 2 + rows.length + 1, header.length, menus,
    function(key) { return (lists[key] || []).length; }, 1, '#fff8ea');
  sheet.autoResizeColumn(1);
}

/**
 * 全拠点合計シート1日分（拠点列付き、両拠点の氏名を全列挙）
 */
function _writeOneDayCombinedSheet(sheet, dateLabel, peopleByGroup) {
  var header = ['氏名', '拠点', dateLabel];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f0e8d8');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  var groups = ['新工場', '本社工場'];
  var menus = _menusForLists(groups.map(function(g) { return peopleByGroup[g]; }));
  var rows = [];
  groups.forEach(function(g) {
    menus.forEach(function(m) {
      (peopleByGroup[g][m.key] || []).forEach(function(n) { rows.push([n, g, m.sym]); });
    });
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  _writeSummaryRows_(sheet, 2 + rows.length + 1, header.length, menus, function(key) {
    return groups.reduce(function(a, g) { return a + (peopleByGroup[g][key] || []).length; }, 0);
  }, 1, '#fff2e0');
  sheet.autoResizeColumn(1);
}

/**
 * 拠点別シート複数日分（氏名×複数日付）
 * @param peopleByDate [ {bento:[], okazu:[], ...}, ... ] 日付と同じ並び順
 */
function _writeMultiDayLocationSheet(sheet, dateLabels, peopleByDate, deptLabel) {
  var n = dateLabels.length;
  var header = ['氏名', '部署'].concat(dateLabels);
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f0e8d8');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  var menus = _menusForLists(peopleByDate);
  var perPerson = {};
  peopleByDate.forEach(function(p, idx) {
    menus.forEach(function(m) {
      (p[m.key] || []).forEach(function(name) {
        if (!perPerson[name]) perPerson[name] = new Array(n).fill('');
        perPerson[name][idx] = m.sym;
      });
    });
  });
  var rows = Object.keys(perPerson).map(function(name) {
    return [name, deptLabel || ''].concat(perPerson[name]);
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  _writeSummaryRows_(sheet, 2 + rows.length + 1, header.length, menus,
    function(key, i) { return (peopleByDate[i][key] || []).length; }, n, '#fff8ea');
  sheet.autoResizeColumn(1);
}

/**
 * 全拠点合計シート複数日分（拠点列付き）
 * @param peopleByDateByGroup [ { '新工場':{bento,...}, '本社工場':{...} }, ...]
 */
function _writeMultiDayCombinedSheet(sheet, dateLabels, peopleByDateByGroup) {
  var n = dateLabels.length;
  var header = ['氏名', '拠点'].concat(dateLabels);
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f0e8d8');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  var groups = ['新工場', '本社工場'];
  var all = [];
  peopleByDateByGroup.forEach(function(bg) { groups.forEach(function(g) { all.push(bg[g]); }); });
  var menus = _menusForLists(all);

  var perPerson = {};
  peopleByDateByGroup.forEach(function(byGroup, idx) {
    groups.forEach(function(g) {
      menus.forEach(function(m) {
        (byGroup[g][m.key] || []).forEach(function(name) {
          var key = name + '|' + g;
          if (!perPerson[key]) perPerson[key] = { name: name, group: g, cells: new Array(n).fill('') };
          perPerson[key].cells[idx] = m.sym;
        });
      });
    });
  });

  // 拠点ごとに並べる: 新工場 → 本社工場
  var rows = [];
  groups.forEach(function(g) {
    Object.keys(perPerson).forEach(function(key) {
      var p = perPerson[key];
      if (p.group === g) rows.push([p.name, g].concat(p.cells));
    });
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  _writeSummaryRows_(sheet, 2 + rows.length + 1, header.length, menus, function(key, i) {
    return groups.reduce(function(a, g) { return a + (peopleByDateByGroup[i][g][key] || []).length; }, 0);
  }, n, '#fff2e0');
  sheet.autoResizeColumn(1);
}

/**
 * 1日分のサンプル添付スプレッドシートを生成（全拠点合計シートも全員の名前付き）
 */
function _buildSampleAdminSpreadsheet(summary, dateStr) {
  var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var ss = SpreadsheetApp.create('お弁当予約_サンプル_' + stamp);
  var people = _peopleListsFromSummary(summary);

  ['新工場', '本社工場'].forEach(function(group) {
    var sheet = ss.insertSheet(group);
    _writeOneDayLocationSheet(sheet, dateStr, people[group], '（サンプル）');
  });

  var totalSheet = ss.insertSheet('全拠点合計');
  _writeOneDayCombinedSheet(totalSheet, dateStr, people);

  var def = ss.getSheetByName('シート1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  return ss;
}

/**
 * 金曜サンプル用: 本日(金)＋翌日(土・暫定)の2日分スプレッドシート
 */
function _buildFridaySampleSpreadsheet(friSummary, satSummary, friDateLabel, satDateLabel) {
  var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var ss = SpreadsheetApp.create('お弁当予約_サンプル金曜_' + stamp);
  var friPeople = _peopleListsFromSummary(friSummary);
  var satPeople = _peopleListsFromSummary(satSummary);
  var labels = [friDateLabel, satDateLabel];

  ['新工場', '本社工場'].forEach(function(group) {
    var sheet = ss.insertSheet(group);
    _writeMultiDayLocationSheet(sheet, labels, [friPeople[group], satPeople[group]], '（サンプル）');
  });

  var totalSheet = ss.insertSheet('全拠点合計');
  _writeMultiDayCombinedSheet(totalSheet, labels, [friPeople, satPeople]);

  var def = ss.getSheetByName('シート1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  return ss;
}

/**
 * 1日分のスプレッドシートを実データの summary から生成（土曜追加変更メール用）
 */
function _buildOneDaySpreadsheetFromSummary(summary, dateStr, baseName) {
  var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var ss = SpreadsheetApp.create((baseName || 'お弁当予約_確定') + '_' + stamp);
  var people = _peopleListsFromSummary(summary);

  ['新工場', '本社工場'].forEach(function(group) {
    var sheet = ss.insertSheet(group);
    _writeOneDayLocationSheet(sheet, dateStr, people[group], '');
  });

  var totalSheet = ss.insertSheet('全拠点合計');
  _writeOneDayCombinedSheet(totalSheet, dateStr, people);

  var def = ss.getSheetByName('シート1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  return ss;
}
