/**
 * MailService.gs - 日次レポートメール送信
 */

var MailService = (function() {

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

  /**
   * 拠点ごとの集計1行を整形
   */
  function _formatGroupLine(label, g) {
    return '▼ ' + label + '：' + g.total + '名（弁当' + g.bento + ' / おかずのみ' + g.okazu + '）';
  }

  /**
   * メール本文を生成
   * @param summary 当日分の集計
   * @param nextDaySatSummary 翌日が出勤土曜の場合の集計（任意）
   */
  function buildMailBody(summary, nextDaySatSummary) {
    var d = new Date(summary.date + 'T00:00:00+09:00');
    var dateLabel = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy年MM月dd日');
    var timeStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm');

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
    lines.push('━━━━━━━━━━━━━━━━━━━━');
    lines.push('【西原屋さん連絡用】お弁当手配');
    lines.push('━━━━━━━━━━━━━━━━━━━━');
    lines.push('');
    var bg = summary.byGroup;
    var shin = bg['新工場'] || { bento: 0, okazu: 0, total: 0, names: [] };
    var honsha = bg['本社工場'] || { bento: 0, okazu: 0, total: 0, names: [] };
    lines.push(_formatGroupLine('新工場', shin));
    lines.push(_formatGroupLine('本社工場', honsha));
    lines.push('');
    lines.push('▼ 全体合計：' + summary.grandTotal.total + '名（弁当' + summary.grandTotal.bento + ' + おかずのみ' + summary.grandTotal.okazu + '）');
    lines.push('');
    lines.push('━━━━━━━━━━━━━━━━━━━━');
    lines.push('【手配者名一覧】');
    lines.push('━━━━━━━━━━━━━━━━━━━━');
    summary.groupOrder.forEach(function(g) {
      var d2 = summary.byGroup[g];
      lines.push('■ ' + g + ':');
      if (d2.names.length === 0) {
        lines.push('　（なし）');
      } else {
        lines.push('　' + d2.names.join('、'));
      }
    });

    // ===== 翌日が出勤土曜の場合: 暫定集計を併記 =====
    if (nextDaySatSummary) {
      var satD = new Date(nextDaySatSummary.date + 'T00:00:00+09:00');
      var satLabel = Utilities.formatDate(satD, 'Asia/Tokyo', 'MM月dd日');
      lines.push('');
      lines.push('━━━━━━━━━━━━━━━━━━━━');
      lines.push('【明日（出勤土曜 ' + satLabel + '）分・暫定】 締切: 本日15:00');
      lines.push('━━━━━━━━━━━━━━━━━━━━');
      var nbg = nextDaySatSummary.byGroup;
      lines.push(_formatGroupLine('新工場', nbg['新工場'] || { bento: 0, okazu: 0, total: 0, names: [] }));
      lines.push(_formatGroupLine('本社工場', nbg['本社工場'] || { bento: 0, okazu: 0, total: 0, names: [] }));
      lines.push('');
      lines.push('▼ 全体合計：' + nextDaySatSummary.grandTotal.total + '名（弁当' + nextDaySatSummary.grandTotal.bento + ' + おかずのみ' + nextDaySatSummary.grandTotal.okazu + '）');
      lines.push('');
      lines.push('【明日分・手配者名一覧】');
      nextDaySatSummary.groupOrder.forEach(function(g) {
        var d3 = nextDaySatSummary.byGroup[g];
        lines.push('■ ' + g + ':');
        lines.push('　' + (d3.names.length === 0 ? '（なし）' : d3.names.join('、')));
      });
      lines.push('');
      lines.push('※ 上記は本日朝時点の暫定値です。本日15時の締切後に変更があれば、別途「追加変更」メールでお知らせします。');
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
    var d = new Date(dateStr + 'T00:00:00+09:00');
    d.setDate(d.getDate() + 1);
    var nextStr = formatDateYmd(d);
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
      minimal.byGroup[g] = { bento: d.bento, okazu: d.okazu, total: d.total, names: d.names.slice() };
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
   * 名前リスト「氏名（弁当）」「氏名（おかずのみ）」を分解
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
    var satD = new Date(satDateStr + 'T00:00:00+09:00');
    var satLabel = Utilities.formatDate(satD, 'Asia/Tokyo', 'yyyy年MM月dd日');
    var timeStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm');
    var lines = [];
    lines.push('総務ご担当者様');
    lines.push('');
    lines.push('お疲れ様です。');
    lines.push('本日朝のメールでお知らせした明日（' + satLabel + '・出勤土曜）分の');
    lines.push('お弁当注文に変更がありました。最終確定状況をご連絡いたします。');
    lines.push('');
    lines.push('■ 対象日: ' + satLabel);
    lines.push('■ 締切時刻: 本日15:00');
    lines.push('■ 確定時刻: ' + timeStr);
    lines.push('');
    lines.push('━━━━━━━━━━━━━━━━━━━━');
    lines.push('【朝メールからの変更点】');
    lines.push('━━━━━━━━━━━━━━━━━━━━');
    lines.push('▼ 追加: ' + diff.added.length + '件');
    diff.added.forEach(function(s) { lines.push('　・' + s); });
    lines.push('▼ 取消: ' + diff.removed.length + '件');
    diff.removed.forEach(function(s) { lines.push('　・' + s); });
    lines.push('▼ 種別変更: ' + diff.changed.length + '件');
    diff.changed.forEach(function(c) { lines.push('　・' + c.name + '：' + c.from + ' → ' + c.to); });
    lines.push('');
    lines.push('━━━━━━━━━━━━━━━━━━━━');
    lines.push('【最終確定後の集計】');
    lines.push('━━━━━━━━━━━━━━━━━━━━');
    var bg = current.byGroup;
    lines.push(_formatGroupLine('新工場', bg['新工場'] || { bento: 0, okazu: 0, total: 0, names: [] }));
    lines.push(_formatGroupLine('本社工場', bg['本社工場'] || { bento: 0, okazu: 0, total: 0, names: [] }));
    lines.push('');
    lines.push('▼ 全体合計：' + current.grandTotal.total + '名（弁当' + current.grandTotal.bento + ' + おかずのみ' + current.grandTotal.okazu + '）');
    lines.push('');
    lines.push('━━━━━━━━━━━━━━━━━━━━');
    lines.push('【最終 手配者名一覧】');
    lines.push('━━━━━━━━━━━━━━━━━━━━');
    current.groupOrder.forEach(function(g) {
      var d = current.byGroup[g];
      lines.push('■ ' + g + ':');
      lines.push('　' + (d.names.length === 0 ? '（なし）' : d.names.join('、')));
    });
    lines.push('');
    lines.push('※ 本メールは出勤土曜の前日15時の締切後に変更があった場合のみ送信されます。');
    lines.push('※ このメールはお弁当予約アプリより自動送信されています。');
    return lines.join('\n');
  }

  /**
   * 出勤土曜の前日15時に呼ばれる追加変更メール送信
   * - 翌日が出勤土曜でなければ何もしない
   * - スナップショットと現在を比較し、差分があるときだけメール送信
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
      if (diff.added.length === 0 && diff.removed.length === 0 && diff.changed.length === 0) {
        Logger.log('Saturday supplemental: no diff for ' + satDate + ', skip');
        _clearSatSnapshot(satDate);
        return;
      }
      var recipients = SheetService.getMailRecipients();
      if (!recipients || recipients.length === 0) {
        Logger.log('Saturday supplemental: no recipients');
        return;
      }
      var satD = new Date(satDate + 'T00:00:00+09:00');
      var subject = '【お弁当注文表・追加変更】明日（出勤土曜 ' + Utilities.formatDate(satD, 'Asia/Tokyo', 'MM月dd日') + '）分の最終確定';
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
   * 日次レポートメール送信
   * - 翌日が出勤土曜の場合は、翌日分の暫定集計を本文に併記しスナップショット保存
   */
  function sendDailyReport() {
    try {
      var today = formatDateYmd(new Date());
      var summary = ReportService.generateReportSummary(today);
      var recipients = SheetService.getMailRecipients();
      if (!recipients || recipients.length === 0) {
        Logger.log('sendDailyReport: no recipients');
        return;
      }

      // 翌日が出勤土曜なら、翌日分も生成して本文に併記＋スナップショット保存
      var satDate = _nextDayIfWorkSat(today);
      var satSummary = null;
      if (satDate) {
        satSummary = ReportService.generateReportSummary(satDate);
        _saveSatSnapshot(satDate, satSummary);
        Logger.log('Saturday snapshot saved for ' + satDate);
      }

      var ss = ReportService.getOrCreateAdminSpreadsheet();
      var blob = null;
      try {
        blob = exportSpreadsheetAsExcel(ss.getId());
      } catch (e) {
        Logger.log('Excel export failed: ' + e.message);
      }

      var dLabel = Utilities.formatDate(new Date(today + 'T00:00:00+09:00'), 'Asia/Tokyo', 'yyyy年MM月dd日');
      var subject = '【お弁当注文表】本日分のお弁当注文一覧（' + dLabel + '）'
                  + (satSummary ? ' ＋翌出勤土曜分（暫定）' : '');
      var body = buildMailBody(summary, satSummary);

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
 * - 朝メール時のスナップショットと差分があれば追加変更メールを送信
 */
function saturdayDeadlineMailJob() {
  try {
    MailService.sendSaturdaySupplementalMail();
  } catch (err) {
    Logger.log('saturdayDeadlineMailJob error: ' + err.message);
  }
}

/**
 * サンプル: 出勤土曜の前日15時の追加変更メール（架空の差分データで送信）
 * - 朝メール時のスナップショット（17名）と、15時時点の最終確定（20名）を架空生成
 * - 追加4件 / 取消1件 / 種別変更1件 を含む差分を生成し、SAMPLE_MAIL_TO 全員に送信
 */
function sendSampleSaturdaySupplementalMail() {
  var today = formatDateYmd(new Date());
  var d = new Date(today + 'T00:00:00+09:00');
  d.setDate(d.getDate() + 1);
  var satDate = formatDateYmd(d);

  // 朝9時時点のスナップショット
  var snapshot = _buildSatSampleSummary(satDate,
    ['田村 修二', '佐藤 健', '鈴木 一郎', '高橋 美咲', '伊藤 大輔', '渡辺 真理', '山本 剛', '中村 結衣'],
    ['橋本 一夫', '石川 幸子'],
    ['後藤 雄一', '長谷川 涼', '村上 理恵', '近藤 雅人', '坂本 葵', '遠藤 健一'],
    ['竹内 大']
  );

  // 15時時点の最終確定（朝から変更あり）
  // - 追加: 小林 光（弁当）, 加藤 智子（弁当）, 前田 翔太（おかず）, 金子 由佳（おかず）
  // - 取消: 遠藤 健一（弁当）
  // - 種別変更: 山本 剛 弁当→おかず
  var current = _buildSatSampleSummary(satDate,
    ['田村 修二', '佐藤 健', '鈴木 一郎', '高橋 美咲', '伊藤 大輔', '渡辺 真理', '中村 結衣', '小林 光', '加藤 智子'],
    ['橋本 一夫', '石川 幸子', '前田 翔太', '山本 剛'],
    ['後藤 雄一', '長谷川 涼', '村上 理恵', '近藤 雅人', '坂本 葵'],
    ['竹内 大', '金子 由佳']
  );

  var diff = MailService.diffSatSummary(snapshot, current);
  var satLabel = Utilities.formatDate(d, 'Asia/Tokyo', 'MM月dd日');
  var subject = '[サンプル]【お弁当注文表・追加変更】明日（出勤土曜 ' + satLabel + '）分の最終確定';
  var body = '※ これはサンプルメールです。架空の差分データ（追加' + diff.added.length
           + '件・取消' + diff.removed.length + '件・種別変更' + diff.changed.length + '件）で生成しています。\n\n'
           + MailService.buildSatSupplementalBody(satDate, current, diff);

  // 添付: 最終確定後の土曜分Excel
  var options = {};
  var tempSsId = null;
  try {
    var tempSs = _buildOneDaySpreadsheetFromSummary(current, satDate, 'お弁当予約_サンプル土曜最終');
    tempSsId = tempSs.getId();
    SpreadsheetApp.flush();
    options.attachments = [MailService.exportSpreadsheetAsExcel(tempSsId)];
  } catch (e) {
    Logger.log('Sample supplemental: attachment generation failed: ' + e.message);
    body += '\n\n[添付生成エラー] ' + e.message;
  }

  try {
    SAMPLE_MAIL_TO.forEach(function(to) {
      try {
        MailApp.sendEmail(to, subject, body, options);
        Logger.log('Sample supplemental mail sent to: ' + to + ' (attachment=' + (options.attachments ? 'yes' : 'no') + ')');
      } catch (e) {
        Logger.log('Sample supplemental send failed to ' + to + ': ' + e.message);
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
 * サンプル: 金曜朝メール（本日(金)＋翌日(出勤土曜)分・暫定 を併記）
 * SAMPLE_MAIL_TO 全員に送信
 */
function sendSampleFridayMail() {
  var today = formatDateYmd(new Date());
  var d = new Date(today + 'T00:00:00+09:00');
  var nextD = new Date(d);
  nextD.setDate(nextD.getDate() + 1);
  var friDate = today;
  var satDate = formatDateYmd(nextD);

  // 金曜の本日分: 通常規模（35名）
  var friSummary = _buildSatSampleSummary(friDate,
    ['田村 修二', '佐藤 健', '鈴木 一郎', '高橋 美咲', '伊藤 大輔', '渡辺 真理', '山本 剛', '中村 結衣', '小林 光', '加藤 智子', '吉田 隆', '山田 花子', '松本 拓也', '井上 由美', '木村 慎一'],
    ['橋本 一夫', '石川 幸子', '前田 翔太'],
    ['後藤 雄一', '長谷川 涼', '村上 理恵', '近藤 雅人', '坂本 葵', '遠藤 健一', '青木 恵', '福田 武', '太田 美香', '西村 龍', '藤井 七海', '岡本 真', '三浦 美咲'],
    ['竹内 大', '金子 由佳', '和田 進', '中川 美鈴']
  );
  // 翌日の出勤土曜分・暫定: 17名
  var satSummary = _buildSatSampleSummary(satDate,
    ['田村 修二', '佐藤 健', '鈴木 一郎', '高橋 美咲', '伊藤 大輔', '渡辺 真理', '山本 剛', '中村 結衣'],
    ['橋本 一夫', '石川 幸子'],
    ['後藤 雄一', '長谷川 涼', '村上 理恵', '近藤 雅人', '坂本 葵', '遠藤 健一'],
    ['竹内 大']
  );

  var friLabel = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy年MM月dd日');
  var subject = '[サンプル]【お弁当注文表】本日分のお弁当注文一覧（' + friLabel + '）＋翌出勤土曜分（暫定）';
  var body = '※ これはサンプルメールです。架空の予約データで「金曜朝メール（翌日出勤土曜あり）」を再現しています。\n'
           + '※ 本日分(' + friSummary.grandTotal.total + '名) ＋ 翌出勤土曜分・暫定(' + satSummary.grandTotal.total + '名)\n\n'
           + MailService.buildMailBody(friSummary, satSummary);

  // 添付: 本日(金)と翌日(土・暫定)の2列スプレッドシート
  var options = {};
  var tempSsId = null;
  try {
    var friColLabel = Utilities.formatDate(d, 'Asia/Tokyo', 'MM月dd日') + '(金)';
    var satColLabel = Utilities.formatDate(nextD, 'Asia/Tokyo', 'MM月dd日') + '(土・暫定)';
    var tempSs = _buildFridaySampleSpreadsheet(friSummary, satSummary, friColLabel, satColLabel);
    tempSsId = tempSs.getId();
    SpreadsheetApp.flush();
    options.attachments = [MailService.exportSpreadsheetAsExcel(tempSsId)];
  } catch (e) {
    Logger.log('Sample Friday: attachment generation failed: ' + e.message);
    body += '\n\n[添付生成エラー] ' + e.message;
  }

  try {
    SAMPLE_MAIL_TO.forEach(function(to) {
      try {
        MailApp.sendEmail(to, subject, body, options);
        Logger.log('Sample Friday mail sent to: ' + to + ' (attachment=' + (options.attachments ? 'yes' : 'no') + ')');
      } catch (e) {
        Logger.log('Sample Friday send failed to ' + to + ': ' + e.message);
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
 * サンプル土曜分の集計データを生成
 */
function _buildSatSampleSummary(dateStr, shinB, shinO, honB, honO) {
  function _names(b, o) {
    return b.map(function(n) { return n + '（弁当）'; })
            .concat(o.map(function(n) { return n + '（おかずのみ）'; }));
  }
  var byGroup = {
    '新工場': {
      bento: shinB.length,
      okazu: shinO.length,
      total: shinB.length + shinO.length,
      names: _names(shinB, shinO)
    },
    '本社工場': {
      bento: honB.length,
      okazu: honO.length,
      total: honB.length + honO.length,
      names: _names(honB, honO)
    }
  };
  return {
    date: dateStr,
    byGroup: byGroup,
    grandTotal: {
      bento: byGroup['新工場'].bento + byGroup['本社工場'].bento,
      okazu: byGroup['新工場'].okazu + byGroup['本社工場'].okazu,
      total: byGroup['新工場'].total + byGroup['本社工場'].total
    },
    groupOrder: ['新工場', '本社工場'],
    _samplePeople: {
      '新工場': { bento: shinB, okazu: shinO },
      '本社工場': { bento: honB, okazu: honO }
    }
  };
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
 * サンプルメール送信（架空データ45名分・指定アドレス1件のみへ送信）
 * GASエディタから sendSampleMail を選んで手動実行する用途
 */
var SAMPLE_MAIL_TO = [
  'imaizumi@lineworks.co.jp',
  'c-matusita@lineworks.co.jp'
];

function sendSampleMail() {
  var today = formatDateYmd(new Date());
  var dLabel = Utilities.formatDate(new Date(today + 'T00:00:00+09:00'), 'Asia/Tokyo', 'yyyy年MM月dd日');
  var summary = _buildSampleSummary(today);
  var subject = '[サンプル]【お弁当注文表】本日分のお弁当注文一覧（' + dLabel + '）';
  var body = '※ これはサンプルメールです。実データではなく架空の予約データ（45名）で生成しています。\n\n'
           + MailService.buildMailBody(summary);

  var options = {};
  var tempSsId = null;
  try {
    var tempSs = _buildSampleAdminSpreadsheet(summary, today);
    tempSsId = tempSs.getId();
    SpreadsheetApp.flush();
    options.attachments = [MailService.exportSpreadsheetAsExcel(tempSsId)];
  } catch (e) {
    Logger.log('Sample mail: attachment generation failed: ' + e.message);
    body += '\n\n[添付生成エラー] ' + e.message;
  }

  try {
    SAMPLE_MAIL_TO.forEach(function(to) {
      try {
        MailApp.sendEmail(to, subject, body, options);
        Logger.log('Sample mail sent to: ' + to + ' (attachment=' + (options.attachments ? 'yes' : 'no') + ')');
      } catch (e) {
        Logger.log('Sample mail send failed to ' + to + ': ' + e.message);
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
 * サンプル用の架空集計データを構築（新工場25名 / 本社工場20名 = 計45名）
 */
function _buildSampleSummary(dateStr) {
  var shinBento = ['田村 修二', '佐藤 健', '鈴木 一郎', '高橋 美咲', '伊藤 大輔', '渡辺 真理', '山本 剛', '中村 結衣', '小林 光', '加藤 智子', '吉田 隆', '山田 花子', '松本 拓也', '井上 由美', '木村 慎一', '林 香織', '清水 浩二', '山口 美穂', '森 賢治', '池田 裕子'];
  var shinOkazu = ['橋本 一夫', '石川 幸子', '前田 翔太', '藤田 麻衣', '岡田 純一'];
  var honshaBento = ['後藤 雄一', '長谷川 涼', '村上 理恵', '近藤 雅人', '坂本 葵', '遠藤 健一', '青木 恵', '福田 武', '太田 美香', '西村 龍', '藤井 七海', '岡本 真', '三浦 美咲', '中島 修', '原 さつき'];
  var honshaOkazu = ['竹内 大', '金子 由佳', '和田 進', '中川 美鈴', '石井 研'];

  function _names(bentoArr, okazuArr) {
    var arr = [];
    bentoArr.forEach(function(n) { arr.push(n + '（弁当）'); });
    okazuArr.forEach(function(n) { arr.push(n + '（おかずのみ）'); });
    return arr;
  }

  var byGroup = {
    '新工場': {
      bento: shinBento.length,
      okazu: shinOkazu.length,
      total: shinBento.length + shinOkazu.length,
      names: _names(shinBento, shinOkazu)
    },
    '本社工場': {
      bento: honshaBento.length,
      okazu: honshaOkazu.length,
      total: honshaBento.length + honshaOkazu.length,
      names: _names(honshaBento, honshaOkazu)
    }
  };
  var grand = {
    bento: byGroup['新工場'].bento + byGroup['本社工場'].bento,
    okazu: byGroup['新工場'].okazu + byGroup['本社工場'].okazu,
    total: byGroup['新工場'].total + byGroup['本社工場'].total
  };
  return {
    date: dateStr,
    byGroup: byGroup,
    grandTotal: grand,
    groupOrder: ['新工場', '本社工場'],
    _samplePeople: {
      '新工場': { bento: shinBento, okazu: shinOkazu },
      '本社工場': { bento: honshaBento, okazu: honshaOkazu }
    }
  };
}

/**
 * "氏名（弁当）" / "氏名（おかずのみ）" を分解
 */
function _parseNTGlobal(nt) {
  var m = String(nt).match(/^(.+?)（(.+?)）$/);
  return m ? { name: m[1], type: m[2] } : { name: String(nt), type: '' };
}

/**
 * サマリの byGroup[group].names から bento/okazu の氏名リストを再構築
 */
function _peopleListsFromSummary(summary) {
  var result = { '新工場': { bento: [], okazu: [] }, '本社工場': { bento: [], okazu: [] } };
  ['新工場', '本社工場'].forEach(function(g) {
    var names = (summary.byGroup[g] && summary.byGroup[g].names) || [];
    names.forEach(function(nt) {
      var p = _parseNTGlobal(nt);
      if (p.type === '弁当') result[g].bento.push(p.name);
      else if (p.type === 'おかずのみ') result[g].okazu.push(p.name);
    });
  });
  return result;
}

/**
 * 拠点別シート1日分（氏名×1日付）を書き込み
 */
function _writeOneDayLocationSheet(sheet, dateLabel, bentoList, okazuList, deptLabel) {
  var header = ['氏名', '部署', dateLabel];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f0e8d8');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  var rows = [];
  bentoList.forEach(function(n) { rows.push([n, deptLabel || '', '○']); });
  okazuList.forEach(function(n) { rows.push([n, deptLabel || '', 'お']); });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  var startRow = 2 + rows.length + 1;
  var summaryRows = [
    ['弁当', '', bentoList.length],
    ['おかずのみ', '', okazuList.length],
    ['合計', '', bentoList.length + okazuList.length]
  ];
  sheet.getRange(startRow, 1, 3, header.length).setValues(summaryRows);
  sheet.getRange(startRow, 1, 3, 1).setFontWeight('bold');
  sheet.getRange(startRow, 1, 3, header.length).setBackground('#fff8ea');
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

  var rows = [];
  ['新工場', '本社工場'].forEach(function(g) {
    peopleByGroup[g].bento.forEach(function(n) { rows.push([n, g, '○']); });
    peopleByGroup[g].okazu.forEach(function(n) { rows.push([n, g, 'お']); });
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  var totalBento = peopleByGroup['新工場'].bento.length + peopleByGroup['本社工場'].bento.length;
  var totalOkazu = peopleByGroup['新工場'].okazu.length + peopleByGroup['本社工場'].okazu.length;
  var startRow = 2 + rows.length + 1;
  var summaryRows = [
    ['弁当', '', totalBento],
    ['おかずのみ', '', totalOkazu],
    ['合計', '', totalBento + totalOkazu]
  ];
  sheet.getRange(startRow, 1, 3, header.length).setValues(summaryRows);
  sheet.getRange(startRow, 1, 3, 1).setFontWeight('bold');
  sheet.getRange(startRow, 1, 3, header.length).setBackground('#fff2e0');
  sheet.autoResizeColumn(1);
}

/**
 * 拠点別シート複数日分（氏名×複数日付）
 */
function _writeMultiDayLocationSheet(sheet, dateLabels, peopleByDate, deptLabel) {
  // peopleByDate: [ { bento:[], okazu:[] }, ... ] 同じ並び順
  var n = dateLabels.length;
  var header = ['氏名', '部署'].concat(dateLabels);
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f0e8d8');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // 全期間で登場した氏名を集約
  var perPerson = {};
  peopleByDate.forEach(function(p, idx) {
    p.bento.forEach(function(name) {
      if (!perPerson[name]) perPerson[name] = new Array(n).fill('');
      perPerson[name][idx] = '○';
    });
    p.okazu.forEach(function(name) {
      if (!perPerson[name]) perPerson[name] = new Array(n).fill('');
      perPerson[name][idx] = 'お';
    });
  });
  var names = Object.keys(perPerson);
  var rows = names.map(function(name) {
    return [name, deptLabel || ''].concat(perPerson[name]);
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  // 集計行
  var startRow = 2 + rows.length + 1;
  var bentoRow = ['弁当', ''], okazuRow = ['おかずのみ', ''], totalRow = ['合計', ''];
  for (var i = 0; i < n; i++) {
    bentoRow.push(peopleByDate[i].bento.length);
    okazuRow.push(peopleByDate[i].okazu.length);
    totalRow.push(peopleByDate[i].bento.length + peopleByDate[i].okazu.length);
  }
  sheet.getRange(startRow, 1, 3, header.length).setValues([bentoRow, okazuRow, totalRow]);
  sheet.getRange(startRow, 1, 3, 1).setFontWeight('bold');
  sheet.getRange(startRow, 1, 3, header.length).setBackground('#fff8ea');
  sheet.autoResizeColumn(1);
}

/**
 * 全拠点合計シート複数日分（拠点列付き）
 */
function _writeMultiDayCombinedSheet(sheet, dateLabels, peopleByDateByGroup) {
  // peopleByDateByGroup: [ { '新工場':{bento,okazu}, '本社工場':{bento,okazu} }, ...]
  var n = dateLabels.length;
  var header = ['氏名', '拠点'].concat(dateLabels);
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f0e8d8');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // (氏名, 拠点) ごとに集約
  var perPerson = {};
  peopleByDateByGroup.forEach(function(byGroup, idx) {
    ['新工場', '本社工場'].forEach(function(g) {
      byGroup[g].bento.forEach(function(name) {
        var key = name + '|' + g;
        if (!perPerson[key]) perPerson[key] = { name: name, group: g, cells: new Array(n).fill('') };
        perPerson[key].cells[idx] = '○';
      });
      byGroup[g].okazu.forEach(function(name) {
        var key = name + '|' + g;
        if (!perPerson[key]) perPerson[key] = { name: name, group: g, cells: new Array(n).fill('') };
        perPerson[key].cells[idx] = 'お';
      });
    });
  });

  // 拠点ごとに並べる: 新工場 → 本社工場
  var rows = [];
  ['新工場', '本社工場'].forEach(function(g) {
    Object.keys(perPerson).forEach(function(key) {
      var p = perPerson[key];
      if (p.group === g) {
        rows.push([p.name, g].concat(p.cells));
      }
    });
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  // 集計行
  var startRow = 2 + rows.length + 1;
  var bentoRow = ['弁当', ''], okazuRow = ['おかずのみ', ''], totalRow = ['合計', ''];
  for (var i = 0; i < n; i++) {
    var b = peopleByDateByGroup[i]['新工場'].bento.length + peopleByDateByGroup[i]['本社工場'].bento.length;
    var o = peopleByDateByGroup[i]['新工場'].okazu.length + peopleByDateByGroup[i]['本社工場'].okazu.length;
    bentoRow.push(b); okazuRow.push(o); totalRow.push(b + o);
  }
  sheet.getRange(startRow, 1, 3, header.length).setValues([bentoRow, okazuRow, totalRow]);
  sheet.getRange(startRow, 1, 3, 1).setFontWeight('bold');
  sheet.getRange(startRow, 1, 3, header.length).setBackground('#fff2e0');
  sheet.autoResizeColumn(1);
}

/**
 * 1日分のサンプル添付スプレッドシートを生成（全拠点合計シートも全員の名前付き）
 */
function _buildSampleAdminSpreadsheet(summary, dateStr) {
  var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var ss = SpreadsheetApp.create('お弁当予約_サンプル_' + stamp);
  var people = summary._samplePeople || _peopleListsFromSummary(summary);

  ['新工場', '本社工場'].forEach(function(group) {
    var sheet = ss.insertSheet(group);
    _writeOneDayLocationSheet(sheet, dateStr, people[group].bento, people[group].okazu, '（サンプル）');
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
  var friPeople = friSummary._samplePeople || _peopleListsFromSummary(friSummary);
  var satPeople = satSummary._samplePeople || _peopleListsFromSummary(satSummary);
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
    _writeOneDayLocationSheet(sheet, dateStr, people[group].bento, people[group].okazu, '');
  });

  var totalSheet = ss.insertSheet('全拠点合計');
  _writeOneDayCombinedSheet(totalSheet, dateStr, people);

  var def = ss.getSheetByName('シート1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  return ss;
}
