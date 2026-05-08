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
      recipients.forEach(function(to) {
        try {
          MailApp.sendEmail(to, subject, body);
          Logger.log('Saturday supplemental sent to: ' + to);
        } catch (e) {
          Logger.log('Saturday supplemental send failed to ' + to + ': ' + e.message);
        }
      });
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
    buildMailBody: buildMailBody
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
function sendSampleMail() {
  var to = 'imaizumi@lineworks.co.jp';
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
    MailApp.sendEmail(to, subject, body, options);
    Logger.log('Sample mail sent to: ' + to + ' (attachment=' + (options.attachments ? 'yes' : 'no') + ')');
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
 * サンプル添付用の一時スプレッドシートを生成
 */
function _buildSampleAdminSpreadsheet(summary, dateStr) {
  var stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  var ss = SpreadsheetApp.create('お弁当予約_サンプル_' + stamp);

  ['新工場', '本社工場'].forEach(function(group) {
    var sheet = ss.insertSheet(group);
    var people = summary._samplePeople[group];
    var header = ['氏名', '部署', dateStr];
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f0e8d8');
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);

    var rows = [];
    people.bento.forEach(function(name) { rows.push([name, '（サンプル）', '○']); });
    people.okazu.forEach(function(name) { rows.push([name, '（サンプル）', 'お']); });

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
    }

    var summaryStartRow = 2 + rows.length + 1;
    var summaryRows = [
      ['弁当', '', people.bento.length],
      ['おかずのみ', '', people.okazu.length],
      ['合計', '', people.bento.length + people.okazu.length]
    ];
    sheet.getRange(summaryStartRow, 1, 3, header.length).setValues(summaryRows);
    sheet.getRange(summaryStartRow, 1, 3, 1).setFontWeight('bold');
    sheet.getRange(summaryStartRow, 1, 3, header.length).setBackground('#fff8ea');
    sheet.autoResizeColumn(1);
  });

  // 全拠点合計シート
  var totalSheet = ss.insertSheet('全拠点合計');
  totalSheet.getRange(1, 1, 1, 4).setValues([['拠点グループ', '弁当', 'おかずのみ', '合計']]);
  totalSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f0e8d8');
  totalSheet.setFrozenRows(1);
  var totalRows = [];
  ['新工場', '本社工場'].forEach(function(g) {
    var d = summary.byGroup[g];
    totalRows.push([g, d.bento, d.okazu, d.total]);
  });
  totalRows.push(['【全合計】', summary.grandTotal.bento, summary.grandTotal.okazu, summary.grandTotal.total]);
  totalSheet.getRange(2, 1, totalRows.length, 4).setValues(totalRows);
  totalSheet.getRange(2 + totalRows.length - 1, 1, 1, 4).setFontWeight('bold').setBackground('#fff2e0');

  // デフォルトの「シート1」を削除
  var def = ss.getSheetByName('シート1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  return ss;
}
