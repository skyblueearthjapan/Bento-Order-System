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
   * メール本文を生成
   */
  function buildMailBody(summary) {
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
    lines.push('▼ 新工場：' + shin.total + '名（弁当' + shin.bento + ' / おかずのみ' + shin.okazu + '）');
    lines.push('▼ 本社工場：' + honsha.total + '名（弁当' + honsha.bento + ' / おかずのみ' + honsha.okazu + '）');
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
    lines.push('');
    lines.push('※ 詳細は添付Excelをご参照ください。');
    lines.push('※ このメールはお弁当予約アプリより自動送信されています。');
    lines.push('※ 内容にご不明点がございましたら、システム管理者までお問い合わせください。');
    return lines.join('\n');
  }

  /**
   * 日次レポートメール送信
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

      var ss = ReportService.getOrCreateAdminSpreadsheet();
      var blob = null;
      try {
        blob = exportSpreadsheetAsExcel(ss.getId());
      } catch (e) {
        Logger.log('Excel export failed: ' + e.message);
      }

      var dLabel = Utilities.formatDate(new Date(today + 'T00:00:00+09:00'), 'Asia/Tokyo', 'yyyy年MM月dd日');
      var subject = '【お弁当注文表】本日分のお弁当注文一覧（' + dLabel + '）';
      var body = buildMailBody(summary);

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
