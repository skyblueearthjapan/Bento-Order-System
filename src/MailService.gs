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
    lines.push('【本日の手配内訳】');
    summary.groupOrder.forEach(function(g) {
      var d2 = summary.byGroup[g];
      lines.push('　' + g + '：弁当' + d2.bento + '個 / おかずのみ' + d2.okazu + '個 / 計' + d2.total + '個（' + d2.names.length + '名）');
    });
    lines.push('');
    lines.push('【合計】');
    lines.push('　弁当' + summary.grandTotal.bento + '個 / おかずのみ' + summary.grandTotal.okazu + '個 / 合計' + summary.grandTotal.total + '個');
    lines.push('');
    lines.push('【手配者名一覧】');
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
