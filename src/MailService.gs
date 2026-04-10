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
    var shinOffice = bg['新工場事務職員'] || { bento: 0, okazu: 0, total: 0, names: [] };
    var shinFactory = bg['新工場'] || { bento: 0, okazu: 0, total: 0, names: [] };
    var honshaOffice = bg['本社工場事務職員'] || { bento: 0, okazu: 0, total: 0, names: [] };
    var honshaFactory = bg['本社工場'] || { bento: 0, okazu: 0, total: 0, names: [] };
    var shinTotalCount = shinOffice.total + shinFactory.total;
    var honshaTotalCount = honshaOffice.total + honshaFactory.total;
    lines.push('▼ 新工場');
    lines.push('　事務所：' + shinOffice.total + '名（弁当' + shinOffice.bento + ' / おかずのみ' + shinOffice.okazu + '）');
    lines.push('　工 場：' + shinFactory.total + '名（弁当' + shinFactory.bento + ' / おかずのみ' + shinFactory.okazu + '）');
    lines.push('　新工場合計：' + shinTotalCount + '名');
    lines.push('');
    lines.push('▼ 本社工場');
    lines.push('　事務所：' + honshaOffice.total + '名(弁当' + honshaOffice.bento + ' / おかずのみ' + honshaOffice.okazu + ')');
    lines.push('　工 場：' + honshaFactory.total + '名(弁当' + honshaFactory.bento + ' / おかずのみ' + honshaFactory.okazu + ')');
    lines.push('　本社工場合計：' + honshaTotalCount + '名');
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
