/**
 * ReportService.gs - 総務管理用スプレッドシートの生成・更新
 */

var ReportService = (function() {

  /**
   * 総務管理用スプレッドシートを取得 or 作成
   */
  function getOrCreateAdminSpreadsheet() {
    var year = new Date().getFullYear();
    var fileName = 'お弁当予約_総務管理_' + year + '年度';
    var folder = DriveApp.getFolderById(ADMIN_OUTPUT_FOLDER_ID);
    var files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      var f = files.next();
      return SpreadsheetApp.openById(f.getId());
    }
    var newSs = SpreadsheetApp.create(fileName);
    var newFile = DriveApp.getFileById(newSs.getId());
    try {
      newFile.moveTo(folder);
    } catch (e) {
      // 旧API フォールバック
      try {
        folder.addFile(newFile);
        DriveApp.getRootFolder().removeFile(newFile);
      } catch (e2) {
        Logger.log('moveTo failed: ' + e2.message);
      }
    }
    return newSs;
  }

  /**
   * 日付から月度キーを生成（16日以降は翌月度扱い）
   * 例: 2026-04-15 → '2026年04月度', 2026-04-16 → '2026年05月度'
   */
  function getMonthKey(dateOrStr) {
    var d = typeof dateOrStr === 'string' ? new Date(dateOrStr + 'T00:00:00+09:00') : dateOrStr;
    var y = d.getFullYear();
    var m = d.getMonth() + 1;  // 1-12
    var day = d.getDate();
    if (day >= RULES.MONTH_START_DAY) {
      m++;
      if (m > 12) { m = 1; y++; }
    }
    return y + '年' + ('0' + m).slice(-2) + '月度';
  }

  /**
   * 月度キーから対象期間 {from, to} を取得（前月16日〜当月15日）
   */
  function getMonthRange(monthKey) {
    var match = monthKey.match(/^(\d{4})年(\d{2})月度$/);
    if (!match) return null;
    var y = parseInt(match[1], 10);
    var m = parseInt(match[2], 10);  // 当月
    // from: 前月16日
    var fromY = y, fromM = m - 1;
    if (fromM < 1) { fromM = 12; fromY--; }
    var from = fromY + '-' + ('0' + fromM).slice(-2) + '-16';
    // to: 当月15日
    var to = y + '-' + ('0' + m).slice(-2) + '-15';
    return { from: from, to: to };
  }

  /**
   * 期間内の日付リストを生成
   */
  function _listDates(fromDate, toDate) {
    var result = [];
    var cur = new Date(fromDate + 'T00:00:00+09:00');
    var end = new Date(toDate + 'T00:00:00+09:00');
    while (cur <= end) {
      result.push(formatDateYmd(cur));
      cur.setDate(cur.getDate() + 1);
    }
    return result;
  }

  /**
   * 日別配達拠点を解決（予約の拠点を優先、なければマスタのデフォルト）
   */
  function resolveLocationForDate(workerCode, dateStr, reservationsMap, workersMap) {
    var key = workerCode + '_' + dateStr;
    var rez = reservationsMap[key];
    if (rez && rez.location) return rez.location;
    var w = workersMap[workerCode];
    return w ? w.location : '';
  }

  /**
   * 予約配列をキー{code}_{date}でマップ化
   */
  function _buildReservationsMap(reservations) {
    var map = {};
    reservations.forEach(function(r) {
      map[r.workerCode + '_' + r.date] = r;
    });
    return map;
  }

  /**
   * 拠点+スタッフ種類でグループ分け
   */
  function _getGroupKey(location, staffType) {
    if (location === LOCATION.SHIN && staffType === STAFF_TYPE.OFFICE) return '新工場事務職員';
    if (location === LOCATION.SHIN && staffType === STAFF_TYPE.FACTORY) return '新工場';
    if (location === LOCATION.HONSHA && staffType === STAFF_TYPE.OFFICE) return '本社工場事務職員';
    if (location === LOCATION.HONSHA && staffType === STAFF_TYPE.FACTORY) return '本社工場';
    return null;
  }

  var GROUP_ORDER = ['新工場事務職員', '新工場', '本社工場事務職員', '本社工場'];

  /**
   * 指定月度の管理シート群を更新
   */
  function updateAdminSheetsForMonth(monthKey) {
    var range = getMonthRange(monthKey);
    if (!range) return;

    var ss = getOrCreateAdminSpreadsheet();
    var workers = SheetService.getWorkers();
    var workersMap = {};
    workers.forEach(function(w) { workersMap[w.code] = w; });
    var reservations = SheetService.getReservationsInRange(range.from, range.to);
    var reservationsMap = _buildReservationsMap(reservations);
    var dates = _listDates(range.from, range.to);

    GROUP_ORDER.forEach(function(group) {
      var sheetName = monthKey + '_' + group;
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      } else {
        sheet.clear();
      }

      // ヘッダー: A=氏名, B〜=日付
      var header = ['氏名'].concat(dates);
      sheet.getRange(1, 1, 1, header.length).setValues([header]);
      sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f0e8d8');
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(1);

      // 該当グループの作業員（日別拠点解決後の配属先で判定）
      // 各日について、その日の配達拠点がこのグループに該当する作業員を含める
      // 簡略化: マスタ拠点+スタッフ種類でまず絞り込み、日別で拠点変更があれば別グループへ動的に
      var rows = [];
      var groupWorkers = workers.filter(function(w) {
        return _getGroupKey(w.location, w.staffType) === group;
      });

      groupWorkers.forEach(function(w) {
        var row = [w.name];
        dates.forEach(function(d) {
          var loc = resolveLocationForDate(w.code, d, reservationsMap, workersMap);
          var g = _getGroupKey(loc, w.staffType);
          // この日、このワーカーがこのグループに属するか
          if (g !== group) {
            row.push('');
            return;
          }
          var rez = reservationsMap[w.code + '_' + d];
          if (!rez) { row.push(''); return; }
          if (rez.orderState === ORDER_STATE.BENTO) row.push('○');
          else if (rez.orderState === ORDER_STATE.OKAZU) row.push('お');
          else row.push('');
        });
        rows.push(row);
      });

      // 日別拠点変更で他グループから来た人も追加
      workers.forEach(function(w) {
        if (groupWorkers.indexOf(w) >= 0) return;  // 既に処理済み
        // この作業員がこの月度のいずれかの日でこのグループに属していればリストに追加
        var hasAny = false;
        var row = [w.name + '※'];  // 一時配属を示す印
        dates.forEach(function(d) {
          var loc = resolveLocationForDate(w.code, d, reservationsMap, workersMap);
          var g = _getGroupKey(loc, w.staffType);
          if (g !== group) { row.push(''); return; }
          var rez = reservationsMap[w.code + '_' + d];
          if (!rez) { row.push(''); return; }
          if (rez.orderState === ORDER_STATE.BENTO) { row.push('○'); hasAny = true; }
          else if (rez.orderState === ORDER_STATE.OKAZU) { row.push('お'); hasAny = true; }
          else row.push('');
        });
        if (hasAny) rows.push(row);
      });

      // 空行スキップ: データのない行は除外してから1回だけ書き込み
      var displayRows = rows.filter(function(r) {
        for (var c = 1; c < r.length; c++) { if (r[c]) return true; }
        return false;
      });
      if (displayRows.length > 0) {
        sheet.getRange(2, 1, displayRows.length, header.length).setValues(displayRows);
      }

      // 集計行: 弁当 / おかずのみ / 合計
      var summaryStartRow = 2 + displayRows.length + 1;
      var summaryRows = [
        ['弁当'].concat(dates.map(function(d, i) {
          var col = i + 1;  // rowsは[氏名, d0, d1, ...]なのでd_iはr[i+1]
          var count = 0;
          displayRows.forEach(function(r) { if (r[col] === '○') count++; });
          return count;
        })),
        ['おかずのみ'].concat(dates.map(function(d, i) {
          var col = i + 1;
          var count = 0;
          displayRows.forEach(function(r) { if (r[col] === 'お') count++; });
          return count;
        })),
        ['合計'].concat(dates.map(function(d, i) {
          var col = i + 1;
          var count = 0;
          displayRows.forEach(function(r) { if (r[col] === '○' || r[col] === 'お') count++; });
          return count;
        }))
      ];
      sheet.getRange(summaryStartRow, 1, 3, header.length).setValues(summaryRows);
      sheet.getRange(summaryStartRow, 1, 3, 1).setFontWeight('bold');
      sheet.getRange(summaryStartRow, 1, 3, header.length).setBackground('#fff8ea');

      // 列幅調整
      sheet.autoResizeColumns(1, 1);
    });

    // ===== 全グループ合計シート =====
    var totalSheetName = monthKey + '_全拠点合計';
    var totalSheet = ss.getSheetByName(totalSheetName);
    if (!totalSheet) {
      totalSheet = ss.insertSheet(totalSheetName);
    } else {
      totalSheet.clear();
    }
    var totalHeader = ['拠点グループ'].concat(dates).concat(['月度合計(弁当)', '月度合計(おかず)', '月度合計']);
    totalSheet.getRange(1, 1, 1, totalHeader.length).setValues([totalHeader]);
    totalSheet.getRange(1, 1, 1, totalHeader.length).setFontWeight('bold').setBackground('#f0e8d8');
    totalSheet.setFrozenRows(1);
    totalSheet.setFrozenColumns(1);

    // 日別・グループ別の集計を1日ずつ計算
    var totalRows = [];
    GROUP_ORDER.forEach(function(group) {
      var monthB = 0, monthO = 0;
      var dayCounts = dates.map(function(dateStr) {
        var bento = 0, okazu = 0;
        workers.forEach(function(w) {
          var loc = resolveLocationForDate(w.code, dateStr, reservationsMap, workersMap);
          var g = _getGroupKey(loc, w.staffType);
          if (g !== group) return;
          var rez = reservationsMap[w.code + '_' + dateStr];
          if (!rez) return;
          if (rez.orderState === ORDER_STATE.BENTO) bento++;
          else if (rez.orderState === ORDER_STATE.OKAZU) okazu++;
        });
        monthB += bento;
        monthO += okazu;
        return bento + okazu;
      });
      totalRows.push([group].concat(dayCounts).concat([monthB, monthO, monthB + monthO]));
    });

    // 総合計行
    var grandRow = ['【全合計】'];
    var grandB = 0, grandO = 0;
    dates.forEach(function(dateStr, i) {
      var dayB = 0, dayO = 0;
      workers.forEach(function(w) {
        var loc = resolveLocationForDate(w.code, dateStr, reservationsMap, workersMap);
        var g = _getGroupKey(loc, w.staffType);
        if (!g) return;
        var rez = reservationsMap[w.code + '_' + dateStr];
        if (!rez) return;
        if (rez.orderState === ORDER_STATE.BENTO) dayB++;
        else if (rez.orderState === ORDER_STATE.OKAZU) dayO++;
      });
      grandB += dayB;
      grandO += dayO;
      grandRow.push(dayB + dayO);
    });
    grandRow.push(grandB);
    grandRow.push(grandO);
    grandRow.push(grandB + grandO);
    totalRows.push(grandRow);

    totalSheet.getRange(2, 1, totalRows.length, totalHeader.length).setValues(totalRows);
    totalSheet.getRange(2 + totalRows.length - 1, 1, 1, totalHeader.length).setFontWeight('bold').setBackground('#fff2e0');

    // 最初の「シート1」があれば削除
    var defaultSheet = ss.getSheetByName('シート1');
    if (defaultSheet && ss.getSheets().length > 1) {
      ss.deleteSheet(defaultSheet);
    }
  }

  /**
   * 本日を含む月度シートを更新（トリガー・手動更新用）
   */
  function updateAdminSheetsForToday() {
    var monthKey = getMonthKey(new Date());
    updateAdminSheetsForMonth(monthKey);
  }

  /**
   * 本日分のメール用集計データを生成
   */
  function generateReportSummary(dateStr) {
    if (!dateStr) dateStr = formatDateYmd(new Date());
    var workers = SheetService.getWorkers();
    var workersMap = {};
    workers.forEach(function(w) { workersMap[w.code] = w; });
    var reservations = SheetService.getReservationsInRange(dateStr, dateStr);
    var reservationsMap = _buildReservationsMap(reservations);

    var byGroup = {};
    GROUP_ORDER.forEach(function(g) {
      byGroup[g] = { bento: 0, okazu: 0, total: 0, names: [] };
    });

    reservations.forEach(function(r) {
      if (!r.orderState) return;
      var w = workersMap[r.workerCode];
      if (!w) return;
      var loc = r.location || w.location;
      var group = _getGroupKey(loc, w.staffType);
      if (!group) return;
      if (r.orderState === ORDER_STATE.BENTO) {
        byGroup[group].bento++;
        byGroup[group].total++;
        byGroup[group].names.push(w.name + '（弁当）');
      } else if (r.orderState === ORDER_STATE.OKAZU) {
        byGroup[group].okazu++;
        byGroup[group].total++;
        byGroup[group].names.push(w.name + '（おかずのみ）');
      }
    });

    var grand = { bento: 0, okazu: 0, total: 0 };
    GROUP_ORDER.forEach(function(g) {
      grand.bento += byGroup[g].bento;
      grand.okazu += byGroup[g].okazu;
      grand.total += byGroup[g].total;
    });

    return { date: dateStr, byGroup: byGroup, grandTotal: grand, groupOrder: GROUP_ORDER };
  }

  return {
    getOrCreateAdminSpreadsheet: getOrCreateAdminSpreadsheet,
    getMonthKey: getMonthKey,
    getMonthRange: getMonthRange,
    updateAdminSheetsForMonth: updateAdminSheetsForMonth,
    updateAdminSheetsForToday: updateAdminSheetsForToday,
    generateReportSummary: generateReportSummary,
    GROUP_ORDER: GROUP_ORDER
  };
})();

/**
 * トリガーから呼ばれるラッパー
 */
function updateAdminSheetsForToday() {
  return ReportService.updateAdminSheetsForToday();
}
