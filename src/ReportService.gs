/**
 * ReportService.gs - 総務管理用スプレッドシートの生成・更新
 */

var ReportService = (function() {

  /**
   * 月度別の総務管理用スプレッドシートを取得 or 作成
   * - monthKey 省略時は本日を含む月度（16日以降は翌月度扱い）
   * - 月度ごとに独立したファイルとなる（例: お弁当予約_総務管理_2026年05月度）
   */
  function getOrCreateAdminSpreadsheet(monthKey) {
    monthKey = monthKey || getMonthKey(new Date());
    var fileName = 'お弁当予約_総務管理_' + monthKey;
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
   * 同一作業員コードの行を1行に絞る（マスタの部署重複登録への対策）
   * マスタDBに「1人が複数部署」のレコードがあると、レポート生成時に
   * 同じ予約が2行に展開されて発注が二重になるので、レポート用途では
   * 作業員コード単位で先頭行のみ採用する。マスタDBは変更しない。
   */
  function _uniqueWorkersByCode(workers) {
    var seen = {};
    var result = [];
    for (var i = 0; i < workers.length; i++) {
      var code = workers[i].code;
      if (!code) continue;
      if (seen[code]) continue;
      seen[code] = true;
      result.push(workers[i]);
    }
    return result;
  }

  /**
   * 拠点でグループ分け（事務所/工場の区別なし、拠点のみで集計）
   */
  function _getGroupKey(location) {
    if (location === LOCATION.SHIN) return '新工場';
    if (location === LOCATION.HONSHA) return '本社工場';
    return null;
  }

  var GROUP_ORDER = ['新工場', '本社工場'];

  // ===== メニューごとの記号・色 =====
  // ※ ファイルの読み込み順に依存しないよう、MENU_DEFS は使うときに読む
  function _symMap() {
    var m = {};
    MENU_DEFS.forEach(function(d) { m[d.key] = d.sym; });
    return m;
  }
  var CELL_BG = { men: '#e4ebf8', don: '#f8e0e0', toku: '#ffe9a8' };

  function _symOf(state) { return _symMap()[state] || ''; }

  /**
   * この月度の Excel に載せるメニュー
   * 弁当・おかずのみは常に。麺・丼・特弁は、予約があるか受付日がある月度だけ
   */
  function _menusForRange(dates, reservations, ctx) {
    return MENU_DEFS.filter(function(m) {
      if (m.kind === 'always') return true;
      if (reservations.some(function(r) { return r.orderState === m.key; })) return true;
      return dates.some(function(d) { return MenuService.isOffered(m.key, d, ctx); });
    });
  }

  /**
   * 指定月度の管理シート群を更新
   * 月度ごとに独立したスプレッドシートを使用するため、シート名は拠点名のみでOK
   */
  function updateAdminSheetsForMonth(monthKey) {
    var range = getMonthRange(monthKey);
    if (!range) return;

    var ss = getOrCreateAdminSpreadsheet(monthKey);
    var workers = _uniqueWorkersByCode(SheetService.getWorkers());
    var workersMap = {};
    workers.forEach(function(w) { workersMap[w.code] = w; });
    var reservations = SheetService.getReservationsInRange(range.from, range.to);
    var reservationsMap = _buildReservationsMap(reservations);
    var dates = _listDates(range.from, range.to);
    var menus = _menusForRange(dates, reservations, MenuService.loadContext());

    GROUP_ORDER.forEach(function(group) {
      var sheetName = group;
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      } else {
        sheet.clear();
      }

      // ヘッダー: A=氏名, B=部署, C〜=日付
      var header = ['氏名', '部署'].concat(dates);
      sheet.getRange(1, 1, 1, header.length).setValues([header]);
      sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f0e8d8');
      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(2);

      // 該当グループの作業員（日別拠点解決後の配属先で判定）
      // マスタ拠点でまず絞り込み、日別で拠点変更があれば別グループへ動的に
      var rows = [];
      var groupWorkers = workers.filter(function(w) {
        return _getGroupKey(w.location) === group;
      });

      function cellsFor(w) {
        var hasAny = false;
        var cells = dates.map(function(d) {
          var loc = resolveLocationForDate(w.code, d, reservationsMap, workersMap);
          if (_getGroupKey(loc) !== group) return '';
          var rez = reservationsMap[w.code + '_' + d];
          var sym = rez ? _symOf(rez.orderState) : '';
          if (sym) hasAny = true;
          return sym;
        });
        return { cells: cells, hasAny: hasAny };
      }

      groupWorkers.forEach(function(w) {
        rows.push([w.name, w.dept].concat(cellsFor(w).cells));
      });

      // 日別拠点変更で他グループから来た人も追加
      workers.forEach(function(w) {
        if (groupWorkers.indexOf(w) >= 0) return;  // 既に処理済み
        var c = cellsFor(w);
        if (c.hasAny) rows.push([w.name + '※', w.dept].concat(c.cells));  // ※ = 一時配属
      });

      // 空行スキップ: データのない行は除外してから1回だけ書き込み
      var displayRows = rows.filter(function(r) {
        for (var c = 2; c < r.length; c++) { if (r[c]) return true; }
        return false;
      });
      if (displayRows.length > 0) {
        sheet.getRange(2, 1, displayRows.length, header.length).setValues(displayRows);
        // 麺・丼・特弁のセルに色を付けて見分けやすくする
        var SYM = _symMap();
        var bgs = displayRows.map(function(r) {
          return r.map(function(v, i) {
            if (i < 2) return null;
            if (v === SYM.men) return CELL_BG.men;
            if (v === SYM.don) return CELL_BG.don;
            if (v === SYM.toku) return CELL_BG.toku;
            return null;
          });
        });
        sheet.getRange(2, 1, displayRows.length, header.length).setBackgrounds(bgs);
      }

      // 集計行: メニューごと + 合計
      var summaryStartRow = 2 + displayRows.length + 1;
      var allSyms = menus.map(function(m) { return m.sym; });
      var summaryRows = menus.map(function(m) {
        return [m.label, ''].concat(dates.map(function(d, i) {
          var col = i + 2;
          var count = 0;
          displayRows.forEach(function(r) { if (r[col] === m.sym) count++; });
          return count;
        }));
      });
      summaryRows.push(['合計', ''].concat(dates.map(function(d, i) {
        var col = i + 2;
        var count = 0;
        displayRows.forEach(function(r) { if (allSyms.indexOf(r[col]) !== -1) count++; });
        return count;
      })));
      sheet.getRange(summaryStartRow, 1, summaryRows.length, header.length).setValues(summaryRows);
      sheet.getRange(summaryStartRow, 1, summaryRows.length, 1).setFontWeight('bold');
      sheet.getRange(summaryStartRow, 1, summaryRows.length, header.length).setBackground('#fff8ea');
      sheet.getRange(summaryStartRow + summaryRows.length - 1, 1, 1, header.length).setBackground('#fff2e0');

      // 列幅調整
      sheet.autoResizeColumns(1, 1);
    });

    // ===== 全拠点合計シート（拠点 × メニューごとの日別個数） =====
    var totalSheetName = '全拠点合計';
    var totalSheet = ss.getSheetByName(totalSheetName);
    if (!totalSheet) {
      totalSheet = ss.insertSheet(totalSheetName);
    } else {
      totalSheet.clear();
    }
    var totalHeader = ['拠点', 'メニュー'].concat(dates).concat(['月度合計']);
    totalSheet.getRange(1, 1, 1, totalHeader.length).setValues([totalHeader]);
    totalSheet.getRange(1, 1, 1, totalHeader.length).setFontWeight('bold').setBackground('#f0e8d8');
    totalSheet.setFrozenRows(1);
    totalSheet.setFrozenColumns(2);

    // counts[group][menu][dateIndex]
    var counts = {};
    GROUP_ORDER.forEach(function(g) {
      counts[g] = {};
      menus.forEach(function(m) { counts[g][m.key] = dates.map(function() { return 0; }); });
    });
    dates.forEach(function(dateStr, i) {
      workers.forEach(function(w) {
        var rez = reservationsMap[w.code + '_' + dateStr];
        if (!rez || !counts[GROUP_ORDER[0]][rez.orderState]) return;
        var g = _getGroupKey(resolveLocationForDate(w.code, dateStr, reservationsMap, workersMap));
        if (!g) return;
        counts[g][rez.orderState][i]++;
      });
    });

    var totalRows = [];
    var subtotalRows = [];  // 「計」行の位置（0始まり）
    function pushGroup(label, getCount) {
      menus.forEach(function(m, mi) {
        var day = dates.map(function(d, i) { return getCount(m.key, i); });
        var sum = day.reduce(function(a, b) { return a + b; }, 0);
        totalRows.push([mi === 0 ? label : '', m.label].concat(day).concat([sum]));
      });
      var dayTotal = dates.map(function(d, i) {
        return menus.reduce(function(a, m) { return a + getCount(m.key, i); }, 0);
      });
      totalRows.push(['', '計'].concat(dayTotal).concat([dayTotal.reduce(function(a, b) { return a + b; }, 0)]));
      subtotalRows.push(totalRows.length - 1);
    }
    GROUP_ORDER.forEach(function(g) {
      pushGroup(g, function(key, i) { return counts[g][key][i]; });
    });
    var grandStart = totalRows.length;
    pushGroup('【全合計】', function(key, i) {
      return GROUP_ORDER.reduce(function(a, g) { return a + counts[g][key][i]; }, 0);
    });

    totalSheet.getRange(2, 1, totalRows.length, totalHeader.length).setValues(totalRows);
    totalSheet.getRange(2, 1, totalRows.length, 1).setFontWeight('bold');
    totalSheet.getRange(2 + grandStart, 1, totalRows.length - grandStart, totalHeader.length).setBackground('#fff2e0');
    subtotalRows.forEach(function(r) {
      totalSheet.getRange(2 + r, 1, 1, totalHeader.length).setFontWeight('bold').setBackground('#fff8ea');
    });
    totalSheet.autoResizeColumns(1, 2);

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

  /** 空の集計（1拠点分） */
  function emptyGroup() {
    var g = { total: 0, names: [], namesBy: {} };
    MENU_DEFS.forEach(function(m) { g[m.key] = 0; g.namesBy[m.key] = []; });
    return g;
  }

  /**
   * 指定日のメール用集計データを生成
   * byGroup[拠点] = { bento, okazu, men, don, toku, total, names:['氏名（弁当）',...], namesBy:{bento:[氏名],...} }
   */
  function generateReportSummary(dateStr) {
    if (!dateStr) dateStr = formatDateYmd(new Date());
    var workers = _uniqueWorkersByCode(SheetService.getWorkers());
    var workersMap = {};
    workers.forEach(function(w) { workersMap[w.code] = w; });
    var reservations = SheetService.getReservationsInRange(dateStr, dateStr);

    var byGroup = {};
    GROUP_ORDER.forEach(function(g) { byGroup[g] = emptyGroup(); });

    reservations.forEach(function(r) {
      if (!r.orderState) return;
      var def = getMenuDef(r.orderState);
      if (!def) return;
      var w = workersMap[r.workerCode];
      if (!w) return;
      var group = _getGroupKey(r.location || w.location);
      if (!group) return;
      var b = byGroup[group];
      b[def.key]++;
      b.total++;
      b.names.push(w.name + '（' + def.label + '）');
      b.namesBy[def.key].push(w.name);
    });

    var grand = { total: 0 };
    MENU_DEFS.forEach(function(m) { grand[m.key] = 0; });
    GROUP_ORDER.forEach(function(g) {
      MENU_DEFS.forEach(function(m) { grand[m.key] += byGroup[g][m.key]; });
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
    emptyGroup: emptyGroup,
    GROUP_ORDER: GROUP_ORDER
  };
})();

/**
 * トリガーから呼ばれるラッパー
 */
function updateAdminSheetsForToday() {
  return ReportService.updateAdminSheetsForToday();
}
