/**
 * お弁当予約アプリ - 設定定数
 * 全てのモジュールから参照される定数を一元管理
 */

// ===== スプレッドシートID =====
// アプリ本体のスプレッドシート（予約データ + 管理シート + マスタキャッシュ）
// GAS拡張として紐づくスプレッドシート
var SOURCE_SPREADSHEET_ID = '1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ';

// 外部マスタシート所在スプレッドシート（作業員マスタ・社内カレンダーマスタのソース）
// 現状は同一スプレッドシート内だが、将来別SSに分離する場合のため定数分離
var EXTERNAL_MASTER_SPREADSHEET_ID = '1iu5HoaknlW1W1HheeYv0jqcRq-aY0SyEE2seQd2pHkQ';

// 総務管理用スプレッドシート保存先フォルダ
var ADMIN_OUTPUT_FOLDER_ID = '1vA0ByiqXNiqufXx_fwSKdoEFB5yJox8B';

// ===== 外部マスタシート（読み取り専用） =====
var MASTER_WORKER_GID = 684189184;      // 作業員マスタ
var MASTER_CALENDAR_GID = 757513238;    // 社内カレンダーマスタ

// ===== アプリ内部シート名 =====
var SHEET_WORKERS = 'workers_cache';           // 作業員マスタのキャッシュ
var SHEET_CALENDAR = 'calendar_cache';         // カレンダーマスタのキャッシュ
var SHEET_RESERVATIONS = 'reservations';       // 予約データ
var SHEET_EMAIL_RECIPIENTS = 'mail_recipients'; // メール送信先管理
var SHEET_NOTICES = 'notices';                 // お知らせ管理（表示内容・期限）
var SHEET_OPERATION_LOG = 'operation_log';     // 操作ログ
var SHEET_CONFIG = 'app_config';               // アプリ設定

// ===== 予約データシート 列構成 =====
var RESERVATION_COL = {
  WORKER_CODE: 1,    // A: 作業員コード
  WORKER_NAME: 2,    // B: 氏名
  DATE: 3,           // C: 日付 (YYYY-MM-DD)
  ORDER_STATE: 4,    // D: 注文状態 ('', 'bento', 'okazu')
  LOCATION: 5,       // E: 配達拠点 ('新工場' | '本社工場')
  UPDATED_AT: 6,     // F: 最終更新日時
  UPDATED_BY: 7      // G: 更新者（Session.getActiveUser().getEmail()）
};
var RESERVATION_HEADERS = ['作業員コード', '氏名', '日付', '注文状態', '配達拠点', '更新日時', '更新者'];

// ===== 作業員マスタ 列構成 =====
var WORKER_COL = {
  CODE: 1,          // A: 作業員コード
  NAME: 2,          // B: 氏名
  DEPT: 3,          // C: 部署
  TASK: 4,          // D: 担当業務（アプリでは未使用）
  LOCATION: 5,      // E: 拠点 ('新工場' | '本社工場')
  STAFF_TYPE: 6     // F: スタッフ種類 ('事務所' | '工場')
};
var WORKER_HEADERS = ['作業員コード', '氏名', '部署', '担当業務', '拠点', 'スタッフ種類'];

// ===== カレンダーマスタ 列構成 =====
var CALENDAR_COL = {
  DATE: 1,       // A: 日付
  TYPE: 2,       // B: 区分 ('休日' | '出勤土曜' | '祝日')
  WEEKDAY: 3,    // C: 曜日
  NOTE: 4        // D: 備考
};
var CALENDAR_HEADERS = ['日付', '区分', '曜日', '備考'];

// ===== カレンダー区分 =====
var CALENDAR_TYPE = {
  HOLIDAY: '休日',
  WORK_SAT: '出勤土曜',
  NATIONAL: '祝日'
};

// ===== 注文状態 =====
var ORDER_STATE = {
  NONE: '',
  BENTO: 'bento',      // 通常弁当 ○
  OKAZU: 'okazu'       // おかずのみ お
};

// ===== 拠点 =====
var LOCATION = {
  SHIN: '新工場',
  HONSHA: '本社工場'
};

// ===== スタッフ種類 =====
var STAFF_TYPE = {
  OFFICE: '事務所',
  FACTORY: '工場'
};

// ===== 業務ルール =====
var RULES = {
  DEADLINE_HOUR: 9,              // 当日締切 9:00
  DEADLINE_MINUTE: 0,
  SATURDAY_DEADLINE_HOUR: 15,    // 出勤土曜の締切（金曜）15:00
  SATURDAY_DEADLINE_MINUTE: 0,
  RESERVATION_WINDOW_DAYS: 14,   // 予約可能期間: 本日から2週間
  MONTH_START_DAY: 16            // 月度開始日: 16日
};

// ===== ポータルURL（社内ポータルGAS Webアプリ） =====
var PORTAL_URL = 'https://script.google.com/a/macros/lineworks-local.info/s/AKfycbx2eyJMOYP9o--GPBuhY-pj071IIR6Kqb_0xALwwNzdLQZux0dIAlL3P9EoCucnzXA/exec';

// ===== アプリ情報 =====
var APP_TITLE = 'お弁当予約';

// ===== お知らせ初期値 =====
var DEFAULT_NOTICE = '【価格改定のお知らせ】2025/10/16注文分より、お弁当の価格が変わりました\n（従来：¥400 → 新価格：¥450 ※麺・丼も同価格）';
var DEFAULT_WARNING = '※ 丼・麺類・幕ノ内等は、このアプリケーションからはご注文いただけません。\nご希望の場合は、前日の15時までに総務までご連絡ください。';

// ===== ヘルパー =====

/**
 * ソーススプレッドシートを開く
 */
function getSourceSpreadsheet() {
  return SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
}

/**
 * 指定名のシートを取得（存在しなければ作成）
 */
function getOrCreateSheet(sheetName, headers) {
  var ss = getSourceSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0e8d8');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

/**
 * 今の日本時間
 */
function nowJst() {
  return new Date();
}

/**
 * YYYY-MM-DD形式にフォーマット
 */
function formatDateYmd(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
}

/**
 * YYYY-MM-DD HH:mm:ss形式にフォーマット
 */
function formatDateTime(date) {
  if (!date) return '';
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}
