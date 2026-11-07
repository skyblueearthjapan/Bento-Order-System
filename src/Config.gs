/**
 * お弁当予約アプリ - 設定定数
 * 全てのモジュールから参照される定数を一元管理
 */

// ===== スプレッドシートID =====
// アプリ本体のスプレッドシート（予約データ + 管理シート + マスタキャッシュ）
// firstTimeSetup() 実行時に全シートが自動生成される
var SOURCE_SPREADSHEET_ID = '1L5JhsWYcmGMs3-tTw1Y-TC_yn1CNybw4SKgg_HaB2ow';

// 外部マスタシート所在スプレッドシート（作業員マスタ・社内カレンダーマスタのソース、読み取り専用）
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
var SHEET_MENU_PERIODS = 'menu_periods';       // 麺・丼の受付期間（総務設定画面から編集）
var SHEET_TOKUBEN = 'tokuben';                 // 特弁（1日限定）の設定（総務設定画面から編集）
var SHEET_ADMIN_USERS = 'admin_users';         // 総務設定を開ける人
var SHEET_ADMIN_LOG = 'admin_log';             // 総務設定の変更履歴

var MENU_PERIOD_HEADERS = ['ID', 'メニュー', '期間の名前', '開始日', '終了日', '有効フラグ', 'チラシファイルID', 'チラシファイル名', '更新日時', '更新者'];
var TOKUBEN_HEADERS = ['ID', '日付', '名前', '価格', '有効フラグ', 'チラシファイルID', 'チラシファイル名', '更新日時', '更新者'];
var ADMIN_USER_HEADERS = ['メールアドレス', '名前', '有効フラグ'];
var ADMIN_LOG_HEADERS = ['日時', 'ユーザー', '内容'];

// 「設定できる人」シートを新しく作るときの初期登録
var INITIAL_ADMIN_USERS = [
  ['imaizumi@lineworks-local.info', '今泉', 'ON'],
  ['admin-01@lineworks-local.info', '管理者', 'ON']
];

// チラシPDFの保存先（ADMIN_OUTPUT_FOLDER_ID の中に自動で作るフォルダ名）
var FLYER_FOLDER_NAME = 'お弁当チラシ';
var FLYER_MAX_BYTES = 10 * 1024 * 1024;   // 1ファイル10MBまで（画面から送れる大きさに余裕を持たせる）

// ===== 予約データシート 列構成 =====
var RESERVATION_COL = {
  WORKER_CODE: 1,    // A: 作業員コード
  WORKER_NAME: 2,    // B: 氏名
  DEPT: 3,           // C: 部署
  STAFF_TYPE: 4,     // D: スタッフ種類 (事務所/工場)
  DATE: 5,           // E: 日付 (YYYY-MM-DD)
  ORDER_STATE: 6,    // F: 注文状態 ('', 'bento', 'okazu')
  LOCATION: 7,       // G: 配達拠点 ('新工場' | '本社工場')
  UPDATED_AT: 8,     // H: 最終更新日時
  UPDATED_BY: 9      // I: 更新者（Session.getActiveUser().getEmail()）
};
var RESERVATION_HEADERS = ['作業員コード', '氏名', '部署', 'スタッフ種類', '日付', '注文状態', '配達拠点', '更新日時', '更新者'];

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
  OKAZU: 'okazu',      // おかずのみ お
  MEN: 'men',          // 麺（受付期間のみ） 麺
  DON: 'don',          // 丼（受付期間のみ） 丼
  TOKU: 'toku'         // 特弁（1日限定） 特
};

// ===== メニュー定義（表示名・Excel記号・価格・受付の種類） =====
// kind: 'always' = いつでも / 'period' = 総務が決めた受付期間のみ / 'day' = 特弁の日のみ
// 麺・丼・特弁（always 以外）の締切は「前の平日の朝8:00」
var MENU_DEFS = [
  { key: 'bento', label: '弁当',       sym: '○',  price: 450, kind: 'always' },
  { key: 'okazu', label: 'おかずのみ', sym: 'お', price: 350, kind: 'always' },
  { key: 'men',   label: '麺',         sym: '麺', price: 450, kind: 'period' },
  { key: 'don',   label: '丼',         sym: '丼', price: 450, kind: 'period' },
  { key: 'toku',  label: '特弁',       sym: '特', price: 0,   kind: 'day' }
];
var MENU_KEYS = MENU_DEFS.map(function(m) { return m.key; });
var EARLY_MENU_KEYS = ['men', 'don', 'toku'];

function getMenuDef(key) {
  for (var i = 0; i < MENU_DEFS.length; i++) {
    if (MENU_DEFS[i].key === key) return MENU_DEFS[i];
  }
  return null;
}

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
  DEADLINE_HOUR: 8,              // 当日締切 8:00
  DEADLINE_MINUTE: 0,
  SATURDAY_DEADLINE_HOUR: 15,    // 出勤土曜の締切（前日金曜）15:00
  SATURDAY_DEADLINE_MINUTE: 0,
  EARLY_DEADLINE_HOUR: 8,        // 麺・丼・特弁の締切: 前の平日の8:00
  RESERVATION_WINDOW_DAYS: 14,   // 予約可能期間: 本日から2週間
  MONTH_START_DAY: 16            // 月度開始日: 16日
};

// ===== ポータルURL（社内ポータルGAS Webアプリ） =====
var PORTAL_URL = 'https://script.google.com/a/macros/lineworks-local.info/s/AKfycbx2eyJMOYP9o--GPBuhY-pj071IIR6Kqb_0xALwwNzdLQZux0dIAlL3P9EoCucnzXA/exec';

// ===== アプリ情報 =====
var APP_TITLE = 'お弁当予約';

// ===== お知らせ初期値 =====
var DEFAULT_NOTICE = '【価格改定のお知らせ】2025/10/16注文分より、お弁当の価格が変わりました\n（従来：¥400 → 新価格：¥450 ※麺・丼も同価格）';
var DEFAULT_WARNING = '※ 幕ノ内等は、このアプリケーションからはご注文いただけません。\nご希望の場合は、前日の15時までに総務までご連絡ください。';

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
