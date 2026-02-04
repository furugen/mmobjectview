/**
 * Config.gs - 設定管理
 *
 * アプリケーション全体の定数定義、環境切替、エラーコード、
 * 日本語メッセージマッピング、データ型変換テーブルを管理する。
 *
 * 設計方針:
 * - Script Properties で外部から変更可能な設定値はgetXxx()関数で取得
 * - 変更不要な定数は var で直接定義
 * - 全ての設定値をこのファイルに集約し、他モジュールからの参照先を明確化
 */

// ============================================================
// アプリケーション基本設定
// ============================================================

/**
 * アプリケーション定数
 */
var APP_CONFIG = {
  /** アドオン名 */
  APP_NAME: 'Salesforce Object Explorer',

  /** アドオンバージョン */
  VERSION: '1.0.0',

  /** サイドバーのタイトル */
  SIDEBAR_TITLE: 'Salesforce オブジェクト展開',

  /** サイドバーの幅（ピクセル）- GASデフォルト300px */
  SIDEBAR_WIDTH: 300,

  /** OAuth2 サービス名（ライブラリ内部で使用） */
  OAUTH_SERVICE_NAME: 'salesforce',

  /** OAuth2 コールバック関数名 */
  OAUTH_CALLBACK_FUNCTION: 'authCallback'
};

// ============================================================
// Salesforce 環境設定
// ============================================================

/**
 * Script Properties のキー定義
 */
var PROP_KEYS = {
  /** Salesforce クライアントID（Connected App の Consumer Key） */
  SF_CLIENT_ID: 'SF_CLIENT_ID',

  /** Salesforce クライアントシークレット（Connected App の Consumer Secret） */
  SF_CLIENT_SECRET: 'SF_CLIENT_SECRET',

  /** Salesforce ログインURL（本番 / Sandbox / カスタムドメイン） */
  SF_LOGIN_URL: 'SF_LOGIN_URL',

  /** Salesforce REST API バージョン */
  SF_API_VERSION: 'SF_API_VERSION'
};

/**
 * デフォルト値
 */
var SF_DEFAULTS = {
  /** デフォルトのログインURL（本番環境） */
  LOGIN_URL: 'https://login.salesforce.com',

  /** デフォルトのAPIバージョン（BOSS判断: v62.0 安定版採用） */
  API_VERSION: 'v62.0',

  /** Salesforce OAuth スコープ */
  OAUTH_SCOPE: 'api refresh_token'
};

/**
 * Salesforce 環境定義
 */
var SF_ENVIRONMENTS = {
  production: {
    label: '本番環境',
    loginUrl: 'https://login.salesforce.com'
  },
  sandbox: {
    label: 'Sandbox',
    loginUrl: 'https://test.salesforce.com'
  }
};

/**
 * Salesforce ログインURLを取得する
 * Script Properties に設定があればそれを使用、なければデフォルト（本番）
 * @return {string} ログインURL
 */
function getLoginUrl_() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty(PROP_KEYS.SF_LOGIN_URL) || SF_DEFAULTS.LOGIN_URL;
}

/**
 * Salesforce API バージョンを取得する
 * Script Properties に設定があればそれを使用、なければデフォルト（v62.0）
 * @return {string} API バージョン（例: "v62.0"）
 */
function getApiVersion_() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty(PROP_KEYS.SF_API_VERSION) || SF_DEFAULTS.API_VERSION;
}

/**
 * Salesforce REST API のベースパスを取得する
 * @return {string} ベースパス（例: "/services/data/v62.0"）
 */
function getApiBasePath_() {
  return '/services/data/' + getApiVersion_();
}

// ============================================================
// スプレッドシート書式設定
// ============================================================

/**
 * スプレッドシート展開時の書式設定
 */
var SHEET_CONFIG = {
  /** ヘッダー行の背景色（Google Blue） */
  HEADER_BG_COLOR: '#4285F4',

  /** ヘッダー行のフォント色（白） */
  HEADER_FONT_COLOR: '#FFFFFF',

  /** ヘッダーのフォントサイズ */
  HEADER_FONT_SIZE: 10,

  /** データ行のフォントサイズ */
  DATA_FONT_SIZE: 10,

  /** 奇数データ行の背景色（ゼブラストライプ） */
  ALT_ROW_COLOR: '#F8F9FA',

  /** ヘッダーの固定行数 */
  FROZEN_ROWS: 1,

  /** シート名 */
  SHEET_NAME: 'フィールド一覧'
};

/**
 * スプレッドシートに展開する24項目のカラム定義
 * SF API Specialist の仕様（docs/SALESFORCE_API_SPEC.md セクション2.3）を正とする
 *
 * 各要素:
 *   header: ヘッダー表示名（日本語）
 *   key: Describe API レスポンスの fields[].xxx に対応するプロパティ名
 *   transform: 値の変換方法（"boolean" → ○/−, "type" → 日本語型名, "array" → カンマ区切り, "picklistValues" → 特殊処理）
 */
var COLUMN_DEFINITIONS = [
  { header: 'API名',           key: 'name',               transform: null },
  { header: 'ラベル',           key: 'label',              transform: null },
  { header: 'データ型',         key: 'type',               transform: 'type' },
  { header: '桁数/長さ',        key: 'length',             transform: 'lengthOrPrecision' },
  { header: '必須',             key: null,                 transform: 'required' },
  { header: '一意',             key: 'unique',             transform: 'boolean' },
  { header: 'カスタム',          key: 'custom',             transform: 'boolean' },
  { header: '作成可',           key: 'createable',         transform: 'boolean' },
  { header: '更新可',           key: 'updateable',         transform: 'boolean' },
  { header: '数式',             key: 'calculated',         transform: 'boolean' },
  { header: '数式内容',          key: 'calculatedFormula',   transform: null },
  { header: '自動採番',          key: 'autoNumber',         transform: 'boolean' },
  { header: '外部ID',           key: 'externalId',         transform: 'boolean' },
  { header: '暗号化',           key: 'encrypted',          transform: 'boolean' },
  { header: '参照先',           key: 'referenceTo',        transform: 'array' },
  { header: 'リレーション名',    key: 'relationshipName',   transform: null },
  { header: '選択リスト値',      key: 'picklistValues',     transform: 'picklistValues' },
  { header: '連動選択リスト',    key: 'dependentPicklist',  transform: 'boolean' },
  { header: '制御項目',          key: 'controllerName',     transform: null },
  { header: 'デフォルト値',      key: 'defaultValue',       transform: 'defaultValue' },
  { header: 'ヘルプテキスト',    key: 'inlineHelpText',     transform: null },
  { header: 'Name項目',         key: 'nameField',          transform: 'boolean' },
  { header: 'フィルタ可',        key: 'filterable',         transform: 'boolean' },
  { header: '制限付き選択リスト', key: 'restrictedPicklist', transform: 'boolean' }
];

// ============================================================
// エラーコード・メッセージ
// ============================================================

/**
 * エラーコード定義
 */
var ERROR_CODES = {
  // 認証系
  AUTH_REQUIRED: 'AUTH_REQUIRED',
  AUTH_EXPIRED: 'AUTH_EXPIRED',
  AUTH_FAILED: 'AUTH_FAILED',

  // Salesforce API 系
  API_ERROR: 'API_ERROR',
  API_RATE_LIMIT: 'API_RATE_LIMIT',
  API_NOT_FOUND: 'API_NOT_FOUND',
  API_FORBIDDEN: 'API_FORBIDDEN',

  // スクリプト系
  SHEET_ERROR: 'SHEET_ERROR',
  INVALID_INPUT: 'INVALID_INPUT',
  UNKNOWN_ERROR: 'UNKNOWN_ERROR'
};

/**
 * エラーメッセージ定義（日本語）
 * ERROR_CODES と 1:1 対応
 */
var ERROR_MESSAGES = {
  AUTH_REQUIRED: 'Salesforceに接続されていません。「Salesforceに接続」ボタンをクリックしてください。',
  AUTH_EXPIRED: '認証の有効期限が切れました。再度接続してください。',
  AUTH_FAILED: '認証に失敗しました。接続設定を確認してください。',
  API_ERROR: 'Salesforce APIの呼び出しに失敗しました。',
  API_RATE_LIMIT: 'APIの利用制限に達しました。しばらく待ってから再試行してください。',
  API_NOT_FOUND: '指定されたオブジェクトが見つかりませんでした。',
  API_FORBIDDEN: 'このリソースへのアクセス権限がありません。Salesforce管理者に確認してください。',
  SHEET_ERROR: 'スプレッドシートの作成に失敗しました。',
  INVALID_INPUT: '入力値が不正です。',
  UNKNOWN_ERROR: '予期しないエラーが発生しました。しばらく待ってから再試行してください。'
};

// ============================================================
// データ型 日本語変換テーブル
// ============================================================

/**
 * Salesforce フィールド型 → 日本語表記の変換テーブル
 * Describe API の fields[].type の値をキーとする
 */
var FIELD_TYPE_LABELS = {
  'id': 'ID',
  'string': 'テキスト',
  'textarea': 'テキストエリア',
  'boolean': 'チェックボックス',
  'int': '整数',
  'double': '数値（小数）',
  'currency': '通貨',
  'percent': 'パーセント',
  'date': '日付',
  'datetime': '日時',
  'time': '時刻',
  'email': 'メール',
  'phone': '電話',
  'url': 'URL',
  'picklist': '選択リスト',
  'multipicklist': '複数選択リスト',
  'reference': '参照関係',
  'masterrecord': 'マスタ詳細',
  'location': '地理位置情報',
  'address': '住所',
  'encryptedstring': '暗号化テキスト',
  'base64': 'Base64',
  'combobox': 'コンボボックス',
  'anyType': '任意型',
  'complexvalue': '複合値'
};

/**
 * Salesforce のフィールド型名を日本語に変換する
 * @param {string} sfType - Salesforce のフィールド型名（例: "string", "double"）
 * @return {string} 日本語の型名（変換テーブルにない場合は原文をそのまま返す）
 */
function translateFieldType(sfType) {
  if (!sfType) return '';
  var lower = sfType.toLowerCase();
  return FIELD_TYPE_LABELS[lower] || sfType;
}

// ============================================================
// ユーティリティ関数（レスポンス生成）
// ============================================================

/**
 * 成功レスポンスを生成する
 * @param {Object} data - レスポンスデータ
 * @param {string} [message] - 成功メッセージ（省略可）
 * @return {Object} 統一レスポンス形式
 */
function createSuccessResponse_(data, message) {
  return {
    success: true,
    data: data || {},
    message: message || ''
  };
}

/**
 * エラーレスポンスを生成する
 * @param {string} code - ERROR_CODES で定義されたエラーコード
 * @param {string} [details] - 技術的詳細（ログ用）
 * @return {Object} 統一レスポンス形式
 */
function createErrorResponse_(code, details) {
  var message = ERROR_MESSAGES[code] || ERROR_MESSAGES.UNKNOWN_ERROR;
  return {
    success: false,
    error: {
      code: code,
      message: message,
      details: details || ''
    }
  };
}
