/**
 * SalesforceApi.gs
 * Salesforce REST API 呼び出しモジュール
 *
 * 責務:
 * - Salesforce REST API への共通リクエスト基盤（認証ヘッダー付与、エラーハンドリング）
 * - Describe Global（オブジェクト一覧取得）
 * - sObject Describe（オブジェクト詳細・フィールド情報取得）
 * - フィールド情報のスプレッドシート展開用整形（24カラム）
 *
 * 依存モジュール:
 * - Auth.gs: getAccessToken_(), getInstanceUrl_(), isAuthorized_(), getSalesforceService_()
 * - Config.gs: APP_CONFIG, ERROR_CODES, ERROR_MESSAGES, PROP_KEYS, FIELD_TYPE_LABELS
 *
 * @author Salesforce API Specialist
 * @version 1.0.0
 * @since 2026-01-30
 */

// ============================================================
// 定数
// ============================================================

/**
 * オブジェクト一覧キャッシュのキー
 * @const {string}
 */
var CACHE_KEY_OBJECTS = 'sf_objects_list';

/**
 * キャッシュ有効期間（秒）: 6時間
 * @const {number}
 */
var CACHE_TTL_SECONDS = 21600;

/**
 * サーバーエラー時の最大リトライ回数
 * @const {number}
 */
var MAX_RETRIES = 3;

/**
 * リトライ時の初期待機時間（ミリ秒）
 * @const {number}
 */
var RETRY_BASE_DELAY_MS = 1000;

/**
 * データ型の日本語ラベルマッピング
 * Config.gs に FIELD_TYPE_LABELS が定義されていない場合のフォールバック
 * @const {Object<string, string>}
 */
var DATA_TYPE_LABELS_FALLBACK = {
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
  'anyType': '任意の型'
};

// ============================================================
// API基盤（共通リクエスト関数）
// ============================================================

/**
 * Salesforce REST API にリクエストを送信する（共通基盤）
 *
 * 機能:
 * - Auth.gs からアクセストークン・インスタンスURLを取得してリクエスト
 * - 401エラー時: トークン自動リフレッシュ → 1回リトライ
 * - 500/503エラー時: 指数バックオフで最大3回リトライ
 * - その他エラー: 統一エラー形式でスロー
 *
 * @param {string} endpoint - APIエンドポイントパス（例: "/services/data/v62.0/sobjects/"）
 * @param {Object} [options] - 追加オプション
 * @param {string} [options.method='get'] - HTTPメソッド
 * @param {Object} [options.payload] - リクエストボディ（POSTの場合）
 * @return {Object} パース済みJSONレスポンス
 * @throws {Object} 統一エラーオブジェクト { code, message, details }
 * @private
 */
function callSalesforceApi_(endpoint, options) {
  // 認証チェック
  if (!isAuthorized_()) {
    throw {
      code: ERROR_CODES.AUTH_REQUIRED,
      message: ERROR_MESSAGES.AUTH_REQUIRED,
      details: 'Not authenticated'
    };
  }

  options = options || {};
  var method = options.method || 'get';
  var payload = options.payload || null;

  var instanceUrl = getInstanceUrl_();
  var url = instanceUrl + endpoint;

  var fetchOptions = {
    method: method,
    headers: {
      'Authorization': 'Bearer ' + getAccessToken_(),
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (payload) {
    fetchOptions.payload = JSON.stringify(payload);
  }

  // 初回リクエスト
  var response = UrlFetchApp.fetch(url, fetchOptions);
  var statusCode = response.getResponseCode();

  // ── 401: セッション期限切れ → リフレッシュして1回リトライ ──
  if (statusCode === 401) {
    console.info('[SalesforceApi] 401検出: トークンリフレッシュを試行');
    try {
      var service = getSalesforceService_();
      service.refresh();

      // リフレッシュ後のトークンとinstance_urlで再試行
      fetchOptions.headers['Authorization'] = 'Bearer ' + getAccessToken_();
      var newInstanceUrl = getInstanceUrl_();
      url = newInstanceUrl + endpoint;

      response = UrlFetchApp.fetch(url, fetchOptions);
      statusCode = response.getResponseCode();
    } catch (refreshError) {
      console.error('[SalesforceApi] トークンリフレッシュ失敗:', refreshError);
      throw {
        code: ERROR_CODES.AUTH_EXPIRED,
        message: ERROR_MESSAGES.AUTH_EXPIRED,
        details: refreshError.message || 'Token refresh failed'
      };
    }
  }

  // ── 500/503: サーバーエラー → 指数バックオフでリトライ ──
  if (statusCode === 500 || statusCode === 503) {
    for (var retry = 1; retry <= MAX_RETRIES; retry++) {
      var delay = RETRY_BASE_DELAY_MS * Math.pow(2, retry - 1); // 1s, 2s, 4s
      console.warn('[SalesforceApi] サーバーエラー(' + statusCode + '): ' + delay + 'ms後にリトライ (' + retry + '/' + MAX_RETRIES + ')');
      Utilities.sleep(delay);

      response = UrlFetchApp.fetch(url, fetchOptions);
      statusCode = response.getResponseCode();

      if (statusCode !== 500 && statusCode !== 503) {
        break;
      }
    }
  }

  // ── 成功レスポンス ──
  if (statusCode >= 200 && statusCode < 300) {
    var responseText = response.getContentText();
    if (!responseText || responseText.trim() === '') {
      return {};
    }
    return JSON.parse(responseText);
  }

  // ── エラーレスポンス処理 ──
  handleSalesforceError_(statusCode, response.getContentText());
}

/**
 * Salesforce APIエラーレスポンスを解析し、統一エラー形式でスローする
 *
 * Salesforce REST API のエラーレスポンス形式:
 * - 標準エラー（400,403,404,500）: 配列 [{ message, errorCode, fields }]
 * - 認証エラー（401）: オブジェクト { error, error_description }
 *
 * @param {number} statusCode - HTTPステータスコード
 * @param {string} responseBody - レスポンスボディ（JSON文字列）
 * @throws {Object} 統一エラーオブジェクト { code, message, details }
 * @private
 */
function handleSalesforceError_(statusCode, responseBody) {
  var sfErrorCode = '';
  var sfMessage = '';

  try {
    var parsed = JSON.parse(responseBody);

    // 配列形式（標準APIエラー）
    if (Array.isArray(parsed) && parsed.length > 0) {
      sfErrorCode = parsed[0].errorCode || '';
      sfMessage = parsed[0].message || '';
    }
    // オブジェクト形式（OAuth/認証エラー）
    else if (parsed.error) {
      sfErrorCode = parsed.error;
      sfMessage = parsed.error_description || parsed.error;
    }
  } catch (e) {
    sfMessage = responseBody;
  }

  // ステータスコードごとにエラーコードをマッピング
  var errorCode;
  var errorMessage;

  switch (statusCode) {
    case 401:
      errorCode = ERROR_CODES.AUTH_EXPIRED;
      errorMessage = ERROR_MESSAGES.AUTH_EXPIRED;
      break;

    case 403:
      if (sfErrorCode === 'REQUEST_LIMIT_EXCEEDED') {
        errorCode = ERROR_CODES.API_RATE_LIMIT;
        errorMessage = ERROR_MESSAGES.API_RATE_LIMIT;
      } else {
        errorCode = ERROR_CODES.API_ERROR;
        errorMessage = 'アクセス権限が不足しています。Salesforce管理者に連絡してください。';
      }
      break;

    case 404:
      errorCode = ERROR_CODES.API_NOT_FOUND;
      errorMessage = ERROR_MESSAGES.API_NOT_FOUND;
      break;

    case 400:
      errorCode = ERROR_CODES.API_ERROR;
      errorMessage = sfMessage || ERROR_MESSAGES.API_ERROR;
      break;

    default:
      errorCode = ERROR_CODES.API_ERROR;
      errorMessage = ERROR_MESSAGES.API_ERROR;
      break;
  }

  console.error('[SalesforceApi] APIエラー: HTTP ' + statusCode + ' / ' + sfErrorCode + ' / ' + sfMessage);

  throw {
    code: errorCode,
    message: errorMessage,
    details: 'HTTP ' + statusCode + ': ' + sfErrorCode + ' - ' + sfMessage
  };
}

// ============================================================
// オブジェクト一覧取得（Describe Global）
// ============================================================

/**
 * Salesforceオブジェクト一覧を取得する（Describe Global）
 *
 * - キャッシュ対応（6時間TTL）
 * - デフォルトで deprecatedAndHidden を除外
 * - フィルタリングオプション対応
 *
 * @param {Object} [filterOptions] - フィルタリングオプション
 * @param {boolean} [filterOptions.showCustomOnly=false] - カスタムオブジェクトのみ表示
 * @param {boolean} [filterOptions.showStandardOnly=false] - 標準オブジェクトのみ表示
 * @param {boolean} [filterOptions.showQueryableOnly=false] - クエリ可能なオブジェクトのみ
 * @param {string} [filterOptions.searchText=''] - 検索テキスト（name または label に部分一致）
 * @param {boolean} [filterOptions.forceRefresh=false] - キャッシュを無視して再取得
 * @return {Array<Object>} オブジェクト情報の配列
 *   各要素: { name, label, labelPlural, keyPrefix, custom, queryable, searchable, createable, updateable }
 */
function getObjectList(filterOptions) {
  filterOptions = filterOptions || {};

  // ── キャッシュ確認 ──
  var allObjects;
  if (!filterOptions.forceRefresh) {
    allObjects = getObjectListFromCache_();
  }

  // ── APIから取得 ──
  if (!allObjects) {
    var apiVersion = getApiVersion_();
    var result = callSalesforceApi_('/services/data/' + apiVersion + '/sobjects/');

    // 必要なプロパティだけ抽出し、deprecatedAndHidden を除外
    allObjects = result.sobjects
      .filter(function(obj) {
        return !obj.deprecatedAndHidden;
      })
      .map(function(obj) {
        return {
          name: obj.name,
          label: obj.label,
          labelPlural: obj.labelPlural,
          keyPrefix: obj.keyPrefix || '',
          custom: obj.custom,
          queryable: obj.queryable,
          searchable: obj.searchable,
          createable: obj.createable,
          updateable: obj.updateable
        };
      });

    // ラベルでソート
    allObjects.sort(function(a, b) {
      return (a.label || '').localeCompare(b.label || '');
    });

    // キャッシュに保存
    putObjectListToCache_(allObjects);
  }

  // ── フィルタリング適用 ──
  var filtered = allObjects;

  if (filterOptions.showCustomOnly) {
    filtered = filtered.filter(function(obj) { return obj.custom === true; });
  } else if (filterOptions.showStandardOnly) {
    filtered = filtered.filter(function(obj) { return obj.custom === false; });
  }

  if (filterOptions.showQueryableOnly) {
    filtered = filtered.filter(function(obj) { return obj.queryable === true; });
  }

  if (filterOptions.searchText) {
    var search = filterOptions.searchText.toLowerCase();
    filtered = filtered.filter(function(obj) {
      return (obj.name && obj.name.toLowerCase().indexOf(search) !== -1) ||
             (obj.label && obj.label.toLowerCase().indexOf(search) !== -1);
    });
  }

  return filtered;
}

/**
 * キャッシュからオブジェクト一覧を取得
 * @return {Array<Object>|null} キャッシュされたオブジェクト配列、またはnull
 * @private
 */
function getObjectListFromCache_() {
  try {
    var cache = CacheService.getUserCache();
    var cached = cache.get(CACHE_KEY_OBJECTS);
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {
    console.warn('[SalesforceApi] キャッシュ読み込みエラー:', e.message);
  }
  return null;
}

/**
 * オブジェクト一覧をキャッシュに保存
 *
 * CacheService の上限は 100KB/エントリ。
 * オブジェクト一覧が大きい場合はキャッシュをスキップする。
 *
 * @param {Array<Object>} objects - オブジェクト情報の配列
 * @private
 */
function putObjectListToCache_(objects) {
  try {
    var json = JSON.stringify(objects);
    // CacheService の 100KB 上限チェック
    if (json.length < 100000) {
      var cache = CacheService.getUserCache();
      cache.put(CACHE_KEY_OBJECTS, json, CACHE_TTL_SECONDS);
    } else {
      console.warn('[SalesforceApi] オブジェクト一覧が大きすぎるためキャッシュをスキップ (' + json.length + ' bytes)');
    }
  } catch (e) {
    console.warn('[SalesforceApi] キャッシュ保存エラー:', e.message);
  }
}

/**
 * オブジェクト一覧キャッシュをクリアする
 * 環境切替時やログアウト時に呼び出す
 */
function clearObjectListCache() {
  try {
    var cache = CacheService.getUserCache();
    cache.remove(CACHE_KEY_OBJECTS);
  } catch (e) {
    console.warn('[SalesforceApi] キャッシュクリアエラー:', e.message);
  }
}

// ============================================================
// オブジェクトDescribe（フィールド情報取得）
// ============================================================

/**
 * 指定オブジェクトの詳細情報（フィールド一覧）を取得する（sObject Describe）
 *
 * @param {string} objectApiName - オブジェクトのAPI名（例: "Account"）
 * @return {Object} オブジェクト詳細
 *   {
 *     objectInfo: { name, label, labelPlural, keyPrefix, custom, recordTypeInfos },
 *     fields: Array<FieldInfo>
 *   }
 *
 * FieldInfo の構造:
 *   { name, label, type, length, precision, scale, nillable, unique,
 *     createable, updateable, custom, calculated, calculatedFormula,
 *     autoNumber, externalId, encrypted, filterable, nameField,
 *     dependentPicklist, controllerName, restrictedPicklist,
 *     referenceTo, relationshipName, picklistValues, defaultValue,
 *     inlineHelpText }
 */
function describeObject(objectApiName) {
  if (!objectApiName) {
    throw {
      code: ERROR_CODES.INVALID_INPUT,
      message: ERROR_MESSAGES.INVALID_INPUT,
      details: 'objectApiName is required'
    };
  }

  var apiVersion = getApiVersion_();
  var result = callSalesforceApi_(
    '/services/data/' + apiVersion + '/sobjects/' + objectApiName + '/describe'
  );

  // オブジェクト基本情報を抽出
  var objectInfo = {
    name: result.name,
    label: result.label,
    labelPlural: result.labelPlural,
    keyPrefix: result.keyPrefix || '',
    custom: result.custom,
    recordTypeInfos: (result.recordTypeInfos || []).map(function(rt) {
      return {
        name: rt.name,
        recordTypeId: rt.recordTypeId,
        active: rt.active,
        defaultRecordTypeMapping: rt.defaultRecordTypeMapping
      };
    })
  };

  // フィールド情報を整形して抽出
  var fields = (result.fields || []).map(function(f) {
    return {
      name: f.name,
      label: f.label,
      type: f.type,
      length: f.length || 0,
      precision: f.precision || 0,
      scale: f.scale || 0,
      nillable: f.nillable,
      unique: f.unique,
      createable: f.createable,
      updateable: f.updateable,
      custom: f.custom,
      calculated: f.calculated,
      calculatedFormula: f.calculatedFormula || '',
      autoNumber: f.autoNumber,
      externalId: f.externalId,
      encrypted: f.encrypted,
      filterable: f.filterable,
      nameField: f.nameField,
      dependentPicklist: f.dependentPicklist,
      controllerName: f.controllerName || '',
      restrictedPicklist: f.restrictedPicklist,
      referenceTo: f.referenceTo || [],
      relationshipName: f.relationshipName || '',
      picklistValues: (f.picklistValues || []).filter(function(pv) {
        return pv.active === true;
      }).map(function(pv) {
        return {
          label: pv.label,
          value: pv.value,
          defaultValue: pv.defaultValue
        };
      }),
      defaultValue: f.defaultValue,
      inlineHelpText: f.inlineHelpText || ''
    };
  });

  return {
    objectInfo: objectInfo,
    fields: fields
  };
}

// ============================================================
// フィールド情報のスプレッドシート展開用整形
// ============================================================

/**
 * フィールド情報をスプレッドシート展開用の2次元配列に変換する
 *
 * SALESFORCE_API_SPEC.md セクション2.3 の24カラム定義に準拠:
 *  1. API名          2. ラベル          3. データ型
 *  4. 桁数/長さ       5. 必須            6. 一意
 *  7. カスタム        8. 作成可          9. 更新可
 * 10. 数式           11. 数式内容        12. 自動採番
 * 13. 外部ID         14. 暗号化          15. 参照先
 * 16. リレーション名   17. 選択リスト値    18. 連動選択リスト
 * 19. 制御項目        20. デフォルト値    21. ヘルプテキスト
 * 22. Name項目       23. フィルタ可       24. 制限付き選択リスト
 *
 * @param {Array<Object>} fields - describeObject() で取得したフィールド配列
 * @return {Object} 展開用データ
 *   {
 *     headers: Array<string>,  // ヘッダー行（24要素）
 *     rows: Array<Array>       // データ行の2次元配列
 *   }
 */
function formatFieldsForSheet(fields) {
  // ── ヘッダー行（24カラム） ──
  var headers = [
    'API名',                   // 1
    'ラベル',                   // 2
    'データ型',                 // 3
    '桁数/長さ',               // 4
    '必須',                     // 5
    '一意',                     // 6
    'カスタム',                 // 7
    '作成可',                   // 8
    '更新可',                   // 9
    '数式',                     // 10
    '数式内容',                 // 11
    '自動採番',                 // 12
    '外部ID',                   // 13
    '暗号化',                   // 14
    '参照先',                   // 15
    'リレーション名',           // 16
    '選択リスト値',             // 17
    '連動選択リスト',           // 18
    '制御項目',                 // 19
    'デフォルト値',             // 20
    'ヘルプテキスト',           // 21
    'Name項目',                // 22
    'フィルタ可',               // 23
    '制限付き選択リスト'        // 24
  ];

  // ── データ行変換 ──
  var rows = fields.map(function(f) {
    return [
      f.name,                                                    // 1. API名
      f.label,                                                   // 2. ラベル
      translateFieldType_(f.type),                               // 3. データ型（日本語）
      formatLengthPrecision_(f),                                 // 4. 桁数/長さ
      formatBoolean_(isRequired_(f)),                            // 5. 必須
      formatBoolean_(f.unique),                                  // 6. 一意
      formatBoolean_(f.custom),                                  // 7. カスタム
      formatBoolean_(f.createable),                              // 8. 作成可
      formatBoolean_(f.updateable),                              // 9. 更新可
      formatBoolean_(f.calculated),                              // 10. 数式
      f.calculatedFormula || '',                                 // 11. 数式内容
      formatBoolean_(f.autoNumber),                              // 12. 自動採番
      formatBoolean_(f.externalId),                              // 13. 外部ID
      formatBoolean_(f.encrypted),                               // 14. 暗号化
      formatReferenceTo_(f.referenceTo),                         // 15. 参照先
      f.relationshipName || '',                                  // 16. リレーション名
      formatPicklistValues_(f.picklistValues),                   // 17. 選択リスト値
      formatBoolean_(f.dependentPicklist),                       // 18. 連動選択リスト
      f.controllerName || '',                                    // 19. 制御項目
      formatDefaultValue_(f.defaultValue),                       // 20. デフォルト値
      f.inlineHelpText || '',                                    // 21. ヘルプテキスト
      formatBoolean_(f.nameField),                               // 22. Name項目
      formatBoolean_(f.filterable),                              // 23. フィルタ可
      formatBoolean_(f.restrictedPicklist)                       // 24. 制限付き選択リスト
    ];
  });

  return {
    headers: headers,
    rows: rows
  };
}

// ============================================================
// ユーザー情報取得
// ============================================================

/**
 * Salesforce ユーザー情報を取得する
 * 接続確認・ユーザー名表示に使用
 *
 * @return {Object} { id, name, email, organizationId }
 */
function getCurrentUserInfo() {
  var result = callSalesforceApi_('/services/oauth2/userinfo');
  return {
    id: result.user_id || result.sub || '',
    name: result.name || '',
    email: result.email || '',
    organizationId: result.organization_id || ''
  };
}

// ============================================================
// ヘルパー関数（Private）
// ============================================================

// getApiVersion_() は Config.gs で定義済み（GAS同一グローバルスコープ）
// 重複定義を避けるため、ここでは定義しない

/**
 * フィールドの型名を日本語に変換する
 *
 * @param {string} sfType - Salesforceのフィールド型名（例: "string", "double"）
 * @return {string} 日本語の型名（例: "テキスト", "数値（小数）"）
 * @private
 */
function translateFieldType_(sfType) {
  if (!sfType) return '';

  // Config.gs に FIELD_TYPE_LABELS が定義されていればそちらを優先
  var labels;
  if (typeof FIELD_TYPE_LABELS !== 'undefined') {
    labels = FIELD_TYPE_LABELS;
  } else {
    labels = DATA_TYPE_LABELS_FALLBACK;
  }

  return labels[sfType] || sfType;
}

/**
 * フィールドが必須かどうかを判定する
 *
 * Salesforce の必須判定ロジック:
 * - nillable === false（NULL不可）
 * - かつ createable === true（作成時に値を設定可能）
 * ※ ID項目やシステム項目は nillable=false だが createable=false のため必須ではない
 *
 * @param {Object} field - フィールド情報オブジェクト
 * @return {boolean} 必須かどうか
 * @private
 */
function isRequired_(field) {
  return !field.nillable && field.createable;
}

/**
 * Boolean値を日本語表記に変換する
 * true → "○", false → "−"
 *
 * @param {boolean} value - 変換対象の値
 * @return {string} "○" または "−"
 * @private
 */
function formatBoolean_(value) {
  return value ? '○' : '−';
}

/**
 * 桁数/長さのフォーマット
 *
 * - テキスト系: length値をそのまま返す（例: "255"）
 * - 数値系（precision > 0）: "precision.scale" 形式（例: "18.2"）
 * - その他: 空文字
 *
 * @param {Object} field - フィールド情報オブジェクト
 * @return {string} フォーマット済み桁数/長さ
 * @private
 */
function formatLengthPrecision_(field) {
  // 数値系フィールド（precision が設定されている場合）
  if (field.precision > 0) {
    return field.precision + '.' + (field.scale || 0);
  }

  // テキスト系フィールド（length が設定されている場合）
  if (field.length > 0) {
    return String(field.length);
  }

  return '';
}

/**
 * 参照先オブジェクトの配列をカンマ区切り文字列に変換
 *
 * @param {Array<string>} referenceTo - 参照先オブジェクト名の配列
 * @return {string} カンマ区切りの文字列（例: "Account, Contact"）
 * @private
 */
function formatReferenceTo_(referenceTo) {
  if (!referenceTo || !Array.isArray(referenceTo) || referenceTo.length === 0) {
    return '';
  }
  return referenceTo.join(', ');
}

/**
 * 選択リスト値を改行区切り文字列に変換
 *
 * 形式: "値ラベル (API値)" を改行で区切り
 * ラベルとAPI値が同じ場合はラベルのみ表示
 *
 * @param {Array<Object>} picklistValues - 選択リスト値の配列
 * @return {string} 改行区切りの文字列
 * @private
 */
function formatPicklistValues_(picklistValues) {
  if (!picklistValues || !Array.isArray(picklistValues) || picklistValues.length === 0) {
    return '';
  }

  return picklistValues.map(function(pv) {
    if (pv.label === pv.value) {
      return pv.value;
    }
    return pv.label + ' (' + pv.value + ')';
  }).join('\n');
}

/**
 * デフォルト値をスプレッドシート表示用にフォーマット
 *
 * @param {*} defaultValue - デフォルト値（任意の型）
 * @return {string} フォーマット済みデフォルト値
 * @private
 */
function formatDefaultValue_(defaultValue) {
  if (defaultValue === null || defaultValue === undefined) {
    return '';
  }
  if (typeof defaultValue === 'boolean') {
    return defaultValue ? 'true' : 'false';
  }
  return String(defaultValue);
}
