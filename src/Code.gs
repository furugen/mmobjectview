/**
 * Code.gs - メインエントリポイント（ファサードパターン）
 *
 * sidebar.html から google.script.run で呼び出される全ての関数を定義する。
 * 各関数は対応するモジュール（Auth.gs, SalesforceApi.gs, SheetBuilder.gs）に処理を委譲し、
 * 統一レスポンス形式 { success, data, error } で結果を返す。
 *
 * GASの制約:
 *   - google.script.run から呼び出せるのはグローバル関数のみ
 *   - 全 .gs ファイルが同一グローバルスコープを共有するため、関数名の衝突に注意
 *   - サイドバーから呼ぶファサード関数は "facade" 接頭辞を付けて一意性を確保
 *     （SalesforceApi.gs の getObjectList, describeObject 等と衝突しないように）
 *
 * サイドバーから呼ばれるグローバル関数一覧:
 *   - checkAuthStatus()
 *   - getAuthorizationUrl()
 *   - authCallback()         ※ OAuth2ライブラリから呼ばれる
 *   - logout()
 *   - switchEnvironment(env)
 *   - fetchObjects(filterOptions)       ← SalesforceApi.gs の getObjectList() を呼ぶ
 *   - expandObjectToSheet(objectName)   ← describeObject() + buildSheet() を呼ぶ
 *   - getEnvironmentInfo()
 */

// ============================================================
// Simple Trigger / メニュー登録
// ============================================================

/**
 * スプレッドシートを開いたときに呼ばれる Simple Trigger
 * アドオンメニューを追加する
 * @param {Object} e - イベントオブジェクト
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('サイドバーを開く', 'showSidebar')
    .addToUi();
}

/**
 * アドオンインストール時に呼ばれるトリガー
 * onOpen と同じ処理を実行する
 * @param {Object} e - イベントオブジェクト
 */
function onInstall(e) {
  onOpen(e);
}

// ============================================================
// サイドバー表示
// ============================================================

/**
 * サイドバーを表示する
 * メニュー「Salesforce」→「サイドバーを開く」から呼び出される
 */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle(APP_CONFIG.SIDEBAR_TITLE);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================
// 認証関連（Auth.gs へ委譲）
// ============================================================

/**
 * 認証状態をチェックする
 * サイドバーの初期表示時に呼び出される
 *
 * @return {Object} 統一レスポンス形式
 *   成功時 data: {
 *     authenticated: boolean,
 *     instanceUrl: string|null,   // 認証済みの場合のみ
 *     environment: string         // "production" | "sandbox"
 *   }
 */
function checkAuthStatus() {
  try {
    console.info('[checkAuthStatus] 開始');
    var service = getSalesforceService_();
    var authenticated = service.hasAccess();
    var data = {
      authenticated: authenticated,
      instanceUrl: null,
      environment: getLoginUrl_() === SF_ENVIRONMENTS.sandbox.loginUrl ? 'sandbox' : 'production'
    };

    if (authenticated) {
      try {
        var token = service.getToken();
        data.instanceUrl = token.instance_url || null;
      } catch (tokenErr) {
        console.warn('トークン情報の取得に失敗しましたが、認証状態は有効です: ' + tokenErr.message);
      }
    }

    return createSuccessResponse_(data);
  } catch (e) {
    console.error('[checkAuthStatus] エラー発生:', e.message);
    console.error('[checkAuthStatus] スタックトレース:', e.stack);
    console.error('[checkAuthStatus] エラー名:', e.name);
    return createErrorResponse_(ERROR_CODES.UNKNOWN_ERROR, e.message);
  }
}

/**
 * Salesforce 認証URLを取得する
 * サイドバーの「Salesforceに接続」ボタンから呼び出される
 *
 * @return {Object} 統一レスポンス形式
 *   成功時 data: { authUrl: string }
 */
function getAuthorizationUrl() {
  try {
    console.info('[getAuthorizationUrl] 開始');
    var service = getSalesforceService_();
    var authUrl = service.getAuthorizationUrl();
    return createSuccessResponse_({ authUrl: authUrl });
  } catch (e) {
    console.error('[getAuthorizationUrl] エラー発生:', e.message);
    console.error('[getAuthorizationUrl] スタックトレース:', e.stack);
    return createErrorResponse_(ERROR_CODES.AUTH_FAILED, e.message);
  }
}

/**
 * OAuth2 コールバック関数
 * Salesforce からのリダイレクト時にOAuth2ライブラリが自動的に呼び出す
 *
 * @param {Object} request - コールバックリクエスト
 * @return {HtmlOutput} 認証結果を表示するHTML
 */
function authCallback(request) {
  try {
    var service = getSalesforceService_();
    var authorized = service.handleCallback(request);

    if (authorized) {
      console.info('Salesforce OAuth2 認証成功');
      return HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html><head><style>'
        + 'body { font-family: "Google Sans", Arial, sans-serif; text-align: center; padding: 40px 20px; color: #202124; }'
        + '.success { color: #1e8e3e; font-size: 24px; margin-bottom: 16px; }'
        + 'p { color: #5f6368; font-size: 14px; }'
        + '</style></head><body>'
        + '<div class="success">✅ 接続成功</div>'
        + '<p>Salesforce への接続に成功しました。<br>このウィンドウを閉じて、スプレッドシートに戻ってください。</p>'
        + '<p style="color:#aaa; font-size:12px;">（3秒後に自動で閉じます）</p>'
        + '<script>setTimeout(function(){ window.close(); }, 3000);</script>'
        + '</body></html>'
      );
    } else {
      console.warn('Salesforce OAuth2 認証: ユーザーが拒否またはエラー');
      return HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html><head><style>'
        + 'body { font-family: "Google Sans", Arial, sans-serif; text-align: center; padding: 40px 20px; color: #202124; }'
        + '.error { color: #d93025; font-size: 24px; margin-bottom: 16px; }'
        + 'p { color: #5f6368; font-size: 14px; }'
        + '</style></head><body>'
        + '<div class="error">❌ 接続失敗</div>'
        + '<p>Salesforce への接続に失敗しました。<br>もう一度お試しください。</p>'
        + '</body></html>'
      );
    }
  } catch (e) {
    console.error('authCallback error:', e.message, e.stack);
    // e.message をHTMLエスケープしてXSS対策
    var safeMessage = String(e.message || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><body>'
      + '<h3 style="color:red;">エラーが発生しました</h3>'
      + '<p>' + safeMessage + '</p>'
      + '</body></html>'
    );
  }
}

/**
 * Salesforce との接続を解除する（ログアウト）
 * サイドバーの「切断」ボタンから呼び出される
 *
 * @return {Object} 統一レスポンス形式
 *   成功時 data: { disconnected: true }
 */
function logout() {
  try {
    resetAuth_();

    // ログアウト時はキャッシュもクリア（再接続時に古いデータが表示されないように）
    try {
      clearObjectListCache();
    } catch (cacheErr) {
      console.warn('ログアウト時のキャッシュクリアをスキップ: ' + cacheErr.message);
    }

    console.info('Salesforce 接続を切断しました');
    return createSuccessResponse_({ disconnected: true }, 'Salesforce との接続を解除しました。');
  } catch (e) {
    console.error('logout error:', e.message, e.stack);
    return createErrorResponse_(ERROR_CODES.UNKNOWN_ERROR, e.message);
  }
}

/**
 * Salesforce の環境を切り替える（本番 ↔ Sandbox）
 * サイドバーの環境切替UIから呼び出される
 *
 * @param {string} env - "production" | "sandbox"
 * @return {Object} 統一レスポンス形式
 *   成功時 data: { environment: string, loginUrl: string }
 */
function switchEnvironment(env) {
  try {
    if (!SF_ENVIRONMENTS[env]) {
      return createErrorResponse_(ERROR_CODES.INVALID_INPUT, '不正な環境指定: ' + env);
    }

    var envConfig = SF_ENVIRONMENTS[env];

    // 既存の接続をリセット（環境変更時はトークン無効化が必要）
    try {
      resetAuth_();
    } catch (resetErr) {
      // リセット失敗は無視（初回利用時など未認証の場合）
      console.warn('環境切替時のリセットをスキップ: ' + resetErr.message);
    }

    // 環境変更時はキャッシュもクリア（異なるSalesforce組織のオブジェクト一覧が混在しないように）
    try {
      clearObjectListCache();
    } catch (cacheErr) {
      console.warn('キャッシュクリアをスキップ: ' + cacheErr.message);
    }

    // 新しいログインURLを保存
    PropertiesService.getScriptProperties()
      .setProperty(PROP_KEYS.SF_LOGIN_URL, envConfig.loginUrl);

    console.info('Salesforce 環境を切り替えました: ' + env + ' (' + envConfig.loginUrl + ')');

    return createSuccessResponse_(
      { environment: env, loginUrl: envConfig.loginUrl },
      envConfig.label + 'に切り替えました。再度接続してください。'
    );
  } catch (e) {
    console.error('switchEnvironment error:', e.message, e.stack);
    return createErrorResponse_(ERROR_CODES.UNKNOWN_ERROR, e.message);
  }
}

// ============================================================
// Salesforce API 関連（SalesforceApi.gs へ委譲）
// ============================================================

/**
 * オブジェクト一覧を取得する（サイドバーから呼び出されるファサード関数）
 *
 * ※ 関数名を "fetchObjects" としている理由:
 *    SalesforceApi.gs に "getObjectList" という同名関数があるため、
 *    GAS のグローバルスコープでの関数名衝突を避けるために別名とした。
 *
 * @param {Object} [filterOptions] - フィルタオプション（サイドバーUIからは null を渡す）
 *   サイドバー側でクライアントフィルタリングを行うため、サーバーには全件取得を依頼。
 * @return {Object} 統一レスポンス形式
 *   成功時 data: {
 *     objects: Array<{ name, label, keyPrefix, custom, queryable, ... }>,
 *     totalCount: number
 *   }
 */
function fetchObjects(filterOptions) {
  try {
    // 認証チェック
    if (!isAuthorized_()) {
      return createErrorResponse_(ERROR_CODES.AUTH_REQUIRED);
    }

    // SalesforceApi.gs: getObjectList(filterOptions) を呼び出し
    // サイドバーでは全件取得 → クライアント側でフィルタリング
    var objects = getObjectList(filterOptions || {});

    return createSuccessResponse_(
      {
        objects: objects,
        totalCount: objects.length
      },
      'オブジェクト一覧を取得しました（' + objects.length + '件）'
    );

  } catch (e) {
    console.error('fetchObjects error:', e.message, e.stack);
    return handleApiError_(e);
  }
}

/**
 * 選択されたオブジェクトのフィールド情報を新規スプレッドシートに展開する
 * サイドバーの「展開」ボタンから呼び出される
 *
 * @param {string} objectApiName - Salesforce オブジェクトの API 名（例: "Account"）
 * @return {Object} 統一レスポンス形式
 *   成功時 data: {
 *     spreadsheetUrl: string,
 *     spreadsheetId: string,
 *     sheetName: string,
 *     objectName: string,
 *     objectLabel: string,
 *     fieldCount: number
 *   }
 */
function expandObjectToSheet(objectApiName) {
  try {
    // 入力バリデーション
    if (!objectApiName || typeof objectApiName !== 'string' || objectApiName.trim() === '') {
      return createErrorResponse_(ERROR_CODES.INVALID_INPUT, 'オブジェクト名が指定されていません');
    }

    // 認証チェック
    if (!isAuthorized_()) {
      return createErrorResponse_(ERROR_CODES.AUTH_REQUIRED);
    }

    var cleanName = objectApiName.trim();

    // ステップ1: Salesforce からオブジェクト詳細（フィールド一覧）を取得
    //   SalesforceApi.gs: describeObject(objectApiName)
    //   戻り値: { objectInfo: {...}, fields: [...] }
    console.info('オブジェクト展開開始: ' + cleanName);
    var describeResult = describeObject(cleanName);

    // ステップ2: 新規スプレッドシートにフィールド情報を展開
    //   SheetBuilder.gs: buildSheet(objectInfo, fields)
    //   戻り値: { spreadsheetUrl, spreadsheetId, sheetName, fieldCount }
    var sheetResult = buildSheet(describeResult.objectInfo, describeResult.fields);

    console.info('オブジェクト展開完了: ' + cleanName + ' (' + sheetResult.fieldCount + 'フィールド)');

    return createSuccessResponse_(
      {
        spreadsheetUrl: sheetResult.spreadsheetUrl,
        spreadsheetId: sheetResult.spreadsheetId,
        sheetName: sheetResult.sheetName,
        objectName: cleanName,
        objectLabel: describeResult.objectInfo.label,
        fieldCount: sheetResult.fieldCount
      },
      describeResult.objectInfo.label + ' (' + cleanName + ') を展開しました（' + sheetResult.fieldCount + 'フィールド）'
    );

  } catch (e) {
    console.error('expandObjectToSheet error:', e.message, e.stack);
    return handleApiError_(e);
  }
}

// ============================================================
// 環境情報取得
// ============================================================

/**
 * 現在の環境設定を取得する
 * サイドバーの環境表示で呼び出される
 *
 * @return {Object} 統一レスポンス形式
 *   成功時 data: {
 *     environment: string,       // "production" | "sandbox"
 *     environmentLabel: string,  // "本番環境" | "Sandbox"
 *     apiVersion: string,        // "v62.0"
 *     appVersion: string         // "1.0.0"
 *   }
 */
function getEnvironmentInfo() {
  try {
    var loginUrl = getLoginUrl_();
    var env = loginUrl === SF_ENVIRONMENTS.sandbox.loginUrl ? 'sandbox' : 'production';
    var envLabel = SF_ENVIRONMENTS[env].label;

    return createSuccessResponse_({
      environment: env,
      environmentLabel: envLabel,
      apiVersion: getApiVersion_(),
      appVersion: APP_CONFIG.VERSION
    });
  } catch (e) {
    console.error('getEnvironmentInfo error:', e.message, e.stack);
    return createErrorResponse_(ERROR_CODES.UNKNOWN_ERROR, e.message);
  }
}

// ============================================================
// 内部ヘルパー関数
// ============================================================

/**
 * API エラーを統一レスポンス形式に変換するヘルパー
 * SalesforceApiError の場合は HTTP ステータスに応じたエラーコードを使用
 *
 * @param {Error} e - 発生したエラー
 * @return {Object} エラーレスポンス
 * @private
 */
function handleApiError_(e) {
  // SalesforceApiError（SalesforceApi.gs で定義）の場合
  if (e.name === 'SalesforceApiError') {
    var code;
    switch (e.statusCode) {
      case 401:
        code = ERROR_CODES.AUTH_EXPIRED;
        break;
      case 403:
        code = (e.errorCode === 'REQUEST_LIMIT_EXCEEDED')
          ? ERROR_CODES.API_RATE_LIMIT
          : ERROR_CODES.API_FORBIDDEN;
        break;
      case 404:
        code = ERROR_CODES.API_NOT_FOUND;
        break;
      default:
        code = ERROR_CODES.API_ERROR;
    }
    return createErrorResponse_(code, e.message);
  }

  // SalesforceApi.gs が throw するオブジェクト形式のエラー
  if (e.code && ERROR_CODES[e.code]) {
    return createErrorResponse_(e.code, e.details || e.message);
  }

  // 認証関連エラー
  if (e.message && e.message.indexOf('Salesforceに接続されていません') >= 0) {
    return createErrorResponse_(ERROR_CODES.AUTH_REQUIRED, e.message);
  }

  // その他の未知のエラー
  return createErrorResponse_(ERROR_CODES.UNKNOWN_ERROR, e.message);
}

// ============================================================
// デバッグ用テスト関数
// ============================================================

/**
 * デバッグ用: 最小限のテスト関数
 * PropertiesServiceやOAuth2など外部依存なし。
 * google.script.run 自体が動作するかの切り分け用。
 */
function testPing() {
  return { success: true, data: { message: 'pong! サーバー応答OK', timestamp: new Date().toISOString() } };
}
