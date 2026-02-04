/**
 * Auth.gs - Salesforce OAuth2 認証
 *
 * apps-script-oauth2 ライブラリを使用して Salesforce への OAuth2.0 Web Server Flow を実装する。
 *
 * 責務:
 * - OAuth2 サービスの構築（getSalesforceService_）
 * - 認証状態の確認（isAuthorized_）
 * - アクセストークンの取得（getAccessToken_、自動リフレッシュ対応）
 * - Salesforce インスタンスURLの取得（getInstanceUrl_）
 * - 認証リセット / ログアウト（resetAuth_）
 *
 * 設計方針:
 * - 関数名末尾に _ を付けたものは内部利用（Code.gs 等から呼ばれるが、サイドバーから直接呼ばれない）
 * - authCallback() のみグローバル関数（OAuth2ライブラリから呼び出されるため Code.gs に配置）
 * - Client ID / Secret は Script Properties から取得（ハードコード厳禁）
 * - トークンは PropertiesService.getUserProperties() + CacheService.getUserCache() に保存
 *
 * 依存: Config.gs（PROP_KEYS, SF_DEFAULTS, APP_CONFIG）
 * 参照: docs/OAUTH2_DESIGN.md
 */

// ============================================================
// OAuth2 サービス構築
// ============================================================

/**
 * Salesforce 用 OAuth2 サービスを構築する
 *
 * apps-script-oauth2 ライブラリの OAuth2.createService() を使用して
 * Salesforce 接続用のサービスオブジェクトを生成する。
 *
 * @return {OAuth2.Service} OAuth2 サービスインスタンス
 * @private
 */
function getSalesforceService_() {
  console.info('[getSalesforceService_] サービス構築開始');
  var scriptProps = PropertiesService.getScriptProperties();
  var loginUrl = getLoginUrl_();
  var clientId = scriptProps.getProperty(PROP_KEYS.SF_CLIENT_ID);
  var clientSecret = scriptProps.getProperty(PROP_KEYS.SF_CLIENT_SECRET);

  console.info('[getSalesforceService_] Client ID設定: ' + (clientId ? '✅ 有り' : '❌ なし'));
  console.info('[getSalesforceService_] Client Secret設定: ' + (clientSecret ? '✅ 有り' : '❌ なし'));
  console.info('[getSalesforceService_] Login URL: ' + loginUrl);

  // Client ID / Secret の設定チェック
  if (!clientId || !clientSecret) {
    console.error('Auth.gs: SF_CLIENT_ID または SF_CLIENT_SECRET が Script Properties に設定されていません');
    throw new Error(
      'Salesforce の接続設定が完了していません。管理者に連絡して、'
      + 'Script Properties に SF_CLIENT_ID と SF_CLIENT_SECRET を設定してください。'
    );
  }

  return OAuth2.createService(APP_CONFIG.OAUTH_SERVICE_NAME)
    // ---- 認可エンドポイント ----
    .setAuthorizationBaseUrl(loginUrl + '/services/oauth2/authorize')
    // ---- トークンエンドポイント ----
    .setTokenUrl(loginUrl + '/services/oauth2/token')

    // ---- クライアント情報（Connected App） ----
    .setClientId(clientId)
    .setClientSecret(clientSecret)

    // ---- コールバック関数名 ----
    // Code.gs で定義された authCallback() が呼ばれる
    .setCallbackFunction(APP_CONFIG.OAUTH_CALLBACK_FUNCTION)

    // ---- トークン保存先（ユーザーごとに隔離） ----
    .setPropertyStore(PropertiesService.getUserProperties())

    // ---- パフォーマンス最適化 ----
    // CacheService を使用して PropertiesService のクォータ消費を抑える
    .setCache(CacheService.getUserCache())
    // LockService で並行アクセス時のトークンリフレッシュを保護
    .setLock(LockService.getUserLock())

    // ---- Salesforce OAuth スコープ ----
    // api: REST API へのアクセス
    // refresh_token: リフレッシュトークンの取得（長期的なアクセス維持）
    .setScope(SF_DEFAULTS.OAUTH_SCOPE)

    // ---- トークンリクエストヘッダー ----
    .setTokenHeaders({
      'Content-Type': 'application/x-www-form-urlencoded'
    });
}

// ============================================================
// 認証状態の確認
// ============================================================

/**
 * Salesforce に認証済みかどうかを確認する
 *
 * OAuth2 サービスの hasAccess() を呼び出してトークンの存在を確認する。
 * ※ hasAccess() は「トークンが存在するか」のみチェックし、有効期限のチェックは行わない。
 *    実際のAPI呼び出し時に 401 が返った場合は getAccessToken_() で自動リフレッシュされる。
 *
 * @return {boolean} 認証済みかどうか
 * @private
 */
function isAuthorized_() {
  try {
    var service = getSalesforceService_();
    return service.hasAccess();
  } catch (e) {
    console.warn('Auth.gs isAuthorized_: 認証状態の確認に失敗: ' + e.message);
    return false;
  }
}

// ============================================================
// アクセストークンの取得
// ============================================================

/**
 * 有効なアクセストークンを取得する
 *
 * OAuth2 ライブラリの getAccessToken() は内部的にトークンの有効期限をチェックし、
 * 期限切れの場合はリフレッシュトークンを使って自動的に新しいアクセストークンを取得する。
 *
 * @return {string} アクセストークン
 * @throws {Error} 未認証またはリフレッシュ失敗時
 * @private
 */
function getAccessToken_() {
  var service = getSalesforceService_();

  if (!service.hasAccess()) {
    console.error('Auth.gs getAccessToken_: Salesforce に接続されていません');
    throw new Error('Salesforceに接続されていません。サイドバーから接続してください。');
  }

  try {
    var token = service.getAccessToken();
    if (!token) {
      throw new Error('アクセストークンが空です');
    }
    return token;
  } catch (e) {
    console.error('Auth.gs getAccessToken_: トークン取得失敗: ' + e.message);

    // リフレッシュトークンが無効化されている可能性がある場合
    // トークンをリセットして再認証を促す
    try {
      service.reset();
      console.info('Auth.gs: 無効なトークンをリセットしました');
    } catch (resetErr) {
      console.warn('Auth.gs: トークンリセット失敗: ' + resetErr.message);
    }

    throw new Error('認証の有効期限が切れました。再度接続してください。');
  }
}

// ============================================================
// インスタンスURL の取得
// ============================================================

/**
 * Salesforce インスタンスURLを取得する
 *
 * トークンレスポンスに含まれる instance_url はAPI呼び出し先のベースURLとなる。
 * 例: "https://na1.salesforce.com", "https://mycompany.my.salesforce.com"
 *
 * ※ login.salesforce.com に直接APIリクエストを送ってはいけない。必ず instance_url を使用する。
 *
 * @return {string} インスタンスURL
 * @throws {Error} 未認証またはトークン情報取得失敗時
 * @private
 */
function getInstanceUrl_() {
  var service = getSalesforceService_();

  if (!service.hasAccess()) {
    throw new Error('Salesforceに接続されていません。');
  }

  var token = service.getToken();
  if (!token || !token.instance_url) {
    console.error('Auth.gs getInstanceUrl_: トークンに instance_url が含まれていません');
    throw new Error('Salesforce インスタンスURLを取得できませんでした。再度接続してください。');
  }

  return token.instance_url;
}

// ============================================================
// 認証リセット / ログアウト
// ============================================================

/**
 * Salesforce 認証をリセットする（ログアウト）
 *
 * OAuth2 サービスのトークンを削除し、キャッシュもクリアする。
 * 環境切替時にも呼び出される（環境変更時は既存のトークンが無効になるため）。
 *
 * @private
 */
function resetAuth_() {
  try {
    var service = getSalesforceService_();
    service.reset();
    console.info('Auth.gs: OAuth2 トークンをリセットしました');
  } catch (e) {
    // サービス構築自体が失敗する場合は、直接 PropertiesService をクリーンアップ
    console.warn('Auth.gs resetAuth_: サービスリセット失敗、手動クリーンアップを実行: ' + e.message);
    try {
      var userProps = PropertiesService.getUserProperties();
      // OAuth2ライブラリは "oauth2.{serviceName}" の形式でキーを保存する
      userProps.deleteProperty('oauth2.' + APP_CONFIG.OAUTH_SERVICE_NAME);
    } catch (cleanupErr) {
      console.error('Auth.gs resetAuth_: 手動クリーンアップも失敗: ' + cleanupErr.message);
    }
  }
}

// ============================================================
// デバッグ用関数（開発時に使用、本番では使わない）
// ============================================================

/**
 * Salesforce 接続情報をログ出力する（開発・デバッグ用）
 * GASエディタの「実行」から直接実行して確認する
 */
function debugConnectionInfo_() {
  console.log('=== Salesforce 接続デバッグ情報 ===');

  var scriptProps = PropertiesService.getScriptProperties();
  console.log('Login URL: ' + (scriptProps.getProperty(PROP_KEYS.SF_LOGIN_URL) || '(未設定 → デフォルト: ' + SF_DEFAULTS.LOGIN_URL + ')'));
  console.log('API Version: ' + (scriptProps.getProperty(PROP_KEYS.SF_API_VERSION) || '(未設定 → デフォルト: ' + SF_DEFAULTS.API_VERSION + ')'));
  console.log('Client ID 設定: ' + (scriptProps.getProperty(PROP_KEYS.SF_CLIENT_ID) ? '✅ 済' : '❌ 未設定'));
  console.log('Client Secret 設定: ' + (scriptProps.getProperty(PROP_KEYS.SF_CLIENT_SECRET) ? '✅ 済' : '❌ 未設定'));

  try {
    var service = getSalesforceService_();
    console.log('接続状態: ' + (service.hasAccess() ? '✅ 接続中' : '❌ 未接続'));

    if (service.hasAccess()) {
      var token = service.getToken();
      console.log('Instance URL: ' + token.instance_url);
      console.log('Token Type: ' + token.token_type);
      if (token.issued_at) {
        console.log('Issued At: ' + new Date(parseInt(token.issued_at)).toLocaleString('ja-JP'));
      }
    }

    console.log('Callback URL: ' + service.getRedirectUri());
  } catch (e) {
    console.log('サービス構築エラー: ' + e.message);
  }

  console.log('===================================');
}
