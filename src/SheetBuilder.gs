/**
 * SheetBuilder.gs
 * スプレッドシート生成・書式設定モジュール
 *
 * 責務:
 * - 既存スプレッドシートへのシート追加（SpreadsheetApp.getActiveSpreadsheet + insertSheet）
 * - ヘッダー行の書き込み・書式設定
 * - フィールドデータの一括書き込み（setValues）
 * - 書式設定（色、フォント、列幅、フィルター、ゼブラストライプ）
 * - サマリー情報の追加
 *
 * 依存モジュール:
 * - Config.gs: SHEET_CONFIG（書式定数）、APP_CONFIG（バージョン情報）
 * - SalesforceApi.gs: formatFieldsForSheet()（フィールドデータの整形）
 *
 * @author Salesforce API Specialist
 * @version 1.1.0
 * @since 2026-01-30
 */

// ============================================================
// メインAPI
// ============================================================

/**
 * オブジェクト情報を現在開いているスプレッドシートに新しいシートとして展開する
 *
 * 処理フロー:
 * 1. 現在のスプレッドシートを取得し、新しいシートを追加
 * 2. ヘッダー行を書き込み
 * 3. フィールドデータを一括書き込み
 * 4. 書式を設定（色、フォント、列幅、フィルター、ゼブラストライプ）
 * 5. サマリー情報を追加
 * 6. 作成結果を返却
 *
 * @param {Object} objectInfo - オブジェクトのメタ情報
 * @param {string} objectInfo.name - オブジェクトAPI名（例: "Account"）
 * @param {string} objectInfo.label - オブジェクト表示ラベル（例: "取引先"）
 * @param {Array<Object>} fields - フィールド情報の配列（describeObject()の戻り値.fields）
 * @return {Object} 作成結果
 *   {
 *     spreadsheetUrl: string,   // スプレッドシートのURL
 *     spreadsheetId: string,    // スプレッドシートID
 *     sheetName: string,        // 作成したシート名
 *     fieldCount: number        // 展開したフィールド数
 *   }
 * @throws {Object} 統一エラーオブジェクト { code, message, details }
 */
function buildSheet(objectInfo, fields) {
  try {
    // フィールドデータを24カラムの2次元配列に整形
    var formattedData = formatFieldsForSheet(fields);
    var headers = formattedData.headers;
    var rows = formattedData.rows;

    // ── 1. 現在のスプレッドシートに新しいシートを追加 ──
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = buildSheetName_(objectInfo, ss);
    console.info('[SheetBuilder] 新しいシートを追加: ' + sheetName);
    var sheet = ss.insertSheet(sheetName);

    // ── 2. ヘッダー行書き込み ──
    buildHeader_(sheet, headers);

    // ── 3. データ書き込み ──
    var dataRowCount = writeFieldData_(sheet, rows);

    // ── 4. 書式設定 ──
    applyFormatting_(sheet, headers.length, dataRowCount);

    // ── 5. サマリー情報追加 ──
    addSummaryInfo_(sheet, objectInfo, dataRowCount, headers.length);

    // ── 結果返却 ──
    var result = {
      spreadsheetUrl: ss.getUrl(),
      spreadsheetId: ss.getId(),
      sheetName: sheetName,
      fieldCount: dataRowCount
    };

    console.info('[SheetBuilder] シート追加完了: ' + sheetName + ' (' + result.spreadsheetUrl + ')');
    return result;

  } catch (e) {
    console.error('[SheetBuilder] シート追加エラー:', e.message || e);

    // 既に統一エラー形式の場合はそのまま投げる
    if (e.code) {
      throw e;
    }

    throw {
      code: typeof ERROR_CODES !== 'undefined' ? ERROR_CODES.SHEET_ERROR : 'SHEET_ERROR',
      message: typeof ERROR_MESSAGES !== 'undefined' ? ERROR_MESSAGES.SHEET_ERROR : 'シートの作成に失敗しました。',
      details: e.message || String(e)
    };
  }
}

// ============================================================
// ヘッダー構築
// ============================================================

/**
 * ヘッダー行を書き込み、書式を設定する
 *
 * - 24カラムのヘッダーを1行目に書き込み
 * - 背景色: #4285F4（Google Blue）
 * - テキスト色: 白
 * - フォント: 太字 10pt
 * - 配置: 中央揃え
 * - 行の固定: 1行目を固定
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Array<string>} headers - ヘッダー文字列の配列（24要素）
 * @private
 */
function buildHeader_(sheet, headers) {
  var colCount = headers.length;

  // ヘッダーデータ書き込み
  var headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange.setValues([headers]);

  // 書式設定定数を取得（Config.gs が存在する場合はそちらを参照）
  var bgColor = getSheetConfig_('HEADER_BG_COLOR', '#4285F4');
  var fontColor = getSheetConfig_('HEADER_FONT_COLOR', '#FFFFFF');
  var fontSize = getSheetConfig_('HEADER_FONT_SIZE', 10);
  var frozenRows = getSheetConfig_('FROZEN_ROWS', 1);

  // 書式設定
  headerRange
    .setBackground(bgColor)
    .setFontColor(fontColor)
    .setFontWeight('bold')
    .setFontSize(fontSize)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  // ヘッダー行の高さを少し広げる
  sheet.setRowHeight(1, 32);

  // 1行目を固定（スクロールしてもヘッダーが見える）
  sheet.setFrozenRows(frozenRows);
}

// ============================================================
// データ書き込み
// ============================================================

/**
 * フィールドデータを一括書き込みする
 *
 * パフォーマンス重視: setValues() で一括書き込み
 * （1セルずつ setValue() すると非常に遅い）
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Array<Array>} rows - データ行の2次元配列
 * @return {number} 書き込んだデータ行数
 * @private
 */
function writeFieldData_(sheet, rows) {
  if (!rows || rows.length === 0) {
    console.warn('[SheetBuilder] 書き込むフィールドデータがありません');
    return 0;
  }

  var rowCount = rows.length;
  var colCount = rows[0].length;

  // 2行目からデータを一括書き込み（1行目はヘッダー）
  var dataRange = sheet.getRange(2, 1, rowCount, colCount);
  dataRange.setValues(rows);

  // データ行のフォントサイズ設定
  var dataFontSize = getSheetConfig_('DATA_FONT_SIZE', 10);
  dataRange.setFontSize(dataFontSize);

  // テキスト折り返し設定（主要カラムのみ）
  // 数式内容(11列目)、選択リスト値(17列目)、ヘルプテキスト(21列目)
  var wrapColumns = [11, 17, 21];
  for (var i = 0; i < wrapColumns.length; i++) {
    var col = wrapColumns[i];
    if (col <= colCount) {
      sheet.getRange(2, col, rowCount, 1).setWrap(true);
    }
  }

  return rowCount;
}

// ============================================================
// 書式設定
// ============================================================

/**
 * シートの書式を設定する
 *
 * - 列幅自動調整
 * - オートフィルター設定
 * - ゼブラストライプ（交互背景色）
 * - ヘッダー下部の罫線
 * - 特定列の幅制限
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {number} colCount - カラム数
 * @param {number} dataRowCount - データ行数
 * @private
 */
function applyFormatting_(sheet, colCount, dataRowCount) {
  if (dataRowCount === 0) return;

  var totalRows = dataRowCount + 1; // ヘッダー + データ行

  // ── 列幅自動調整 ──
  for (var c = 1; c <= colCount; c++) {
    sheet.autoResizeColumn(c);
  }

  // ── 特定列の最大幅を制限 ──
  // 数式内容(11), 選択リスト値(17), ヘルプテキスト(21) は長い場合があるため制限
  var maxWidthColumns = {
    11: 300,  // 数式内容
    17: 250,  // 選択リスト値
    21: 250   // ヘルプテキスト
  };

  for (var colNum in maxWidthColumns) {
    var maxWidth = maxWidthColumns[colNum];
    var intCol = parseInt(colNum, 10);
    if (intCol <= colCount) {
      var currentWidth = sheet.getColumnWidth(intCol);
      if (currentWidth > maxWidth) {
        sheet.setColumnWidth(intCol, maxWidth);
      }
    }
  }

  // ── 特定列の最小幅を設定 ──
  // API名(1), ラベル(2) は狭すぎないようにする
  var minWidthColumns = {
    1: 120,   // API名
    2: 120,   // ラベル
    3: 100    // データ型
  };

  for (var colNum2 in minWidthColumns) {
    var minWidth = minWidthColumns[colNum2];
    var intCol2 = parseInt(colNum2, 10);
    if (intCol2 <= colCount) {
      var currentWidth2 = sheet.getColumnWidth(intCol2);
      if (currentWidth2 < minWidth) {
        sheet.setColumnWidth(intCol2, minWidth);
      }
    }
  }

  // ── オートフィルター設定 ──
  var filterRange = sheet.getRange(1, 1, totalRows, colCount);
  try {
    sheet.getFilter();  // 既存フィルタがある場合は何もしない
  } catch (e) {
    // フィルタが無い場合は設定
  }
  filterRange.createFilter();

  // ── ゼブラストライプ（偶数行に淡いグレー背景） ──
  var altRowColor = getSheetConfig_('ALT_ROW_COLOR', '#F8F9FA');
  for (var row = 2; row <= totalRows; row++) {
    // 偶数行（データの偶数番目: row 3, 5, 7... → 0始まりで偶数）
    if (row % 2 === 1) { // row=3（データ2行目）, row=5（データ4行目）...
      sheet.getRange(row, 1, 1, colCount).setBackground(altRowColor);
    }
  }

  // ── ヘッダー下部の罫線（太め） ──
  var headerBottomRange = sheet.getRange(1, 1, 1, colCount);
  headerBottomRange.setBorder(
    null,   // top
    null,   // left
    true,   // bottom
    null,   // right
    null,   // vertical
    null,   // horizontal
    '#1A73E8',   // 色（Googleブルーの濃い版）
    SpreadsheetApp.BorderStyle.SOLID_MEDIUM  // 太め実線
  );
}

// ============================================================
// サマリー情報
// ============================================================

/**
 * データの下にサマリー情報を追加する
 *
 * 空行を1行空けて、オブジェクト名、総フィールド数、生成日時、APIバージョン等を記載
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Object} objectInfo - オブジェクト情報
 * @param {number} dataRowCount - データ行数
 * @param {number} colCount - カラム数
 * @private
 */
function addSummaryInfo_(sheet, objectInfo, dataRowCount, colCount) {
  var summaryStartRow = dataRowCount + 3; // ヘッダー(1) + データ(n) + 空行(1) → n+3

  // サマリー色
  var summaryLabelColor = '#5F6368'; // グレー
  var summaryValueColor = '#202124'; // ダークグレー

  // サマリー項目の定義
  var summaryItems = [
    ['オブジェクト名:', objectInfo.label + ' (' + objectInfo.name + ')'],
    ['総フィールド数:', dataRowCount + ' 項目'],
    ['生成日時:', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')],
    ['APIバージョン:', getApiVersionForSummary_()],
    ['生成ツール:', getAppNameForSummary_()]
  ];

  for (var i = 0; i < summaryItems.length; i++) {
    var row = summaryStartRow + i;
    var labelCell = sheet.getRange(row, 1);
    var valueCell = sheet.getRange(row, 2);

    labelCell
      .setValue(summaryItems[i][0])
      .setFontColor(summaryLabelColor)
      .setFontWeight('bold')
      .setFontSize(9);

    valueCell
      .setValue(summaryItems[i][1])
      .setFontColor(summaryValueColor)
      .setFontSize(9);
  }
}

// ============================================================
// ヘルパー関数（Private）
// ============================================================

/**
 * 新規シートの名前を生成する（重複対応付き）
 *
 * 命名規則: "{オブジェクトラベル}({オブジェクトAPI名})"
 * 例: "取引先(Account)"
 *
 * 同名シートが既に存在する場合は連番を付ける:
 * "取引先(Account) (2)", "取引先(Account) (3)", ...
 *
 * GASのシート名は31文字制限があるため、超過時は切り詰める。
 *
 * @param {Object} objectInfo - オブジェクト情報
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - 対象スプレッドシート
 * @return {string} 一意なシート名
 * @private
 */
function buildSheetName_(objectInfo, ss) {
  var label = objectInfo.label || objectInfo.name;
  var name = objectInfo.name || '';
  var baseName = label + '(' + name + ')';

  // GASのシート名31文字制限に対応
  var maxLen = 31;
  if (baseName.length > maxLen) {
    baseName = baseName.substring(0, maxLen - 1) + '…';
  }

  // 既存シート名の一覧を取得
  var existingSheets = ss.getSheets().map(function(s) { return s.getName(); });

  // 重複チェック: 同名が無ければそのまま返す
  if (existingSheets.indexOf(baseName) === -1) {
    return baseName;
  }

  // 同名シートが存在する場合、連番を付与して一意な名前を生成
  for (var i = 2; i <= 100; i++) {
    var candidate = baseName + ' (' + i + ')';
    // 連番付きで31文字を超える場合はベース名を切り詰め
    if (candidate.length > maxLen) {
      var suffix = ' (' + i + ')';
      candidate = baseName.substring(0, maxLen - suffix.length) + suffix;
    }
    if (existingSheets.indexOf(candidate) === -1) {
      return candidate;
    }
  }

  // 万が一100回重複した場合はタイムスタンプ付き
  return baseName.substring(0, maxLen - 14) + '_' + Date.now();
}

/**
 * SHEET_CONFIG 定数から値を取得するヘルパー
 * Config.gs が存在しない場合はデフォルト値を返す
 *
 * @param {string} key - SHEET_CONFIG のキー
 * @param {*} defaultValue - デフォルト値
 * @return {*} 設定値
 * @private
 */
function getSheetConfig_(key, defaultValue) {
  if (typeof SHEET_CONFIG !== 'undefined' && SHEET_CONFIG[key] !== undefined) {
    return SHEET_CONFIG[key];
  }
  return defaultValue;
}

/**
 * サマリー用にAPIバージョンを取得する
 * @return {string} APIバージョン文字列
 * @private
 */
function getApiVersionForSummary_() {
  try {
    var override = PropertiesService.getScriptProperties().getProperty('SF_API_VERSION');
    if (override) return override;
  } catch (e) { /* ignore */ }

  if (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.SF_API_VERSION) {
    return APP_CONFIG.SF_API_VERSION;
  }

  return 'v62.0';
}

/**
 * サマリー用にアプリケーション名を取得する
 * @return {string} アプリケーション名
 * @private
 */
function getAppNameForSummary_() {
  if (typeof APP_CONFIG !== 'undefined' && APP_CONFIG.APP_NAME) {
    return APP_CONFIG.APP_NAME + ' ' + (APP_CONFIG.VERSION || '');
  }
  return 'Salesforce Object Explorer';
}
