# Google Workspace Marketplace 申請チェックリスト

MMObjectView (App ID: 415072229976) のMarketplace公開に向けた対応手順。
各手順が完了したらチェックを入れる。

---

## ステップ1: スコープの整合性確認

3箇所のスコープ設定が**完全に一致**している必要がある。

対象スコープ（3つ）:
- `https://www.googleapis.com/auth/spreadsheets.currentonly`
- `https://www.googleapis.com/auth/script.container.ui`
- `https://www.googleapis.com/auth/script.external_request`

### チェック項目

- [ ] **1-1.** `src/appsscript.json` の `oauthScopes` に上記3つが記載されている
  - URL: GASエディタ > プロジェクトの設定 > appsscript.json を表示
- [ ] **1-2.** GCP Console の OAuth 同意画面にスコープ3つを追加した
  - URL: https://console.cloud.google.com/apis/credentials/consent
  - 「スコープを追加または削除」から上記3つを追加
- [ ] **1-3.** Marketplace SDK のスコープ設定に同じ3つを記載した
  - URL: GCP Console > API とサービス > Google Workspace Marketplace SDK > アプリの構成
  - 「OAuth Scopes」欄に上記3つを入力
- [ ] **1-4.** 3箇所すべてのスコープが完全一致していることを目視確認した

---

## ステップ2: OAuth 同意画面の設定確認

- [ ] **2-1.** ユーザーの種類が「外部」に設定されている
- [ ] **2-2.** アプリ名が「MMObjectView」になっている
- [ ] **2-3.** ユーザーサポートメールが設定されている
- [ ] **2-4.** アプリのホームページURLが設定されている（`index.html` をホスティングしているURL）
- [ ] **2-5.** プライバシーポリシーURLが設定されている
- [ ] **2-6.** 利用規約URLが設定されている（プライバシーポリシーと同じページでもOK）
- [ ] **2-7.** 承認済みドメインにホスティング先のドメインが追加されている

---

## ステップ3: ウェブサイト（index.html）の確認

- [ ] **3-1.** Marketplaceへのリンクやボタンが含まれていない
- [ ] **3-2.** 「Coming Soon」「近日公開」等の文言が含まれていない
- [ ] **3-3.** プライバシーポリシーにスコープの使用目的が記載されている
- [ ] **3-4.** 問い合わせ先が記載されている
- [ ] **3-5.** ウェブサイトがHTTPSでアクセスできる

---

## ステップ4: OAuth 検証の申請

- [ ] **4-1.** GCP Console の OAuth 同意画面で「確認のために送信（Submit for Verification）」をクリックした
  - URL: https://console.cloud.google.com/apis/credentials/consent
- [ ] **4-2.** スコープの使用目的を英語で記載して送信した
  - `spreadsheets.currentonly`: To create new sheets and write Salesforce object field metadata into the currently open spreadsheet
  - `script.container.ui`: To display a sidebar UI for users to authenticate with Salesforce and select objects to explore
  - `script.external_request`: To make REST API calls to Salesforce for OAuth2 authentication and retrieving object/field metadata
- [ ] **4-3.** Googleからの確認メール（追加資料の要求等）に対応した
- [ ] **4-4.** OAuth 検証が承認された（Trust & Security チームからの承認メールを受信）

---

## ステップ5: Marketplace 再申請

**※ステップ4が完了するまで再申請しないこと**

- [ ] **5-1.** OAuth 検証が完了していることを再確認した
- [ ] **5-2.** GASエディタからアドオンを再公開（Publish）した
- [ ] **5-3.** Marketplace Review Team からの審査結果を受信した
- [ ] **5-4.** Marketplace に公開された

---

## 参考リンク

| リンク | 用途 |
|-------|------|
| https://console.cloud.google.com/apis/credentials/consent | OAuth 同意画面の設定 |
| https://developers.google.com/apps-script/guides/client-verification | Apps Script のクライアント検証ガイド |
| https://support.google.com/cloud/answer/7454865 | OAuth 検証要件 |
| https://support.google.com/cloud/answer/10311615 | OAuth 同意画面の設定ヘルプ |
| https://developers.google.com/workspace/marketplace/enable-configure-sdk | Marketplace SDK の設定方法 |

---

## メモ欄

対応中に気づいた点や、Googleからの返信内容をここに記録する。

-
