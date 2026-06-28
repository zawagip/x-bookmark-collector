# x-bookmark-collector

XブックマークからFANZA URLを収集するGASパイプライン

## 構成
- gas/bookmarks.js : ブックマーク取得・スプシ書き込み
- Cloudflare Workers (ryulink) でX OAuth2中継

## セットアップ
1. https://ryulink.zawagipyask.workers.dev/x-oauth/redirect で認証
2. GAS ScriptPropertiesに設定:
   - WORKERS_ADMIN_SECRET
   - X_USER_ID
3. fetchAndStoreBookmarks() を実行

## トリガー
1時間おきに自動実行
