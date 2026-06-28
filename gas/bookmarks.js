const WORKERS_BASE = 'https://ryulink.zawagipyask.workers.dev';
const SHEET_ID = '1uKSe2MVq3-xE6sYJLt9vkameF93-4-uB2yYebYGQvJY';
const RAW_SHEET_NAME = 'raw_data';

function fetchAndStoreBookmarks() {
  const token = getAccessToken_();
  const userId = PropertiesService.getScriptProperties().getProperty('X_USER_ID');
  if (!userId) throw new Error('X_USER_ID が ScriptProperties に未設定');

  const allBookmarks = [];
  let nextToken = null;
  let page = 0;

  do {
    page++;
    let paramStr = 'max_results=100'
      + '&tweet.fields=created_at,public_metrics,author_id,entities,conversation_id'
      + '&user.fields=username,name'
      + '&expansions=author_id';
    if (nextToken) paramStr += '&pagination_token=' + encodeURIComponent(nextToken);

    const res = UrlFetchApp.fetch(
      'https://api.twitter.com/2/users/' + userId + '/bookmarks?' + paramStr,
      {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true,
      }
    );

    const status = res.getResponseCode();
    if (status === 401) {
      Logger.log('401: tokenリフレッシュを試みます');
      refreshToken_();
      throw new Error('tokenをリフレッシュしました。もう一度実行してください。');
    }
    if (status !== 200) {
      throw new Error('ブックマーク取得失敗 (' + status + '): ' + res.getContentText());
    }

    const json = JSON.parse(res.getContentText());
    const tweets = json.data || [];
    const users = (json.includes && json.includes.users || []).reduce(function(m, u) {
      m[u.id] = u; return m;
    }, {});

    tweets.forEach(function(t) {
      const user = users[t.author_id] || {};
      allBookmarks.push({
        tweetId: t.id,
        authorId: t.author_id,
        username: user.username || '',
        text: t.text,
        createdAt: t.created_at,
        likeCount: (t.public_metrics && t.public_metrics.like_count) || 0,
        retweetCount: (t.public_metrics && t.public_metrics.retweet_count) || 0,
        conversationId: t.conversation_id,
        urls: extractFanzaUrls_(t),
      });
    });

    nextToken = (json.meta && json.meta.next_token) || null;
    Logger.log('page ' + page + ': ' + tweets.length + '件取得 (累計 ' + allBookmarks.length + '件)');

    if (nextToken) Utilities.sleep(1000);
  } while (nextToken && page < 50);

  writeToSheet_(allBookmarks);
  Logger.log('完了: 合計 ' + allBookmarks.length + '件をスプシに書き込み');
}

function getAccessToken_() {
  const adminSecret = PropertiesService.getScriptProperties().getProperty('WORKERS_ADMIN_SECRET');
  if (!adminSecret) throw new Error('WORKERS_ADMIN_SECRET が未設定');

  const res = UrlFetchApp.fetch(WORKERS_BASE + '/x-oauth/token', {
    headers: { 'X-Admin-Secret': adminSecret },
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() === 404) {
    throw new Error('tokenが見つかりません。\nブラウザで ' + WORKERS_BASE + '/x-oauth/redirect を開いて認証してください。');
  }
  if (res.getResponseCode() !== 200) {
    throw new Error('token取得失敗: ' + res.getContentText());
  }

  const data = JSON.parse(res.getContentText());

  if (data.expiresAt && Date.now() > data.expiresAt - 5 * 60 * 1000) {
    Logger.log('tokenが期限切れ間近 → リフレッシュします');
    return refreshToken_();
  }

  return data.accessToken;
}

function refreshToken_() {
  const adminSecret = PropertiesService.getScriptProperties().getProperty('WORKERS_ADMIN_SECRET');

  const res = UrlFetchApp.fetch(WORKERS_BASE + '/x-oauth/refresh', {
    method: 'POST',
    headers: { 'X-Admin-Secret': adminSecret },
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    throw new Error('リフレッシュ失敗 (' + res.getResponseCode() + '): ' + res.getContentText() + '\n手動再認証が必要です: ' + WORKERS_BASE + '/x-oauth/redirect');
  }

  const data = JSON.parse(res.getContentText());
  Logger.log('tokenリフレッシュ成功');
  return data.accessToken;
}

function extractFanzaUrls_(tweet) {
  const urls = (tweet.entities && tweet.entities.urls) || [];
  return urls
    .map(function(u) { return u.expanded_url || u.url; })
    .filter(function(u) { return /fanza\.jp|dmm\.co\.jp|al\.fanza\.co\.jp|ryulink\.link/.test(u); })
    .join(',');
}

function writeToSheet_(bookmarks) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(RAW_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RAW_SHEET_NAME);
  }

  const headers = [
    'tweetId', 'authorId', 'username', 'text',
    'createdAt', 'likeCount', 'retweetCount',
    'conversationId', 'fanzaUrls', 'importedAt',
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  const existingIds = new Set();
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
      .getValues()
      .forEach(function(r) { existingIds.add(String(r[0])); });
  }

  const now = new Date().toISOString();
  const newRows = bookmarks
    .filter(function(b) { return !existingIds.has(b.tweetId); })
    .map(function(b) {
      return [
        b.tweetId, b.authorId, b.username, b.text,
        b.createdAt, b.likeCount, b.retweetCount,
        b.conversationId, b.urls, now,
      ];
    });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length)
      .setValues(newRows);
    Logger.log(newRows.length + '件の新規ブックマークを追記');
  } else {
    Logger.log('新規ブックマークなし（重複スキップ）');
  }
}

function getUserId() {
  const token = getAccessToken_();
  const res = UrlFetchApp.fetch('https://api.twitter.com/2/users/me', {
    headers: { Authorization: 'Bearer ' + token }
  });
  Logger.log(res.getContentText());
}

function doGet(e) {
  const status = e.parameter.oauth;
  if (status === 'success') {
    return HtmlService.createHtmlOutput(
      '<h2>✅ X認証が完了しました</h2>' +
      '<p>このタブを閉じてGASエディタに戻り、fetchAndStoreBookmarks() を実行してください。</p>'
    );
  }
  return HtmlService.createHtmlOutput('<h2>GAS WebApp</h2><p>パラメータなし</p>');
}
