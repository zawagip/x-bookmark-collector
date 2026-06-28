const WORKERS_BASE = 'https://ryulink.zawagipyask.workers.dev';
const SHEET_ID = '1y5l03kvHT74qr3jmNDv2AVsS-6hrZpNwvHctw8q76JU';
const RAW_SHEET_NAME = 'raw_data';
const WORKS_SHEET_NAME = 'works';

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
      const fanzaUrls = extractFanzaUrls_(t);
      allBookmarks.push({
        tweetId: t.id,
        authorId: t.author_id,
        username: user.username || '',
        text: t.text,
        createdAt: t.created_at,
        likeCount: (t.public_metrics && t.public_metrics.like_count) || 0,
        retweetCount: (t.public_metrics && t.public_metrics.retweet_count) || 0,
        impressionCount: (t.public_metrics && t.public_metrics.impression_count) || 0,
        conversationId: t.conversation_id,
        fanzaUrls: fanzaUrls,
        cid: extractCid_(fanzaUrls),
      });
    });

    nextToken = (json.meta && json.meta.next_token) || null;
    Logger.log('page ' + page + ': ' + tweets.length + '件取得 (累計 ' + allBookmarks.length + '件)');

    if (nextToken) Utilities.sleep(1000);
  } while (nextToken && page < 50);

  writeToRawSheet_(allBookmarks);
  writeToWorksSheet_(allBookmarks);
  enrichWorksWithFanza();
  Logger.log('完了: 合計 ' + allBookmarks.length + '件をスプシに書き込み');
}

// FANZA APIでworksシートのタイトル・サムネ・価格を補完
function enrichWorksWithFanza() {
  const props = PropertiesService.getScriptProperties();
  const apiId = props.getProperty('FANZA_API_ID');
  const affiliateId = props.getProperty('FANZA_AFFILIATE_ID');
  if (!apiId || !affiliateId) throw new Error('FANZA_API_ID または FANZA_AFFILIATE_ID が未設定');

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(WORKS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  // タイトルが空の行のみ処理
  const emptyRows = [];
  data.forEach(function(row, i) {
    if (!row[1]) { // タイトル列が空
      emptyRows.push({ rowIndex: i + 2, cid: String(row[0]) });
    }
  });

  if (emptyRows.length === 0) {
    Logger.log('works: FANZA補完対象なし');
    return;
  }

  Logger.log('works: FANZA補完対象 ' + emptyRows.length + '件');

  emptyRows.forEach(function(item) {
    const cid = item.cid;
    if (!cid) return;

    // floorを判定（d_で始まるならdoujin、それ以外はebookコミック）
    const isDoujin = cid.match(/^d_/);
    const service = isDoujin ? 'doujin' : 'ebook';
    const floorParam = isDoujin ? 'digital_doujin' : 'comic';

    const url = 'https://api.dmm.com/affiliate/v3/ItemList'
      + '?api_id=' + apiId
      + '&affiliate_id=' + affiliateId
      + '&site=FANZA'
      + '&service=' + service
      + '&floor=' + floorParam
      + '&cid=' + cid
      + '&output=json';

    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) {
        Logger.log('FANZA API失敗 cid=' + cid + ' status=' + res.getResponseCode());
        return;
      }

      const json = JSON.parse(res.getContentText());
      const items = json.result && json.result.items;
      if (!items || items.length === 0) {
        Logger.log('FANZA API: 作品なし cid=' + cid);
        return;
      }

      const work = items[0];
      const title = work.title || '';
      const thumb = (work.imageURL && work.imageURL.large) || '';
      const price = (work.prices && work.prices.price) || '';

      sheet.getRange(item.rowIndex, 2).setValue(title);
      sheet.getRange(item.rowIndex, 3).setValue(thumb);
      sheet.getRange(item.rowIndex, 4).setValue(price);

      Logger.log('補完完了: ' + cid + ' / ' + title);
      Utilities.sleep(500); // API負荷対策

    } catch(e) {
      Logger.log('FANZA API例外 cid=' + cid + ': ' + e.message);
    }
  });
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
    .filter(function(u) { return /fanza\.jp|dmm\.co\.jp|al\.fanza\.co\.jp|al\.dmm\.co\.jp/.test(u); })
    .join(',');
}

function extractCid_(fanzaUrls) {
  if (!fanzaUrls) return '';

  const urls = fanzaUrls.split(',');
  const cids = [];

  urls.forEach(function(url) {
    if (!url) return;
    const decoded = decodeURIComponent(url);

    const cidMatch = decoded.match(/[?&\/]cid[=\/]([a-z0-9_]+)/i);
    if (cidMatch) {
      cids.push(cidMatch[1]);
      return;
    }

    const productMatch = decoded.match(/\/product\/([a-z0-9]+)\//i);
    if (productMatch) {
      cids.push(productMatch[1]);
      return;
    }
  });

  return cids.filter(function(v, i, a) { return a.indexOf(v) === i; }).join(',');
}

function writeToRawSheet_(bookmarks) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(RAW_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(RAW_SHEET_NAME);

  const headers = [
    'ツイートID', '著者ID', 'ユーザー名', 'ツイート本文',
    '投稿日時', 'いいね数', 'リツイート数', 'インプレッション数',
    '会話ID', 'FANZA URL', '取得日時', 'CID',
  ];

  if (sheet.getLastRow() === 0) sheet.appendRow(headers);

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
        b.createdAt, b.likeCount, b.retweetCount, b.impressionCount,
        b.conversationId, b.fanzaUrls, now, b.cid,
      ];
    });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
    Logger.log('raw_data: ' + newRows.length + '件追記');
  } else {
    Logger.log('raw_data: 新規なし');
  }
}

function writeToWorksSheet_(bookmarks) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(WORKS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(WORKS_SHEET_NAME);

  const headers = [
    'CID', 'タイトル', 'サムネURL', '価格',
    '元ツイートID', 'いいね数', 'RT数', 'インプレッション数', '初回取得日時',
  ];

  if (sheet.getLastRow() === 0) sheet.appendRow(headers);

  const existingCids = new Set();
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
      .getValues()
      .forEach(function(r) { existingCids.add(String(r[0])); });
  }

  const now = new Date().toISOString();
  const newRows = [];

  bookmarks.forEach(function(b) {
    if (!b.cid) return;
    b.cid.split(',').forEach(function(cid) {
      cid = cid.trim();
      if (!cid || existingCids.has(cid)) return;
      existingCids.add(cid);
      newRows.push([
        cid, '', '', '',
        b.tweetId, b.likeCount, b.retweetCount, b.impressionCount, now,
      ]);
    });
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
    Logger.log('works: ' + newRows.length + '件追記');
  } else {
    Logger.log('works: 新規なし');
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
