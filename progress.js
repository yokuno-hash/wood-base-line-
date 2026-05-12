// ==========================================
// 進捗管理表 LINE自然文クエリ
// ==========================================
// シート名: 「進捗管理表」
// 想定列: 店舗名 | 施工月 | 区分 | 担当者 | 依頼日 | 図面チェック | 発注 | 見積 | ...
// 実装方針:
//   - ユーザー質問文から対象列・月・店舗・担当者を推定
//   - ヘッダー文字列はファジーマッチ（部分一致＋共通部分文字列）
//   - ○系の値は完了。空欄/△/確認中/×/未/保留/対応中/差戻し/要確認 は通知対象
// ==========================================

var PROGRESS_SHEET_NAME = '進捗管理表';

// 進捗管理表シートを取得（別スプシ対応）
// スクリプトプロパティ PROGRESS_SPREADSHEET_ID が設定されていればそのスプシを開く。
// 未設定なら既存のメインスプシ（SPREADSHEET_ID）を使う。
// シート名は PROGRESS_SHEET_NAME_OVERRIDE があればそちら、なければ PROGRESS_SHEET_NAME。
function getProgressSheet() {
  var props      = PropertiesService.getScriptProperties();
  var extId      = props.getProperty('PROGRESS_SPREADSHEET_ID');
  var sheetName  = props.getProperty('PROGRESS_SHEET_NAME_OVERRIDE') || PROGRESS_SHEET_NAME;
  try {
    var ss = extId ? SpreadsheetApp.openById(extId) : getSS();
    return ss.getSheetByName(sheetName);
  } catch (err) {
    console.error('getProgressSheet error:', err.message);
    return null;
  }
}

// 完了とみなす値（trimして照合）
var PROGRESS_DONE_VALUES = ['○', '〇', '丸', '済', '済み', '完了', 'ok', 'OK', 'done', 'Done', '✓', '✔'];

// 完了判定
function isProgressDone(value) {
  var v = String(value == null ? '' : value).trim();
  if (!v) return false;
  var lc = v.toLowerCase();
  for (var i = 0; i < PROGRESS_DONE_VALUES.length; i++) {
    var d = PROGRESS_DONE_VALUES[i];
    if (v === d || lc === String(d).toLowerCase()) return true;
  }
  return false;
}

function progressStatusLabel(value) {
  var v = String(value == null ? '' : value).trim();
  return v ? v : '未入力';
}

function normalizeHeader(s) {
  // NFKC で半角カナ→全角カナ・全角英数→半角英数を正規化
  return String(s || '')
    .normalize('NFKC')
    .replace(/[\s　・/／\\\-_（）()\[\]【】]/g, '')
    .toLowerCase();
}

// 複数の候補名を順に試して見つかったヘッダーインデックスを返す
function findMatchingHeaderMulti(headers, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var idx = findMatchingHeader(headers, candidates[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

// 質問文中にフレーズが含まれるか（NFKC正規化して比較）
function textContainsPhrase(text, phrase) {
  var t = String(text || '').normalize('NFKC');
  var p = String(phrase || '').normalize('NFKC');
  return p && t.indexOf(p) !== -1;
}

// ヘッダー配列から keyword に最も近い列インデックスを返す（見つからなければ-1）
function findMatchingHeader(headers, keyword) {
  if (!keyword || !headers || !headers.length) return -1;
  var k = normalizeHeader(keyword);
  if (!k) return -1;
  var best = -1, bestScore = 0;
  for (var i = 0; i < headers.length; i++) {
    var h = normalizeHeader(headers[i]);
    if (!h) continue;
    var score = 0;
    if (h === k)                              score = 100;
    else if (h.includes(k) || k.includes(h))  score = 80;
    else {
      var sub = longestCommonSubstring(h, k);
      if (sub >= 3)      score = 55 + sub * 3;
      else if (sub >= 2) score = 30;
    }
    if (score > bestScore) { bestScore = score; best = i; }
  }
  return bestScore >= 50 ? best : -1;
}

// 質問文 → 対象列・フィルタ条件を抽出
function analyzeProgressQuestion(text, headers) {
  var analysis = {
    targetColumns: [],    // 注目するチェック列名（複数可）
    monthFilter: null,    // 5 → 5月
    storeFilter: null,    // 「○○店」
    assigneeFilter: null, // 「田中」
    statusFilter: null,   // 「確認中」など特定状態
    mode: 'specific',     // 'specific' or 'all'
  };
  var t = String(text || '');

  // 月抽出
  var mm = t.match(/(\d{1,2})月/);
  if (mm) analysis.monthFilter = parseInt(mm[1], 10);
  if (/今月/.test(t)) analysis.monthFilter = new Date().getMonth() + 1;
  if (/来月/.test(t)) {
    var nm = new Date(); nm.setMonth(nm.getMonth() + 1);
    analysis.monthFilter = nm.getMonth() + 1;
  }

  // 特定状態キーワード
  var statusKeywords = ['確認中', '対応中', '保留', '差戻し', '差し戻し', '要確認', '△'];
  for (var si = 0; si < statusKeywords.length; si++) {
    if (t.indexOf(statusKeywords[si]) !== -1 && /(だけ|のみ|ある|教え|表示|どこ)/.test(t)) {
      analysis.statusFilter = statusKeywords[si];
      break;
    }
  }

  // 担当者「〇〇さん」
  var am = t.match(/([一-龥ぁ-んァ-ヶーA-Za-z]{1,8})さん/);
  if (am) analysis.assigneeFilter = am[1];

  // 店舗「〇〇店」（区分の「美容/理容」だけは除外）
  var sm = t.match(/([一-龥ぁ-んァ-ヶーA-Za-z0-9]{1,12})店/);
  if (sm && sm[1] && !/^(美容|理容)$/.test(sm[1])) {
    analysis.storeFilter = sm[1];
  }

  // 全体モード
  if (/全部|全て|まとめて|一覧/.test(t)) analysis.mode = 'all';

  // 対象列推定：フレーズ→候補ヘッダー（複数可）。最初に見つかったヘッダーを採用。
  // 「発注」「納品」「見積」など複数列に対応する語は、複数列を同時に対象にする。
  var phraseMap = [
    { phrases: ['図面チェック', '図面確認', '図面'],
      headers: ['図面チェック'], multi: false },
    { phrases: ['依頼日', '依頼した日', '依頼の日', '依頼'],
      headers: ['依頼日'], multi: false },
    { phrases: ['制作一覧', '一覧入力', '一覧'],
      headers: ['制作一覧入力'], multi: false },
    { phrases: ['安陳見積もり', '安陳見積', '安陳の見積'],
      headers: ['安陳見積もり', '安陳見積'], multi: false },
    { phrases: ['安陳入力', '安陳'],
      headers: ['安陳入力'], multi: false },
    { phrases: ['スクショ見積', 'スクショ'],
      headers: ['スクショ見積'], multi: false },
    { phrases: ['金額調整', '金額確認', '金額'],
      headers: ['金額調整確認'], multi: false },
    { phrases: ['パイオニア見積', 'パイオニア', 'ﾊﾟｲｵﾆｱ'],
      headers: ['ﾊﾟｲｵﾆｱ見積', 'パイオニア見積'], multi: false },
    { phrases: ['発注書'],
      headers: ['発注書'], multi: false },
    { phrases: ['制作発注'],
      headers: ['制作発注'], multi: false },
    { phrases: ['発注'],
      headers: ['制作発注', '発注書'], multi: true },
    { phrases: ['納品日時', '納品日', '納品時間'],
      headers: ['納品日時'], multi: false },
    { phrases: ['納品場所'],
      headers: ['納品場所'], multi: false },
    { phrases: ['納品'],
      headers: ['納品日時', '納品場所'], multi: true },
    { phrases: ['完工日', '完工'],
      headers: ['完工日'], multi: false },
    { phrases: ['見積もり', '見積り', '見積'],
      headers: ['見積'], multi: false },
    { phrases: ['搬入'],
      headers: ['搬入'], multi: false },
    { phrases: ['設置'],
      headers: ['設置'], multi: false },
    { phrases: ['担当者', '担当'],
      headers: ['担当者'], multi: false },
    { phrases: ['区分', '美容', '理容', '理/美', '美/理'],
      headers: ['区分', '理/美', '美/理'], multi: false },
  ];

  var picked = {}; // 重複防止
  for (var pi = 0; pi < phraseMap.length; pi++) {
    var entry = phraseMap[pi];
    var hit = false;
    for (var pj = 0; pj < entry.phrases.length; pj++) {
      if (textContainsPhrase(t, entry.phrases[pj])) { hit = true; break; }
    }
    if (!hit) continue;

    if (entry.multi) {
      // 複数列同時に対象（例：発注 → 制作発注 + 発注書）
      entry.headers.forEach(function(hn) {
        var ix = findMatchingHeader(headers, hn);
        if (ix !== -1 && !picked[headers[ix]]) { picked[headers[ix]] = true; analysis.targetColumns.push(headers[ix]); }
      });
    } else {
      var ix = findMatchingHeaderMulti(headers, entry.headers);
      if (ix !== -1 && !picked[headers[ix]]) { picked[headers[ix]] = true; analysis.targetColumns.push(headers[ix]); }
    }
  }

  return analysis;
}

// 行抽出
// rows: シートの全データ行（ヘッダー除く）
// headers: ヘッダー文字列配列
function extractProgressItemsByQuestion(rows, headers, analysis) {
  var idxStore   = findMatchingHeaderMulti(headers, ['店舗名', '店舗', '店名']);
  var idxMonth   = findMatchingHeader(headers, '施工月');
  var idxKubun   = findMatchingHeaderMulti(headers, ['区分', '理/美', '美/理', '理美', '美理']);
  var idxAssign  = findMatchingHeaderMulti(headers, ['担当者', '担当']);
  var idxReqDate = findMatchingHeader(headers, '依頼日');

  // チェック列の範囲：図面チェック以降。なければ依頼日の次以降。それもなければ全列。
  var checkStart = findMatchingHeader(headers, '図面チェック');
  if (checkStart === -1 && idxReqDate !== -1) checkStart = idxReqDate + 1;
  if (checkStart === -1) checkStart = 0;

  // 日付系列は「空欄のみ通知」（値が入っていれば完了扱い）
  var dateColIdxs = {};
  ['依頼日', '納品日時', '完工日'].forEach(function(n) {
    var i = findMatchingHeader(headers, n);
    if (i !== -1) dateColIdxs[i] = true;
  });

  var results = [];

  rows.forEach(function(row) {
    var store = idxStore !== -1 ? String(row[idxStore] || '').trim() : '';
    if (!store) return;

    // 月フィルタ（Dateオブジェクト・「5月」「5/1」「2025/5/1」いずれも対応）
    if (analysis.monthFilter !== null && idxMonth !== -1) {
      var rawMonth = row[idxMonth];
      var monthVal = null;
      if (rawMonth instanceof Date) {
        monthVal = rawMonth.getMonth() + 1;
      } else {
        var mvStr = String(rawMonth == null ? '' : rawMonth);
        var monthMatch = mvStr.match(/(\d{1,2})月/) || mvStr.match(/\/(\d{1,2})\b/) || mvStr.match(/^(\d{1,2})$/);
        if (monthMatch) monthVal = parseInt(monthMatch[1], 10);
      }
      if (monthVal !== analysis.monthFilter) return;
    }
    // 店舗フィルタ
    if (analysis.storeFilter && store.indexOf(analysis.storeFilter) === -1) return;
    // 担当者フィルタ
    if (analysis.assigneeFilter && idxAssign !== -1) {
      if (String(row[idxAssign] || '').indexOf(analysis.assigneeFilter) === -1) return;
    }

    function fmtCell(v) {
      if (v == null) return '';
      if (v instanceof Date) {
        return Utilities.formatDate(v, 'Asia/Tokyo', 'M/d');
      }
      return String(v).trim();
    }
    function fmtMonthCell(v) {
      if (v == null) return '';
      if (v instanceof Date) return (v.getMonth() + 1) + '月';
      var s = String(v).trim();
      if (/^\d{1,2}$/.test(s)) return s + '月';
      return s;
    }
    var item = {
      store:       store,
      month:       idxMonth   !== -1 ? fmtMonthCell(row[idxMonth])      : '',
      kubun:       idxKubun   !== -1 ? fmtCell(row[idxKubun])            : '',
      assignee:    idxAssign  !== -1 ? fmtCell(row[idxAssign])           : '',
      requestDate: idxReqDate !== -1 ? fmtCell(row[idxReqDate])          : '',
      pending: [],
    };

    function pushPending(colIdx, colName) {
      var v = row[colIdx];
      // 日付列：値があれば完了、空欄なら未入力として通知
      if (dateColIdxs[colIdx]) {
        var hasVal = (v instanceof Date) || String(v == null ? '' : v).trim() !== '';
        if (!hasVal && !analysis.statusFilter) {
          item.pending.push({ column: colName, value: '未入力' });
        }
        return;
      }
      if (isProgressDone(v)) return;
      var label = progressStatusLabel(v);
      if (analysis.statusFilter && label !== analysis.statusFilter) return;
      item.pending.push({ column: colName, value: label });
    }

    if (analysis.targetColumns.length) {
      analysis.targetColumns.forEach(function(colName) {
        var ci = headers.indexOf(colName);
        if (ci !== -1) pushPending(ci, colName);
      });
    } else {
      // 全チェック列を見る（属性列は除外）
      var skipCols = {};
      [idxStore, idxMonth, idxKubun, idxAssign, idxReqDate].forEach(function(i) { if (i !== -1) skipCols[i] = true; });
      for (var ci2 = checkStart; ci2 < headers.length; ci2++) {
        if (skipCols[ci2]) continue;
        pushPending(ci2, headers[ci2]);
      }
    }

    if (item.pending.length) results.push(item);
  });

  return results;
}

// 回答テキスト生成
function formatProgressQuestionAnswer(results, analysis) {
  if (!results.length) return '該当する未完了・要確認の店舗はありません。';

  var headerLine = '';
  if (analysis.statusFilter) {
    headerLine = '【状態「' + analysis.statusFilter + '」の項目】';
  } else if (analysis.targetColumns.length === 1) {
    headerLine = '【' + analysis.targetColumns[0] + '：未完了・要確認】';
  } else if (analysis.monthFilter !== null) {
    headerLine = '【' + analysis.monthFilter + '月施工分：未完了・要確認】';
  } else {
    headerLine = '【未完了・要確認】';
  }

  var MAX = 20;
  var lines = [headerLine];
  results.slice(0, MAX).forEach(function(it) {
    lines.push('');
    lines.push('■ ' + it.store);
    if (it.month)       lines.push('施工月：' + it.month);
    if (it.kubun)       lines.push('区分：' + it.kubun);
    if (it.assignee)    lines.push('担当：' + it.assignee);
    if (it.requestDate && analysis.targetColumns.indexOf('依頼日') === -1) lines.push('依頼日：' + it.requestDate);

    if (analysis.targetColumns.length === 1 && it.pending.length === 1) {
      lines.push('状態：' + it.pending[0].value);
    } else if (it.pending.length === 1 && analysis.targetColumns.length === 0 && analysis.statusFilter) {
      lines.push(it.pending[0].column + '：' + it.pending[0].value);
    } else {
      lines.push('未完了：');
      it.pending.forEach(function(p) { lines.push('・' + p.column + '：' + p.value); });
    }
  });
  if (results.length > MAX) lines.push('\n…ほか ' + (results.length - MAX) + ' 件');
  return lines.join('\n');
}

// 進捗質問エントリ：質問でなければnull、質問なら回答テキスト返却
function answerProgressQuestion(text) {
  var sheet = getProgressSheet();
  if (!sheet || sheet.getLastRow() < 2) return null;

  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h == null ? '' : h).trim(); });
  var rows    = data.slice(1);

  var analysis = analyzeProgressQuestion(text, headers);

  // トリガー判定：列マッチ・月・店舗・担当者・状態のいずれか or 明示キーワード
  var triggered =
       analysis.targetColumns.length > 0
    || analysis.monthFilter   !== null
    || analysis.storeFilter   !== null
    || analysis.assigneeFilter!== null
    || analysis.statusFilter  !== null
    || /未完了|まとめて|一覧|全部|全て|進捗管理|残ってる|残ってない|○ついてない|まる(つい|入っ)てない/.test(text);

  if (!triggered) return null;

  var results = extractProgressItemsByQuestion(rows, headers, analysis);
  return formatProgressQuestionAnswer(results, analysis);
}

// LINE event 入口（handleMentionCommand / handleDM から呼ぶ）
// 返り値: true=回答済み, false=対象外
function handleProgressQuestionRequest(event, messageText) {
  try {
    var answer = answerProgressQuestion(messageText);
    if (!answer) return false;
    sendLineReply(event.replyToken, answer);
    return true;
  } catch (err) {
    console.error('handleProgressQuestionRequest error:', err.message);
    return false;
  }
}

// ==========================================
// テスト
// ==========================================
function testProgressQuestion() {
  var sheet = getProgressSheet();
  if (!sheet) { console.error('進捗管理表シートが見つかりません（PROGRESS_SPREADSHEET_ID / シート名を確認）'); return; }
  console.log('対象スプシ:', sheet.getParent().getName(), '/ シート:', sheet.getName());
  var cases = [
    '図面チェックできてない店舗ある？',
    '依頼日入ってない店舗ある？',
    '5月の図面チェックまだのところ教えて',
    '確認中の店舗だけ教えて',
    '6月施工で未完了ある？',
    '田中さん担当で未完了ある？',
    '未完了全部教えて',
  ];
  cases.forEach(function(q) {
    console.log('Q: ' + q);
    var ans = answerProgressQuestion(q);
    console.log(ans || '（対象外と判定）');
    console.log('---');
  });
}

// 別スプシを進捗管理表として登録（GASエディタから1回実行）
// 例: setupProgressSpreadsheet('1zzA2qSoKZoTBp81BvH4Vl36TdJUuy1F4prjzcKUX-uE', '進捗管理表');
function setupProgressSpreadsheet(spreadsheetId, sheetName) {
  if (!spreadsheetId) { console.error('spreadsheetId が必要です'); return; }
  var props = PropertiesService.getScriptProperties();
  props.setProperty('PROGRESS_SPREADSHEET_ID', String(spreadsheetId));
  if (sheetName) props.setProperty('PROGRESS_SHEET_NAME_OVERRIDE', String(sheetName));

  // 動作確認：開けるか・対象シートが存在するか
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var target = sheetName || PROGRESS_SHEET_NAME;
    var sh = ss.getSheetByName(target);
    if (!sh) {
      console.error('スプシは開けましたが「' + target + '」シートが見つかりません。シート名を確認してください。');
      console.log('利用可能シート:', ss.getSheets().map(function(s){ return s.getName(); }).join(', '));
      return;
    }
    console.log('✅ 設定完了:', ss.getName(), '/', sh.getName(), '/ 行数:', sh.getLastRow());
  } catch (err) {
    console.error('開けませんでした。共有設定（GAS実行アカウントに閲覧権限）と ID を確認してください。', err.message);
  }
}

// 進捗管理表の現在の接続先を確認
function showProgressSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  console.log('PROGRESS_SPREADSHEET_ID:', props.getProperty('PROGRESS_SPREADSHEET_ID') || '(未設定 → メインスプシを使用)');
  console.log('PROGRESS_SHEET_NAME_OVERRIDE:', props.getProperty('PROGRESS_SHEET_NAME_OVERRIDE') || '(未設定 → ' + PROGRESS_SHEET_NAME + ')');
  var sh = getProgressSheet();
  if (sh) console.log('解決後:', sh.getParent().getName(), '/', sh.getName(), '/ 行数:', sh.getLastRow());
  else    console.log('シート解決失敗');
}

// ヘッダーマッチのみ単体テスト（シート不要）
function testFindMatchingHeader() {
  var headers = ['店舗名', '施工月', '美容/理容区分', '担当者', '依頼日', '図面チェック', '発注', '見積', '納品'];
  var cases = [
    ['図面',     '図面チェック'],
    ['図面確認', '図面チェック'],
    ['担当',     '担当者'],
    ['依頼した日', '依頼日'],
    ['美容',     '美容/理容区分'],
    ['区分',     '美容/理容区分'],
    ['発注',     '発注'],
    ['存在しない列', null],
  ];
  cases.forEach(function(c) {
    var idx = findMatchingHeader(headers, c[0]);
    var got = idx === -1 ? null : headers[idx];
    console.log((got === c[1] ? '✅' : '❌') + ' "' + c[0] + '" → ' + got + ' (期待:' + c[1] + ')');
  });
}
