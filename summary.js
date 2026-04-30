// ==========================================
// WOODBASE 案件サマリー自動生成システム
// コード.js の getConfig() / getSheet() / callGemini() を共有
// ==========================================

// ===== 定数 =====
var DOC_SUMMARY_SECTION = '■ 最新サマリー（自動更新）';
var DOC_DIFF_SECTION    = '■ 差分（前回からの変化）';
var DOC_HISTORY_SECTION = '■ 履歴';
var DOC_SEPARATOR       = '──────────────────────';
var DOC_SUBFOLDER_NAME  = '案件サマリー';
var MAX_MESSAGES        = 30;  // Geminiに渡すメッセージ最大件数
var MAX_TASKS           = 20;
var MAX_EVENTS          = 10;
var MAX_HISTORY_DAYS    = 30;  // 履歴保持日数

// ==========================================
// SECTION 1: エントリーポイント
// ==========================================

// 指定案件のサマリーを生成してDocsに反映（LINEメンション・手動実行用）
function generateAndUpdateSummary(projectName) {
  if (!projectName) { console.error('案件名が必要です'); return null; }

  console.log('サマリー生成開始:', projectName);
  var summary = generateProjectSummary(projectName);
  if (!summary) { console.error('サマリー生成失敗:', projectName); return null; }

  updateProjectDoc(projectName, summary);
  console.log('Docs更新完了:', projectName);
  return summary;
}

// 全進行中案件を一括更新（日次トリガー用：毎朝7時）
function runDailySummaryUpdate() {
  var projects = getActiveProjects(50);
  if (!projects.length) { console.log('進行中の案件なし'); return; }

  projects.forEach(function(name) {
    try {
      generateAndUpdateSummary(name);
      Utilities.sleep(1500); // API制限対策
    } catch (err) {
      console.error('サマリー更新エラー[' + name + ']:', err.message);
    }
  });

  console.log('全案件サマリー更新完了:', projects.length + '件');
}

// ==========================================
// SECTION 2: サマリー生成（Gemini）
// ==========================================
function generateProjectSummary(projectName) {
  var messages = getProjectMessages(projectName);
  var tasks    = getProjectTasks(projectName);
  var events   = getProjectEvents(projectName);

  if (!messages.length && !tasks.length && !events.length) {
    return '（登録データなし）';
  }

  var config    = getConfig();
  var today     = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日');
  var dataBlock = buildDataBlock(messages, tasks, events);

  var prompt =
    'あなたは建築会社WOODBASEの専属秘書AIです。\n' +
    '今日：' + today + '\n\n' +
    '以下のデータをもとに「' + projectName + '」の案件サマリーを生成してください。\n\n' +
    '【厳守ルール】\n' +
    '・箇条書きのみ（前置き・結論文・説明文は出力しない）\n' +
    '・推測禁止（データに明記されていない情報は書かない）\n' +
    '・以下は完全に無視する：挨拶・雑談・感想・意味のない短文（例：「了解」「OK」「はい」など）\n' +
    '・各セクション最大5項目。該当データがなければそのセクションを省略\n' +
    '・「なし」「不明」などの補完記述は禁止\n' +
    '・古い情報より直近の情報を優先する\n' +
    '・状況が変わっている場合は最新を採用する\n' +
    '・タスクは以下の優先順で表示する：①期限が近い ②未完了 ③他人に影響する\n\n' +
    '【出力フォーマット（このまま出力・変更禁止）】\n' +
    '▼ 案件概要\n' +
    '・（案件の基本情報・目的・規模など）\n\n' +
    '▼ 現在の状況\n' +
    '・（最新の進捗・完了事項など）\n\n' +
    '▼ 進行中のタスク\n' +
    '・（未完了タスク・担当者・期日）\n\n' +
    '▼ 課題・リスク\n' +
    '・（問題・懸念事項・未解決事項）\n\n' +
    '▼ 次のアクション\n' +
    '・（直近でやるべきこと・期限）\n\n' +
    '【データ】\n' + dataBlock;

  var result = callGemini(config.GEMINI_API_KEY, prompt, 0.15);

  // Gemini失敗時はルールベースフォールバック
  return result || buildFallbackSummary(tasks, events);
}

// Geminiに渡すデータブロックを構築
function buildDataBlock(messages, tasks, events) {
  var lines = [];

  if (messages.length) {
    lines.push('--- メッセージ（直近' + messages.length + '件） ---');
    messages.forEach(function(m) {
      lines.push('[' + m[0] + '] ' + m[2] + ': ' + String(m[3]).slice(0, 200));
    });
    lines.push('');
  }

  if (tasks.length) {
    lines.push('--- タスク ---');
    tasks.forEach(function(t) {
      var dl  = t[4] ? String(t[4]).slice(0, 10) : '期日未定';
      var st  = t[5] || '不明';
      lines.push('・' + t[2] + '（担当:' + t[3] + ' / 期日:' + dl + ' / ' + st + '）');
    });
    lines.push('');
  }

  if (events.length) {
    lines.push('--- 予定 ---');
    events.forEach(function(e) {
      var dt  = e[3] ? String(e[3]).slice(0, 10) : '';
      var tm  = e[4] ? ' ' + e[4] : '';
      var loc = e[6] ? '（' + e[6] + '）' : '';
      lines.push('・' + dt + tm + ' ' + e[2] + loc);
    });
    lines.push('');
  }

  return lines.join('\n');
}

// Gemini失敗時のルールベースサマリー
function buildFallbackSummary(tasks, events) {
  var lines = ['（Gemini生成失敗 - 簡易サマリー）\n'];

  var activeTasks = tasks.filter(function(t) { return t[5] !== 'done' && t[5] !== '完了'; });
  if (activeTasks.length) {
    lines.push('▼ 進行中のタスク');
    activeTasks.slice(0, 5).forEach(function(t) {
      var dl = t[4] ? String(t[4]).slice(0, 10) : '期日未定';
      lines.push('・' + t[2] + '（' + t[3] + ' / ' + dl + '）');
    });
  }

  if (events.length) {
    lines.push('\n▼ 直近の予定');
    events.slice(0, 5).forEach(function(e) {
      lines.push('・' + String(e[3]).slice(0, 10) + ' ' + e[2]);
    });
  }

  return lines.join('\n') || '（データ不足）';
}

// ==========================================
// SECTION 3: データ取得
// ==========================================

// メッセージログシート列: [0]=日時, [1]=グループID, [2]=送信者, [3]=メッセージ
function getProjectMessages(projectName) {
  var sheet = getSheet('メッセージログ');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues().slice(1);

  return data
    .filter(function(r) {
      var text = String(r[3] || '');
      // 雑談・短文・定型返事を除外
      if (text.length < 8) return false;
      var NOISE = ['了解', 'はい', 'ありがとう', 'お疲れ', 'よろしく', 'おk', 'OK'];
      if (text.length <= 12 && NOISE.some(function(w) { return text.includes(w); })) return false;
      return isRelatedToProject(text, projectName);
    })
    // 長文優先（内容が濃いメッセージを優先）
    .sort(function(a, b) { return String(b[3]).length - String(a[3]).length; })
    .slice(0, MAX_MESSAGES);
}

// タスク管理シート列: [0]=task_id, [1]=案件名, [2]=内容, [3]=担当者, [4]=期日, [5]=ステータス, [6]=作成日時, [7]=group_id
function getProjectTasks(projectName) {
  var sheet = getSheet('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  return sheet.getDataRange().getValues().slice(1)
    .filter(function(r) { return String(r[1] || '') === projectName && r[2]; })
    .slice(0, MAX_TASKS);
}

// スケジュール管理シート列: [0]=登録日時, [1]=案件名, [2]=予定タイトル, [3]=日付, [4]=開始時間, ..., [6]=場所, [7]=参加者
function getProjectEvents(projectName) {
  var sheet = getSheet('スケジュール管理');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  // 過去30日以内 + 未来の予定を取得
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - MAX_HISTORY_DAYS);
  var cutoffStr = fmtDate(cutoff);

  return sheet.getDataRange().getValues().slice(1)
    .filter(function(r) {
      if (String(r[1] || '') !== projectName || !r[3]) return false;
      var d = r[3] instanceof Date ? fmtDate(r[3]) : String(r[3]).slice(0, 10);
      return d >= cutoffStr;
    })
    .sort(function(a, b) {
      var da = a[3] instanceof Date ? fmtDate(a[3]) : String(a[3]).slice(0, 10);
      var db = b[3] instanceof Date ? fmtDate(b[3]) : String(b[3]).slice(0, 10);
      return da > db ? -1 : 1;
    })
    .slice(0, MAX_EVENTS);
}

// メッセージが案件に関連するか（案件名・略称の一致チェック）
function isRelatedToProject(text, projectName) {
  var clean = text.replace(/\s/g, '');
  var name  = projectName.replace(/\s/g, '');
  if (clean.includes(name)) return true;

  // プロジェクト管理シートの略称も確認
  var data = getProjectData();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][1] || '') === projectName) {
      var abbr = String(data[i][0] || '').replace(/\s/g, '');
      if (abbr.length >= 2 && clean.includes(abbr)) return true;
      break;
    }
  }
  return false;
}

// ==========================================
// SECTION 3.5: 差分サマリー生成
// ==========================================

// Docsの「■ 最新サマリー」セクションからテキストを取得
function getPreviousSummaryFromDoc(projectName) {
  try {
    var folder = getSummaryFolder();
    var title  = '【' + projectName + '】案件サマリー';
    var files  = folder.getFilesByName(title);
    if (!files.hasNext()) return '';

    var body     = DocumentApp.openById(files.next().getId()).getBody();
    var paras    = body.getParagraphs();
    var startIdx = -1;
    var endIdx   = -1;

    for (var i = 0; i < paras.length; i++) {
      var t = paras[i].getText().trim();
      if (t === DOC_SUMMARY_SECTION && startIdx === -1) { startIdx = i; continue; }
      if ((t === DOC_DIFF_SECTION || t === DOC_HISTORY_SECTION) && startIdx !== -1) { endIdx = i; break; }
    }

    if (startIdx === -1) return '';
    if (endIdx   === -1) endIdx = paras.length;

    var lines = [];
    for (var j = startIdx + 1; j < endIdx; j++) {
      lines.push(paras[j].getText());
    }
    return lines.join('\n').trim();
  } catch (e) {
    console.warn('前回サマリー取得失敗:', e.message);
    return '';
  }
}

// 差分サマリーをGeminiで生成
function generateDiffSummary(projectName) {
  var prevSummary = getPreviousSummaryFromDoc(projectName);
  if (!prevSummary) return null; // 前回がなければ差分なし

  var messages = getProjectMessages(projectName);
  var tasks    = getProjectTasks(projectName);
  var config   = getConfig();

  var taskBlock = tasks.map(function(t) {
    var dl = t[4] ? String(t[4]).slice(0, 10) : '期日未定';
    return '・' + t[2] + '（担当:' + t[3] + ' / 期日:' + dl + ' / ' + (t[5] || '不明') + '）';
  }).join('\n');

  var msgBlock = messages.slice(0, 20).map(function(m) {
    return '[' + m[0] + '] ' + m[2] + ': ' + String(m[3]).slice(0, 150);
  }).join('\n');

  var prompt =
    'あなたはプロジェクトマネージャーです。\n' +
    '以下の「前回サマリー」と「最新データ」を厳密に比較し、"変化した事実のみ"を抽出してください。\n\n' +
    '【最重要ルール】\n' +
    '・変化していない内容は絶対に出力しない\n' +
    '・推測・補完は禁止\n' +
    '・必ず「前回 → 今回」の差分として判断する\n' +
    '・同じ内容でも表現違いは"同一"とみなす\n' +
    '・各項目最大3件、なければ「なし」と記載\n\n' +
    '【出力フォーマット（厳守）】\n' +
    '【新規タスク】\n・（新しく追加されたもの）\n\n' +
    '【完了タスク】\n・（完了したもの）\n\n' +
    '【状態変化】\n・何が → どう変わったか\n\n' +
    '【新規課題】\n・問題内容\n\n' +
    '【アクション変化】\n・変更前 → 変更後\n\n' +
    '【前回サマリー】\n' + prevSummary + '\n\n' +
    '【最新メッセージ（時系列）】\n' + (msgBlock || '（なし）') + '\n\n' +
    '【最新タスク一覧】\n' + (taskBlock || '（なし）');

  return callGemini(config.GEMINI_API_KEY, prompt, 0.1) || null;
}

// Docsの「■ 差分」セクションを更新
function updateDiffSection(body, diff) {
  var paras    = body.getParagraphs();
  var startIdx = -1;
  var endIdx   = -1;

  for (var i = 0; i < paras.length; i++) {
    var t = paras[i].getText().trim();
    if (t === DOC_DIFF_SECTION  && startIdx === -1) { startIdx = i; continue; }
    if (t === DOC_HISTORY_SECTION && startIdx !== -1) { endIdx = i; break; }
  }

  if (startIdx === -1) return; // セクションなければスキップ

  if (endIdx === -1) endIdx = paras.length;
  for (var j = endIdx - 1; j > startIdx; j--) {
    body.removeChild(body.getParagraphs()[j]);
  }

  var insertAt = startIdx + 1;
  var lines    = diff.split('\n');
  lines.push('');
  for (var k = lines.length - 1; k >= 0; k--) {
    body.insertParagraph(insertAt, lines[k]);
  }
}

// ==========================================
// SECTION 4: Docs管理
// ==========================================

// 案件名でDocsを取得または新規作成（DRIVE_FOLDER_ID 内のサブフォルダ「案件サマリー」に保存）
function getOrCreateProjectDoc(projectName) {
  var folder = getSummaryFolder();
  var title  = '【' + projectName + '】案件サマリー';

  var files = folder.getFilesByName(title);
  if (files.hasNext()) {
    return DocumentApp.openById(files.next().getId());
  }

  // 新規作成
  var doc  = DocumentApp.create(title);
  var file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch (e) {} // マイドライブから除去（エラー無視）

  initDocStructure(doc.getBody(), projectName);
  doc.saveAndClose();
  console.log('新規Docs作成:', title);

  return DocumentApp.openById(doc.getId());
}

// DRIVE_FOLDER_ID 配下に「案件サマリー」サブフォルダを取得または作成
function getSummaryFolder() {
  var config = getConfig();
  if (!config.DRIVE_FOLDER_ID) throw new Error('DRIVE_FOLDER_IDが設定されていません');
  var parent   = DriveApp.getFolderById(config.DRIVE_FOLDER_ID);
  var existing = parent.getFoldersByName(DOC_SUBFOLDER_NAME);
  return existing.hasNext() ? existing.next() : parent.createFolder(DOC_SUBFOLDER_NAME);
}

// Docsの初期構造（案件名・最新サマリー・履歴セクション）
function initDocStructure(body, projectName) {
  body.clear();

  var h1 = body.appendParagraph('【' + projectName + '】');
  h1.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendParagraph('');

  var h2a = body.appendParagraph(DOC_SUMMARY_SECTION);
  h2a.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('（まだサマリーがありません）');
  body.appendParagraph('');

  var h2d = body.appendParagraph(DOC_DIFF_SECTION);
  h2d.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('（差分はまだありません）');
  body.appendParagraph('');

  var h2b = body.appendParagraph(DOC_HISTORY_SECTION);
  h2b.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('');
}

// Docsを更新（最新サマリー上書き＋履歴追記）
function updateProjectDoc(projectName, summary) {
  var doc;
  try {
    doc = getOrCreateProjectDoc(projectName);
  } catch (err) {
    console.error('Docs取得失敗[' + projectName + ']:', err.message);
    return;
  }

  var body  = doc.getBody();
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  // 差分生成（前回サマリーと比較）
  var diff = generateDiffSummary(projectName);

  replaceLatestSummary(body, summary);
  if (diff) updateDiffSection(body, diff);
  appendOrUpdateHistory(body, today, summary);

  doc.saveAndClose();
  console.log('Docs更新完了:', projectName, today);
}

// ==========================================
// SECTION 5: Docs編集ロジック
// ==========================================

// 「■ 最新サマリー」セクションの内容を毎回上書き
function replaceLatestSummary(body, summary) {
  var paras    = body.getParagraphs();
  var startIdx = -1;
  var endIdx   = -1;

  for (var i = 0; i < paras.length; i++) {
    var t = paras[i].getText().trim();
    if (t === DOC_SUMMARY_SECTION && startIdx === -1) { startIdx = i; continue; }
    if (t === DOC_HISTORY_SECTION  && startIdx !== -1) { endIdx = i; break; }
  }

  if (startIdx === -1) { console.error('最新サマリーセクションが見つかりません'); return; }
  // ■ 履歴 が見つからない場合はドキュメント末尾まで
  if (endIdx === -1) endIdx = paras.length;

  // startIdx+1 〜 endIdx-1 を逆順で削除
  for (var j = endIdx - 1; j > startIdx; j--) {
    body.removeChild(body.getParagraphs()[j]);
  }

  // 新サマリーを ■ 最新サマリー の直後に挿入
  var insertAt = startIdx + 1;
  var lines    = summary.split('\n');
  lines.push(''); // ■ 履歴 との間に空行

  // 逆順で insertParagraph すると正順になる
  for (var k = lines.length - 1; k >= 0; k--) {
    body.insertParagraph(insertAt, lines[k]);
  }
}

// 「■ 履歴」セクションに日付付きで追記（同日なら上書き）
function appendOrUpdateHistory(body, date, summary) {
  var paras   = body.getParagraphs();
  var histIdx = -1;

  for (var i = 0; i < paras.length; i++) {
    if (paras[i].getText().trim() === DOC_HISTORY_SECTION) { histIdx = i; break; }
  }

  if (histIdx === -1) { console.error('履歴セクションが見つかりません'); return; }

  // 同日エントリの検索
  paras = body.getParagraphs(); // 再取得（replaceLatestSummaryで変化している可能性）
  var sameDayIdx  = -1;
  var nextDateIdx = -1;

  for (var j = histIdx + 1; j < paras.length; j++) {
    var t = paras[j].getText().trim();
    if (t === date && sameDayIdx === -1)                             { sameDayIdx = j; continue; }
    if (sameDayIdx !== -1 && /^\d{4}-\d{2}-\d{2}$/.test(t))        { nextDateIdx = j; break; }
  }

  if (sameDayIdx !== -1) {
    // 同日エントリを削除して再挿入
    var deleteEnd = nextDateIdx !== -1 ? nextDateIdx : body.getParagraphs().length;
    for (var k = deleteEnd - 1; k >= sameDayIdx; k--) {
      body.removeChild(body.getParagraphs()[k]);
    }
  }

  // 履歴セクションの直後に挿入（最新が上に来る）
  insertHistoryBlock(body, histIdx + 1, date, summary);

  // 古い履歴の削除（MAX_HISTORY_DAYS日より古いエントリを除去）
  pruneOldHistory(body);
}

// 履歴ブロックを指定位置に挿入（日付 → 内容 → 区切り線）
function insertHistoryBlock(body, insertAt, date, summary) {
  var lines = [date, ''].concat(summary.split('\n')).concat([DOC_SEPARATOR, '']);
  for (var i = lines.length - 1; i >= 0; i--) {
    body.insertParagraph(insertAt, lines[i]);
  }
}

// MAX_HISTORY_DAYS日より古い履歴エントリを削除
function pruneOldHistory(body) {
  var cutoff    = new Date();
  cutoff.setDate(cutoff.getDate() - MAX_HISTORY_DAYS);
  var cutoffStr = fmtDate(cutoff);

  var paras   = body.getParagraphs();
  var histIdx = -1;

  for (var i = 0; i < paras.length; i++) {
    if (paras[i].getText().trim() === DOC_HISTORY_SECTION) { histIdx = i; break; }
  }
  if (histIdx === -1) return;

  // 古い日付以降の行を検出して削除
  var deleteFrom = -1;
  for (var j = histIdx + 1; j < paras.length; j++) {
    var t = paras[j].getText().trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(t) && t < cutoffStr) { deleteFrom = j; break; }
  }

  if (deleteFrom !== -1) {
    for (var k = paras.length - 1; k >= deleteFrom; k--) {
      body.removeChild(body.getParagraphs()[k]);
    }
    console.log('古い履歴を削除:', cutoffStr + '以前');
  }
}

// ==========================================
// SECTION 6: LINEメンション統合
// コード.js の handleMentionCommand から呼ばれる
// ==========================================

// メンションテキストからDocsサマリー更新依頼か判定
function isDocSummaryRequest(text) {
  var docKw  = ['Docs', 'docs', 'ドキュメント', 'サマリー', '資料'];
  var actKw  = ['更新', '作成', 'まとめ', '生成'];
  return docKw.some(function(k) { return text.includes(k); }) &&
         actKw.some(function(k) { return text.includes(k); });
}

// メンションテキストから対象案件を推定してDocsを更新
// handleMentionCommand から呼ぶ
function handleDocSummaryRequest(replyToken, text, groupId, sender) {
  // 案件識別（テキスト + グループID）
  var proj = identifyProject(text, groupId, '');

  if (proj.confidence < 50) {
    // 案件不明 → 案件選択UIを表示
    var projects = getActiveProjects(11);
    var items    = projects.map(function(name) {
      var label = name.length > 20 ? name.slice(0, 19) + '…' : name;
      return { type: 'action', action: { type: 'postback', label: label, data: 'action=doc_summary&p=' + encodeURIComponent(name) } };
    });
    sendQuickReply(replyToken, 'どの案件のDocsを更新しますか？', items);
    return;
  }

  // 案件確定 → 生成・更新
  var summary = generateAndUpdateSummary(proj.name);
  if (summary) {
    sendLineReply(replyToken, '✅ ' + proj.name + ' のDocsを更新しました。');
  } else {
    sendLineReply(replyToken, '❌ Docs更新に失敗しました。データが不足している可能性があります。');
  }
}

// postback action=doc_summary のハンドリング（handlePostback から呼ぶ）
function handleDocSummaryPostback(ev, params) {
  var projectName = decodeURIComponent(params.p || '');
  if (!projectName) { sendLineReply(ev.replyToken, '案件名が取得できませんでした。'); return; }

  var summary = generateAndUpdateSummary(projectName);
  if (summary) {
    sendLineReply(ev.replyToken, '✅ ' + projectName + ' のDocsを更新しました。');
  } else {
    sendLineReply(ev.replyToken, '❌ Docs更新に失敗しました。');
  }
}

// ==========================================
// SECTION 7: トリガー設定
// ==========================================
function setupSummaryTrigger() {
  var fns = ['runDailySummaryUpdate'];
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return fns.indexOf(t.getHandlerFunction()) !== -1; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('runDailySummaryUpdate').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert('完了：毎朝7時に全案件のDocsサマリーを自動更新します。');
}

// ==========================================
// SECTION 8: テスト
// ==========================================

// データ取得件数の確認
function testDataFetch() {
  var projectName = '雨のち晴れクリニック'; // 実在する案件名に変更
  console.log('=== データ取得テスト:', projectName, '===');
  console.log('メッセージ:', getProjectMessages(projectName).length, '件');
  console.log('タスク:',     getProjectTasks(projectName).length,    '件');
  console.log('予定:',       getProjectEvents(projectName).length,   '件');
}

// サマリー生成のみ（Docs書き込みなし）
function testGenerateSummary() {
  var projectName = '雨のち晴れクリニック';
  console.log('=== サマリー生成テスト:', projectName, '===');
  var summary = generateProjectSummary(projectName);
  console.log(summary);
}

// Docsの作成・更新フルテスト
function testUpdateDoc() {
  var projectName = '雨のち晴れクリニック';
  console.log('=== Docs更新テスト:', projectName, '===');
  generateAndUpdateSummary(projectName);
  console.log('完了：DriveのDOCS_SUBFOLDER_NAMEフォルダを確認してください');
}

// フォルダ・Docs接続確認
function testDocsConnection() {
  try {
    var folder = getSummaryFolder();
    console.log('サマリーフォルダ確認OK:', folder.getName(), '/', folder.getId());
  } catch (err) {
    console.error('フォルダ接続失敗:', err.message);
  }
}
