// ==========================================
// WOODBASE 専属秘書AIシステム v2.0
// ==========================================

// ---------- 設定 ----------
function getConfig() {
  const p = PropertiesService.getScriptProperties();
  return {
    LINE_CHANNEL_ACCESS_TOKEN : p.getProperty('LINE_CHANNEL_ACCESS_TOKEN'),
    LINE_CHANNEL_SECRET       : p.getProperty('LINE_CHANNEL_SECRET'),
    GEMINI_API_KEY            : p.getProperty('GEMINI_API_KEY'),
    SPREADSHEET_ID            : p.getProperty('SPREADSHEET_ID'),
    INTERNAL_GROUP_ID         : p.getProperty('INTERNAL_GROUP_ID'),   // 社内LINEグループID
    DRIVE_FOLDER_ID           : p.getProperty('DRIVE_FOLDER_ID'),     // Googleドライブ保存先フォルダID
    CALENDAR_ID               : p.getProperty('CALENDAR_ID'),         // 登録先カレンダーID
  };
}

// ==========================================
// 1. LINE Webhook受信（メイン入口）
// ==========================================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    }
    const config = getConfig();
    const body   = e.postData.contents;
    const sig    = e.parameter['x-line-signature'] || (e.headers && e.headers['x-line-signature']);

    // GASはHTTPヘッダーを受け取れないためLINE署名検証はスキップ（GASの仕様上の制限）

    const events = JSON.parse(body).events;

    for (const event of events) {
      const groupId    = event.source.groupId || event.source.roomId || event.source.userId;
      const timestamp  = new Date(event.timestamp);
      const senderName = getMemberNameByUserId(event.source.userId) || '不明';

      // --- ファイル・画像の保存 ---
      if (event.type === 'message' && ['image', 'file', 'video'].includes(event.message.type)) {
        saveFileToDrive(event.message.id, event.message.fileName || event.message.type, groupId, timestamp);
        continue;
      }

      // --- 1対1メッセージ：メンバー登録 or 個人タスク照会 ---
      if (event.source.type === 'user' && event.type === 'message' && event.message.type === 'text') {
        const userId   = event.source.userId;
        const isNew    = registerMember(userId);
        if (isNew) {
          sendLineMessage(userId, '登録しました！\nタスクが割り当てられたとき個人通知が届きます。\n\n「残タスクは？」「今週の予定は？」などと送ると確認できます。');
        } else {
          // 会話履歴に保存
          saveChatHistory(userId, senderName, event.message.text, timestamp);

          const memberName = getMemberNameByUserId(userId);

          // 完了報告
          if (isCompletionReport(event.message.text)) {
            handleCompletion(event.message.text, memberName, event.replyToken);
          // 議事録・箇条書き要約リクエスト
          } else if (isSummaryRequest(event.message.text)) {
            const summary = summarizeChat(userId);
            sendLineReply(event.replyToken, `【会話まとめ】\n${summary}`);
          } else {
            const answer = answerQueryForMember(event.message.text, memberName);
            sendLineMessage(userId, answer);
          }
        }
        continue;
      }

      // --- テキストメッセージ ---
      if (event.type !== 'message' || event.message.type !== 'text') continue;

      const text = event.message.text;

      // 会話履歴に保存
      saveChatHistory(groupId, senderName, text, timestamp);

      const isGroup = event.source.type === 'group' || event.source.type === 'room';

      // 議事録・箇条書き要約リクエスト（グループでは@メンション必須）
      if (isSummaryRequest(text, event.message.mention, isGroup)) {
        const summary = summarizeChat(groupId);
        sendLineReply(event.replyToken, `【会話まとめ】\n${summary}`);
        continue;
      }

      // グループでの完了報告（@メンション＋完了キーワード）
      if (isGroup && event.message.mention?.mentionees?.length > 0 && isCompletionReport(text)) {
        handleCompletion(text, senderName, event.replyToken);
        continue;
      }

      // 秘書AIへの質問
      if (isQuery(text)) {
        const answer = answerQuery(text);
        notifyBoth(groupId, answer);
        continue;
      }

      // 短いメッセージ・絵文字のみはGemini呼び出しをスキップ
      if (shouldSkipExtraction(text)) continue;

      // タスク＋スケジュール同時抽出
      const extracted = extractAll(text, groupId, timestamp, senderName);
      if (!extracted) continue;

      // タスク処理
      if (extracted.tasks && extracted.tasks.length > 0) {
        for (const task of extracted.tasks) {
          if (task.projectName) registerNewProject(task.projectName, groupId);
          writeTaskToSheet(task);
          notifyBoth(groupId, buildTaskNotification(task));
          // 担当者に個人DM通知
          const assigneeId = getMemberUserId(task.assignee);
          if (assigneeId) sendLineMessage(assigneeId, `【あなたへのタスク】\n${buildTaskNotification(task)}`);
        }
      }

      // スケジュール処理
      if (extracted.schedules && extracted.schedules.length > 0) {
        for (const schedule of extracted.schedules) {
          if (schedule.projectName) registerNewProject(schedule.projectName, groupId);
          writeScheduleToSheet(schedule);
          addToCalendar(schedule);
          notifyBoth(groupId, buildScheduleNotification(schedule));
        }
      }
    }

    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    console.error('doPost error:', err);
    return ContentService.createTextOutput('Error').setMimeType(ContentService.MimeType.TEXT);
  }
}

// LINE署名検証
function verifySignature(body, signature, secret) {
  if (!signature || !secret) return false;
  const hash = Utilities.computeHmacSha256Signature(body, secret);
  const hashBytes = hash.map(b => (b < 0 ? b + 256 : b));
  return Utilities.base64Encode(hashBytes) === signature;
}

// ==========================================
// 2. Gemini：タスク＋スケジュール同時抽出
// ==========================================

// プロジェクト管理シートの全データを取得（名寄せ用）
function getProjectData() {
  const config = getConfig();
  const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const sheet  = ss.getSheetByName('プロジェクト管理');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getDataRange().getValues().slice(1); // ヘッダー除く
}

// GASコードで名寄せ：生のプロジェクト名 → 正式名称に変換
function normalizeProjectName(rawName) {
  if (!rawName) return rawName;

  const data = getProjectData();
  if (data.length === 0) return rawName;

  const clean = rawName.replace(/[【】\s]/g, ''); // 【】や空白を除去して比較

  // ① 完全一致（略称 or 正式名称）
  for (const row of data) {
    const abbr   = String(row[0] || '').replace(/\s/g, '');
    const formal = String(row[1] || '').replace(/\s/g, '');
    if (!formal) continue;
    if (abbr === clean || formal === clean) return row[1];
  }

  // ② 部分一致（どちらかが一方を含む）
  const candidates = [];
  for (const row of data) {
    const abbr   = String(row[0] || '');
    const formal = String(row[1] || '');
    if (!formal) continue;
    if (formal.includes(clean) || clean.includes(abbr) || abbr.includes(clean) || clean.includes(formal)) {
      candidates.push(row[1]);
    }
  }

  if (candidates.length === 1) return candidates[0];

  // ③ 候補が複数 → Geminiに候補だけ渡して絞る（リストが短いので精度◎）
  if (candidates.length > 1) {
    return resolveProjectNameWithGemini(rawName, candidates);
  }

  // ④ マッチなし → そのまま返す（新規プロジェクトとして扱う）
  return rawName;
}

// 候補が複数あるときだけGeminiで絞る（候補リストが短いので高精度）
function resolveProjectNameWithGemini(rawName, candidates) {
  const config = getConfig();
  const prompt = `以下のプロジェクト名候補の中から、「${rawName}」に最も近いものを1つだけ選んで、その名称のみを出力してください。余計な説明は不要です。\n\n候補：\n${candidates.map((c, i) => `${i + 1}. ${c}`).join('\n')}`;
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${config.GEMINI_API_KEY}`;
  try {
    const res  = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0 } }),
      muteHttpExceptions: true
    });
    const answer = JSON.parse(res.getContentText()).candidates?.[0]?.content?.parts?.[0]?.text?.trim() || '';
    return candidates.includes(answer) ? answer : candidates[0];
  } catch (err) {
    return candidates[0];
  }
}

function extractAll(message, groupId, timestamp, senderName) {
  const config = getConfig();
  const today  = Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyy年MM月dd日');

  const prompt = `あなたは建築会社WOODBASEの専属秘書AIです。
以下のLINEメッセージからタスクとスケジュールをすべて抽出してください。
今日の日付は${today}です。
このメッセージの送信者は「${senderName}」さんです。

【担当者の決め方（重要）】
パターンA：単独実行タスク（資料作成・提出・見積もり・確認・調査など、メンションされた人が一人で完結する作業）
  → assignee にはメンションされた人のみを登録する。送信者は含めない。

パターンB：双方向・参加型（打ち合わせ・会議・面談・現場同行・現場確認など、送信者と受信者が同じ場・時間を共有する内容）
  → assignee またはスケジュールの attendees に、メンションされた人と送信者（${senderName}）の両方をカンマ区切りで登録する。

【プロジェクト名の抽出ルール】
- メッセージに含まれるプロジェクト名・物件名・現場名をそのまま抽出すること（【】があれば中身を使う）
- 変換や統一は不要。抽出した名前をそのまま出力すること

【最重要ルール】
- メッセージに複数のタスク・スケジュールが含まれる場合、必ず配列（Array）を使い、それぞれ独立したオブジェクトとして分割・出力すること。絶対に1件にまとめないこと。
- 例：「AとBをお願いします」→ tasksに2件の独立したオブジェクトを出力する

【抽出ルール】
- 名前＋動詞・依頼 → タスク（「〇〇さん、△△をお願いします」等）
- 「承知しました」「かしこまりました」→ タスク受諾として記録
- 日時＋場所・内容 → スケジュール（打ち合わせ、現場確認、検査、引渡し等）
- 期日ワード（今週中/〇日まで/〇月〇日/来週）→ 期日にセット
- タスクもスケジュールもない場合は両方nullを返す

【出力形式（JSON厳守・配列必須）】
{
  "tasks": [
    {
      "projectName": "プロジェクト名（抽出したまま）",
      "assignee": "担当者名",
      "taskContent": "タスク内容",
      "deadline": "yyyy-MM-dd（不明は空文字）",
      "status": "未着手"
    }
  ] または null,
  "schedules": [
    {
      "projectName": "プロジェクト名（抽出したまま）",
      "title": "予定タイトル",
      "date": "yyyy-MM-dd",
      "startTime": "HH:mm（不明は空文字）",
      "endTime": "HH:mm（不明は空文字）",
      "location": "場所（不明は空文字）",
      "attendees": "参加者（不明は空文字）",
      "description": "詳細"
    }
  ] または null
}

LINEメッセージ：
${message}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${config.GEMINI_API_KEY}`;

  try {
    const res    = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.1 } }),
      muteHttpExceptions: true
    });
    const text   = JSON.parse(res.getContentText()).candidates?.[0]?.content?.parts?.[0]?.text || '';
    const match  = text.match(/\{[\s\S]*\}/);
    if (!match) return null;

    const parsed = JSON.parse(match[0]);
    const result = { tasks: [], schedules: [] };

    if (Array.isArray(parsed.tasks)) {
      result.tasks = parsed.tasks.map(t => ({
        ...t,
        projectName: normalizeProjectName(t.projectName),
        datetime: timestamp,
        groupId
      }));
    }
    if (Array.isArray(parsed.schedules)) {
      result.schedules = parsed.schedules.map(s => ({
        ...s,
        projectName: normalizeProjectName(s.projectName),
        datetime: timestamp,
        groupId
      }));
    }
    return result;

  } catch (err) {
    console.error('extractAll error:', err);
    return null;
  }
}

// ==========================================
// 3. タスク → スプレッドシート書き込み
// ==========================================
function writeTaskToSheet(task) {
  const ss    = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const sheet = ss.getSheetByName('タスク管理');
  const dt    = Utilities.formatDate(new Date(task.datetime), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  sheet.appendRow([
    dt,
    task.projectName  || '',
    task.assignee     || '',
    task.taskContent  || '',
    task.deadline     || '',
    task.status       || '未着手',
    task.groupId      || ''
  ]);
}

// ==========================================
// 4. スケジュール → スプレッドシート書き込み
// ==========================================
function writeScheduleToSheet(schedule) {
  const ss    = SpreadsheetApp.openById(getConfig().SPREADSHEET_ID);
  const sheet = ss.getSheetByName('スケジュール管理');
  const dt    = Utilities.formatDate(new Date(schedule.datetime), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  sheet.appendRow([
    dt,
    schedule.projectName || '',
    schedule.title       || '',
    schedule.date        || '',
    schedule.startTime   || '',
    schedule.endTime     || '',
    schedule.location    || '',
    schedule.attendees   || '',
    schedule.description || '',
    schedule.groupId     || ''
  ]);
}

// ==========================================
// 5. スケジュール → Google Calendar登録
// ==========================================
function addToCalendar(schedule) {
  try {
    if (!schedule.date) return;

    if (!config.CALENDAR_ID) {
      console.log('CALENDAR_ID未設定のためカレンダー登録スキップ');
      return;
    }
    const calendar = CalendarApp.getCalendarById(config.CALENDAR_ID);
    const title    = `【${schedule.projectName || 'WOODBASE'}】${schedule.title}`;
    const date     = new Date(schedule.date);

    if (schedule.startTime) {
      const [sh, sm] = schedule.startTime.split(':').map(Number);
      const start    = new Date(date);
      start.setHours(sh, sm, 0);

      const end = new Date(start);
      if (schedule.endTime) {
        const [eh, em] = schedule.endTime.split(':').map(Number);
        end.setHours(eh, em, 0);
      } else {
        end.setHours(sh + 1, sm, 0); // デフォルト1時間
      }

      calendar.createEvent(title, start, end, {
        location    : schedule.location    || '',
        description : schedule.description || ''
      });
    } else {
      // 終日イベント
      calendar.createAllDayEvent(title, date, {
        location    : schedule.location    || '',
        description : schedule.description || ''
      });
    }

    console.log('カレンダー登録:', title);
  } catch (err) {
    console.error('addToCalendar error:', err);
  }
}

// ==========================================
// 6. 画像・ファイル → Google Drive保存
// ==========================================
function saveFileToDrive(messageId, fileName, groupId, timestamp) {
  try {
    const config  = getConfig();
    const content = UrlFetchApp.fetch(`https://api-data.line.me/v2/bot/message/${messageId}/content`, {
      headers: { 'Authorization': `Bearer ${config.LINE_CHANNEL_ACCESS_TOKEN}` },
      muteHttpExceptions: true
    });

    if (content.getResponseCode() !== 200) return;

    const blob    = content.getBlob();
    const dateStr = Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyyMMdd');
    const safeName = `${dateStr}_${fileName || messageId}`;

    // プロジェクト名をグループIDから取得
    const projectName = getProjectNameByGroupId(groupId) || '未分類';

    // プロジェクトフォルダを取得または作成
    const projectFolder = getOrCreateFolder(config.DRIVE_FOLDER_ID, projectName);

    blob.setName(safeName);
    projectFolder.createFile(blob);

    console.log('ファイル保存:', projectName + '/' + safeName);

    notifyBoth(groupId, `【ファイル保存】\nプロジェクト：${projectName}\nファイル：${safeName}\nGoogleドライブに保存しました。`);

  } catch (err) {
    console.error('saveFileToDrive error:', err);
  }
}

// フォルダを取得、なければ作成
function getOrCreateFolder(parentFolderId, folderName) {
  const parent   = DriveApp.getFolderById(parentFolderId);
  const existing = parent.getFoldersByName(folderName);
  if (existing.hasNext()) return existing.next();
  console.log('新規フォルダ作成:', folderName);
  return parent.createFolder(folderName);
}

// グループIDからプロジェクト名を取得（プロジェクト管理シートを参照）
function getProjectNameByGroupId(groupId) {
  const config = getConfig();
  const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const sheet  = ss.getSheetByName('プロジェクト管理');
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    // C列（施主グループID）またはD列（業者グループID）と一致したら正式名称を返す
    if (data[i][2] === groupId || data[i][3] === groupId) {
      return data[i][1] || data[i][0]; // 正式名称 or 略称
    }
  }
  return null;
}

// 新規プロジェクトをプロジェクト管理シートに自動登録
function registerNewProject(projectName, groupId) {
  const config = getConfig();
  const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const sheet  = ss.getSheetByName('プロジェクト管理');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  // 既に登録済みか確認
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === projectName || data[i][0] === projectName) return;
  }

  // 新規登録
  sheet.appendRow([projectName, projectName, groupId, '', '進行中', '自動登録']);
  console.log('新規プロジェクト登録:', projectName);

  // Googleドライブにフォルダも自動作成
  getOrCreateFolder(config.DRIVE_FOLDER_ID, projectName);
}

// ==========================================
// 7. メンバー管理
// ==========================================

// LINEプロフィールを取得してメンバー管理シートに登録（新規登録ならtrue）
function registerMember(userId) {
  const config = getConfig();
  const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const sheet  = ss.getSheetByName('メンバー管理');
  if (!sheet) return false;

  // 既に登録済みか確認
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userId) return false; // 既存
  }

  // LINEプロフィール取得
  let displayName = '未設定';
  try {
    const res = UrlFetchApp.fetch(`https://api.line.me/v2/bot/profile/${userId}`, {
      headers: { 'Authorization': `Bearer ${config.LINE_CHANNEL_ACCESS_TOKEN}` },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() === 200) {
      displayName = JSON.parse(res.getContentText()).displayName || '未設定';
    }
  } catch(e) {}

  sheet.appendRow([displayName, userId, '社内', '']);
  console.log('メンバー登録:', displayName, userId);
  return true; // 新規登録
}

// userIdから名前を取得
function getMemberNameByUserId(userId) {
  const config = getConfig();
  const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const sheet  = ss.getSheetByName('メンバー管理');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userId) return data[i][0] || null;
  }
  return null;
}

// 個人向けタスク・スケジュール照会
function answerQueryForMember(question, memberName) {
  const config = getConfig();
  const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);

  const now      = new Date();
  const today    = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日');
  const todayYmd = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  const tomorrow = Utilities.formatDate(new Date(now.getTime() + 86400000), 'Asia/Tokyo', 'yyyy-MM-dd');

  // 自分のタスクのみ抽出
  const taskSheet = ss.getSheetByName('タスク管理');
  const taskRows  = taskSheet.getLastRow() > 1
    ? taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, 6).getValues()
    : [];
  const myTasks = taskRows.filter(r => r[5] !== '完了' && r[3] && (!memberName || r[2].includes(memberName.replace('さん',''))));

  // スケジュール全件
  const schedSheet = ss.getSheetByName('スケジュール管理');
  const schedRows  = schedSheet.getLastRow() > 1
    ? schedSheet.getRange(2, 1, schedSheet.getLastRow() - 1, 9).getValues()
    : [];

  const taskList = myTasks.map((r, i) => {
    const dl = r[4] ? Utilities.formatDate(new Date(r[4]), 'Asia/Tokyo', 'M/d') : '期日未定';
    return `${i + 1}. [${r[1]}] ${r[3]}（期日:${dl}・${r[5]}）`;
  }).join('\n') || 'なし';

  const schedList = schedRows.map((r, i) => {
    const attendees = r[7] ? `　参加者：${r[7]}` : '';
    return `${i + 1}. [${r[1]}] ${r[2]} ${r[3]} ${r[4]}〜 場所：${r[6]}${attendees}`;
  }).join('\n') || 'なし';

  const prompt = `あなたは建築会社WOODBASEの専属秘書AIです。
${memberName ? `相手は「${memberName}」さんです。` : ''}
以下のデータを元に質問に簡潔に日本語で答えてください。

【現在日時】
今日：${today}（${todayYmd}）
明日：${tomorrow}
※「明日」「来週」「今週」などは上記を基準に具体的な日付で判断すること。

【${memberName || 'あなた'}の未完了タスク】
${taskList}

【スケジュール一覧】
※「参加者」には、その予定に関わる人物が記録されている。「誰と？」「担当者は？」と聞かれた場合は必ずこの参加者情報を参照して回答すること。
${schedList}

質問：${question}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${config.GEMINI_API_KEY}`;
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2 } }),
      muteHttpExceptions: true
    });
    return JSON.parse(res.getContentText()).candidates?.[0]?.content?.parts?.[0]?.text || 'うまく答えられませんでした。';
  } catch (err) {
    return 'エラーが発生しました。';
  }
}

// 名前からユーザーIDを取得
function getMemberUserId(name) {
  if (!name) return null;
  const config = getConfig();
  const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const sheet  = ss.getSheetByName('メンバー管理');
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].includes(name.replace('さん', '').trim())) {
      return data[i][1] || null;
    }
  }
  return null;
}

// ==========================================
// 7-5. Gemini事前フィルター
// ==========================================
function shouldSkipExtraction(text) {
  if (!text || text.length <= 3) return true;
  // 日本語・英数字を含まない（絵文字・記号のみ）
  if (!/[ぁ-んァ-ヶー一-龯a-zA-Z0-9]/.test(text)) return true;
  // 短い定型返事（10文字以下）
  if (text.length <= 10) {
    const skipWords = ['了解', 'はい', 'おk', 'ok', 'OK', 'なるほど', 'ありがとう', 'お疲れ', 'お世話', 'よろしく', 'わかりました', 'わかった'];
    if (skipWords.some(w => text.includes(w))) return true;
  }
  return false;
}

// ==========================================
// 8. 秘書AI：質問判定
// ==========================================
function isQuery(message) {
  const keywords = ['残タスク', 'タスクは', 'タスクある', '何件', '進捗', 'どうなってる', '教えて', 'スケジュール', '今週', '来週', '予定'];
  return keywords.some(k => message.includes(k));
}

// ==========================================
// 8-0. タスク完了報告
// ==========================================

// 完了報告の判定
function isCompletionReport(text) {
  const keywords = ['完了', '終わりました', 'できました', '終わった', '完成しました', 'やりました'];
  return keywords.some(k => text.includes(k));
}

// タスクを完了に更新（担当者名＋メッセージからマッチング）
function handleCompletion(text, memberName, replyToken) {
  const config = getConfig();
  const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const sheet  = ss.getSheetByName('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) {
    sendLineReply(replyToken, 'タスクが見つかりませんでした。');
    return;
  }

  const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const cleanName = (memberName || '').replace('さん', '').trim();

  // 担当者の未完了タスクを抽出
  const myTasks = [];
  for (let i = 0; i < data.length; i++) {
    const assignee = String(data[i][2] || '');
    const status   = String(data[i][5] || '');
    if (status !== '完了' && cleanName && assignee.includes(cleanName)) {
      myTasks.push({ rowIndex: i + 2, content: data[i][3], project: data[i][1] });
    }
  }

  if (myTasks.length === 0) {
    sendLineReply(replyToken, '未完了のタスクが見つかりませんでした。');
    return;
  }

  let target = myTasks[0];

  // タスクが複数ある場合はGeminiでメッセージと照合
  if (myTasks.length > 1) {
    const list   = myTasks.map((t, i) => `${i + 1}. [${t.project}] ${t.content}`).join('\n');
    const prompt = `以下のタスク一覧の中から、このメッセージ「${text}」が完了報告しているタスクを1つ選び、番号だけ答えてください。\n\n${list}`;
    const url    = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${config.GEMINI_API_KEY}`;
    try {
      const res    = UrlFetchApp.fetch(url, {
        method: 'post', contentType: 'application/json',
        payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0 } }),
        muteHttpExceptions: true
      });
      const answer = JSON.parse(res.getContentText()).candidates?.[0]?.content?.parts?.[0]?.text?.trim() || '';
      const num    = parseInt(answer);
      if (num >= 1 && num <= myTasks.length) target = myTasks[num - 1];
    } catch (err) {
      console.error('handleCompletion Gemini error:', err);
    }
  }

  // スプレッドシートのステータスを完了に更新
  sheet.getRange(target.rowIndex, 6).setValue('完了');
  console.log('タスク完了:', target.content);

  sendLineReply(replyToken, `✅ 完了しました！\nプロジェクト：${target.project}\nタスク：${target.content}`);
}

// 議事録（箇条書き要約）リクエスト判定
// グループでは @メンション必須、1対1では不要
function isSummaryRequest(message, mention, isGroup) {
  const hasKeyword = ['まとめて', '箇条書き', '会話内容', '要約して', '議事録'].some(k => message.includes(k));
  if (!hasKeyword) return false;
  if (isGroup) {
    // グループは @メンションが必須
    return !!(mention && mention.mentionees && mention.mentionees.length > 0);
  }
  // 1対1はキーワードだけでOK
  return true;
}

// ==========================================
// 8-2. 会話履歴の保存・取得・要約
// ==========================================

// テキストメッセージを会話履歴シートに保存
function saveChatHistory(groupId, senderName, text, timestamp) {
  try {
    const config = getConfig();
    const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    let sheet    = ss.getSheetByName('会話履歴');
    if (!sheet) {
      sheet = ss.insertSheet('会話履歴');
      sheet.appendRow(['日時', 'グループID', '送信者', 'メッセージ']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#9FC5E8').setFontColor('#FFFFFF').setHorizontalAlignment('center');
      sheet.setColumnWidth(1, 140); sheet.setColumnWidth(2, 200);
      sheet.setColumnWidth(3, 100); sheet.setColumnWidth(4, 400);
    }
    const dt = Utilities.formatDate(timestamp, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    sheet.appendRow([dt, groupId, senderName, text]);
  } catch (err) {
    console.error('saveChatHistory error:', err);
  }
}

// 会話履歴を取得（groupIdで絞り込み、直近N件）
function getChatHistory(groupId, limit) {
  try {
    const config = getConfig();
    const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const sheet  = ss.getSheetByName('会話履歴');
    if (!sheet || sheet.getLastRow() <= 1) return [];

    const data     = sheet.getDataRange().getValues();
    const filtered = data.slice(1).filter(r => r[1] === groupId);
    return filtered.slice(-limit); // 直近limit件
  } catch (err) {
    console.error('getChatHistory error:', err);
    return [];
  }
}

// Geminiで会話履歴を箇条書きに要約
function summarizeChat(groupId) {
  const config  = getConfig();
  const history = getChatHistory(groupId, 50); // 直近50件

  if (history.length === 0) return '会話履歴がまだ記録されていません。';

  const historyText = history.map(r => `[${r[0]}] ${r[2]}：${r[3]}`).join('\n');

  const prompt = `あなたは建築会社WOODBASEの専属秘書AIです。
以下のLINEグループの会話履歴を、箇条書きで簡潔にまとめてください。
決定事項・依頼・タスク・スケジュール・重要な情報を中心に整理してください。
余計な前置きは不要です。箇条書きのみ出力してください。

【会話履歴】
${historyText}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${config.GEMINI_API_KEY}`;
  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2 } }),
      muteHttpExceptions: true
    });
    return JSON.parse(res.getContentText()).candidates?.[0]?.content?.parts?.[0]?.text || 'まとめられませんでした。';
  } catch (err) {
    console.error('summarizeChat error:', err);
    return 'エラーが発生しました。';
  }
}

// ==========================================
// 9. 秘書AI：スプレッドシートを読んで回答
// ==========================================
function answerQuery(question) {
  const config = getConfig();
  const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);

  // タスク一覧
  const taskSheet  = ss.getSheetByName('タスク管理');
  const taskRows   = taskSheet.getLastRow() > 1
    ? taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, 6).getValues()
    : [];
  const activeTasks = taskRows.filter(r => r[5] !== '完了' && r[3]);

  // スケジュール一覧
  const schedSheet  = ss.getSheetByName('スケジュール管理');
  const schedRows   = schedSheet.getLastRow() > 1
    ? schedSheet.getRange(2, 1, schedSheet.getLastRow() - 1, 9).getValues()
    : [];

  const taskList = activeTasks.map((r, i) => {
    const dl = r[4] ? Utilities.formatDate(new Date(r[4]), 'Asia/Tokyo', 'M/d') : '期日未定';
    return `${i + 1}. [${r[1]}] ${r[2]}：${r[3]}（${dl}・${r[5]}）`;
  }).join('\n') || 'なし';

  const schedList = schedRows.map((r, i) => {
    const attendees = r[7] ? `　参加者：${r[7]}` : '';
    return `${i + 1}. [${r[1]}] ${r[2]} ${r[3]} ${r[4]}〜 場所：${r[6]}${attendees}`;
  }).join('\n') || 'なし';

  const now     = new Date();
  const today   = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日');
  const todayYmd = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  const tomorrow = Utilities.formatDate(new Date(now.getTime() + 86400000), 'Asia/Tokyo', 'yyyy-MM-dd');

  const prompt = `あなたは建築会社WOODBASEの専属秘書AIです。
以下のデータを元に質問に簡潔に日本語で答えてください。

【現在日時】
今日：${today}（${todayYmd}）
明日：${tomorrow}
※「明日」「来週」「今週」などの相対的な表現は、上記の現在日時を基準に具体的な日付（yyyy-MM-dd）に変換して判断すること。

【未完了タスク一覧】
${taskList}

【スケジュール一覧】
※「参加者」には、その予定に関わる人物が記録されている。「誰と？」「担当者は？」と聞かれた場合は必ずこの参加者情報を参照して回答すること。
${schedList}

質問：${question}`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${config.GEMINI_API_KEY}`;

  try {
    const res  = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2 } }),
      muteHttpExceptions: true
    });
    return JSON.parse(res.getContentText()).candidates?.[0]?.content?.parts?.[0]?.text || 'うまく答えられませんでした。';
  } catch (err) {
    console.error('answerQuery error:', err);
    return 'エラーが発生しました。';
  }
}

// ==========================================
// 9. 期日チェック（毎朝9時）
// ==========================================
function checkDeadlines() {
  const config = getConfig();
  const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const sheet  = ss.getSheetByName('タスク管理');

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const data  = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (const row of data) {
    const [, project, assignee, task, deadlineRaw, status, groupId] = row;
    if (!deadlineRaw || status === '完了') continue;

    const deadline = new Date(deadlineRaw);
    deadline.setHours(0, 0, 0, 0);

    const diff = Math.round((deadline - today) / 86400000);
    const dlStr = Utilities.formatDate(deadline, 'Asia/Tokyo', 'M月d日');

    let msg = null;

    if (diff === 3) {
      msg = `【リマインド】\n${assignee}さん、タスクの期日が3日後です。\nプロジェクト：${project}\nタスク：${task}\n期日：${dlStr}`;
    } else if (diff === 0) {
      msg = `【本日期日】\n${assignee}さん、以下のタスクが本日期日です。\nプロジェクト：${project}\nタスク：${task}`;
    } else if (diff < 0 && status !== '完了') {
      msg = `【期日超過】\n${assignee}さん、期日を過ぎたタスクがあります。\nプロジェクト：${project}\nタスク：${task}\n期日：${dlStr}（${Math.abs(diff)}日超過）`;
    } else if (status === '未着手' && diff > 0) {
      msg = `【未着手確認】\n${assignee}さん、未着手のタスクがあります。\nプロジェクト：${project}\nタスク：${task}\n期日：${dlStr}`;
    }

    if (msg) {
      notifyBoth(groupId, msg);
      // 担当者に個人DM通知
      const assigneeId = getMemberUserId(assignee);
      if (assigneeId) sendLineMessage(assigneeId, msg);
    }
  }

  // --- スケジュール前日リマインド ---
  const schedSheet = ss.getSheetByName('スケジュール管理');
  if (schedSheet && schedSheet.getLastRow() > 1) {
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = Utilities.formatDate(tomorrow, 'Asia/Tokyo', 'yyyy-MM-dd');

    const schedData = schedSheet.getRange(2, 1, schedSheet.getLastRow() - 1, 10).getValues();
    for (const row of schedData) {
      const [, project, title, date, startTime, , location, attendees, , schedGroupId] = row;
      if (!date) continue;
      const dateStr = date instanceof Date
        ? Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd')
        : String(date).substring(0, 10);
      if (dateStr !== tomorrowStr) continue;

      const time = startTime ? ` ${startTime}〜` : '';
      const msg  = `【明日の予定】\nプロジェクト：${project || '未定'}\n予定：${title}\n日時：${dateStr}${time}\n場所：${location || '未定'}\n参加者：${attendees || '未定'}`;
      notifyBoth(schedGroupId || config.INTERNAL_GROUP_ID, msg);
    }
  }
}

// ==========================================
// 9-2. 週次レポート（毎週月曜8時）
// ==========================================
function sendWeeklyReport() {
  const config = getConfig();
  const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const now    = new Date();

  // 今週月曜〜日曜の範囲
  const dayOfWeek = now.getDay() === 0 ? 7 : now.getDay();
  const monday    = new Date(now); monday.setDate(now.getDate() - dayOfWeek + 1); monday.setHours(0,0,0,0);
  const sunday    = new Date(monday); sunday.setDate(monday.getDate() + 6);
  const mondayStr = Utilities.formatDate(monday, 'Asia/Tokyo', 'M/d');
  const sundayStr = Utilities.formatDate(sunday, 'Asia/Tokyo', 'M/d');
  const mondayYmd = Utilities.formatDate(monday, 'Asia/Tokyo', 'yyyy-MM-dd');
  const sundayYmd = Utilities.formatDate(sunday, 'Asia/Tokyo', 'yyyy-MM-dd');

  // 未完了タスク
  const taskSheet  = ss.getSheetByName('タスク管理');
  const taskData   = taskSheet.getLastRow() > 1 ? taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, 6).getValues() : [];
  const activeTasks = taskData.filter(r => r[5] !== '完了' && r[3]);
  const taskList   = activeTasks.length > 0
    ? activeTasks.map(r => {
        const dl = r[4] ? Utilities.formatDate(new Date(r[4]), 'Asia/Tokyo', 'M/d') : '期日未定';
        return `・[${r[1]}] ${r[2]}：${r[3]}（${dl}）`;
      }).join('\n')
    : 'なし';

  // 今週のスケジュール
  const schedSheet = ss.getSheetByName('スケジュール管理');
  const schedData  = schedSheet.getLastRow() > 1 ? schedSheet.getRange(2, 1, schedSheet.getLastRow() - 1, 9).getValues() : [];
  const weekScheds = schedData.filter(r => {
    if (!r[3]) return false;
    const d = r[3] instanceof Date ? Utilities.formatDate(r[3], 'Asia/Tokyo', 'yyyy-MM-dd') : String(r[3]).substring(0, 10);
    return d >= mondayYmd && d <= sundayYmd;
  });
  const schedList  = weekScheds.length > 0
    ? weekScheds.map(r => {
        const d    = r[3] instanceof Date ? Utilities.formatDate(r[3], 'Asia/Tokyo', 'M/d') : String(r[3]).substring(5, 10).replace('-', '/');
        const time = r[4] ? ` ${r[4]}〜` : '';
        return `・${d}${time} [${r[1]}] ${r[2]}`;
      }).join('\n')
    : 'なし';

  const msg = `【週次レポート ${mondayStr}〜${sundayStr}】\n\n▼ 今週のスケジュール\n${schedList}\n\n▼ 未完了タスク一覧\n${taskList}`;
  sendLineMessage(config.INTERNAL_GROUP_ID, msg);
}

// ==========================================
// 9-3. 完了タスクのアーカイブ（毎週日曜23時）
// ==========================================
function archiveCompletedTasks() {
  const config = getConfig();
  const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  const sheet  = ss.getSheetByName('タスク管理');
  if (!sheet || sheet.getLastRow() <= 1) return;

  // アーカイブシートを取得または作成
  let archiveSheet = ss.getSheetByName('完了タスク');
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet('完了タスク');
    archiveSheet.appendRow(['日時', 'プロジェクト名', '担当者名', 'タスク内容', '期日', 'ステータス', 'グループID']);
    archiveSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#B7B7B7').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  }

  const data     = sheet.getDataRange().getValues();
  const toDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][5] === '完了') {
      archiveSheet.appendRow(data[i]);
      toDelete.push(i + 1);
    }
  }
  toDelete.forEach(rowNum => sheet.deleteRow(rowNum));
  console.log(`アーカイブ完了: ${toDelete.length}件`);
}

// ==========================================
// 10. 通知ヘルパー
// ==========================================
function notifyBoth(sourceGroupId, message) {
  const config = getConfig();
  if (sourceGroupId) sendLineMessage(sourceGroupId, message);
  if (config.INTERNAL_GROUP_ID && config.INTERNAL_GROUP_ID !== sourceGroupId) {
    sendLineMessage(config.INTERNAL_GROUP_ID, message);
  }
}

// replyTokenを使った返信（送信元に確実に届く）
function sendLineReply(replyToken, message) {
  if (!replyToken) return;
  const config = getConfig();
  // LINE は1メッセージ5000文字制限
  const text = message.length > 4900 ? message.slice(0, 4900) + '…' : message;
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'post',
      headers: {
        'Content-Type'  : 'application/json',
        'Authorization' : `Bearer ${config.LINE_CHANNEL_ACCESS_TOKEN}`
      },
      payload: JSON.stringify({ replyToken, messages: [{ type: 'text', text }] }),
      muteHttpExceptions: true
    });
  } catch (err) {
    console.error('sendLineReply error:', err);
  }
}

function sendLineMessage(targetId, message) {
  if (!targetId || targetId === 'Uxxxxxxxxxx') return;
  const config = getConfig();

  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      headers: {
        'Content-Type'  : 'application/json',
        'Authorization' : `Bearer ${config.LINE_CHANNEL_ACCESS_TOKEN}`
      },
      payload: JSON.stringify({ to: targetId, messages: [{ type: 'text', text: message }] }),
      muteHttpExceptions: true
    });
  } catch (err) {
    console.error('sendLineMessage error:', err);
  }
}

// ==========================================
// 11. 通知メッセージ生成
// ==========================================
function buildTaskNotification(task) {
  const dl = task.deadline
    ? Utilities.formatDate(new Date(task.deadline), 'Asia/Tokyo', 'M月d日')
    : '未定';
  return `【タスク登録】\nプロジェクト：${task.projectName || '不明'}\n担当：${task.assignee || '未定'}\n内容：${task.taskContent}\n期日：${dl}`;
}

function buildScheduleNotification(schedule) {
  const time = schedule.startTime ? ` ${schedule.startTime}〜` : '';
  return `【スケジュール登録】\nプロジェクト：${schedule.projectName || '不明'}\n予定：${schedule.title}\n日時：${schedule.date}${time}\n場所：${schedule.location || '未定'}`;
}

// ==========================================
// 12. 初期セットアップ
// ==========================================
function setup() {
  const config = getConfig();
  const ss     = SpreadsheetApp.openById(config.SPREADSHEET_ID);

  // --- タスク管理シート ---
  setupSheet(ss, 'タスク管理', ['日時', 'プロジェクト名', '担当者名', 'タスク内容', '期日', 'ステータス', 'グループID'], '#4A86E8', [140,160,100,280,100,80,160]);
  const taskSheet = ss.getSheetByName('タスク管理');
  taskSheet.getRange(2, 6, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['未着手', '対応中', '完了'], true).build()
  );

  // --- スケジュール管理シート ---
  setupSheet(ss, 'スケジュール管理', ['登録日時', 'プロジェクト名', '予定タイトル', '日付', '開始時間', '終了時間', '場所', '参加者', '詳細', 'グループID'], '#E67C73', [140,160,200,100,80,80,160,160,200,160]);

  // --- メンバー管理シート ---
  setupSheet(ss, 'メンバー管理', ['名前', 'LINE ユーザーID', '役割', '備考'], '#34A853', [120,220,120,200]);
  const memberSheet = ss.getSheetByName('メンバー管理');
  if (memberSheet.getLastRow() <= 1) {
    memberSheet.appendRow(['濱田', 'Uxxxxxxxxxx', '社内', '']);
    memberSheet.appendRow(['織田', 'Uxxxxxxxxxx', '社内', '']);
  }

  // --- プロジェクト管理シート ---
  setupSheet(ss, 'プロジェクト管理', ['略称', '正式名称', 'グループID（施主）', 'グループID（業者）', 'ステータス', '備考'], '#F6B26B', [80,200,180,180,80,200]);
  const projSheet = ss.getSheetByName('プロジェクト管理');
  if (projSheet.getLastRow() <= 1) {
    projSheet.appendRow(['雨晴れ', '雨のち晴れクリニック', '', '', '進行中', '']);
  }

  // --- 会話履歴シート ---
  setupSheet(ss, '会話履歴', ['日時', 'グループID', '送信者', 'メッセージ'], '#9FC5E8', [140, 200, 100, 400]);

  SpreadsheetApp.getUi().alert('セットアップ完了！\n\n次の手順：\n1. メンバー管理シートにLINEユーザーIDを入力\n2. プロジェクト管理シートにプロジェクト情報を入力\n3. スクリプトプロパティにINTERNAL_GROUP_IDとDRIVE_FOLDER_IDを追加\n4. createTriggersを実行');
}

function setupSheet(ss, name, headers, color, widths) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clearContents();
  const bandings = sheet.getBandings();
  bandings.forEach(b => b.remove());
  sheet.appendRow(headers);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold').setBackground(color).setFontColor('#FFFFFF').setHorizontalAlignment('center');
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange(2, 1, 1000, headers.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
}

// ==========================================
// 13. トリガー設定
// ==========================================
function createTriggers() {
  // 既存トリガーを全削除して再作成
  const targets = ['checkDeadlines', 'sendWeeklyReport', 'archiveCompletedTasks'];
  ScriptApp.getProjectTriggers()
    .filter(t => targets.includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 毎朝9時：タスク期日チェック＋スケジュール前日リマインド
  ScriptApp.newTrigger('checkDeadlines').timeBased().everyDays(1).atHour(9).create();
  // 毎週月曜8時：週次レポート
  ScriptApp.newTrigger('sendWeeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  // 毎週日曜23時：完了タスクをアーカイブ
  ScriptApp.newTrigger('archiveCompletedTasks').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(23).create();

  SpreadsheetApp.getUi().alert('完了：\n・毎朝9時 タスクリマインド＋スケジュール前日通知\n・毎週月曜8時 週次レポート\n・毎週日曜23時 完了タスクアーカイブ');
}

// ==========================================
// 14. テスト用関数
// ==========================================
function testExtractAll() {
  const msg    = '【雨晴れ】@織田さん 4/10までに見積もり提出お願いします。4/8(火)14時から現場確認です。';
  const result = extractAll(msg, 'test-group-id', new Date());
  console.log('抽出結果:', JSON.stringify(result, null, 2));
}

function testFullFlow() {
  const msg    = '【雨晴れ】@織田さん 4/10までに見積もり提出お願いします。';
  const result = extractAll(msg, 'test-group-id', new Date());
  if (!result) { console.log('抽出失敗'); return; }
  if (result.tasks) result.tasks.forEach(t => writeTaskToSheet(t));
  if (result.schedules) result.schedules.forEach(s => { writeScheduleToSheet(s); addToCalendar(s); });
  console.log('フルフロー完了');
}

function testAnswerQuery() {
  console.log(answerQuery('雨晴れの残タスクは？'));
}

// グループIDをログに出力（Webhookが来たときに確認用）
function logGroupId(groupId) {
  console.log('グループID:', groupId);
}
