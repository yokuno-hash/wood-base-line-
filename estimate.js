// ==========================================
// 工程管理表 → 見積書 月別シート自動生成
// ==========================================
// 設定シート（管理用スプシ内）:
//   - 見積連携設定: 工程⇔見積のペア定義
//   - 見積単価設定: 単価ルール
//   - 見積要確認:   自動処理できなかった行のログ
// ==========================================

var ESTIMATE_LINK_SETTING_SHEET = '見積連携設定';
var ESTIMATE_RATE_SHEET         = '見積単価設定';
var ESTIMATE_CONFIRM_SHEET      = '見積要確認';

var ESTIMATE_LINK_HEADERS    = ['設定ID','名称','工程管理表ID','工程シート名','見積書ID','テンプレートシート名','出力シート名形式','担当者表示','有効'];
var ESTIMATE_RATE_HEADERS    = ['設定ID','分類','条件キーワード','単価','単位','備考'];
var ESTIMATE_CONFIRM_HEADERS = ['作成日時','設定ID','対象シート名','物件','図面名','分類','理由'];

// テンプレートのセル位置（テンプレートに合わせて要調整）
// 1ページ目（物件集計）
var COVER_ITEM_START_ROW = 15;  // 明細開始行
var COVER_COL_NAME       = 2;   // B 名称
var COVER_COL_QTY        = 5;   // E 数量
var COVER_COL_UNIT       = 6;   // F 単位
var COVER_COL_AMOUNT     = 7;   // G 金額
var COVER_COL_NOTE       = 8;   // H 備考
var COVER_MAX_ROWS       = 30;  // 1ページに入る最大行

// 2ページ目（明細）
var DETAIL_START_ROW      = 50;  // 明細開始行
var DETAIL_COL_NAME       = 2;   // B
var DETAIL_COL_QTY        = 5;   // E
var DETAIL_COL_UNIT       = 6;   // F
var DETAIL_COL_UNITPRICE  = 7;   // G
var DETAIL_COL_AMOUNT     = 8;   // H
var DETAIL_COL_NOTE       = 9;   // I

var TAX_RATE = 0.10;

// ==========================================
// SECTION: 設定シート操作
// ==========================================

function getEstimateLinkSettings() {
  var sh = getSheet(ESTIMATE_LINK_SETTING_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).trim(); });
  var idx = {};
  ESTIMATE_LINK_HEADERS.forEach(function(h){ idx[h] = headers.indexOf(h); });
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[idx['設定ID']]) continue;
    var enabled = r[idx['有効']];
    if (enabled === false || String(enabled).toUpperCase() === 'FALSE') continue;
    rows.push({
      settingId:       String(r[idx['設定ID']]).trim(),
      name:            String(r[idx['名称']]).trim(),
      processSsId:     String(r[idx['工程管理表ID']]).trim(),
      processSheet:    String(r[idx['工程シート名']]).trim(),
      estimateSsId:    String(r[idx['見積書ID']]).trim(),
      templateSheet:   String(r[idx['テンプレートシート名']]).trim(),
      sheetNameFormat: String(r[idx['出力シート名形式']]).trim() || 'YY-M',
      assigneeLabel:   String(r[idx['担当者表示']]).trim(),
    });
  }
  return rows;
}

function getEstimateRateSettings() {
  var sh = getSheet(ESTIMATE_RATE_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).trim(); });
  var idx = {};
  ESTIMATE_RATE_HEADERS.forEach(function(h){ idx[h] = headers.indexOf(h); });
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[idx['設定ID']]) continue;
    rows.push({
      settingId: String(r[idx['設定ID']]).trim(),
      category:  String(r[idx['分類']]).trim(),
      keyword:   String(r[idx['条件キーワード']]).trim(),
      rate:      Number(r[idx['単価']]) || 0,
      unit:      String(r[idx['単位']]).trim() || '枚',
      note:      String(r[idx['備考']]).trim(),
    });
  }
  return rows;
}

// ==========================================
// SECTION: メッセージ解析
// ==========================================

function isEstimateCreateRequest(text) {
  if (!text) return false;
  // 「見積作って／作成」「見積書作って／作成」など
  return /見積(書)?(を)?(作って|作成|作っといて|つくって|お願い)/.test(text) ||
         (/見積(書)?/.test(text) && /(作って|作成|つくって|お願い)/.test(text));
}

function findEstimateSettingFromMessage(text) {
  if (!text) return null;
  var settings = getEstimateLinkSettings();
  if (!settings.length) return null;
  // 名称・設定IDの完全/部分一致を順に試す
  var t = String(text).normalize('NFKC');
  // 名称優先（長いほうから）
  var byName = settings.slice().sort(function(a,b){ return b.name.length - a.name.length; });
  for (var i = 0; i < byName.length; i++) {
    if (byName[i].name && t.indexOf(byName[i].name) !== -1) return byName[i];
  }
  for (var j = 0; j < settings.length; j++) {
    if (settings[j].settingId && t.toLowerCase().indexOf(settings[j].settingId.toLowerCase()) !== -1) return settings[j];
  }
  // 1件しか登録されていなければそれを使う
  if (settings.length === 1) return settings[0];
  return null;
}

// 対象年月を抽出。年指定なしは現在年（1月〜3月は要注意 - 必要に応じてここを調整）
function parseEstimateTargetDate(text) {
  if (!text) return null;
  var t = String(text).normalize('NFKC');
  var year  = null;
  var month = null;
  // 「2026年5月」「2026/5」
  var ym1 = t.match(/(\d{4})年(\d{1,2})月/);
  var ym2 = t.match(/(\d{4})[\/\-](\d{1,2})/);
  if (ym1)      { year = parseInt(ym1[1], 10); month = parseInt(ym1[2], 10); }
  else if (ym2) { year = parseInt(ym2[1], 10); month = parseInt(ym2[2], 10); }
  else {
    // 「5月」のみ
    var m = t.match(/(\d{1,2})月/);
    if (m) {
      month = parseInt(m[1], 10);
      year  = new Date().getFullYear();
      // 注意: 1月の指定で12月以降のデータを期待する等、年跨ぎが必要な場合は呼び出し側で調整
    }
  }
  if (!month || month < 1 || month > 12) return null;
  if (!year)  year = new Date().getFullYear();
  return { year: year, month: month };
}

function buildEstimateSheetName(year, month) {
  return String(year).slice(-2) + '-' + month;
}

// ==========================================
// SECTION: 工程管理表 読み取り
// ==========================================

function getProcessRowsBySetting(setting) {
  var ss = SpreadsheetApp.openById(setting.processSsId);
  var sh = ss.getSheetByName(setting.processSheet);
  if (!sh) throw new Error('工程シート「' + setting.processSheet + '」が見つかりません');
  if (sh.getLastRow() < 2) return { headers: [], rows: [] };
  var data = sh.getDataRange().getValues();
  // ヘッダー行検出（「物件」を含む最初の行）
  var headerRow = 0;
  for (var i = 0; i < Math.min(5, data.length); i++) {
    var has = data[i].some(function(c){ return String(c).indexOf('物件') !== -1; });
    if (has) { headerRow = i; break; }
  }
  return { headers: data[headerRow].map(function(h){ return String(h).trim(); }), rows: data.slice(headerRow + 1) };
}

function normalizeProcessRows(processData) {
  var headers = processData.headers;
  function colIdx(name) {
    for (var i = 0; i < headers.length; i++) if (headers[i] === name) return i;
    // ファジー
    for (var j = 0; j < headers.length; j++) if (headers[j] && headers[j].indexOf(name) !== -1) return j;
    return -1;
  }
  var iProp   = colIdx('物件');
  var iDate   = colIdx('提出日');
  var iDraw   = colIdx('図面名');
  var iCnt    = colIdx('枚数');
  var iTime   = colIdx('製図時間');
  var iCat    = colIdx('分類');
  var iNo     = colIdx('物件番号');
  var iAuthor = colIdx('製図者');

  var result = [];
  processData.rows.forEach(function(r){
    var prop = iProp !== -1 ? String(r[iProp] || '').trim() : '';
    if (!prop) return; // 物件名空欄はスキップ
    var dateVal = iDate !== -1 ? r[iDate] : '';
    var month = null, day = null, year = null;
    if (dateVal instanceof Date) {
      year = dateVal.getFullYear(); month = dateVal.getMonth() + 1; day = dateVal.getDate();
    } else {
      var s = String(dateVal || '').normalize('NFKC');
      var m1 = s.match(/(\d{1,2})月(\d{1,2})日/);
      var m2 = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
      if (m2)      { year = parseInt(m2[1],10); month = parseInt(m2[2],10); day = parseInt(m2[3],10); }
      else if (m1) { month = parseInt(m1[1],10); day = parseInt(m1[2],10); }
    }
    result.push({
      property:    prop,
      propertyNo:  iNo     !== -1 ? String(r[iNo] || '').trim() : '',
      submitDate:  dateVal,
      submitYear:  year,
      submitMonth: month,
      submitDay:   day,
      drawingName: iDraw   !== -1 ? String(r[iDraw] || '').trim() : '',
      count:       iCnt    !== -1 ? Number(r[iCnt])  || 0 : 0,
      time:        iTime   !== -1 ? Number(r[iTime]) || 0 : 0,
      category:    iCat    !== -1 ? String(r[iCat] || '').trim() : '',
      author:      iAuthor !== -1 ? String(r[iAuthor] || '').trim() : '',
    });
  });
  return result;
}

function filterRowsByTargetMonth(rows, year, month) {
  return rows.filter(function(r){
    if (r.submitMonth !== month) return false;
    if (r.submitYear !== null && r.submitYear !== year) return false;
    return true;
  });
}

function groupRowsByProperty(rows) {
  var map = {};
  var order = [];
  rows.forEach(function(r){
    var key = r.propertyNo + '|' + r.property;
    if (!map[key]) { map[key] = { propertyNo: r.propertyNo, property: r.property, rows: [] }; order.push(key); }
    map[key].rows.push(r);
  });
  return order.map(function(k){ return map[k]; });
}

// ==========================================
// SECTION: 単価判定
// ==========================================

function getEstimateRate(settingId, category, drawingName, rateSettings) {
  if (!rateSettings) rateSettings = getEstimateRateSettings();
  var cands = rateSettings.filter(function(r){
    if (r.settingId && r.settingId !== settingId) return false;
    if (r.category && r.category !== category)   return false;
    return true;
  });
  if (!cands.length) return null;
  // キーワード一致を優先（長いキーワードほど具体的）
  var draw = String(drawingName || '');
  var specific = cands.filter(function(r){ return r.keyword && r.keyword !== 'デフォルト' && draw.indexOf(r.keyword) !== -1; });
  if (specific.length) {
    specific.sort(function(a, b){ return b.keyword.length - a.keyword.length; });
    return specific[0];
  }
  // デフォルト
  var def = cands.filter(function(r){ return !r.keyword || r.keyword === 'デフォルト'; });
  if (def.length) return def[0];
  return null;
}

// ==========================================
// SECTION: 集計計算
// ==========================================

function calculateEstimateData(groups, rateSettings, settingId) {
  var properties   = [];
  var confirmItems = [];
  var subtotal     = 0;

  groups.forEach(function(g){
    var totalAmount = 0;
    var newSheets   = 0;
    var fixHours    = 0;
    var detailItems = [];

    g.rows.forEach(function(r){
      var amount  = 0;
      var qty     = 0;
      var unit    = '';
      var unitPrice = 0;
      var confirm = '';

      if (!r.category) {
        confirm = '分類が空欄';
      } else if (r.category === '修正') {
        if (!r.time) { confirm = '製図時間が空欄'; }
        else {
          qty = r.time; unit = 'H'; unitPrice = 5000; amount = qty * unitPrice;
          fixHours += r.time;
        }
      } else if (r.category === '新規') {
        if (!r.count) { confirm = '枚数が空欄'; }
        else {
          var rate = getEstimateRate(settingId, '新規', r.drawingName, rateSettings);
          if (!rate || !rate.rate) {
            confirm = '新規単価が見つかりません';
            qty = r.count; unit = '枚';
          } else {
            qty = r.count; unit = rate.unit || '枚'; unitPrice = rate.rate; amount = qty * unitPrice;
            newSheets += r.count;
          }
        }
      } else if (r.category === '検討のみ' || r.category === 'WBF') {
        confirm = r.category + '：請求対象不明';
        qty = r.count || r.time || 0;
        unit = r.count ? '枚' : (r.time ? 'H' : '');
      } else {
        confirm = '未対応の分類「' + r.category + '」';
      }

      if (confirm) {
        confirmItems.push({
          settingId:   settingId,
          property:    g.property,
          drawingName: r.drawingName,
          category:    r.category,
          reason:      confirm,
        });
      }

      detailItems.push({
        name:      r.drawingName + (r.category ? '　' + r.category : ''),
        qty:       qty,
        unit:      unit,
        unitPrice: unitPrice,
        amount:    amount,
      });
      totalAmount += amount;
    });

    // 名称生成（【物件番号】物件名 (新規〇枚／修正〇時間)）
    var suffixParts = [];
    if (newSheets > 0) suffixParts.push('新規' + newSheets + '枚');
    if (fixHours  > 0) suffixParts.push('修正' + fixHours + '時間');
    var propLabel = g.property;
    // 「様邸」「邸」「PJ」が既にあれば重複させない（明示的な追加はしない方針）
    var coverName = '【' + (g.propertyNo || '') + '】' + propLabel + (suffixParts.length ? '（' + suffixParts.join('／') + '）' : '');

    properties.push({
      propertyNo:  g.propertyNo,
      property:    g.property,
      coverName:   coverName,
      amount:      totalAmount,
      newSheets:   newSheets,
      fixHours:    fixHours,
      detailItems: detailItems,
    });
    subtotal += totalAmount;
  });

  var tax   = Math.round(subtotal * TAX_RATE);
  var total = subtotal + tax;
  return {
    properties:   properties,
    confirmItems: confirmItems,
    subtotal:     subtotal,
    tax:          tax,
    total:        total,
    rowCount:     groups.reduce(function(s, g){ return s + g.rows.length; }, 0),
    groupCount:   groups.length,
  };
}

// ==========================================
// SECTION: シート操作
// ==========================================

function copyEstimateTemplateSheet(estimateSpreadsheetId, templateSheetName, newSheetName) {
  var ss = SpreadsheetApp.openById(estimateSpreadsheetId);
  if (ss.getSheetByName(newSheetName)) {
    return { existing: true, ss: ss, sheet: ss.getSheetByName(newSheetName) };
  }
  var tpl = ss.getSheetByName(templateSheetName);
  if (!tpl) throw new Error('テンプレートシート「' + templateSheetName + '」が見つかりません');
  var copy = tpl.copyTo(ss);
  copy.setName(newSheetName);
  // 末尾→先頭側に寄せる
  ss.setActiveSheet(copy);
  return { existing: false, ss: ss, sheet: copy };
}

// シート全体から指定テキストを含むセルを探す
function findCellByText(sheet, searchText) {
  var data = sheet.getDataRange().getValues();
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < data[r].length; c++) {
      if (String(data[r][c] || '').indexOf(searchText) !== -1) return { row: r + 1, col: c + 1 };
    }
  }
  return null;
}

function writeEstimateCover(targetSheet, estimateData, setting) {
  var startRow = COVER_ITEM_START_ROW;
  estimateData.properties.forEach(function(p, i){
    var row = startRow + i;
    if (i >= COVER_MAX_ROWS) return; // 上限超過は明細ページに任せる
    targetSheet.getRange(row, COVER_COL_NAME).setValue(p.coverName);
    targetSheet.getRange(row, COVER_COL_QTY).setValue(1);
    targetSheet.getRange(row, COVER_COL_UNIT).setValue('式');
    targetSheet.getRange(row, COVER_COL_AMOUNT).setValue(p.amount);
    targetSheet.getRange(row, COVER_COL_NOTE).setValue(i === 0 ? (setting.assigneeLabel || '') : '同上');
  });
}

function writeEstimateDetails(targetSheet, estimateData, setting) {
  var row = DETAIL_START_ROW;
  estimateData.properties.forEach(function(p){
    // 物件見出し
    targetSheet.getRange(row, DETAIL_COL_NAME).setValue('【' + (p.propertyNo || '') + '】' + p.property);
    targetSheet.getRange(row, DETAIL_COL_NAME).setFontWeight('bold');
    row++;
    p.detailItems.forEach(function(d){
      targetSheet.getRange(row, DETAIL_COL_NAME).setValue(d.name);
      targetSheet.getRange(row, DETAIL_COL_QTY).setValue(d.qty || '');
      targetSheet.getRange(row, DETAIL_COL_UNIT).setValue(d.unit || '');
      targetSheet.getRange(row, DETAIL_COL_UNITPRICE).setValue(d.unitPrice || '');
      targetSheet.getRange(row, DETAIL_COL_AMOUNT).setValue(d.amount || '');
      row++;
    });
    row++; // 物件間の空行
  });
}

// 「小計」「消費税」「合計」「税込合計」をテキスト検索で見つけて右隣に金額を書く
function writeEstimateTotals(targetSheet, estimateData) {
  function setNextRight(label, value) {
    var found = findCellByText(targetSheet, label);
    if (!found) return;
    targetSheet.getRange(found.row, found.col + 1).setValue(value);
  }
  setNextRight('小計',     estimateData.subtotal);
  setNextRight('消費税',   estimateData.tax);
  setNextRight('税込合計', estimateData.total);
  // 「合計」と「税込合計」両方ある場合、税込合計を優先するため税込合計を後に書く必要があるが
  // findCellByTextは最初のセルを返すので、「合計」が「税込合計」より上にあると誤上書きの可能性あり
  setNextRight('合計', estimateData.total);
}

function writeEstimateConfirmItems(confirmItems, targetSheetName) {
  if (!confirmItems || !confirmItems.length) return;
  var sh = getSheet(ESTIMATE_CONFIRM_SHEET);
  if (!sh) { sh = getSS().insertSheet(ESTIMATE_CONFIRM_SHEET); sh.appendRow(ESTIMATE_CONFIRM_HEADERS); }
  var now = fmtDT(new Date());
  confirmItems.forEach(function(c){
    sh.appendRow([now, c.settingId, targetSheetName, c.property, c.drawingName, c.category, c.reason]);
  });
}

// ==========================================
// SECTION: メイン処理
// ==========================================

function handleEstimateCreateRequest(event, messageText) {
  try {
    if (!isEstimateCreateRequest(messageText)) return false;

    var setting = findEstimateSettingFromMessage(messageText);
    if (!setting) {
      sendLineReply(event.replyToken, '⚠️ 対象の見積設定が特定できませんでした。\n「見積連携設定」シートの名称をメッセージに含めてください。');
      return true;
    }

    var ym = parseEstimateTargetDate(messageText);
    if (!ym) {
      sendLineReply(event.replyToken, '⚠️ 対象年月が読み取れませんでした。「5月分」「2026年5月」のように指定してください。');
      return true;
    }

    var sheetName = buildEstimateSheetName(ym.year, ym.month);

    // 既存チェック
    var preCheckSs = SpreadsheetApp.openById(setting.estimateSsId);
    if (preCheckSs.getSheetByName(sheetName)) {
      sendLineReply(event.replyToken, '⚠️ 既に「' + sheetName + '」が存在します。\n上書きしないため処理を中止しました。\n\n見積書：' + preCheckSs.getUrl());
      return true;
    }

    // 工程データ取得 → 抽出
    var processData = getProcessRowsBySetting(setting);
    var allRows     = normalizeProcessRows(processData);
    var targetRows  = filterRowsByTargetMonth(allRows, ym.year, ym.month);
    if (!targetRows.length) {
      sendLineReply(event.replyToken, '⚠️ 工程管理表に「' + ym.year + '年' + ym.month + '月」のデータがありませんでした。');
      return true;
    }

    var groups       = groupRowsByProperty(targetRows);
    var rateSettings = getEstimateRateSettings();
    var data         = calculateEstimateData(groups, rateSettings, setting.settingId);

    // テンプレコピー
    var copyRes = copyEstimateTemplateSheet(setting.estimateSsId, setting.templateSheet, sheetName);
    var target  = copyRes.sheet;

    // 反映
    writeEstimateCover(target, data, setting);
    writeEstimateDetails(target, data, setting);
    writeEstimateTotals(target, data);
    writeEstimateConfirmItems(data.confirmItems, sheetName);

    var reply = formatEstimateCreatedReply({
      setting: setting, year: ym.year, month: ym.month, sheetName: sheetName,
      data: data, spreadsheetUrl: copyRes.ss.getUrl(),
    });
    sendLineReply(event.replyToken, reply);
    return true;
  } catch (err) {
    console.error('handleEstimateCreateRequest error:', err.message, err.stack);
    try { sendLineReply(event.replyToken, '⚠️ 見積書作成中にエラーが発生しました：' + err.message); } catch (e) {}
    return true; // 他のハンドラに流さない
  }
}

function formatEstimateCreatedReply(r) {
  var lines = [];
  lines.push(r.setting.name + ' ' + r.sheetName + ' の見積書を作成しました。');
  lines.push('');
  lines.push('対象：' + r.year + '年' + r.month + '月分');
  lines.push('対象件数：' + r.data.rowCount + '件');
  lines.push('対象物件数：' + r.data.groupCount + '件');
  lines.push('小計：¥' + r.data.subtotal.toLocaleString('ja-JP'));
  lines.push('消費税：¥' + r.data.tax.toLocaleString('ja-JP'));
  lines.push('税込合計：¥' + r.data.total.toLocaleString('ja-JP'));
  if (r.data.confirmItems.length) {
    lines.push('');
    lines.push('要確認：' + r.data.confirmItems.length + '件');
    r.data.confirmItems.slice(0, 10).forEach(function(c){
      lines.push('・' + c.property + ' / ' + c.drawingName + ' / ' + c.category + '：' + c.reason);
    });
    if (r.data.confirmItems.length > 10) lines.push('…ほか ' + (r.data.confirmItems.length - 10) + ' 件');
  }
  lines.push('');
  lines.push('見積書：' + r.spreadsheetUrl);
  return lines.join('\n');
}

// ==========================================
// SECTION: セットアップ
// ==========================================

function setupEstimateSheets() {
  var ss = getSS();
  [ESTIMATE_LINK_SETTING_SHEET, ESTIMATE_RATE_SHEET, ESTIMATE_CONFIRM_SHEET].forEach(function(name){
    if (ss.getSheetByName(name)) return;
    var sh = ss.insertSheet(name);
    var headers = name === ESTIMATE_LINK_SETTING_SHEET ? ESTIMATE_LINK_HEADERS
                : name === ESTIMATE_RATE_SHEET         ? ESTIMATE_RATE_HEADERS
                : ESTIMATE_CONFIRM_HEADERS;
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#D9EAD3');
  });
  console.log('見積関連シート作成完了:', ESTIMATE_LINK_SETTING_SHEET, '/', ESTIMATE_RATE_SHEET, '/', ESTIMATE_CONFIRM_SHEET);
}

// 東邦家具のデフォルト設定を投入（存在しなければ追加のみ・既存値は上書きしない）
function setupEstimateForToho() {
  setupEstimateSheets();

  var link = getSheet(ESTIMATE_LINK_SETTING_SHEET);
  var data = link.getDataRange().getValues();
  var hasToho = false;
  for (var i = 1; i < data.length; i++) if (String(data[i][0]).trim() === 'toho') { hasToho = true; break; }
  if (!hasToho) {
    link.appendRow(['toho', '東邦家具', '1MHHVpkdNjTihse-CX4sOZbeGKi71ncKD', '工程管理表',
                    '1Fr5_KPcvTGa7WP7PpQ7glIQglaZgeuv_', 'テンプレート', 'YY-M', '担当者：米田様', true]);
    console.log('「東邦家具」連携設定を追加しました');
  } else {
    console.log('「東邦家具」連携設定は既に存在します');
  }

  var rate = getSheet(ESTIMATE_RATE_SHEET);
  var rateData = rate.getDataRange().getValues();
  function ensureRate(category, keyword, price, unit, note) {
    for (var i = 1; i < rateData.length; i++) {
      if (String(rateData[i][0]).trim() === 'toho' && String(rateData[i][1]).trim() === category && String(rateData[i][2]).trim() === keyword) return;
    }
    rate.appendRow(['toho', category, keyword, price, unit, note]);
  }
  ensureRate('修正', 'デフォルト', 5000,  'H', '修正は時間単価');
  ensureRate('新規', 'デフォルト', 10000, '枚', '新規デフォルト');
  ensureRate('新規', 'カウンター', 15000, '枚', 'カウンター系');
  ensureRate('新規', '建具',       20000, '枚', '建具系');
  console.log('東邦家具の単価設定を追加（既存はそのまま）');
}

// ==========================================
// SECTION: 診断
// ==========================================

// テンプレートシートの構造を確認（セル位置調整のため）
function diagnoseEstimateTemplate(settingId) {
  var settings = getEstimateLinkSettings();
  var setting = null;
  for (var i = 0; i < settings.length; i++) if (settings[i].settingId === (settingId || 'toho')) { setting = settings[i]; break; }
  if (!setting) { console.error('設定ID未登録:', settingId); return; }

  var ss = SpreadsheetApp.openById(setting.estimateSsId);
  var tpl = ss.getSheetByName(setting.templateSheet);
  if (!tpl) { console.error('テンプレシート未存在:', setting.templateSheet); return; }
  console.log('対象テンプレ:', ss.getName(), '/', tpl.getName(), '行数:', tpl.getLastRow(), '列数:', tpl.getLastColumn());

  // 「小計」「消費税」「合計」「税込合計」の位置
  ['小計','消費税','税込合計','合計','名称','品名','項目'].forEach(function(label){
    var loc = findCellByText(tpl, label);
    if (loc) console.log('  「' + label + '」位置: 行' + loc.row + ' 列' + loc.col + ' (R' + loc.row + 'C' + loc.col + ')');
  });

  // 先頭25行を表示
  var n = Math.min(25, tpl.getLastRow());
  if (n > 0) {
    var d = tpl.getRange(1, 1, n, tpl.getLastColumn()).getValues();
    for (var i = 0; i < n; i++) {
      console.log('行' + (i + 1) + ':', JSON.stringify(d[i]));
    }
  }
}

// LINE送信を伴わないドライラン
function testEstimateDryRun() {
  var setting = findEstimateSettingFromMessage('東邦家具 5月分の見積書作って');
  console.log('setting:', setting && setting.name);
  var ym = parseEstimateTargetDate('東邦家具 5月分の見積書作って');
  console.log('ym:', ym);
  if (!setting || !ym) return;
  var processData = getProcessRowsBySetting(setting);
  console.log('工程ヘッダー:', processData.headers);
  console.log('工程行数:', processData.rows.length);
  var all = normalizeProcessRows(processData);
  console.log('正規化後:', all.length);
  var filtered = filterRowsByTargetMonth(all, ym.year, ym.month);
  console.log('対象月:', filtered.length);
  var groups = groupRowsByProperty(filtered);
  console.log('物件数:', groups.length);
  var data = calculateEstimateData(groups, getEstimateRateSettings(), setting.settingId);
  console.log('小計:', data.subtotal, '税:', data.tax, '合計:', data.total);
  console.log('要確認:', data.confirmItems.length, '件');
  data.confirmItems.slice(0, 5).forEach(function(c){ console.log(' -', c.property, c.drawingName, c.category, c.reason); });
}
