// ===== POS レジ GAS スクリプト (main用) =====
const SS_ID = '1WopPkAIZJSLjc9kr7C5hKvWCPYN7CYmOiXOvko0Q4iI'; // デフォルト（フォールバック用）

// PropertiesServiceからSS_IDを取得（切り替え対応）
function getSSId() {
  return PropertiesService.getScriptProperties().getProperty('CURRENT_SS_ID') || SS_ID;
}
function getTargetSS() {
  return SpreadsheetApp.openById(getSSId());
}

// 列構成 A(1)〜P(16):
// A:日時, B:売上合計, C:商品名, D:カテゴリ, E:数量, F:単価, G:小計,
// H:人数, I:割引, J:支払方法, K:年齢層, L:国籍, M:天気, N:メモ, O:売上ID, P:端末名
var HEADERS = ['日時','売上合計','商品名','カテゴリ','数量','単価','小計','人数','割引','支払方法','年齢層','国籍','天気','メモ','売上ID','端末名'];

// 集計列 (データ16列 + スペーサーQ(17) の後)
const SUMMARY_COL = 18; // R列: 商品別
const AGE_COL     = 23; // W列: 年齢層別
const NAT_COL     = 27; // AA列: 国籍別

// ==================== エントリーポイント ====================

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === 'getCarryOver') {
      return getCarryOverTotal(e.parameter.date || '');
    }
    var d = (e && e.parameter && e.parameter.d) ? JSON.parse(decodeURIComponent(e.parameter.d)) : null;
    if (!d) return ok();
    if (d.type === 'test') return ok();
    var result = handleRequest(d);
    if (result) {
      return ContentService.createTextOutput(JSON.stringify({status:'ok', data:result}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ok();
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error', message:err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var d = (e && e.postData) ? JSON.parse(e.postData.contents) : null;
    if (!d) return ok();
    if (d.type === 'test') return ok();
    var result = handleRequest(d);
    if (result) {
      return ContentService.createTextOutput(JSON.stringify({status:'ok', data:result}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ok();
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error', message:err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function ok() {
  return ContentService.createTextOutput(JSON.stringify({status:'ok'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function writeSyncStatus(ss, msg) {
  var s = ss.getSheetByName('同期ログ') || ss.insertSheet('同期ログ');
  s.getRange(1, 1).setValue(msg);
  SpreadsheetApp.flush();
}

// 前日繰越合計を返す（GETリクエスト用）
function getCarryOverTotal(dateStr) {
  var total = 0;
  try {
    var ss = getTargetSS();
    var sheet = ss.getSheetByName(dateStr + '売上');
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow >= 4) {
        var bVals = sheet.getRange(4, 2,  lastRow - 3, 1).getValues();
        var nVals = sheet.getRange(4, 14, lastRow - 3, 1).getValues();
        for (var i = 0; i < bVals.length; i++) {
          if (bVals[i][0] && String(nVals[i][0]).indexOf('繰越') !== -1) {
            total += Number(bVals[i][0]) || 0;
          }
        }
      }
    }
  } catch(err) {}
  return ContentService.createTextOutput(JSON.stringify({status:'ok', total: total}))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleRequest(data) {
  var type = data.type;
  if      (type === 'add_rows')        addRowsWithCheck(data.rows);
  else if (type === 'delete_rows')     deleteRows(data.sale_id);
  else if (type === 'replace_rows')    replaceRows(data.sale_id, data.rows);
  else if (type === 'clear_sheets')    clearSheets();
  else if (type === 'monthly_summary') return { sheet: createMonthlySummary(data.year, data.month) };
  else if (type === 'bulk_sync') {
    if (data.clear) clearSheets();
    var sales = data.sales || [];
    for (var i = 0; i < sales.length; i++) { if (sales[i] && sales[i].length) addRows(sales[i]); }
  }
  else if (type === 'bulk_merge') {
    var ss = getTargetSS();
    var sheets = ss.getSheets();
    var bSales = data.sales || [];
    var dateLabel = (bSales.length > 0 && bSales[0] && bSales[0].length > 0) ? sheetNameFromRows(bSales[0]) : '同期';
    writeSyncStatus(ss, '🔄 ' + dateLabel + ' 削除中...');
    for (var i = 0; i < bSales.length; i++) {
      var bRows = bSales[i];
      if (!bRows || !bRows.length) continue;
      var bId = String(bRows[0][14]);
      if (bId) { for (var j = 0; j < sheets.length; j++) deleteRowsFromSheet(sheets[j], bId); }
    }
    writeSyncStatus(ss, '🔄 ' + dateLabel + ' 書込中...');
    for (var i = 0; i < bSales.length; i++) { if (bSales[i] && bSales[i].length) addRows(bSales[i]); }
    writeSyncStatus(ss, '✅ ' + dateLabel + ' 完了 ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm:ss'));
    sortSheetsByDate(getTargetSS());
  }
  else if (type === 'get_all_sales') {
    var ssLog = getTargetSS();
    writeSyncStatus(ssLog, '🔄 全データ読込中...');
    var rows = getAllSalesRows();
    writeSyncStatus(ssLog, '✅ 読込完了 ' + rows.length + '行 ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm:ss'));
    return { rows: rows };
  }
  else if (type === 'switch_ss') {
    if (!data.ss_id) return { error: 'ss_id missing' };
    PropertiesService.getScriptProperties().setProperty('CURRENT_SS_ID', data.ss_id);
    return { switched: true, ss_id: data.ss_id };
  }
  else if (type === 'get_current_ss') {
    return { ss_id: getSSId(), is_default: (getSSId() === SS_ID) };
  }
  else if (type === 'create_invoice') {
    var sheetName = createInvoiceSheet(data);
    if (data.record_sale && data.rows && data.rows.length) addRows(data.rows);
    return { sheet: sheetName };
  }
  return null;
}

function getAllSalesRows() {
  var ss = getTargetSS();
  var sheets = ss.getSheets();
  var allRows = [];
  for (var s = 0; s < sheets.length; s++) {
    var name = sheets[s].getName();
    if (name.indexOf('月別集計') !== -1) continue;
    if (!name.match(/売上$/)) continue;
    var lastRow = sheets[s].getLastRow();
    if (lastRow < 4) continue;
    var rows = sheets[s].getRange(4, 1, lastRow - 3, 16).getValues();
    for (var i = 0; i < rows.length; i++) allRows.push(rows[i]);
  }
  return allRows;
}

// 同じ売上IDが当日シートに存在する場合はスキップ（重複防止）
function addRowsWithCheck(rows) {
  if (!rows || !rows.length) return;
  var saleId = String(rows[0][14] || '');
  if (!saleId || saleId === 'undefined') { addRows(rows); return; }
  var ss = getTargetSS();
  var sheetName = sheetNameFromRows(rows);
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 4) {
      var idVals = sheet.getRange(4, 15, lastRow - 3, 1).getValues();
      for (var j = 0; j < idVals.length; j++) {
        if (String(idVals[j][0]) === saleId) return;
      }
    }
  }
  addRows(rows);
}

// 全シートの重複行（同じ売上ID）を削除（初出のみ残す）
function cleanDuplicates() {
  var ss = getTargetSS();
  var sheets = ss.getSheets();
  var seen = {};
  var totalDeleted = 0;
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var name = sheet.getName();
    if (name.indexOf('月別集計') !== -1) continue;
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) continue;
    var idVals = sheet.getRange(4, 15, lastRow - 3, 1).getValues();
    var toDelete = [];
    for (var j = 0; j < idVals.length; j++) {
      var id = String(idVals[j][0] || '');
      if (!id || id === 'undefined') continue;
      if (seen[id]) { toDelete.push(4 + j); } else { seen[id] = true; }
    }
    for (var k = toDelete.length - 1; k >= 0; k--) {
      sheet.deleteRow(toDelete[k]);
      totalDeleted++;
    }
    if (toDelete.length > 0) updateSummary(sheet);
  }
  Logger.log('cleanDuplicates: ' + totalDeleted + '行削除');
  return totalDeleted;
}

function sortSheetsByDate(ss) {
  var sheets = ss.getSheets();
  var dated = [], other = [];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    var m = name.match(/^(?:(\d{4})\/)?(\d+)\/(\d+)/);
    if (m) {
      var yr = m[1] ? parseInt(m[1]) : new Date().getFullYear();
      dated.push({ sheet: sheets[i], key: yr * 10000 + parseInt(m[2]) * 100 + parseInt(m[3]) });
    } else {
      other.push(sheets[i]);
    }
  }
  dated.sort(function(a, b) { return b.key - a.key; }); // 新しい日付が前
  var pos = 1;
  for (var i = 0; i < dated.length; i++) {
    ss.setActiveSheet(dated[i].sheet);
    ss.moveActiveSheet(pos++);
  }
  for (var i = 0; i < other.length; i++) {
    ss.setActiveSheet(other[i]);
    ss.moveActiveSheet(pos++);
  }
}

// ==================== シート名 ====================

function sheetNameFromRows(rows) {
  var dtStr = String(rows[0][0] || '');
  var slash = dtStr.indexOf('/');
  var space = dtStr.indexOf(' ');
  if (slash > 0 && space > slash) {
    return dtStr.substring(0, space) + '売上';
  }
  var now = new Date();
  return (now.getMonth()+1) + '/' + now.getDate() + '売上';
}

// ==================== ヘッダー設定 ====================

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeaders(sheet) {
  // Row1・Row2は常に更新（B2入力値は触らない）
  setTotalsFormulas(sheet);
  setupCashInputRow(sheet);
  // Row3ヘッダーと固定行は初回のみ
  if (sheet.getFrozenRows() < 3) {
    var hRange = sheet.getRange(3, 1, 1, HEADERS.length);
    hRange.setValues([HEADERS]);
    hRange.setFontWeight('bold').setBackground('#f5f5f5').setHorizontalAlignment('center');
    sheet.setFrozenRows(3);
  }
}

// Row1: 集計計算式
function setTotalsFormulas(sheet) {
  sheet.getRange(1, 1, 1, 17).clearContent().clearFormat();

  // 総件数・総売上・総人数: 背景なし
  sheet.getRange(1, 1).setValue('総件数');
  sheet.getRange(1, 2).setFormula('=COUNTA(B4:B)');
  sheet.getRange(1, 3).setValue('総売上');
  sheet.getRange(1, 4).setFormula('=SUM(B4:B)');
  sheet.getRange(1, 4).setNumberFormat('#,##0');
  sheet.getRange(1, 5).setValue('総人数');
  sheet.getRange(1, 6).setFormula('=SUM(H4:H)');
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setHorizontalAlignment('center');

  // クレジット: クレジット色
  sheet.getRange(1, 7).setValue('クレジット');
  sheet.getRange(1, 8).setFormula('=SUMIF(J4:J,"クレジットカード",B4:B)');
  sheet.getRange(1, 8).setNumberFormat('#,##0');
  sheet.getRange(1, 7, 1, 2).setFontWeight('bold').setBackground('#f0f8ff').setHorizontalAlignment('center');

  // 電子決済: 電子決済色
  sheet.getRange(1, 9).setValue('電子決済');
  sheet.getRange(1, 10).setFormula('=SUMIF(J4:J,"電子決済",B4:B)');
  sheet.getRange(1, 10).setNumberFormat('#,##0');
  sheet.getRange(1, 9, 1, 2).setFontWeight('bold').setBackground('#fdf5ff').setHorizontalAlignment('center');

  // 現金: 背景なし（白）
  sheet.getRange(1, 11).setValue('現金');
  sheet.getRange(1, 12).setFormula('=SUMIF(J4:J,"現金",B4:B)');
  sheet.getRange(1, 12).setNumberFormat('#,##0');
  sheet.getRange(1, 11, 1, 2).setFontWeight('bold').setHorizontalAlignment('center');
}

// Row2: レジ現金入力行
function setupCashInputRow(sheet) {
  sheet.getRange(2, 1, 1, 5).setHorizontalAlignment('center');
  sheet.getRange(2, 1).setValue('全レジ現金-売上現金（手入力）').setFontWeight('bold');
  sheet.getRange(2, 2).setBackground('#fff9c4').setFontWeight('bold');
  // C2: 全レジ現金 − 現金売上 = 差額
  sheet.getRange(2, 3).setFormula('=IF(B2="","",B2-SUMIF(J4:J,"現金",B4:B))');
  sheet.getRange(2, 3).setNumberFormat('+#,##0;-#,##0;');
  // D2: ラベル, E2: 対10万差額（ゼロ=0、常に赤太字）
  sheet.getRange(2, 4).setValue('対10万差額').setFontWeight('bold');
  sheet.getRange(2, 5).setFormula('=IF(B2="","",B2-100000)');
  sheet.getRange(2, 5).setNumberFormat('+#,##0;-#,##0;0').setFontColor('#B22222').setFontWeight('bold');
  // G2: クレ+電子ラベル, H2: 合計金額
  sheet.getRange(2, 7).setValue('クレ+電子').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 8).setFormula('=SUMIF(J4:J,"クレジットカード",B4:B)+SUMIF(J4:J,"電子決済",B4:B)');
  sheet.getRange(2, 8).setNumberFormat('#,##0').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#e8e4ff');
  // K2: 繰越合計ラベル, L2: SUMIF（繰越を含む行の売上合計）
  sheet.getRange(2, 11).setValue('繰越合計').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#fff3cd');
  sheet.getRange(2, 12).setFormula('=SUMIF(N4:N,"*繰越*",B4:B)');
  sheet.getRange(2, 12).setNumberFormat('#,##0').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#fff3cd');
}

// ==================== 行操作 ====================

function addRows(rows) {
  if (!rows || !rows.length) return;
  var ss = getTargetSS();
  var sheetName = sheetNameFromRows(rows);
  var sheet = getOrCreateSheet(ss, sheetName);
  ensureHeaders(sheet);

  var lastRow = sheet.getLastRow();
  var startRow = Math.max(lastRow + 1, 4);
  var colCount = rows[0].length;

  sheet.getRange(startRow, 1, rows.length, colCount).setValues(rows);
  sheet.getRange(startRow, 1, rows.length, colCount).setHorizontalAlignment('center');

  for (var i = 0; i < rows.length; i++) {
    applyPaymentColors(sheet, startRow + i, rows[i][9], rows[i][13]); // J列: 支払方法, N列: メモ
  }

  applyGroupBorders(sheet, startRow, rows.length, colCount);
  SpreadsheetApp.flush();
  setDataColumnWidths(sheet);
  updateSummary(sheet);
  sortSheetsByDate(ss);
}

function replaceRows(saleId, rows) {
  if (!saleId) return;
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e) {}
  try {
    var ss = getTargetSS();
    if (rows && rows.length) {
      var sheetName = sheetNameFromRows(rows);
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) deleteRowsFromSheet(sheet, saleId);
    } else {
      var sheets = ss.getSheets();
      for (var i = 0; i < sheets.length; i++) deleteRowsFromSheet(sheets[i], saleId);
    }
    if (rows && rows.length) addRows(rows);
  } finally {
    lock.releaseLock();
  }
}

function deleteRows(saleId) {
  if (!saleId) return;
  var ss = getTargetSS();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    deleteRowsFromSheet(sheets[i], saleId);
  }
}

function deleteRowsFromSheet(sheet, saleId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return false;
  var idCol = 15; // O列: 売上ID
  var vals = sheet.getRange(4, idCol, lastRow - 3, 1).getValues();
  var toDelete = [];
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(saleId)) toDelete.push(4 + i);
  }
  for (var j = toDelete.length - 1; j >= 0; j--) {
    sheet.deleteRow(toDelete[j]);
  }
  if (toDelete.length > 0) updateSummary(sheet);
  return toDelete.length > 0;
}

function clearSheets() {
  var ss = getTargetSS();
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var lastRow = sheet.getLastRow();
    if (lastRow >= 4) sheet.deleteRows(4, lastRow - 3);
    sheet.setFrozenRows(0);
    setTotalsFormulas(sheet);
    setupCashInputRow(sheet);
    var hRange = sheet.getRange(3, 1, 1, HEADERS.length);
    hRange.setValues([HEADERS]).setFontWeight('bold').setBackground('#f5f5f5').setHorizontalAlignment('center');
    sheet.setFrozenRows(3);
    setDataColumnWidths(sheet);
    updateSummary(sheet);
  }
}

// ==================== 書式 ====================

function setDataColumnWidths(sheet) {
  sheet.setColumnWidth(1,  240);  // A: 日時 + Row2ラベル
  sheet.setColumnWidth(2,   85);  // B: 売上合計
  sheet.setColumnWidth(3,  150);  // C: 商品名
  sheet.setColumnWidth(4,  100);  // D: カテゴリ
  sheet.setColumnWidth(5,   65);  // E: 数量
  sheet.setColumnWidth(6,   70);  // F: 単価
  sheet.setColumnWidth(7,   85);  // G: 小計
  sheet.setColumnWidth(8,   55);  // H: 人数
  sheet.setColumnWidth(9,  100);  // I: 割引
  sheet.setColumnWidth(10, 130);  // J: 支払方法
  sheet.setColumnWidth(11, 125);  // K: 年齢層
  sheet.setColumnWidth(12,  85);  // L: 国籍
  sheet.setColumnWidth(13,  60);  // M: 天気
  sheet.setColumnWidth(14, 120);  // N: メモ
  sheet.setColumnWidth(15, 155);  // O: 売上ID
  sheet.setColumnWidth(16,  80);  // P: 端末名
  sheet.setColumnWidth(17,  20);  // Q: スペーサー
}

function setSummaryColumnWidths(sheet) {
  sheet.setColumnWidth(18, 150);  // R: 商品名
  sheet.setColumnWidth(19, 100);  // S: カテゴリ
  sheet.setColumnWidth(20,  60);  // T: 数量
  sheet.setColumnWidth(21,  90);  // U: 金額合計
  sheet.setColumnWidth(22,  20);  // V: スペーサー
  sheet.setColumnWidth(23,  80);  // W: 年齢層
  sheet.setColumnWidth(24,  90);  // X: 件数（組）
  sheet.setColumnWidth(25,  60);  // Y: 人数
  sheet.setColumnWidth(26,  20);  // Z: スペーサー
  sheet.setColumnWidth(27,  80);  // AA: 国籍
  sheet.setColumnWidth(28,  90);  // AB: 件数（組）
  sheet.setColumnWidth(29,  60);  // AC: 人数
}

function applyGroupBorders(sheet, startRow, rowCount, colCount) {
  var range = sheet.getRange(startRow, 1, rowCount, colCount);
  range.setBorder(true, true, true, true, false, null, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  range.setBorder(true, true, true, true, null,  null, '#333333', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function applyPaymentColors(sheet, row, payment, memo) {
  var bg = '';
  if (String(memo || '').indexOf('繰越') !== -1) {
    bg = '#fff3cd';
  } else if (payment === 'クレジットカード') {
    bg = '#f0f8ff';
  } else if (payment === '電子決済') {
    bg = '#fdf5ff';
  }
  if (bg) sheet.getRange(row, 1, 1, 16).setBackground(bg);
}

// ==================== 集計テーブル ====================

function updateSummary(sheet) {
  var lastDataRow = sheet.getLastRow();
  var clearRows = Math.max(lastDataRow, 10);
  sheet.getRange(1, SUMMARY_COL, clearRows, 12).clearContent().clearFormat();

  if (lastDataRow < 4) {
    SpreadsheetApp.flush();
    setSummaryColumnWidths(sheet);
    return;
  }

  var data = sheet.getRange(4, 1, lastDataRow - 3, 16).getValues();

  var itemMap = {};
  var ageMap  = {};
  var natMap  = {};
  var txSeen  = {};

  for (var i = 0; i < data.length; i++) {
    var r        = data[i];
    var itemName = r[2];   // C: 商品名
    var qty      = r[4];   // E: 数量
    var unitPrice= r[5];   // F: 単価
    var cat      = r[3];   // D: カテゴリ
    var jinzu    = r[7];   // H: 人数
    var ageStr   = r[10];  // K: 年齢層
    var natStr   = r[11];  // L: 国籍
    var txId     = r[14];  // O: 売上ID

    if (itemName) {
      var mKey = itemName + '\t' + cat;
      if (!itemMap[mKey]) itemMap[mKey] = {name: itemName, cat: cat, qty: 0, total: 0};
      itemMap[mKey].qty   += (Number(qty)       || 0);
      itemMap[mKey].total += (Number(unitPrice)  || 0) * (Number(qty) || 0);
    }

    if (txId && !txSeen[txId]) {
      txSeen[txId] = true;
      var people = Number(jinzu) || 0;

      if (ageStr) {
        var ages = String(ageStr).split('・');
        for (var a = 0; a < ages.length; a++) {
          var age = ages[a].trim();
          if (!age) continue;
          if (!ageMap[age]) ageMap[age] = {groups: 0, people: 0};
          ageMap[age].groups++;
          ageMap[age].people += people;
        }
      }

      if (natStr) {
        var nats = String(natStr).split('・');
        for (var n = 0; n < nats.length; n++) {
          var nat = nats[n].trim();
          if (!nat) continue;
          if (!natMap[nat]) natMap[nat] = {groups: 0, people: 0};
          natMap[nat].groups++;
          natMap[nat].people += people;
        }
      }
    }
  }

  var hBg = '#e8f5e9';

  // --- 商品別 (R列=18) ---
  var row = 1;
  sheet.getRange(row, SUMMARY_COL, 1, 4)
    .setValues([['商品名','カテゴリ','数量','金額合計']])
    .setBackground(hBg).setFontWeight('bold').setHorizontalAlignment('center');
  row++;
  var iKeys = Object.keys(itemMap);
  for (var k = 0; k < iKeys.length; k++) {
    var v = itemMap[iKeys[k]];
    sheet.getRange(row, SUMMARY_COL, 1, 4)
      .setValues([[v.name, v.cat, v.qty, v.total]])
      .setHorizontalAlignment('center');
    row++;
  }

  // --- 年齢層別 (W列=23) ---
  var aRow = 1;
  sheet.getRange(aRow, AGE_COL, 1, 3)
    .setValues([['年齢層','件数（組）','人数']])
    .setBackground(hBg).setFontWeight('bold').setHorizontalAlignment('center');
  aRow++;
  var aKeys = Object.keys(ageMap);
  for (var k = 0; k < aKeys.length; k++) {
    var v = ageMap[aKeys[k]];
    sheet.getRange(aRow, AGE_COL, 1, 3)
      .setValues([[aKeys[k], v.groups, v.people]])
      .setHorizontalAlignment('center');
    aRow++;
  }

  // --- 国籍別 (AA列=27) ---
  var nRow = 1;
  sheet.getRange(nRow, NAT_COL, 1, 3)
    .setValues([['国籍','件数（組）','人数']])
    .setBackground(hBg).setFontWeight('bold').setHorizontalAlignment('center');
  nRow++;
  var nKeys = Object.keys(natMap);
  for (var k = 0; k < nKeys.length; k++) {
    var v = natMap[nKeys[k]];
    sheet.getRange(nRow, NAT_COL, 1, 3)
      .setValues([[nKeys[k], v.groups, v.people]])
      .setHorizontalAlignment('center');
    nRow++;
  }

  SpreadsheetApp.flush();
  setSummaryColumnWidths(sheet);
}

// ==================== 月別集計 ====================

// 毎月1日朝6時に前月集計を自動生成するタイマーから呼ばれる
function runMonthlySummary() {
  var now = new Date();
  var month = now.getMonth(); // 0-indexed → 0=1月なので前月はそのまま使える
  var year  = now.getFullYear();
  if (month === 0) { month = 12; year -= 1; } // 1月1日なら前年12月
  createMonthlySummary(year, month);
}

// シート名一覧をログに出力してデバッグ確認用（GASエディタから手動実行）
function debugSheetNames() {
  var ss = getTargetSS();
  var sheets = ss.getSheets();
  var names = sheets.map(function(s){ return s.getName(); });
  Logger.log(names.join('\n'));
}

// GASの「トリガー」画面で一度だけ実行するとタイマーが登録される
function setupMonthlyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === 'runMonthlySummary') {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }
  ScriptApp.newTrigger('runMonthlySummary')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📅 集計')
    .addItem('月別集計を作成...', 'promptMonthlySummary')
    .addSeparator()
    .addItem('📂 新しい月のSSに切り替え...', 'promptSwitchSpreadsheet')
    .addToUi();
}

function promptSwitchSpreadsheet() {
  var currentId = getSSId();
  var isDefault = (currentId === SS_ID);
  var html = '<style>'
    + 'body{font-family:sans-serif;padding:14px;margin:0;font-size:13px;}'
    + 'input{width:100%;padding:9px;border:1px solid #ccc;border-radius:6px;font-size:12px;box-sizing:border-box;margin-top:4px;}'
    + '.cur{background:#f5f5f5;border-radius:6px;padding:8px 10px;font-size:11px;color:#555;word-break:break-all;margin-bottom:10px;}'
    + '.btn{width:100%;padding:10px;margin-top:8px;border:none;border-radius:6px;font-size:13px;font-weight:700;cursor:pointer;}'
    + '#res{margin-top:8px;font-weight:bold;min-height:18px;}'
    + '</style>'
    + '<div class="cur">現在のSS ID:<br><b>' + currentId + '</b>' + (isDefault ? '（デフォルト）' : '') + '</div>'
    + '<label>新しいスプレッドシートのID</label>'
    + '<input id="ssid" placeholder="スプレッドシートのIDを貼り付け" />'
    + '<button class="btn" style="background:#2d5016;color:#fff;" onclick="doSwitch()">切り替える</button>'
    + '<div id="res"></div>'
    + '<button class="btn" style="background:#eee;color:#333;" onclick="google.script.host.close()">閉じる</button>'
    + '<script>function extractId(s){var m=s.match(/\\/spreadsheets\\/d\\/([a-zA-Z0-9_-]+)/);return m?m[1]:s.trim();}function doSwitch(){'
    + 'var raw=document.getElementById("ssid").value.trim();var id=extractId(raw);'
    + 'if(!id){document.getElementById("res").textContent="⚠ URLまたはIDを入力してください";return;}'
    + 'document.getElementById("res").textContent="⏳ 切り替え中...";'
    + 'google.script.run'
    + '.withSuccessHandler(function(r){document.getElementById("res").textContent="✅ 切り替えました";setTimeout(google.script.host.close,1500);})'
    + '.withFailureHandler(function(e){document.getElementById("res").textContent="❌ "+e.message;})'
    + '.switchSpreadsheetFromDialog(id);'
    + '}<\/script>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(360).setHeight(280), '📂 SSの切り替え'
  );
}

function switchSpreadsheetFromDialog(newSsId) {
  if (!newSsId) throw new Error('IDが空です');
  // 存在確認
  SpreadsheetApp.openById(newSsId);
  PropertiesService.getScriptProperties().setProperty('CURRENT_SS_ID', newSsId);
}

function promptMonthlySummary() {
  var now = new Date();
  var btns = '';
  for (var i = 0; i < 6; i++) {
    var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    var y = d.getFullYear(), m = d.getMonth() + 1;
    btns += '<button class="mb" onclick="run(' + y + ',' + m + ')">' + y + '年' + m + '月</button>';
  }
  var html = '<style>'
    + 'body{font-family:sans-serif;padding:12px;margin:0;}'
    + '.mb{display:block;width:100%;padding:11px;margin:5px 0;font-size:14px;cursor:pointer;border:1px solid #bbb;border-radius:6px;background:#f8f9fa;text-align:left;}'
    + '.mb:hover{background:#e8f0e0;border-color:#2d5016;}'
    + '#res{margin-top:10px;font-weight:bold;min-height:20px;color:#2d5016;}'
    + '#cls{display:block;width:100%;padding:9px;margin-top:8px;font-size:13px;cursor:pointer;border:1px solid #ccc;border-radius:6px;background:#eee;}'
    + '</style>'
    + btns
    + '<div id="res"></div>'
    + '<button id="cls" onclick="google.script.host.close()">閉じる</button>'
    + '<script>function run(y,m){'
    + 'document.getElementById("res").textContent="⏳ 集計中...";'
    + 'google.script.run'
    + '.withSuccessHandler(function(s){document.getElementById("res").textContent="✅ "+s+" 作成完了";})'
    + '.withFailureHandler(function(e){document.getElementById("res").textContent="❌ "+e.message;})'
    + '.createMonthlySummary(y,m);'
    + '}<\/script>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(260).setHeight(430), '📅 月別集計'
  );
}

function debugDiscValues(yearMonth) {
  var m=(yearMonth||'2026/5').match(/(\d{4})[\/\-](\d{1,2})/);
  var year=parseInt(m[1]),month=parseInt(m[2]);
  var ss=getTargetSS(),sheets=ss.getSheets();
  var txData={};
  for (var s=0;s<sheets.length;s++){
    var name=sheets[s].getName();
    if (name.indexOf('月別集計')!==-1) continue;
    var match=name.match(/^(?:\d{4}\/)?(\d+)\/(\d+)/);
    if (!match||parseInt(match[1])!==month) continue;
    var sheet=sheets[s],lastRow=sheet.getLastRow();
    var row2=lastRow>=2?sheet.getRange(2,1,1,3).getValues()[0]:[];
    var isOld=String(row2[1]||'').indexOf('曜日')!==-1;
    var dataStart=isOld?3:4,co=isOld?1:0;
    if (lastRow<dataStart) continue;
    var data=sheet.getRange(dataStart,1,lastRow-dataStart+1,17).getValues();
    var dateStr=month+'/'+parseInt(match[2]);
    for (var i=0;i<data.length;i++){
      var r=data[i];
      var txId=String(r[14+co]||'');
      if (!txId||txId==='undefined') continue;
      if (!txData[txId]){
        txData[txId]={date:dateStr,amount:Number(r[1+co])||0,disc:String(r[8+co]||'')};
      } else if (String(r[8+co]||'')!=='') {
        txData[txId].disc=String(r[8+co]);
      }
    }
  }
  var discCounts={};
  Object.keys(txData).forEach(function(txId){
    var d=txData[txId],key=d.date+'  disc=「'+d.disc+'」  amount='+d.amount;
    discCounts[key]=(discCounts[key]||0)+1;
  });
  var lines=Object.keys(discCounts).sort();
  Logger.log('=== 割引フィールド一覧（取引単位）===\n'+lines.map(function(k){return discCounts[k]+'件  '+k;}).join('\n'));
}

function debugDate(yearMonthDay) {
  var m=(yearMonthDay||'2026/5/3').match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  var year=parseInt(m[1]),month=parseInt(m[2]),day=parseInt(m[3]);
  var ss=getTargetSS(),sheets=ss.getSheets();
  for (var s=0;s<sheets.length;s++){
    var name=sheets[s].getName();
    if (name.indexOf('月別集計')!==-1) continue;
    var match=name.match(/^(?:\d{4}\/)?(\d+)\/(\d+)/);
    if (!match||parseInt(match[1])!==month||parseInt(match[2])!==day) continue;
    var sheet=sheets[s],lastRow=sheet.getLastRow();
    var row2=lastRow>=2?sheet.getRange(2,1,1,3).getValues()[0]:[];
    var isOld=String(row2[1]||'').indexOf('曜日')!==-1;
    var dataStart=isOld?3:4,co=isOld?1:0;
    if (lastRow<dataStart) continue;
    var data=sheet.getRange(dataStart,1,lastRow-dataStart+1,17).getValues();
    var txData={};
    for (var i=0;i<data.length;i++){
      var r=data[i],txId=String(r[14+co]||'');
      if (!txId||txId==='undefined') continue;
      var amount=Number(r[1+co])||0;
      var itemName=String(r[2+co]||'');
      var qty=Number(r[4+co])||0,unitPrice=Number(r[5+co])||0;
      if (!txData[txId]) txData[txId]={amount:0,itemTotal:0,items:[],rowCount:0};
      txData[txId].rowCount++;
      if (amount>0&&txData[txId].amount===0) txData[txId].amount=amount;
      if (itemName){
        txData[txId].items.push(itemName+'×'+qty+'@'+unitPrice+'='+unitPrice*qty);
        txData[txId].itemTotal+=unitPrice*qty;
      }
    }
    var lines=['=== '+name+' 取引詳細 ==='],totalAmt=0,totalItems=0;
    Object.keys(txData).forEach(function(txId){
      var t=txData[txId],diff=t.itemTotal-t.amount;
      totalAmt+=t.amount;totalItems+=t.itemTotal;
      var flag=diff!==0?'  ★差異='+diff:'';
      lines.push('txId='+txId+' 支払='+t.amount+' 品目計='+t.itemTotal+flag+' ('+t.rowCount+'行)');
      t.items.forEach(function(it){lines.push('  '+it);});
    });
    lines.push('--- 合計: 支払='+totalAmt+' 品目計='+totalItems+' 差='+(totalItems-totalAmt));
    Logger.log(lines.join('\n'));
    return;
  }
  Logger.log('シートが見つかりません: '+yearMonthDay);
}

function debugSonoTa(yearMonth) {
  var m=(yearMonth||'2026/5').match(/(\d{4})[\/\-](\d{1,2})/);
  var year=parseInt(m[1]),month=parseInt(m[2]);
  var ss=getTargetSS(),sheets=ss.getSheets();
  var counts={};
  for (var s=0;s<sheets.length;s++){
    var name=sheets[s].getName();
    if (name.indexOf('月別集計')!==-1) continue;
    var match=name.match(/^(?:\d{4}\/)?(\d+)\/(\d+)/);
    if (!match||parseInt(match[1])!==month) continue;
    var sheet=sheets[s],lastRow=sheet.getLastRow();
    var row2=lastRow>=2?sheet.getRange(2,1,1,3).getValues()[0]:[];
    var isOld=String(row2[1]||'').indexOf('曜日')!==-1;
    var dataStart=isOld?3:4,co=isOld?1:0;
    if (lastRow<dataStart) continue;
    var data=sheet.getRange(dataStart,1,lastRow-dataStart+1,17).getValues();
    for (var i=0;i<data.length;i++){
      var r=data[i];
      var txId=String(r[14+co]||'');
      if (!txId||txId==='undefined') continue;
      var itemName=String(r[2+co]||''),cat=String(r[3+co]||'');
      if (!itemName) continue;
      var pg=getParentGroup(cat);
      if (pg==='その他'){
        var key=cat+'  →  '+itemName;
        counts[key]=(counts[key]||0)+1;
      }
    }
  }
  var lines=Object.keys(counts).sort();
  Logger.log('=== その他 カテゴリ一覧 ===\n'+lines.map(function(k){return counts[k]+'件  '+k;}).join('\n'));
}

function getParentGroupCross(pg, itemName) {
  if (pg === '空中ウォーク' || pg === 'ツリーハウス昼') {
    if (itemName.indexOf('55分') !== -1) return pg + ' 55分';
    if (itemName.indexOf('115分') !== -1) return pg + ' 115分';
    if (itemName.indexOf('1DAY') !== -1) return pg + ' 1DAY';
  }
  return pg;
}

function getParentGroup(cat) {
  if (cat.indexOf('ドーム') !== -1) return 'ドーム';
  if (cat.indexOf('空中ウォーク') !== -1) return '空中ウォーク';
  if (cat.indexOf('空中テント') !== -1) return '空中テント';
  if (cat.indexOf('ツリーハウス') !== -1) return (cat.indexOf('宿泊') !== -1) ? 'ツリーハウス宿泊' : 'ツリーハウス昼';
  if (cat.indexOf('BBQ') !== -1 || cat.indexOf('バーベキュー') !== -1) return 'BBQスペース';
  if (cat.indexOf('ハンモック') !== -1) return 'ハンモック';
  if (cat.indexOf('消耗品') !== -1) return '消耗品';
  if (cat === '備品レンタル') return '備品レンタル';
  if (cat === '受付飲食') return '受付飲食';
  if (cat === 'グッズ') return 'グッズ';
  if (cat === 'ピクニック飲食') return 'ピクニック飲食';
  if (cat === '無人販売') return '無人販売';
  if (cat === 'ペット') return 'ペット';
  return 'その他（カテゴリ・商品登録なし）';
}

var CROSS_ORDER = ['空中ウォーク 55分','空中ウォーク 115分','空中ウォーク 1DAY','ツリーハウス昼 55分','ツリーハウス昼 115分','ツリーハウス昼 1DAY','BBQスペース','ツリーハウス宿泊','ドーム','空中テント','ハンモック','備品レンタル','受付飲食','グッズ','ピクニック飲食','無人販売','ペット','その他（カテゴリ・商品登録なし）'];
var CAT_ORDER   = ['空中ウォーク','ツリーハウス昼','BBQスペース','ツリーハウス宿泊','ドーム','空中テント','ハンモック','消耗品','備品レンタル','受付飲食','グッズ','ピクニック飲食','無人販売','ペット','その他（カテゴリ・商品登録なし）'];
function sortByOrder(keys, order) {
  keys.sort(function(a, b) {
    var ai = order.indexOf(a), bi = order.indexOf(b);
    if (ai >= 0 && bi >= 0) return ai - bi;
    if (ai >= 0) return -1;
    if (bi >= 0) return 1;
    return a.localeCompare(b, 'ja');
  });
}

var PEOPLE_GROUPS = {'ドーム':true,'空中テント':true,'空中ウォーク':true,'空中ウォーク 55分':true,'空中ウォーク 115分':true,'空中ウォーク 1DAY':true,'ツリーハウス昼':true,'ツリーハウス昼 55分':true,'ツリーハウス昼 115分':true,'ツリーハウス昼 1DAY':true,'BBQスペース':true,'ツリーハウス宿泊':true,'ハンモック':true};
var QTY_GROUPS = {'BBQスペース':true};

var HOLIDAYS_2026 = {'1/1':1,'1/12':1,'2/11':1,'2/23':1,'3/20':1,'4/29':1,'5/3':1,'5/4':1,'5/5':1,'5/6':1,'7/20':1,'8/11':1,'9/21':1,'9/22':1,'9/23':1,'10/12':1,'11/3':1,'11/23':1};
var HOLIDAYS_2025 = {'1/1':1,'1/13':1,'2/11':1,'2/23':1,'2/24':1,'3/20':1,'4/29':1,'5/3':1,'5/4':1,'5/5':1,'5/6':1,'7/21':1,'8/11':1,'9/15':1,'9/23':1,'10/13':1,'11/3':1,'11/23':1,'11/24':1};
function isHoliday(year, month, day) {
  var map = year===2026?HOLIDAYS_2026:year===2025?HOLIDAYS_2025:{};
  return !!map[month+'/'+day];
}
function getDow(year, month, day) {
  return ['日','月','火','水','木','金','土'][new Date(year, month-1, day).getDay()];
}
function dateDowLabel(ds, year, month, day) {
  return ds+'('+getDow(year,month,day)+')'+(isHoliday(year,month,day)?'祝':'');
}
function colLetter(n) {
  var s=''; while(n>0){var r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);} return s;
}

// 月別集計シートを作成して返す
function createMonthlySummary(year, month) {
  var ss = getTargetSS();
  var sheets = ss.getSheets();

  var monthSheets=[];
  for (var s=0;s<sheets.length;s++) {
    var name=sheets[s].getName();
    if (name.indexOf('月別集計')!==-1) continue;
    var match=name.match(/^(?:\d{4}\/)?(\d+)\/(\d+)/);
    if (match&&parseInt(match[1])===month) monthSheets.push({sheet:sheets[s],day:parseInt(match[2])});
  }
  monthSheets.sort(function(a,b){return a.day-b.day;});

  var dailyMap={},dailyCatMap={},dailyCrossMap={},muraDailyCatMap={},passDailyCatMap={},otaDailyCatMap={},stayDailyCatMap={};
  var parentMap={},payMap={},discMap={},ageMap={},natMap={};
  var txSeen={},txDiscMap={},txHasItems={},txUniqueRows={},txProcessedRows={};
  var mura={count:0,people:0,total:0},pass={count:0,people:0,total:0},ota={count:0,people:0,total:0},stay={count:0,people:0,total:0};
  var grandTotal=0,grandCount=0,grandPeople=0;

  for (var si=0;si<monthSheets.length;si++) {
    var sheet=monthSheets[si].sheet,lastRow=sheet.getLastRow();
    var row2=lastRow>=2?sheet.getRange(2,1,1,3).getValues()[0]:[];
    var isOld=String(row2[1]||'').indexOf('曜日')!==-1;
    var dataStart=isOld?3:4,co=isOld?1:0;
    if (lastRow<dataStart) continue;
    var data=sheet.getRange(dataStart,1,lastRow-dataStart+1,17).getValues();

    var pass1RowSeen={};
    for (var pi=0;pi<data.length;pi++){
      var pr=data[pi],ptxId=String(pr[14+co]||'');
      if (!ptxId||ptxId==='undefined') continue;
      var pdisc=String(pr[8+co]||'');
      if (pdisc&&pdisc!==''&&pdisc!=='false') txDiscMap[ptxId]=pdisc;
      var pItemName=String(pr[2+co]||'');
      if (pItemName){
        txHasItems[ptxId]=true;
        var pUnitPrice=Number(pr[5+co])||0,pQty=Number(pr[4+co])||0;
        var pKey=ptxId+'|'+pItemName+'|'+pUnitPrice+'|'+pQty;
        if (!pass1RowSeen[pKey]){
          pass1RowSeen[pKey]=true;
          if (!txUniqueRows[ptxId]) txUniqueRows[ptxId]=0;
          txUniqueRows[ptxId]++;
        }
      }
    }

    var dateStr=month+'/'+monthSheets[si].day;
    if (!dailyMap[dateStr])         dailyMap[dateStr]         = {total:0,count:0,people:0,cash:0,card:0,elec:0,muraAmt:0,passAmt:0,otaAmt:0,stayAmt:0};
    if (!dailyCatMap[dateStr])      dailyCatMap[dateStr]      = {};
    if (!dailyCrossMap[dateStr])    dailyCrossMap[dateStr]    = {};
    if (!muraDailyCatMap[dateStr])  muraDailyCatMap[dateStr]  = {};
    if (!passDailyCatMap[dateStr])  passDailyCatMap[dateStr]  = {};
    if (!otaDailyCatMap[dateStr])   otaDailyCatMap[dateStr]   = {};
    if (!stayDailyCatMap[dateStr])  stayDailyCatMap[dateStr]  = {};

    for (var i=0;i<data.length;i++) {
      var r=data[i];
      var txId=String(r[14+co]||'');
      if (!txId||txId==='undefined') continue;
      var amount=Number(r[1+co])||0,itemName=String(r[2+co]||''),cat=String(r[3+co]||'');
      var qty=Number(r[4+co])||0,unitPrice=Number(r[5+co])||0,people=Number(r[7+co])||0;
      var disc=String(r[8+co]||''),payment=String(r[9+co]||'');
      var ageStr=String(r[10+co]||''),natStr=String(r[11+co]||'');

      var txDisc=txDiscMap[txId]||'';
      var isMura=(txDisc==='mura'||txDisc.indexOf('村')!==-1);
      var isPass=(txDisc==='pass_day'||txDisc==='pass_night'||txDisc.indexOf('パス')!==-1||txDisc.indexOf('pass')!==-1);
      var isOTA=(txDisc==='rakuten'||txDisc==='jalan'||txDisc==='sou'||txDisc.indexOf('楽天')!==-1||txDisc.indexOf('じゃらん')!==-1||txDisc.indexOf('ブッキング')!==-1||txDisc.indexOf('booking')!==-1||txDisc.indexOf('Booking')!==-1||txDisc.indexOf('そうエクスペリエンス')!==-1||txDisc.indexOf('そう体験')!==-1||txDisc.indexOf('sou')!==-1);
      var isStay=(txDisc==='stay_guest'||txDisc.indexOf('宿泊')!==-1);

      if (itemName&&itemName!=='') {
        if (!txProcessedRows[txId]) txProcessedRows[txId]=0;
        if (txProcessedRows[txId]>=(txUniqueRows[txId]||0)) { /* 重複行スキップ */ }
        else {
        txProcessedRows[txId]++;
        var pg=getParentGroup(cat);
        if (!parentMap[pg]) parentMap[pg]={count:0,people:0,total:0,txSet:{},cats:{}};
        var pe=parentMap[pg];
        if (!pe.txSet[txId]){pe.txSet[txId]=true;pe.count++;pe.people+=people;}
        pe.total+=unitPrice*qty;
        if (!pe.cats[cat]) pe.cats[cat]={count:0,people:0,total:0,txSet:{},items:{}};
        var ce=pe.cats[cat];
        if (!ce.txSet[txId]){ce.txSet[txId]=true;ce.count++;ce.people+=people;}
        ce.total+=unitPrice*qty;
        if (!ce.items[itemName]) ce.items[itemName]={name:itemName,qty:0,total:0,people:0,txSet:{}};
        ce.items[itemName].qty+=qty; ce.items[itemName].total+=unitPrice*qty;
        if (!ce.items[itemName].txSet[txId]){ce.items[itemName].txSet[txId]=true;ce.items[itemName].people+=people;}

        if (!dailyCatMap[dateStr][pg]) dailyCatMap[dateStr][pg]={people:0,total:0,txSet:{}};
        var dce=dailyCatMap[dateStr][pg];
        if (!dce.txSet[txId]){dce.txSet[txId]=true;dce.people+=people;} dce.total+=unitPrice*qty;

        var crossKey=getParentGroupCross(pg,itemName);
        if (!dailyCrossMap[dateStr][crossKey]) dailyCrossMap[dateStr][crossKey]={people:0,qty:0,total:0,txSet:{}};
        var dcx=dailyCrossMap[dateStr][crossKey];
        if (!dcx.txSet[txId]){dcx.txSet[txId]=true;dcx.people+=people;} dcx.total+=unitPrice*qty; dcx.qty+=qty;

        if (isMura){
          if (!muraDailyCatMap[dateStr][pg]) muraDailyCatMap[dateStr][pg]={people:0,total:0,txSet:{}};
          var mdce=muraDailyCatMap[dateStr][pg];
          if (!mdce.txSet[txId]){mdce.txSet[txId]=true;mdce.people+=people;} mdce.total+=unitPrice*qty;
        }
        if (isPass){
          if (!passDailyCatMap[dateStr][pg]) passDailyCatMap[dateStr][pg]={people:0,total:0,txSet:{}};
          var pdce=passDailyCatMap[dateStr][pg];
          if (!pdce.txSet[txId]){pdce.txSet[txId]=true;pdce.people+=people;} pdce.total+=unitPrice*qty;
        }
        if (isOTA){
          if (!otaDailyCatMap[dateStr][pg]) otaDailyCatMap[dateStr][pg]={people:0,total:0,txSet:{}};
          var odce=otaDailyCatMap[dateStr][pg];
          if (!odce.txSet[txId]){odce.txSet[txId]=true;odce.people+=people;} odce.total+=unitPrice*qty;
        }
        if (isStay){
          if (!stayDailyCatMap[dateStr][pg]) stayDailyCatMap[dateStr][pg]={people:0,total:0,txSet:{}};
          var sdce=stayDailyCatMap[dateStr][pg];
          if (!sdce.txSet[txId]){sdce.txSet[txId]=true;sdce.people+=people;} sdce.total+=unitPrice*qty;
        }
        }
      }

      if (!txSeen[txId]) {
        txSeen[txId]=true;
        dailyMap[dateStr].total+=amount;dailyMap[dateStr].count+=1;dailyMap[dateStr].people+=people;
        if (payment==='現金') dailyMap[dateStr].cash+=amount;
        else if (payment==='クレジットカード') dailyMap[dateStr].card+=amount;
        else if (payment==='電子決済') dailyMap[dateStr].elec+=amount;
        grandTotal+=amount;grandCount+=1;grandPeople+=people;
        if (!payMap[payment]) payMap[payment]={total:0,count:0};
        payMap[payment].total+=amount;payMap[payment].count+=1;
        var dKey=(disc&&disc!==''&&disc!=='false')?disc:'なし';
        if (!discMap[dKey]) discMap[dKey]={count:0,total:0};
        discMap[dKey].count++;discMap[dKey].total+=amount;
        if (isMura){mura.count++;mura.people+=people;mura.total+=amount;dailyMap[dateStr].muraAmt+=amount;}
        if (isPass){pass.count++;pass.people+=people;pass.total+=amount;dailyMap[dateStr].passAmt+=amount;}
        if (isOTA){ota.count++;ota.people+=people;ota.total+=amount;dailyMap[dateStr].otaAmt+=amount;}
        if (isStay){stay.count++;stay.people+=people;stay.total+=amount;dailyMap[dateStr].stayAmt+=amount;}
        if (ageStr){ageStr.split('・').forEach(function(ag){ag=ag.trim();if(!ag)return;if(!ageMap[ag])ageMap[ag]={groups:0,people:0};ageMap[ag].groups++;ageMap[ag].people+=people;});}
        if (natStr){natStr.split('・').forEach(function(nat){nat=nat.trim();if(!nat)return;if(!natMap[nat])natMap[nat]={groups:0,people:0};natMap[nat].groups++;natMap[nat].people+=people;});}
        if (!txHasItems[txId]&&amount>0){
          var sotKey='その他（カテゴリ・商品登録なし）';
          if (!dailyCatMap[dateStr][sotKey]) dailyCatMap[dateStr][sotKey]={people:0,total:0,txSet:{}};
          if (!dailyCatMap[dateStr][sotKey].txSet[txId]){dailyCatMap[dateStr][sotKey].txSet[txId]=true;dailyCatMap[dateStr][sotKey].total+=amount;}
          if (!dailyCrossMap[dateStr][sotKey]) dailyCrossMap[dateStr][sotKey]={people:0,total:0,txSet:{}};
          if (!dailyCrossMap[dateStr][sotKey].txSet[txId]){dailyCrossMap[dateStr][sotKey].txSet[txId]=true;dailyCrossMap[dateStr][sotKey].total+=amount;}
          if (!parentMap[sotKey]) parentMap[sotKey]={count:0,people:0,total:0,txSet:{},cats:{}};
          if (!parentMap[sotKey].txSet[txId]){parentMap[sotKey].txSet[txId]=true;parentMap[sotKey].count++;parentMap[sotKey].total+=amount;}
          if (isMura){if(!muraDailyCatMap[dateStr][sotKey])muraDailyCatMap[dateStr][sotKey]={people:0,total:0,txSet:{}};if(!muraDailyCatMap[dateStr][sotKey].txSet[txId]){muraDailyCatMap[dateStr][sotKey].txSet[txId]=true;muraDailyCatMap[dateStr][sotKey].total+=amount;}}
          if (isPass){if(!passDailyCatMap[dateStr][sotKey])passDailyCatMap[dateStr][sotKey]={people:0,total:0,txSet:{}};if(!passDailyCatMap[dateStr][sotKey].txSet[txId]){passDailyCatMap[dateStr][sotKey].txSet[txId]=true;passDailyCatMap[dateStr][sotKey].total+=amount;}}
          if (isOTA){if(!otaDailyCatMap[dateStr][sotKey])otaDailyCatMap[dateStr][sotKey]={people:0,total:0,txSet:{}};if(!otaDailyCatMap[dateStr][sotKey].txSet[txId]){otaDailyCatMap[dateStr][sotKey].txSet[txId]=true;otaDailyCatMap[dateStr][sotKey].total+=amount;}}
          if (isStay){if(!stayDailyCatMap[dateStr][sotKey])stayDailyCatMap[dateStr][sotKey]={people:0,total:0,txSet:{}};if(!stayDailyCatMap[dateStr][sotKey].txSet[txId]){stayDailyCatMap[dateStr][sotKey].txSet[txId]=true;stayDailyCatMap[dateStr][sotKey].total+=amount;}}
        }
      }
    }
  }

  var pgAllKeys=Object.keys(parentMap);
  var pgKeysCat=pgAllKeys.slice(); sortByOrder(pgKeysCat,CAT_ORDER);
  var crossKeySet={};
  Object.keys(dailyCrossMap).forEach(function(ds){Object.keys(dailyCrossMap[ds]).forEach(function(k){crossKeySet[k]=true;});});
  var TIME_SPLIT_PARENTS={'空中ウォーク':true,'ツリーハウス昼':true};
  var pgKeysCross=Object.keys(crossKeySet).filter(function(k){return !TIME_SPLIT_PARENTS[k];}); sortByOrder(pgKeysCross,CROSS_ORDER);

  var sheetName=year+'年'+month+'月 月別集計';
  var out=ss.getSheetByName(sheetName);
  if (out){
    var existingCharts=out.getCharts();
    for (var ec=0;ec<existingCharts.length;ec++) out.removeChart(existingCharts[ec]);
    out.clear();
  } else {
    out=ss.insertSheet(sheetName,0);
  }
  var neededCols=Math.max(50,5+pgKeysCross.reduce(function(s,k){return s+(PEOPLE_GROUPS[k]?2:1);},0)+10);
  var currentCols=out.getMaxColumns();
  if (currentCols<neededCols) out.insertColumnsAfter(currentCols,neededCols-currentCols);

  var row=1;
  var C='center';
  var BG_HEAD='#2d5016',BG_SEC='#4a7c2f',BG_TOTAL='#e8f0e0',BG_EVEN='#fafafa',BG_SUB='#f0f4e8',BG_CAT='#d4e6c3',BG_CAT2='#eaf3e0';

  out.getRange(row,1,1,8).merge().setValue(year+'年'+month+'月 月別集計')
    .setFontSize(15).setFontWeight('bold').setBackground(BG_HEAD).setFontColor('#fff').setHorizontalAlignment(C);
  row++;
  var avg=grandPeople>0?Math.round(grandTotal/grandPeople):0;
  out.getRange(row,1,1,8).setValues([['総売上',grandTotal,'件数',grandCount,'人数',grandPeople,'客単価（人）',avg]])
    .setBackground('#f0f4e8').setFontWeight('bold').setHorizontalAlignment(C);
  out.getRange(row,2).setNumberFormat('#,##0');out.getRange(row,8).setNumberFormat('#,##0');
  row+=2;

  function secHead(label,cols,align){
    out.getRange(row,1,1,cols).merge().setValue(label)
      .setFontWeight('bold').setBackground(BG_SEC).setFontColor('#fff').setHorizontalAlignment(align||C);row++;
  }
  function colHead(headers){
    out.getRange(row,1,1,headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#f5f5f5').setHorizontalAlignment(C);row++;
  }

  secHead('【日別売上】　※村民割引など割引差引後の実支払金額',7);
  colHead(['日付','売上合計','件数','人数','現金','カード','電子決済']);
  var days=Object.keys(dailyMap).sort(function(a,b){return parseInt(a.split('/')[1])-parseInt(b.split('/')[1]);});
  var dailyDataStart=row;
  for (var d=0;d<days.length;d++) {
    var dd=dailyMap[days[d]],dayN=parseInt(days[d].split('/')[1]);
    out.getRange(row,1,1,7).setValues([[dateDowLabel(days[d],year,month,dayN),dd.total,dd.count,dd.people,dd.cash,dd.card,dd.elec]]).setHorizontalAlignment(C);
    out.getRange(row,2).setNumberFormat('#,##0');out.getRange(row,5,1,3).setNumberFormat('#,##0');
    if (d%2===0) out.getRange(row,1,1,7).setBackground(BG_EVEN);row++;
  }
  var dailyDataEnd=row-1;
  out.getRange(row,1).setValue('合計');
  for (var c=2;c<=7;c++) out.getRange(row,c).setFormula('=SUM('+colLetter(c)+dailyDataStart+':'+colLetter(c)+dailyDataEnd+')').setNumberFormat('#,##0');
  out.getRange(row,1,1,7).setFontWeight('bold').setBackground(BG_TOTAL).setHorizontalAlignment(C);row+=2;

  function writeCrossTable(label,filteredDays,catMap,usedPgKeys,showAllDays,dailyActMap){
    var n=usedPgKeys.length;
    var eCols=dailyActMap?2:0;
    var catColStart=4+eCols;
    var catCols=[];
    var totalCatCols=0;
    for (var pk=0;pk<n;pk++){var ncc=PEOPLE_GROUPS[usedPgKeys[pk]]?2:1;catCols.push(ncc);totalCatCols+=ncc;}
    var dcCols=(catColStart-1)+totalCatCols;
    var catSalesCols={};
    var csoff=catColStart;
    for (var pk=0;pk<n;pk++){if(catCols[pk]===2)csoff++;catSalesCols[csoff]=true;csoff++;}
    out.getRange(row,1,1,7).merge().setValue(label).setFontWeight('bold').setBackground(BG_SEC).setFontColor('#fff').setHorizontalAlignment(C);
    if (dcCols>7) out.getRange(row,8,1,dcCols-7).setBackground(BG_SEC);
    row++;
    var h1=['日付'];
    if (dailyActMap){h1.push('合計人数','合計売上（定価）','実決済額','定価合計 − 実決済額 ※');}
    else{h1.push('合計','');}
    for (var pk=0;pk<n;pk++){h1.push(usedPgKeys[pk]);if(catCols[pk]===2)h1.push('');}
    out.getRange(row,1,1,dcCols).setValues([h1]).setFontWeight('bold').setBackground('#f5f5f5').setHorizontalAlignment(C);
    if (!dailyActMap) out.getRange(row,2,1,2).merge();
    var moff=catColStart;
    for (var pk=0;pk<n;pk++){if(catCols[pk]===2)out.getRange(row,moff,1,2).merge();moff+=catCols[pk];}
    row++;
    var h2=[''];
    if (dailyActMap){for(var ei=0;ei<catColStart-2;ei++)h2.push('');}else{h2.push('人数','売上');}
    for (var pk=0;pk<n;pk++){if(catCols[pk]===2)h2.push(QTY_GROUPS[usedPgKeys[pk]]?'数量':'人数');h2.push('売上');}
    out.getRange(row,1,1,dcCols).setValues([h2]).setFontWeight('bold').setBackground('#f5f5f5').setHorizontalAlignment(C);
    if (dailyActMap){
      out.getRange(row,5)
        .setValue('※ ±は重複・データ差異・調整等')
        .setFontSize(8).setFontColor('#888888').setBackground('#ffffff').setHorizontalAlignment('left').setFontWeight('normal');
    }
    row++;
    var ctStart=row;
    for (var d=0;d<filteredDays.length;d++){
      var ds=filteredDays[d],dayN=parseInt(ds.split('/')[1]);
      var rowPeople=0,rowSales=0,hasData=false,catVals=[];
      for (var pk=0;pk<n;pk++){
        var e=(catMap[ds]&&catMap[ds][usedPgKeys[pk]])?catMap[ds][usedPgKeys[pk]]:null;
        var pp=e?e.people:0,ps=e?e.total:0;
        if(catCols[pk]===2)catVals.push(QTY_GROUPS[usedPgKeys[pk]]?(e?e.qty:0):pp);catVals.push(ps);
        rowPeople+=pp;rowSales+=ps;if(ps>0)hasData=true;
      }
      if (!showAllDays&&!hasData) continue;
      var dispPeople=dailyActMap&&dailyActMap[ds]?dailyActMap[ds].people:rowPeople;
      var rowArr=[dateDowLabel(ds,year,month,dayN),dispPeople];
      if (dailyActMap){
        var act=dailyActMap[ds]?dailyActMap[ds].total:0;
        rowArr.push(rowSales);rowArr.push(act);rowArr.push(rowSales-act);
      } else {
        rowArr.push(rowSales);
      }
      for (var pv=0;pv<catVals.length;pv++) rowArr.push(catVals[pv]);
      out.getRange(row,1,1,rowArr.length).setValues([rowArr]).setHorizontalAlignment(C);
      for (var nc=3;nc<catColStart;nc++) out.getRange(row,nc).setNumberFormat('#,##0');
      for (var sc in catSalesCols) out.getRange(row,Number(sc)).setNumberFormat('#,##0');
      if (d%2===0) out.getRange(row,1,1,rowArr.length).setBackground(BG_EVEN);row++;
    }
    var ctEnd=row-1;
    out.getRange(row,1).setValue('合計');
    if (ctEnd>=ctStart){
      for (var c=2;c<=dcCols;c++){
        out.getRange(row,c).setFormula('=SUM('+colLetter(c)+ctStart+':'+colLetter(c)+ctEnd+')');
        var isFmt=(c>=3&&c<catColStart)||catSalesCols[c];
        if (isFmt) out.getRange(row,c).setNumberFormat('#,##0');
      }
    } else {
      for (var c=2;c<=dcCols;c++) out.getRange(row,c).setValue(0);
    }
    out.getRange(row,1,1,dcCols).setFontWeight('bold').setBackground(BG_TOTAL).setHorizontalAlignment(C);row+=2;
  }

  writeCrossTable('【日別・カテゴリ別売上】　※単価×数量（定価合計・割引前）',days,dailyCrossMap,pgKeysCross,true,dailyMap);

  var muraPgSet={},passPgSet={};
  Object.keys(muraDailyCatMap).forEach(function(ds){Object.keys(muraDailyCatMap[ds]).forEach(function(pg){muraPgSet[pg]=true;});});
  Object.keys(passDailyCatMap).forEach(function(ds){Object.keys(passDailyCatMap[ds]).forEach(function(pg){passPgSet[pg]=true;});});
  var muraPgKeys=Object.keys(muraPgSet); sortByOrder(muraPgKeys,CROSS_ORDER);
  var passPgKeys=Object.keys(passPgSet); sortByOrder(passPgKeys,CROSS_ORDER);
  writeCrossTable('【村民割引（日別・カテゴリ別）】  合計'+mura.count+'件 '+mura.people+'人 '+mura.total.toLocaleString()+'円',days,muraDailyCatMap,muraPgKeys,false);
  writeCrossTable('【年間パス（日別・カテゴリ別）】  合計'+pass.count+'件 '+pass.people+'人 '+pass.total.toLocaleString()+'円',days,passDailyCatMap,passPgKeys,false);

  out.getRange(row,1,1,7).merge().setValue('【団体割引（日別・カテゴリ別）】  合計0件 0人 0円').setFontWeight('bold').setBackground(BG_SEC).setFontColor('#fff').setHorizontalAlignment(C);
  row++;
  out.getRange(row,1,1,7).merge().setValue('（団体割引設定が追加された際に自動表示されます）').setFontColor('#888888').setBackground('#f5f5f5').setHorizontalAlignment(C);
  row+=2;

  secHead('【支払方法別】',3);colHead(['支払方法','売上合計','件数']);
  var pKeys=Object.keys(payMap);
  for (var p=0;p<pKeys.length;p++){
    var pv=payMap[pKeys[p]];
    out.getRange(row,1,1,3).setValues([[pKeys[p],pv.total,pv.count]]).setHorizontalAlignment(C);
    out.getRange(row,2).setNumberFormat('#,##0');
    if(p%2===0)out.getRange(row,1,1,3).setBackground(BG_EVEN);row++;
  }
  row++;

  secHead('【カテゴリ別・商品別売上】',5);
  var catSectionBodyStart=row;
  out.getRange(row,1,1,5).setValues([['カテゴリ / 商品名','件数','人数','人数','売上合計']])
    .setBackground('#f5f5f5').setHorizontalAlignment(C);
  out.getRange(row,1).setFontWeight('bold');
  out.getRange(row,2).setFontWeight('normal');
  out.getRange(row,3).setFontWeight('bold');
  out.getRange(row,4).setFontWeight('bold');
  out.getRange(row,5).setFontWeight('bold');
  row++;

  for (var pi=0;pi<pgKeysCat.length;pi++){
    var pg=pgKeysCat[pi],pe=parentMap[pg];
    var showPeople=!!PEOPLE_GROUPS[pg];
    out.getRange(row,1,1,5).setValues([[pg,pe.count,pe.people,'',pe.total]])
      .setFontWeight('bold').setBackground(BG_CAT).setFontColor('#1a3a08').setHorizontalAlignment(C);
    out.getRange(row,5).setNumberFormat('#,##0');row++;

    var catNames=Object.keys(pe.cats);
    catNames.sort(function(a,b){return pe.cats[b].total-pe.cats[a].total;});
    for (var ci=0;ci<catNames.length;ci++){
      var ck=catNames[ci],ce=pe.cats[ck];
      var showCatRow=ck!==''&&catNames.length>1;
      if (showCatRow){
        out.getRange(row,1,1,5).setValues([['　'+ck,ce.count,ce.people,'',ce.total]])
          .setFontWeight('bold').setBackground(BG_CAT2).setFontColor('#333').setHorizontalAlignment(C);
        out.getRange(row,5).setNumberFormat('#,##0');row++;
      }
      var itKeys=Object.keys(ce.items);
      itKeys.sort(function(a,b){return ce.items[b].total-ce.items[a].total;});
      var indent=showCatRow?'　　':'　';
      for (var ij=0;ij<itKeys.length;ij++){
        var iv=ce.items[itKeys[ij]];
        var col4Val=showPeople?iv.people:iv.qty;
        out.getRange(row,1,1,5).setValues([[indent+iv.name,'','',col4Val,iv.total]]).setHorizontalAlignment(C);
        out.getRange(row,5).setNumberFormat('#,##0');
        if(ij%2===0)out.getRange(row,1,1,5).setBackground(BG_EVEN);row++;
      }
    }
  }

  var chartCats=pgKeysCat.map(function(pg){return {name:pg,total:parentMap[pg].total};});
  var chartTotal=chartCats.reduce(function(s,c){return s+c.total;},0);
  chartCats.sort(function(a,b){return b.total-a.total;});
  var chartDataRow=catSectionBodyStart;
  out.getRange(chartDataRow,7).setValue('カテゴリ').setFontWeight('bold');
  out.getRange(chartDataRow,8).setValue('割合(%)').setFontWeight('bold');
  chartDataRow++;
  var catChartDataStart=chartDataRow;
  for (var ci=0;ci<chartCats.length;ci++){
    out.getRange(chartDataRow,7).setValue(chartCats[ci].name);
    var pct=chartTotal>0?Math.round(chartCats[ci].total/chartTotal*1000)/10:0;
    out.getRange(chartDataRow,8).setValue(pct).setNumberFormat('0.0');
    chartDataRow++;
  }

  SpreadsheetApp.flush();
  try {
    var catChartRange=out.getRange(catChartDataStart-1,7,chartCats.length+1,2);
    var catBarChart=out.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(catChartRange)
      .setOption('title','カテゴリ別売上割合')
      .setOption('legend',{position:'none'})
      .setOption('hAxis',{title:'%',format:'0.0'})
      .setOption('width',420).setOption('height',chartCats.length*28+80)
      .setPosition(catSectionBodyStart,10,0,0)
      .build();
    out.insertChart(catBarChart);
  } catch(e) {
    Logger.log('グラフ作成エラー: '+e.message);
  }
  row++;

  secHead('【割引・予約サイト別】',3);colHead(['割引/予約サイト','件数','売上合計']);
  var dkKeys=Object.keys(discMap);
  for (var dk=0;dk<dkKeys.length;dk++){
    var dv=discMap[dkKeys[dk]];
    out.getRange(row,1,1,3).setValues([[dkKeys[dk],dv.count,dv.total]]).setHorizontalAlignment(C);
    out.getRange(row,3).setNumberFormat('#,##0');
    if(dk%2===0)out.getRange(row,1,1,3).setBackground(BG_EVEN);row++;
  }
  row++;

  secHead('【年齢層別】',3);colHead(['年齢層','件数（組）','人数']);
  var aKeys=Object.keys(ageMap);
  for (var ak=0;ak<aKeys.length;ak++){
    var av=ageMap[aKeys[ak]];
    out.getRange(row,1,1,3).setValues([[aKeys[ak],av.groups,av.people]]).setHorizontalAlignment(C);
    if(ak%2===0)out.getRange(row,1,1,3).setBackground(BG_EVEN);row++;
  }
  row++;

  secHead('【国籍別】',3);colHead(['国籍','件数（組）','人数']);
  var nkKeys=Object.keys(natMap);
  for (var nk=0;nk<nkKeys.length;nk++){
    var nv=natMap[nkKeys[nk]];
    out.getRange(row,1,1,3).setValues([[nkKeys[nk],nv.groups,nv.people]]).setHorizontalAlignment(C);
    if(nk%2===0)out.getRange(row,1,1,3).setBackground(BG_EVEN);row++;
  }

  out.setColumnWidth(1,110);out.setColumnWidth(2,65);out.setColumnWidth(3,90);
  var maxCol=Math.max(8,3+pgKeysCross.length*2);
  for (var wc=4;wc<=maxCol;wc++) out.setColumnWidth(wc,wc%2===0?60:90);
  out.setColumnWidth(7,140);out.setColumnWidth(8,90);

  SpreadsheetApp.flush();
  return sheetName;
}

// ==================== 請求書・見積書 ====================

function getNextInvoiceNumber(ss, prefix) {
  var sheets = ss.getSheets();
  var max = 0;
  for (var i = 0; i < sheets.length; i++) {
    var m = sheets[i].getName().match(new RegExp('^' + prefix + '-(\\d{8})-(\\d+)$'));
    if (m) { var n = parseInt(m[2]); if (n > max) max = n; }
  }
  return String(max + 1).padStart(3, '0');
}

function createInvoiceSheet(data) {
  var ss = getTargetSS();
  var docType = data.doc_type === 'estimate' ? '見積書' : '請求書';
  var prefix = data.doc_type === 'estimate' ? '見積書' : '請求書';
  var billDate = data.bill_date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  var dateKey = billDate.replace(/\//g, '');
  var num = getNextInvoiceNumber(ss, prefix);
  var sheetName = prefix + '-' + dateKey + '-' + num;
  var out = ss.insertSheet(sheetName, 0);
  var C = 'center', L = 'left', R = 'right';
  var BORDER = '#aaaaaa';
  out.getRange(1,1,1,6).merge().setValue('【' + docType + '】').setFontSize(20).setFontWeight('bold').setHorizontalAlignment(C).setVerticalAlignment('middle');
  out.setRowHeight(1, 50);
  out.getRange(2,4).setValue('発行日').setFontWeight('bold').setHorizontalAlignment(R);
  out.getRange(2,5,1,2).merge().setValue(billDate).setHorizontalAlignment(L);
  var recipient = data.recipient || '　';
  out.getRange(3,1,1,3).merge().setValue(recipient + '　様').setFontSize(14).setFontWeight('bold').setHorizontalAlignment(L).setVerticalAlignment('middle');
  out.setRowHeight(3, 32);
  var issuer = ['合同会社フェレリ　代表　ジョラン フェレリ','〒637-1441　奈良県吉野郡十津川村大字小川112','21世紀の森　空中の村','TEL：0746-62-0567','MAIL：info@kuuchuu-no-mura.com'];
  for (var ii = 0; ii < issuer.length; ii++) {
    out.getRange(3+ii,4,1,3).merge().setValue(issuer[ii]).setFontSize(10).setHorizontalAlignment(R);
  }
  var msgRow = 3 + issuer.length + 1;
  var msg = docType === '見積書' ? '下記の通りお見積申し上げます。' : '下記の通りご請求申し上げます。';
  out.getRange(msgRow,1,1,6).merge().setValue(msg).setFontColor('#555').setHorizontalAlignment(L);
  var hRow = msgRow + 1;
  out.getRange(hRow,1,1,4).setValues([['品名','数量','単価','金額']]).setFontWeight('bold').setHorizontalAlignment(C).setBorder(true,true,true,true,true,true,BORDER,SpreadsheetApp.BorderStyle.SOLID);
  out.getRange(hRow,1).setHorizontalAlignment(L);
  var items = data.items || [];
  var dataRow = hRow + 1;
  var subtotal = 0;
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var lineTotal = it.price * it.qty;
    subtotal += lineTotal;
    out.getRange(dataRow,1).setValue(it.name).setHorizontalAlignment(L);
    out.getRange(dataRow,2).setValue(it.qty).setHorizontalAlignment(C);
    out.getRange(dataRow,3).setValue(it.price).setNumberFormat('#,##0').setHorizontalAlignment(R);
    out.getRange(dataRow,4).setValue(lineTotal).setNumberFormat('#,##0').setHorizontalAlignment(R);
    out.getRange(dataRow,1,1,4).setBorder(false,true,true,true,false,false,BORDER,SpreadsheetApp.BorderStyle.SOLID);
    dataRow++;
  }
  while (dataRow < hRow + 6) {
    out.getRange(dataRow,1,1,4).setBorder(false,true,true,true,false,false,BORDER,SpreadsheetApp.BorderStyle.SOLID);
    dataRow++;
  }
  var totalRow = dataRow + 1;
  var total = data.total || subtotal;
  var discount = subtotal - total;
  out.getRange(totalRow,3).setValue('小計').setFontWeight('bold').setHorizontalAlignment(R);
  out.getRange(totalRow,4).setValue(subtotal).setNumberFormat('#,##0').setHorizontalAlignment(R);
  if (discount > 0) {
    totalRow++;
    out.getRange(totalRow,3).setValue('割引').setFontWeight('bold').setHorizontalAlignment(R);
    out.getRange(totalRow,4).setValue(-discount).setNumberFormat('#,##0').setHorizontalAlignment(R).setFontColor('#c0392b');
  }
  totalRow++;
  out.getRange(totalRow,1,1,2).merge().setValue('お支払金額（税込）').setFontSize(12).setFontWeight('bold').setHorizontalAlignment(L).setVerticalAlignment('middle');
  out.getRange(totalRow,3,1,2).merge().setValue('¥ ' + total.toLocaleString()).setFontSize(15).setFontWeight('bold').setHorizontalAlignment(R).setVerticalAlignment('middle');
  out.setRowHeight(totalRow, 32);
  totalRow += 2;
  out.getRange(totalRow,1).setValue('【備考】').setFontWeight('bold');
  out.getRange(totalRow,2,1,5).merge().setValue(data.memo||'').setWrap(true);
  out.setColumnWidth(1, 230); out.setColumnWidth(2, 60); out.setColumnWidth(3, 90); out.setColumnWidth(4, 90); out.setColumnWidth(5, 90); out.setColumnWidth(6, 90);
  SpreadsheetApp.flush();
  return sheetName;
}
