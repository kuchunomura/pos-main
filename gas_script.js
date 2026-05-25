// ===== POS レジ GAS スクリプト (main用) =====
const SS_ID = '13AgulTcqyxRZLF9nejl2TJ3qqpLKlrFprlPX25fU-kg';

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

function handleRequest(data) {
  var type = data.type;
  if      (type === 'add_rows')        addRows(data.rows);
  else if (type === 'delete_rows')     deleteRows(data.sale_id);
  else if (type === 'replace_rows')    replaceRows(data.sale_id, data.rows);
  else if (type === 'clear_sheets')    clearSheets();
  else if (type === 'monthly_summary') return { sheet: createMonthlySummary(data.year, data.month) };
  else if (type === 'bulk_sync') {
    if (data.clear) clearSheets();
    var sales = data.sales || [];
    for (var i = 0; i < sales.length; i++) { if (sales[i] && sales[i].length) addRows(sales[i]); }
  }
  return null;
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
}

// ==================== 行操作 ====================

function addRows(rows) {
  if (!rows || !rows.length) return;
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheetName = sheetNameFromRows(rows);
  var sheet = getOrCreateSheet(ss, sheetName);
  ensureHeaders(sheet);

  var lastRow = sheet.getLastRow();
  var startRow = Math.max(lastRow + 1, 4);
  var colCount = rows[0].length;

  sheet.getRange(startRow, 1, rows.length, colCount).setValues(rows);
  sheet.getRange(startRow, 1, rows.length, colCount).setHorizontalAlignment('center');

  for (var i = 0; i < rows.length; i++) {
    applyPaymentColors(sheet, startRow + i, rows[i][9]); // J列(index9): 支払方法
  }

  applyGroupBorders(sheet, startRow, rows.length, colCount);
  SpreadsheetApp.flush();
  setDataColumnWidths(sheet);
  updateSummary(sheet);
}

function replaceRows(saleId, rows) {
  if (!saleId) return;
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    deleteRowsFromSheet(sheets[i], saleId);
  }
  if (rows && rows.length) addRows(rows);
}

function deleteRows(saleId) {
  if (!saleId) return;
  var ss = SpreadsheetApp.openById(SS_ID);
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
  var ss = SpreadsheetApp.openById(SS_ID);
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

function applyPaymentColors(sheet, row, payment) {
  if (payment === 'クレジットカード') {
    sheet.getRange(row, 1, 1, 16).setBackground('#f0f8ff');
  } else if (payment === '電子決済') {
    sheet.getRange(row, 1, 1, 16).setBackground('#fdf5ff');
  }
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
  var ss = SpreadsheetApp.openById(SS_ID);
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

// 月別集計シートを作成して返す
function createMonthlySummary(year, month) {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheets = ss.getSheets();

  // 該当月のシートを収集
  // 対応形式: "M/D売上" / "YYYY/M/D売上" / "M/D 売上"（スペースあり）
  var monthSheets = [];
  for (var s = 0; s < sheets.length; s++) {
    var name = sheets[s].getName();
    var match = name.match(/^(?:\d{4}\/)?(\d+)\/(\d+)[^0-9\/]*売上$/);
    if (match && parseInt(match[1]) === month) {
      monthSheets.push({ sheet: sheets[s], day: parseInt(match[2]) });
    }
  }
  monthSheets.sort(function(a, b) { return a.day - b.day; });

  // 集計用マップ
  var dailyMap = {};
  var itemMap  = {};
  var payMap   = {};
  var discMap  = {};
  var ageMap   = {};
  var natMap   = {};
  var txSeen   = {};
  var grandTotal = 0, grandCount = 0, grandPeople = 0;

  for (var si = 0; si < monthSheets.length; si++) {
    var sheet   = monthSheets[si].sheet;
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) continue;
    var data    = sheet.getRange(4, 1, lastRow - 3, 16).getValues();
    var dateStr = month + '/' + monthSheets[si].day;
    if (!dailyMap[dateStr]) dailyMap[dateStr] = {total:0,count:0,people:0,cash:0,card:0,elec:0};

    for (var i = 0; i < data.length; i++) {
      var r        = data[i];
      var txId     = String(r[14]);  // O: 売上ID
      var amount   = Number(r[1]);   // B: 売上合計
      var itemName = String(r[2]);   // C: 商品名
      var cat      = String(r[3]);   // D: カテゴリ
      var qty      = Number(r[4]);   // E: 数量
      var unitPrice= Number(r[5]);   // F: 単価
      var people   = Number(r[7]);   // H: 人数
      var disc     = String(r[8]);   // I: 割引
      var payment  = String(r[9]);   // J: 支払方法
      var ageStr   = String(r[10]);  // K: 年齢層
      var natStr   = String(r[11]);  // L: 国籍

      // 商品別（行ごと）
      if (itemName && itemName !== '') {
        var mKey = itemName + '\t' + cat;
        if (!itemMap[mKey]) itemMap[mKey] = {name:itemName, cat:cat, qty:0, total:0};
        itemMap[mKey].qty   += qty;
        itemMap[mKey].total += unitPrice * qty;
      }

      // 取引単位（txId ごとに1回）
      if (!txSeen[txId]) {
        txSeen[txId] = true;

        // 日別
        dailyMap[dateStr].total  += amount;
        dailyMap[dateStr].count  += 1;
        dailyMap[dateStr].people += people;
        if      (payment === '現金')           dailyMap[dateStr].cash += amount;
        else if (payment === 'クレジットカード') dailyMap[dateStr].card += amount;
        else if (payment === '電子決済')        dailyMap[dateStr].elec += amount;

        grandTotal  += amount;
        grandCount  += 1;
        grandPeople += people;

        // 支払方法別
        if (!payMap[payment]) payMap[payment] = {total:0, count:0};
        payMap[payment].total += amount;
        payMap[payment].count += 1;

        // 割引別
        var dKey = (disc && disc !== '' && disc !== 'false') ? disc : 'なし';
        if (!discMap[dKey]) discMap[dKey] = {count:0, total:0};
        discMap[dKey].count += 1;
        discMap[dKey].total += amount;

        // 年齢層別
        if (ageStr && ageStr !== '') {
          var ages = ageStr.split('・');
          for (var a = 0; a < ages.length; a++) {
            var ag = ages[a].trim();
            if (!ag) continue;
            if (!ageMap[ag]) ageMap[ag] = {groups:0, people:0};
            ageMap[ag].groups++;
            ageMap[ag].people += people;
          }
        }

        // 国籍別
        if (natStr && natStr !== '') {
          var nats = natStr.split('・');
          for (var n = 0; n < nats.length; n++) {
            var nat = nats[n].trim();
            if (!nat) continue;
            if (!natMap[nat]) natMap[nat] = {groups:0, people:0};
            natMap[nat].groups++;
            natMap[nat].people += people;
          }
        }
      }
    }
  }

  // シート作成（既存なら上書き）
  var sheetName = year + '年' + month + '月 月別集計';
  var out = ss.getSheetByName(sheetName);
  if (out) { out.clear(); } else { out = ss.insertSheet(sheetName, 0); }

  var row = 1;
  var BG_HEAD  = '#2d5016';
  var BG_SEC   = '#4a7c2f';
  var BG_TOTAL = '#e8f0e0';
  var BG_EVEN  = '#fafafa';

  // タイトル行
  out.getRange(row, 1, 1, 8).merge()
    .setValue(year + '年' + month + '月 月別集計')
    .setFontSize(15).setFontWeight('bold')
    .setBackground(BG_HEAD).setFontColor('#fff').setHorizontalAlignment('center');
  row++;

  // 月合計行
  var avg = grandPeople > 0 ? Math.round(grandTotal / grandPeople) : 0;
  var totRow = [['総売上', grandTotal, '件数', grandCount, '人数', grandPeople, '客単価（人）', avg]];
  out.getRange(row, 1, 1, 8).setValues(totRow)
    .setBackground('#f0f4e8').setFontWeight('bold').setHorizontalAlignment('center');
  out.getRange(row, 2).setNumberFormat('#,##0');
  out.getRange(row, 8).setNumberFormat('#,##0');
  row += 2;

  // ── セクションヘッダー用ヘルパー ──
  function secHead(label, cols) {
    out.getRange(row, 1, 1, cols).merge()
      .setValue(label).setFontWeight('bold')
      .setBackground(BG_SEC).setFontColor('#fff');
    row++;
  }
  function colHead(headers) {
    out.getRange(row, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#f5f5f5').setHorizontalAlignment('center');
    row++;
  }

  // ── 日別売上 ──
  secHead('【日別売上】', 7);
  colHead(['日付','売上合計','件数','人数','現金','カード','電子決済']);
  var days = Object.keys(dailyMap).sort(function(a,b){ return parseInt(a.split('/')[1])-parseInt(b.split('/')[1]); });
  var dt=[0,0,0,0,0,0];
  for (var d = 0; d < days.length; d++) {
    var dd = dailyMap[days[d]];
    out.getRange(row,1,1,7).setValues([[days[d],dd.total,dd.count,dd.people,dd.cash,dd.card,dd.elec]]);
    out.getRange(row,2).setNumberFormat('#,##0');
    out.getRange(row,5,1,3).setNumberFormat('#,##0');
    if (d%2===0) out.getRange(row,1,1,7).setBackground(BG_EVEN);
    dt[0]+=dd.total;dt[1]+=dd.count;dt[2]+=dd.people;dt[3]+=dd.cash;dt[4]+=dd.card;dt[5]+=dd.elec;
    row++;
  }
  out.getRange(row,1,1,7).setValues([['合計',dt[0],dt[1],dt[2],dt[3],dt[4],dt[5]]])
    .setFontWeight('bold').setBackground(BG_TOTAL).setHorizontalAlignment('center');
  out.getRange(row,2).setNumberFormat('#,##0');
  out.getRange(row,5,1,3).setNumberFormat('#,##0');
  row += 2;

  // ── 支払方法別 ──
  secHead('【支払方法別】', 3);
  colHead(['支払方法','売上合計','件数']);
  var pKeys = Object.keys(payMap);
  for (var p = 0; p < pKeys.length; p++) {
    var pv = payMap[pKeys[p]];
    out.getRange(row,1,1,3).setValues([[pKeys[p],pv.total,pv.count]]);
    out.getRange(row,2).setNumberFormat('#,##0');
    if (p%2===0) out.getRange(row,1,1,3).setBackground(BG_EVEN);
    row++;
  }
  row++;

  // ── 商品別売上 ──
  secHead('【商品別売上（金額順）】', 4);
  colHead(['商品名','カテゴリ','数量','売上合計']);
  var iKeys = Object.keys(itemMap);
  iKeys.sort(function(a,b){ return itemMap[b].total - itemMap[a].total; });
  for (var ii = 0; ii < iKeys.length; ii++) {
    var iv = itemMap[iKeys[ii]];
    out.getRange(row,1,1,4).setValues([[iv.name,iv.cat,iv.qty,iv.total]]);
    out.getRange(row,4).setNumberFormat('#,##0');
    if (ii%2===0) out.getRange(row,1,1,4).setBackground(BG_EVEN);
    row++;
  }
  row++;

  // ── 割引・予約サイト別 ──
  secHead('【割引・予約サイト別】', 3);
  colHead(['割引/予約サイト','件数','売上合計']);
  var dkKeys = Object.keys(discMap);
  for (var dk = 0; dk < dkKeys.length; dk++) {
    var dv = discMap[dkKeys[dk]];
    out.getRange(row,1,1,3).setValues([[dkKeys[dk],dv.count,dv.total]]);
    out.getRange(row,3).setNumberFormat('#,##0');
    if (dk%2===0) out.getRange(row,1,1,3).setBackground(BG_EVEN);
    row++;
  }
  row++;

  // ── 年齢層別 ──
  secHead('【年齢層別】', 3);
  colHead(['年齢層','件数（組）','人数']);
  var aKeys = Object.keys(ageMap);
  for (var ak = 0; ak < aKeys.length; ak++) {
    var av = ageMap[aKeys[ak]];
    out.getRange(row,1,1,3).setValues([[aKeys[ak],av.groups,av.people]]);
    if (ak%2===0) out.getRange(row,1,1,3).setBackground(BG_EVEN);
    row++;
  }
  row++;

  // ── 国籍別 ──
  secHead('【国籍別】', 3);
  colHead(['国籍','件数（組）','人数']);
  var nkKeys = Object.keys(natMap);
  for (var nk = 0; nk < nkKeys.length; nk++) {
    var nv = natMap[nkKeys[nk]];
    out.getRange(row,1,1,3).setValues([[nkKeys[nk],nv.groups,nv.people]]);
    if (nk%2===0) out.getRange(row,1,1,3).setBackground(BG_EVEN);
    row++;
  }

  // 列幅
  out.setColumnWidth(1, 130);
  out.setColumnWidth(2, 110);
  out.setColumnWidth(3,  80);
  out.setColumnWidth(4,  80);
  out.setColumnWidth(5, 100);
  out.setColumnWidth(6, 100);
  out.setColumnWidth(7, 100);
  out.setColumnWidth(8, 110);

  SpreadsheetApp.flush();
  return sheetName;
}
