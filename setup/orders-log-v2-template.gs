/**
 * One Stop Travel & Tour — order log + profit in Google Sheets
 * Sheets: Orders (one row per order), Items (one row per tour option / vehicle),
 *         Costs (supplier cost per tour option — edit here), Settings (your profit share), Summary.
 * Profit = Sale − Cost. My share = Profit × Settings!B1 (30%).
 */
const SHEET_ID = '1oeNSPUlHHa4QhUJYueT9_kkMSh7kWk0R6OrhLpXuAE8'; // One Stop — Orders (fjordoflanta@gmail.com)

const ORDER_HEADERS = ['Received', 'Order ID', 'Status', 'First name', 'Last name', 'Phone / WhatsApp', 'E-mail',
  'Staying at', 'Items', 'Total, THB', 'Private on request', 'Comment', 'Paid, THB', 'Notes (internal)', 'Full order text',
  'Profit, THB', 'My share, THB'];
const ITEM_HEADERS = ['Received', 'Order ID', 'Order status', 'Customer', 'Type', 'Tour / vehicle', 'Option',
  'Service date', 'Qty / days', 'Unit sale, THB', 'Sale, THB', 'Unit cost, THB', 'Cost, THB', 'Profit, THB',
  'My share, THB', 'Details', 'Key'];
const COST_HEADERS = ['Key', 'Tour / vehicle', 'Option', 'Sale price, THB', 'Cost price, THB'];
const STATUSES = ['New', 'Confirmed', 'Paid', 'Completed', 'Cancelled'];
const COSTS = __COSTS__;

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function setup() {
  const ss = ss_();
  const orders = sheet_(ss, 'Orders', ORDER_HEADERS);
  const items = sheet_(ss, 'Items', ITEM_HEADERS);
  header_(orders, ORDER_HEADERS); header_(items, ITEM_HEADERS);
  orders.getRange('C2:C5000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  orders.setColumnWidth(15, 360);
  orders.getRange('O2:O5000').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  orders.getRange('P2:Q5000').setNumberFormat('#,##0');
  items.getRange('I2:O5000').setNumberFormat('#,##0');
  items.hideColumns(17);

  // Settings
  let st = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  if (st.getLastRow() === 0) {
    st.getRange('A1:B1').setValues([['My share of profit', 0.3]]);
    st.getRange('B1').setNumberFormat('0%').setFontColor('#0000FF').setBackground('#FFF2CC');
    st.getRange('A3').setValue('Change B1 to change your share; all sheets recalculate.');
    st.setColumnWidth(1, 220);
  }

  // Costs
  let cs = ss.getSheetByName('Costs') || ss.insertSheet('Costs');
  if (cs.getLastRow() === 0) {
    cs.appendRow(COST_HEADERS);
    cs.getRange(1, 1, 1, COST_HEADERS.length).setFontWeight('bold').setBackground('#1E1A16').setFontColor('#ffffff');
    cs.setFrozenRows(1);
    cs.getRange(2, 1, COSTS.length, 5).setValues(COSTS);
    cs.getRange(2, 4, COSTS.length, 2).setNumberFormat('#,##0');
    cs.getRange(2, 5, COSTS.length, 1).setFontColor('#0000FF');
    cs.setColumnWidth(1, 150); cs.setColumnWidth(2, 380); cs.setColumnWidth(3, 160);
  }

  // Summary
  let sm = ss.getSheetByName('Summary') || ss.insertSheet('Summary');
  sm.clear();
  sm.getRange('A1:F1').setValues([['Order status', 'Sale, THB', 'Cost, THB', 'Profit, THB', 'My share, THB', 'Item lines']]);
  sm.getRange('A1:F1').setFontWeight('bold').setBackground('#1E1A16').setFontColor('#ffffff');
  STATUSES.forEach(function (s, i) {
    const r = i + 2;
    sm.getRange(r, 1, 1, 6).setFormulas([['="' + s + '"',
      '=SUMIFS(Items!K:K,Items!C:C,A' + r + ')', '=SUMIFS(Items!M:M,Items!C:C,A' + r + ')',
      '=SUMIFS(Items!N:N,Items!C:C,A' + r + ')', '=SUMIFS(Items!O:O,Items!C:C,A' + r + ')',
      '=COUNTIFS(Items!C:C,A' + r + ')']]);
  });
  sm.getRange('A8:F8').setFormulas([['="All except Cancelled"', '=SUM(B2:B5)', '=SUM(C2:C5)', '=SUM(D2:D5)', '=SUM(E2:E5)', '=SUM(F2:F5)']]);
  sm.getRange('A8:F8').setFontWeight('bold').setBackground('#E8EEF0');
  sm.getRange('B2:E8').setNumberFormat('#,##0');
  sm.getRange('A10').setValue('Profit counts only item lines that have a cost price in the Costs sheet. Fill in missing costs (blank cells) to include them.');
  sm.setColumnWidth(1, 200);

  ['Sheet1', 'Лист1'].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() <= 1 && ss.getSheets().length > 2) ss.deleteSheet(s);
  });
  ss.setActiveSheet(orders);
}

function header_(sh, headers) {
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1E1A16').setFontColor('#ffffff');
}

function sheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) { header_(sh, headers); sh.setFrozenRows(1); }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = ss_();
    const orders = sheet_(ss, 'Orders', ORDER_HEADERS);
    const items = sheet_(ss, 'Items', ITEM_HEADERS);
    const c = d.customer || {};
    const when = d.at ? new Date(d.at) : new Date();
    const list = (d.items || []).map(function (i) { return i.title; }).join('\n');

    const ids = orders.getRange(2, 2, Math.max(orders.getLastRow() - 1, 1), 1).getValues().flat();
    if (ids.indexOf(d.id) !== -1) return ok_();

    orders.appendRow([when, d.id, 'New', c.first || '', c.last || '', "'" + (c.phone || ''), c.email || '',
      c.stay || '', list, Number(d.total) || 0, d.onRequest ? 'Yes' : '', c.comment || '', '', '', d.text || '']);
    const r = orders.getLastRow();
    orders.getRange(r, 3).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
    orders.getRange(r, 1).setNumberFormat('dd.mm.yyyy hh:mm');
    orders.getRange(r, 10).setNumberFormat('#,##0');
    orders.getRange(r, 16, 1, 2).setFormulas([[
      '=SUMIFS(Items!N:N,Items!B:B,B' + r + ')', '=SUMIFS(Items!O:O,Items!B:B,B' + r + ')']]).setNumberFormat('#,##0');

    const cust = (c.first || '') + ' ' + (c.last || '');
    (d.items || []).forEach(function (i) {
      const lines = (i.lines && i.lines.length) ? i.lines :
        [{ key: '', option: '', qty: 1, price: i.amount === '' ? '' : Number(i.amount) || 0 }];
      lines.forEach(function (L) {
        items.appendRow([when, d.id, '', cust, i.type, i.title, L.option || '', i.date || '',
          Number(L.qty) || 0, L.price === '' ? '' : Number(L.price) || 0, '', '', '', '', '', i.details || '', L.key || '']);
        const n = items.getLastRow();
        items.getRange(n, 3).setFormula('=IFERROR(INDEX(Orders!C:C,MATCH(B' + n + ',Orders!B:B,0)),"")');
        items.getRange(n, 11, 1, 5).setFormulas([[
          '=IF(J' + n + '="","",I' + n + '*J' + n + ')',
          '=IFERROR(IF(VLOOKUP(Q' + n + ',Costs!A:E,5,FALSE)="","",VLOOKUP(Q' + n + ',Costs!A:E,5,FALSE)),"")',
          '=IF(L' + n + '="","",I' + n + '*L' + n + ')',
          '=IF(OR(K' + n + '="",M' + n + '=""),"",K' + n + '-M' + n + ')',
          '=IF(N' + n + '="","",N' + n + '*Settings!$B$1)']]);
        items.getRange(n, 1).setNumberFormat('dd.mm.yyyy hh:mm');
        items.getRange(n, 9, 1, 7).setNumberFormat('#,##0');
      });
    });
    return ok_();
  } finally {
    lock.releaseLock();
  }
}

function ok_() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}
