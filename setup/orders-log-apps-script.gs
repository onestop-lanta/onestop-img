/**
 * One Stop Travel & Tour — order log in Google Sheets
 *
 * Setup (once, ~5 minutes):
 * 1. Create a new Google Sheet (e.g. "One Stop — Orders") in fjordoflanta@gmail.com.
 * 2. Extensions → Apps Script. Delete the sample code and paste this whole file. Save.
 * 3. Run the function "setup" once (▶ Run). Allow the permissions it asks for.
 * 4. Deploy → New deployment → type "Web app":
 *      Execute as: Me
 *      Who has access: Anyone
 *    → Deploy → copy the Web app URL (ends with /exec).
 * 5. Send the URL and the link to the Sheet to Claude (or paste them into data.js:
 *    ordersSheetUrl = Web app URL, ordersSheetLink = Sheet link).
 *
 * Every order from the website then appears here automatically:
 *  - sheet "Orders": one row per order (customer, contacts, total, status)
 *  - sheet "Items":  one row per tour / bike / car in the order
 * Change "Status" with the dropdown: New → Confirmed → Paid → Completed / Cancelled.
 */

const ORDER_HEADERS = ['Received', 'Order ID', 'Status', 'First name', 'Last name', 'Phone / WhatsApp', 'E-mail',
  'Staying at', 'Items', 'Total, THB', 'Private on request', 'Comment', 'Paid, THB', 'Notes (internal)', 'Full order text'];
const ITEM_HEADERS = ['Received', 'Order ID', 'Customer', 'Type', 'Tour / vehicle', 'Service date', 'Days', 'People', 'Details', 'Amount, THB'];
const STATUSES = ['New', 'Confirmed', 'Paid', 'Completed', 'Cancelled'];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orders = sheet_(ss, 'Orders', ORDER_HEADERS);
  sheet_(ss, 'Items', ITEM_HEADERS);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build();
  orders.getRange('C2:C5000').setDataValidation(rule);
  orders.setColumnWidth(15, 360);
  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Лист1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(s1);
}

function sheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1E1A16').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orders = sheet_(ss, 'Orders', ORDER_HEADERS);
    const items = sheet_(ss, 'Items', ITEM_HEADERS);
    const c = d.customer || {};
    const when = d.at ? new Date(d.at) : new Date();
    const list = (d.items || []).map(function (i) { return i.title; }).join('\n');

    // avoid duplicates if the same order is sent twice
    const ids = orders.getRange(2, 2, Math.max(orders.getLastRow() - 1, 1), 1).getValues().flat();
    if (ids.indexOf(d.id) !== -1) return ok_();

    orders.appendRow([when, d.id, 'New', c.first || '', c.last || '', "'" + (c.phone || ''), c.email || '',
      c.stay || '', list, Number(d.total) || 0, d.onRequest ? 'Yes' : '', c.comment || '', '', '', d.text || '']);
    const r = orders.getLastRow();
    orders.getRange(r, 3).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
    orders.getRange(r, 1).setNumberFormat('dd.mm.yyyy hh:mm');
    orders.getRange(r, 10).setNumberFormat('#,##0');

    (d.items || []).forEach(function (i) {
      items.appendRow([when, d.id, (c.first || '') + ' ' + (c.last || ''), i.type, i.title, i.date || '', i.days || '',
        i.people || '', i.details || '', i.amount === '' ? 'On request' : Number(i.amount) || 0]);
      items.getRange(items.getLastRow(), 1).setNumberFormat('dd.mm.yyyy hh:mm');
    });
    return ok_();
  } finally {
    lock.releaseLock();
  }
}

function ok_() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}
