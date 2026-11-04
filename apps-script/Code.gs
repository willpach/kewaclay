/**
 * Mug pre-orders → Google Sheet
 *
 * Setup (about 5 minutes):
 * 1. Make a new Google Sheet (e.g. "Mug Orders").
 * 2. In the Sheet: Extensions → Apps Script. Delete what's there and paste this whole file. Save.
 * 3. Click Deploy → New deployment → gear icon → Web app.
 *      Execute as: Me
 *      Who has access: Anyone
 *    Click Deploy and allow access when Google asks.
 * 4. Copy the Web app URL (ends in /exec) and paste it into CONFIG.endpoint in index.html.
 *
 * Every order becomes one row in the "Orders" tab.
 */

const HEADERS = [
  "Received", "Order #", "Name", "Phone", "Email", "Instagram/Facebook",
  "# Mugs", "Mugs", "Total", "Payment choice", "Pay now", "Due at pickup",
  "Get mugs by", "Shipping address", "Cash App name", "Gift?", "Notes",
  "Paid? (you fill in)", "Picked up? (you fill in)"
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Orders");
    if (!sheet) {
      sheet = ss.insertSheet("Orders");
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    }
    const p = e.parameter || {};
    sheet.appendRow([
      new Date(), p.orderId, p.name, p.phone, p.email, p.social,
      Number(p.mugCount || 0), p.mugs, Number(p.total || 0), p.payChoice,
      Number(p.dueNow || 0), Number(p.dueAtPickup || 0),
      p.fulfillment, p.address, p.cashappName, p.gift, p.notes, "", ""
    ]);
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}
