/**
 * Mug pre-orders + waitlist → Google Sheet
 *
 * Setup:
 * 1. In the Sheet: Extensions → Apps Script. Replace everything with this file. Save.
 * 2. First time: Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 *    Updating: Deploy → Manage deployments → pencil (Edit) → Version: New version → Deploy.
 *    (Editing keeps the same web app URL.)
 *
 * Orders go to the first tab (or a tab named "Orders").
 * Waitlist sign-ups go to a tab named "Waitlist" (created automatically).
 */

const HEADERS = [
  "Received", "Order #", "Name", "Phone", "Email", "Instagram/Facebook",
  "# Mugs", "Mugs", "Total", "Payment choice", "Pay now", "Due at pickup",
  "Get mugs by", "Shipping address", "Cash App name", "Gift?", "Notes",
  "Paid? (you fill in)", "Amount paid (you fill in)", "Picked up? (you fill in)"
];

const WAITLIST_HEADERS = [
  "Received", "Name", "Email", "Phone", "Designs", "How many", "Size", "Notes", "Emailed? (you fill in)"
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (p.type === "waitlist") {
      let w = ss.getSheetByName("Waitlist");
      if (!w) w = ss.insertSheet("Waitlist");
      if (w.getLastRow() === 0) {
        w.appendRow(WAITLIST_HEADERS);
        w.setFrozenRows(1);
        w.getRange(1, 1, 1, WAITLIST_HEADERS.length).setFontWeight("bold");
      }
      w.appendRow([new Date(), p.name, p.email, p.phone, p.designs, p.count, p.size, p.notes, ""]);
      return ContentService.createTextOutput("ok");
    }

    // Uses a tab named "Orders" if there is one, otherwise the first tab.
    const sheet = ss.getSheetByName("Orders") || ss.getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    }
    sheet.appendRow([
      new Date(), p.orderId, p.name, p.phone, p.email, p.social,
      Number(p.mugCount || 0), p.mugs, Number(p.total || 0), p.payChoice,
      Number(p.dueNow || 0), Number(p.dueAtPickup || 0),
      p.fulfillment, p.address, p.cashappName, p.gift, p.notes
    ]);
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}
