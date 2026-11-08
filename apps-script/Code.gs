/**
 * KewaClay — orders, waitlist and inventory → Google Sheet   (v3)
 *
 * Updating: Extensions → Apps Script → paste this whole file → Save →
 *   Deploy → Manage deployments → pencil (Edit) → Version: New version → Deploy.
 *   (Editing keeps the same web app URL.)
 * First time only: pick "setupInventory" in the function menu at the top and click Run.
 *   It creates the Inventory tab with your bowls already filled in.
 *
 * Tabs:
 *   first tab (or "Orders") — pre-orders and one-of-a-kind purchases
 *   "Waitlist"              — waitlist sign-ups
 *   "Inventory"             — one-of-a-kind pieces shown in "Available now"
 *                             Status: Available (shows with Buy button), Sold (shows SOLD), anything else = hidden
 */

const HEADERS = [
  "Received", "Order #", "Name", "Phone", "Email", "Instagram/Facebook",
  "# Mugs", "Mugs", "Total", "Payment choice", "Pay now", "Due at pickup",
  "Get mugs by", "Shipping address", "Cash App name", "Gift?", "Notes"
];
const WAITLIST_HEADERS = [
  "Received", "Name", "Email", "Phone", "Designs", "How many", "Size", "Notes", "Emailed? (you fill in)"
];
const INVENTORY_HEADERS = [
  "ID", "Item", "Type", "Description", "Size", "Price", "Cost", "Photos",
  "Status", "Sold to", "Order #", "Date sold", "Notes"
];

function setupInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let inv = ss.getSheetByName("Inventory");
  if (inv) return;
  inv = ss.insertSheet("Inventory");
  inv.appendRow(INVENTORY_HEADERS);
  inv.appendRow(["B-01", "Dinosaur bowl with stepped designs", "Bowl",
    "Smiling dinosaurs with Mesa Verde–style stepped designs, black on white. Decorative (not for food). Want it glazed food safe? Ask me about adding glaze.",
    "6 in wide · 3.5 in tall", 120, "", "bowl1.jpg", "Available", "", "", "", ""]);
  inv.appendRow(["B-02", "Dinosaur bowl with triangle designs", "Bowl",
    "Smiling dinosaurs parading under bold black triangles. Decorative (not for food). Want it glazed food safe? Ask me about adding glaze.",
    "6 in wide · 3.5 in tall", 120, "", "bowl2.jpg", "Available", "", "", "", ""]);
  inv.appendRow(["B-03", "Tall dinosaur bowl (surface crack)", "Bowl",
    "A taller bowl with a wavy rim, a dinosaur and stepped designs. Has a small surface crack (tap the photo to see it), so it's priced lower. Decorative (not for food).",
    "5 in wide · 5.5 in tall", 75, "", "bowl3.jpg, bowl3-crack.jpg", "Available", "", "", "", ""]);
  inv.setFrozenRows(1);
  inv.getRange(1, 1, 1, INVENTORY_HEADERS.length).setFontWeight("bold");
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.type === "inventory") {
    const inv = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventory");
    const items = [];
    if (inv && inv.getLastRow() > 1) {
      const rows = inv.getRange(2, 1, inv.getLastRow() - 1, INVENTORY_HEADERS.length).getValues();
      rows.forEach(r => {
        if (!r[0]) return;
        items.push({ id: String(r[0]), item: r[1], type: r[2], description: r[3], size: r[4],
                     price: Number(r[5]) || 0, photos: r[7], status: r[8] });
      });
    }
    return ContentService.createTextOutput(JSON.stringify({ items: items }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput("ok");
}

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

    const sheet = ss.getSheetByName("Orders") || ss.getSheets()[0];

    if (p.type === "item_order") {
      const inv = ss.getSheetByName("Inventory");
      if (!inv) return ContentService.createTextOutput("no inventory");
      const ids = inv.getRange(2, 1, Math.max(inv.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]));
      const i = ids.indexOf(String(p.itemId));
      if (i < 0) return ContentService.createTextOutput("not found");
      const row = i + 2;
      const status = String(inv.getRange(row, 9).getValue());
      if (!/^available$/i.test(status)) return ContentService.createTextOutput("sold");
      inv.getRange(row, 9, 1, 4).setValues([["Sold", p.name, p.orderId, new Date()]]);
      const price = Number(p.price || 0);
      sheet.appendRow([
        new Date(), p.orderId, p.name, p.phone, p.email, "",
        1, "1) " + p.item + " (" + p.itemId + ") — $" + price, price, "Paid in full · one-of-a-kind",
        price, 0, p.fulfillment, p.address, p.cashappName, "", p.notes
      ]);
      return ContentService.createTextOutput("ok");
    }

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
