/**
 * Built on Our Block: entry receiver
 * Saves each entry to this Google Sheet, emails info@efpcs.org,
 * and sends the entrant a confirmation.
 */

const NOTIFY_EMAIL = "info@efpcs.org";
const SHEET_NAME = "Entries";
const FIELDS = [
  ["Submitted", null],
  ["Entrant or group", "name"],
  ["Contact email", "email"],
  ["Category", "category"],
  ["Parent / guardian / teacher", "guardian"],
  ["Video title", "title"],
  ["Video link", "link"],
  ["Business", "business"],
  ["Location shown", "location"],
  ["Historic or current", "era"],
  ["Why / sources", "why"]
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = (e && e.parameter) || {};
    if (!p.name || !p.email || !p.link) return reply_("missing");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(FIELDS.map(f => f[0]));
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, FIELDS.length).setFontWeight("bold");
    }
    const row = FIELDS.map(f => f[1] ? String(p[f[1]] || "").slice(0, 2000) : new Date());
    sheet.appendRow(row);

    const lines = FIELDS.slice(1).map((f, i) => f[0] + ": " + row[i + 1]).join("\n");

    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      replyTo: p.email,
      subject: "Built on Our Block Entry: " + p.name + " - " + (p.title || ""),
      body: "A new contest entry was submitted.\n\n" + lines +
            "\n\nAll entries: " + ss.getUrl()
    });

    MailApp.sendEmail({
      to: p.email,
      replyTo: NOTIFY_EMAIL,
      name: "Enduring Families Project: Children's Stories",
      subject: "We received your Built on Our Block entry",
      body: "Thank you for entering Built on Our Block: A Black Business Short Video Contest!\n\n" +
            "We received your video \"" + (p.title || "") + "\" (" + p.link + ").\n\n" +
            "Selections are announced November 29, and winners are notified on Giving Tuesday, December 1.\n" +
            "Please keep your video posted (Public or Unlisted) until winners are announced.\n\n" +
            "Questions? Email info@efpcs.org.\n\n" +
            "Enduring Families Project: Children's Stories\n#BuiltOnOurBlock"
    });

    return reply_("ok");
  } catch (err) {
    return reply_("error");
  } finally {
    lock.releaseLock();
  }
}

function reply_(status) {
  return ContentService.createTextOutput(JSON.stringify({ status: status }))
    .setMimeType(ContentService.MimeType.JSON);
}
