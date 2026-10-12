/**
 * atulnaik.in contact form → Google Sheet + email notification.
 *
 * Paste this into the Apps Script editor of the "AN - Portfolio Form Responses"
 * sheet (Extensions → Apps Script), then Deploy → New deployment → Web app,
 * Execute as: Me, Who has access: Anyone. Put the /exec URL in index.html
 * (FORM_ENDPOINT).
 */

const NOTIFY_EMAIL = "atulnaikofficial@gmail.com";
const SHEET_NAME = "Responses";
const HEADERS = ["Received at", "Name", "Company", "Process", "Today", "Goal", "Email", "Phone", "Page"];
const LIMITS = { name: 120, company: 160, process: 1000, today: 1000, goal: 1000, email: 200, phone: 30, page: 200 };

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    // Honeypot: real visitors never see or fill this field.
    if (d.website) return json({ ok: true });

    const f = {};
    Object.keys(LIMITS).forEach(k => { f[k] = String(d[k] || "").trim().slice(0, LIMITS[k]); });

    if (!f.name || !f.process || !/^\S+@\S+\.\S+$/.test(f.email) || !/^\+?[\d\s()-]{7,20}$/.test(f.phone)) {
      return json({ ok: false, error: "invalid" });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sheet_().appendRow([new Date(), f.name, f.company, f.process, f.today, f.goal, f.email, f.phone, f.page].map(safe_));
    } finally {
      lock.releaseLock();
    }

    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      replyTo: f.email,
      subject: "New enquiry from " + f.name + (f.company ? " (" + f.company + ")" : ""),
      body:
        "New message from the atulnaik.in contact form\n\n" +
        "Name: " + f.name + "\n" +
        "Company: " + (f.company || "-") + "\n" +
        "Email: " + f.email + "\n" +
        "Phone: " + f.phone + "\n\n" +
        "Struggling with: " + f.process + "\n" +
        "Manages it today with: " + (f.today || "-") + "\n" +
        "Would like to: " + (f.goal || "-") + "\n\n" +
        "Reply to this email to answer them directly.\n" +
        "All responses: " + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    });

    return json({ ok: true });
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: "server" });
  }
}

// Lets you open the /exec URL in a browser to check the deployment is live.
function doGet() {
  return json({ ok: true, status: "Form endpoint is live" });
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

// Stop text like "=HYPERLINK(...)" being run as a formula in the sheet.
function safe_(v) {
  return typeof v === "string" && /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
