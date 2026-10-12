/**
 * atulnaik.in forms → Google Sheet + email notification.
 *   Contact form (index.html)          → "Responses" tab
 *   Privacy request form (privacy.html) → "Privacy requests" tab
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

const PRIVACY_SHEET = "Privacy requests";
const PRIVACY_HEADERS = ["Received at", "Email", "Request", "Details", "Status"];
const PRIVACY_TYPES = ["See my data", "Correct my data", "Delete my data", "Withdraw consent", "Other question"];

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    // Honeypot: real visitors never see or fill this field.
    if (d.website) return json({ ok: true });
    if (d.form === "privacy") return privacy_(d);

    const f = {};
    Object.keys(LIMITS).forEach(k => { f[k] = String(d[k] || "").trim().slice(0, LIMITS[k]); });

    if (!f.name || !f.process || !/^\S+@\S+\.\S+$/.test(f.email) || !/^\+?[\d\s()-]{7,20}$/.test(f.phone)) {
      return json({ ok: false, error: "invalid" });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sheet_(SHEET_NAME, HEADERS).appendRow([new Date(), f.name, f.company, f.process, f.today, f.goal, f.email, f.phone, f.page].map(safe_));
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

function privacy_(d) {
  const email = String(d.email || "").trim().slice(0, 200);
  const request = String(d.request || "").trim();
  const details = String(d.details || "").trim().slice(0, 2000);
  if (!/^\S+@\S+\.\S+$/.test(email) || PRIVACY_TYPES.indexOf(request) < 0 || !details) {
    return json({ ok: false, error: "invalid" });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet_(PRIVACY_SHEET, PRIVACY_HEADERS).appendRow([new Date(), email, request, details, "Open"].map(safe_));
  } finally {
    lock.releaseLock();
  }

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    replyTo: email,
    subject: "Privacy request: " + request,
    body:
      "New privacy request from the atulnaik.in privacy page\n\n" +
      "Email: " + email + "\n" +
      "Request: " + request + "\n\n" +
      details + "\n\n" +
      "Please respond within 30 days. Reply to this email to answer them directly.\n" +
      "All requests: " + SpreadsheetApp.getActiveSpreadsheet().getUrl()
  });

  return json({ ok: true });
}

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
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
