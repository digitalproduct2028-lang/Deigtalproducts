/**
 * DigitalProducts.design – Free sample leads
 * --------------------------------------------------------------
 * Paste this whole file into Google Apps Script (Extensions > Apps Script)
 * inside the Google Sheet where you want to save leads.
 * Then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Full steps are in SETUP-GUIDE.md.
 */

// ===== 1. YOUR SETTINGS =====
const SHEET_NAME   = "Leads";                    // tab name (created automatically)
const BRAND_NAME   = "DigitalProducts.design";
const WEBSITE_URL  = "https://digitalproducts.design/";
const REPLY_TO     = "digitalproduct2028@gmail.com";

// Sample links are only emailed if they start with one of these.
// This stops strangers from using your Gmail to send other links.
const ALLOWED_LINK_PREFIXES = [
  "https://digitalproducts.design/",
  "https://www.digitalproducts.design/",
  "https://drive.google.com/"
];
// Used when the website sends a link that isn't allowed (e.g. while testing on your computer)
const DEFAULT_SAMPLE_LINK = "https://drive.google.com/file/d/1f1fRjtgZLSaecHdl-04G1cks0VaKAVJQ/view?usp=sharing";

const HEADERS = ["Date & Time", "Type", "Email / Mobile", "Country", "Product", "Updates OK?", "Email sent", "Page", "Sample link"];

// ===== 2. MAIN CODE (no need to edit) =====
const EMAIL_RE  = /^[a-z0-9](?:[a-z0-9._%+-]{0,62}[a-z0-9])?@(gmail\.com|googlemail\.com|ymail\.com|rocketmail\.com|yahoo\.(?:com|in|co\.in|co\.uk|ca|com\.au|fr|de|com\.sg|com\.my|com\.ph))$/i;
const MOBILE_RE = /^\+[1-9]\d{6,14}$/;

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    // spam trap filled = robot: pretend success, save nothing
    if (data.website) return reply({ ok: true });

    const type    = data.type === "mobile" ? "mobile" : "email";
    const contact = String(data.contact || "").trim().toLowerCase().slice(0, 120);
    const product = clean(data.product, 120);

    if (type === "email" && !EMAIL_RE.test(contact)) return reply({ ok: false, error: "Please use a Gmail or Yahoo address." });
    if (type === "mobile" && !MOBILE_RE.test(contact)) return reply({ ok: false, error: "Please enter a valid mobile number with country code." });

    let link = String(data.sampleLink || "");
    if (!ALLOWED_LINK_PREFIXES.some(p => link.indexOf(p) === 0)) link = DEFAULT_SAMPLE_LINK;

    let emailSent = "—";
    if (type === "email") {
      const options = { htmlBody: emailHtml(product, link), name: BRAND_NAME };
      if (REPLY_TO) options.replyTo = REPLY_TO;
      MailApp.sendEmail(contact, "Your free sample: " + product, plainText(product, link), options);
      emailSent = "Yes";
    }

    getSheet().appendRow([
      new Date(),
      type === "email" ? "Email" : "Mobile",
      safe(contact),
      clean(data.country, 60),
      product,
      data.consent === "Yes" ? "Yes" : "No",
      emailSent,
      clean(data.pageUrl, 200),
      link
    ]);

    return reply({ ok: true });
  } catch (err) {
    console.error(err);
    return reply({ ok: false, error: "Server error. Please try again." });
  } finally {
    lock.releaseLock();
  }
}

// Open the Web App URL in your browser to check it is live.
function doGet() {
  return ContentService.createTextOutput("DigitalProducts.design lead form is working ✅");
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#0b1f6b").setFontColor("#ffffff");
    sh.setFrozenRows(1);
  }
  return sh;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function clean(v, max) { return safe(String(v || "").replace(/[\r\n\t]+/g, " ").trim().slice(0, max)); }
// stops text starting with = + - @ from being treated as a formula
function safe(v) { v = String(v); return /^[=+\-@]/.test(v) ? "'" + v : v; }
function escHtml(s) { return String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }

function plainText(product, link) {
  return "Hi!\n\nThanks for your interest in " + product + ".\nDownload your free sample here: " + link +
    "\n\nPrint it, try it with your child, and see how they like it.\nSee all worksheets: " + WEBSITE_URL + "\n\n– " + BRAND_NAME;
}

function emailHtml(product, link) {
  const p = escHtml(product.replace(/^'/, ""));
  return `
  <div style="background:#eaf4ff;padding:24px 12px;font-family:Arial,Helvetica,sans-serif">
    <div style="max-width:520px;margin:0 auto;background:#ffffff;border-radius:20px;overflow:hidden">
      <div style="background:#0b1f6b;padding:20px 24px;color:#ffffff;font-size:20px;font-weight:bold">${BRAND_NAME}</div>
      <div style="padding:28px 24px;color:#1d2747;font-size:16px;line-height:1.6">
        <h1 style="margin:0 0 10px;font-size:26px;color:#0b1f6b">Your free sample is here! 🎉</h1>
        <p style="margin:0 0 18px">Thanks for your interest in <b>${p}</b>. Tap the button to download your sample worksheet.</p>
        <p style="margin:0 0 24px"><a href="${escHtml(link)}" style="display:inline-block;background:#ff6a13;color:#ffffff;text-decoration:none;font-weight:bold;padding:14px 26px;border-radius:12px">Download sample worksheet</a></p>
        <p style="margin:0 0 6px">Print it, try it with your child, and see how they like it.</p>
        <p style="margin:0">Want the full bundle? <a href="${WEBSITE_URL}#products" style="color:#1360f0;font-weight:bold">See all worksheets</a></p>
      </div>
      <div style="padding:16px 24px;background:#f4f8ff;color:#6b7593;font-size:12px">You received this because you asked for a free sample on ${BRAND_NAME}. Reply to this email if you'd like us to delete your details.</div>
    </div>
  </div>`;
}
