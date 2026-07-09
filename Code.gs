// ─────────────────────────────────────────────────────────────────────────────
// Tiger Merchant Application — Google Apps Script Backend
//
// SETUP INSTRUCTIONS:
//   1. Go to script.google.com and create a new project
//   2. Paste this entire file into Code.gs
//   3. Click Deploy → New deployment → Web app
//   4. Execute as: tigerapplication@icans.ai (sign in with that account)
//   5. Who has access: Anyone
//   6. Copy the Web App URL
//   7. In Vercel dashboard → tiger-merchant project → Settings → Environment Variables
//      Add: APPS_SCRIPT_URL = <the Web App URL you copied>
//   8. Redeploy Vercel (or push a commit) to pick up the new env var
// ─────────────────────────────────────────────────────────────────────────────

const SHEET_ID             = '1ScqVmUxIhHnwOrpc3jhsdMUx44Wmbarz_qkkl_kF0zc';
const ONBOARDING_EMAIL     = 'onboarding@icans.ai';
const ZACH_EMAIL           = 'zach@tigerprocessing.com';
const FROM_NAME            = 'Tiger Application Portal';

const BUSINESS_TYPE_LABELS = {
  sole:    'Individual / Sole Proprietorship',
  llc:     'LLC',
  scorp:   'S Corporation',
  ccorp:   'C Corporation',
  partner: 'Partnership',
  nonprofit: 'Non-profit',
};

// ── iCore backfill ────────────────────────────────────────────────────────────
// Run backfillToICore() ONCE manually from the Apps Script editor to replay all
// submissions from June 1 2025 onward into iCore (checklist + notes).
// Sheet columns: A=Date, B=DBA, C=LegalName, D=OwnerName, E=Phone, F=Email,
//   G=Address, H=BillingAddr, I=Website, J=BizType, K=StartDate,
//   L=Statements, M=OwnerName, N=OwnerDOB, O=Ownership%, P=Title,
//   Q=OwnerPhone, R=OwnerAddr, S=Owner2Name, T=Owner2DOB, U=Owner2%,
//   V=Owner2Title, W=Owner2Phone, X=Owner2Addr, Y=GrossYearlySales,
//   Z=MonthlyCCVol, AA=Trucks, AB=Dumpsters

const ICORE_WEBHOOK_URL    = 'https://icore.icans.ai/api/webhooks/tiger-merchant';
const ICORE_WEBHOOK_SECRET = 'a153e4694e1eac08d0e32f824b94edaa64200f7d57ba6ac45e0ad78958327e74';

function backfillToICore() {
  const sheet   = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const rows    = sheet.getDataRange().getValues();
  const cutoff  = new Date('2026-06-01T00:00:00');
  var sent = 0, skipped = 0, errors = 0;

  for (var i = 1; i < rows.length; i++) {  // skip header row 0
    var r = rows[i];
    var dateVal = r[0];  // Column A — Date Added
    var rowDate = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(rowDate) || rowDate < cutoff) { skipped++; continue; }

    // Split "First Last" owner names on first space
    var ownerFull  = String(r[3] || '').trim();
    var ownerParts = ownerFull.split(/\s+/);
    var o1First    = ownerParts[0] || '';
    var o1Last     = ownerParts.slice(1).join(' ') || '';

    var owner2Full  = String(r[18] || '').trim();  // S
    var owner2Parts = owner2Full.split(/\s+/);
    var o2First     = owner2Parts[0] || '';
    var o2Last      = owner2Parts.slice(1).join(' ') || '';
    var hasOwner2   = !!owner2Full;

    // Reconstruct payload from sheet columns (best-effort — no SSN/banking)
    var payload = {
      submittedAt:          String(r[0]  || ''),
      dba:                  String(r[1]  || ''),
      legalName:            String(r[2]  || ''),
      owner1FirstName:      o1First,
      owner1LastName:       o1Last,
      contactPhone:         String(r[4]  || ''),
      owner1Email:          String(r[5]  || ''),
      address1:             String(r[6]  || ''),
      billingAddress:       String(r[7]  || ''),
      website:              String(r[8]  || ''),
      businessType:         String(r[9]  || ''),
      businessStartDate:    String(r[10] || ''),
      whereToSendStatements:String(r[11] || ''),
      owner1Dob:            String(r[13] || ''),
      owner1Ownership:      String(r[14] || '').replace('%',''),
      owner1Title:          String(r[15] || ''),
      owner1Phone:          String(r[16] || ''),
      owner1Address:        String(r[17] || ''),
      hasSecondOwner:       hasOwner2,
      owner2FirstName:      o2First,
      owner2LastName:       o2Last,
      owner2Dob:            String(r[19] || ''),
      owner2Ownership:      String(r[20] || '').replace('%',''),
      owner2Title:          String(r[21] || ''),
      owner2Phone:          String(r[22] || ''),
      owner2Address:        String(r[23] || ''),
      grossYearlySales:     String(r[24] || ''),
      monthlyVolume:        String(r[25] || ''),
      numTrucks:            String(r[26] || ''),
      numDumpsters:         String(r[27] || ''),
      files: [],
    };

    try {
      var response = UrlFetchApp.fetch(ICORE_WEBHOOK_URL, {
        method: 'post',
        contentType: 'application/json',
        headers: { 'x-webhook-secret': ICORE_WEBHOOK_SECRET },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });
      var code = response.getResponseCode();
      var body = response.getContentText().slice(0, 200);
      Logger.log('Row ' + (i+1) + ' [' + (payload.dba || payload.legalName) + ']: HTTP ' + code + ' — ' + body);
      if (code >= 200 && code < 300) { sent++; } else { errors++; }
    } catch (err) {
      Logger.log('Row ' + (i+1) + ' ERROR: ' + err.toString());
      errors++;
    }

    Utilities.sleep(300); // be polite — avoid rate limits
  }

  Logger.log('Backfill complete — sent: ' + sent + ', skipped (before June 1): ' + skipped + ', errors: ' + errors);
}

// ── Auth test — run this ONCE in the editor to authorize Gmail ────────────────
function testGmail() {
  GmailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    'Tiger App Script — Gmail auth confirmed',
    'Gmail is authorized. You can delete this test.'
  );
  Logger.log('Test email sent to ' + Session.getActiveUser().getEmail());
}

// ── SSN helpers ───────────────────────────────────────────────────────────────
// Returns last-4 masked: •••-••-1234
function maskSsn(ssn) {
  if (!ssn) return '';
  var digits = String(ssn).replace(/\D/g, '');
  if (digits.length < 4) return '•••-••-' + digits;
  return '•••-••-' + digits.slice(-4);
}

// ── Entry point ───────────────────────────────────────────────────────────────
function doPost(e) {
  var sheetError = null;
  var emailError = null;

  try {
    const data = JSON.parse(e.postData.contents);

    try {
      saveToSheet(data);
    } catch (err) {
      sheetError = err.toString();
      Logger.log('Sheet error: ' + sheetError);
    }

    try {
      sendEmails(data);
    } catch (err) {
      emailError = err.toString();
      Logger.log('Email error: ' + emailError);
    }

    if (sheetError) {
      return ok({ success: false, error: 'Sheet: ' + sheetError });
    }
    if (emailError) {
      // Sheet saved — return a partial success so the form still advances
      // but include the email error for visibility
      return ok({ success: true, emailError: emailError });
    }
    return ok({ success: true });
  } catch (err) {
    Logger.log('Parse error: ' + err.toString());
    return ok({ success: false, error: err.toString() });
  }
}

// OPTIONS preflight (called directly, not via proxy)
function doOptions() {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON);
}

function ok(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Google Sheets ─────────────────────────────────────────────────────────────
function saveToSheet(d) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];

  const ownerName  = join([d.owner1FirstName, d.owner1LastName]);
  const owner2Name = d.hasSecondOwner ? join([d.owner2FirstName, d.owner2LastName]) : '';
  const owner3Name = d.hasThirdOwner  ? join([d.owner3FirstName, d.owner3LastName]) : '';
  const address    = join([d.address1, d.city, d.state, d.zip], ', ');
  const bizType    = BUSINESS_TYPE_LABELS[d.businessType] || d.businessType || '';

  sheet.appendRow([
    d.submittedAt           || new Date().toLocaleString('en-US'), // A  Date Added
    d.dba || d.legalName,                                          // B  Business Name
    d.legalName             || '',                                 // C  Legal Company Name
    ownerName,                                                     // D  Full Name
    d.contactPhone          || d.owner1Phone || '',                // E  Phone
    d.owner1Email           || '',                                 // F  Email
    address,                                                       // G  Business Address
    d.billingAddress        || '',                                 // H  Billing Address
    d.website               || '',                                 // I  Website
    bizType,                                                       // J  Business Type
    d.businessStartDate     || '',                                 // K  Business Start Date
    d.whereToSendStatements || '',                                 // L  Where to send Statements
    ownerName,                                                     // M  Owner Name
    d.owner1Dob             || '',                                 // N  Owner DOB
    d.owner1Ownership       ? d.owner1Ownership + '%' : '',        // O  Ownership %
    d.owner1Title           || '',                                 // P  Company Title
    d.owner1Phone           || d.contactPhone || '',               // Q  Owner Phone
    d.owner1Address         || d.address1 || '',                   // R  Owner Address
    owner2Name,                                                    // S  Second Owner Name
    d.owner2Dob             || '',                                 // T  Second Owner DOB
    d.owner2Ownership       ? d.owner2Ownership + '%' : '',        // U  Second Owner %
    d.owner2Title           || '',                                 // V  Second Owner Title
    d.owner2Phone           || '',                                 // W  Second Owner Phone
    d.owner2Address         || '',                                 // X  Second Owner Address
    d.grossYearlySales      || '',                                 // Y  Gross Yearly Sales
    d.monthlyVolume         || '',                                 // Z  Yearly Credit Card Sales
    d.numTrucks             || '',                                 // AA Number of Trucks
    d.numDumpsters          || '',                                 // AB Number of Dumpsters
    d.signature ? 'Signed ✓' : '',                                 // AC Signature captured
  ]);
}

// ── Emails ────────────────────────────────────────────────────────────────────
function sendEmails(d) {
  const attachments  = buildAttachments(d.files || []);
  const inlineImages = {};

  // Signature — attach as signature.png (so Tiger can drop it into their PDF)
  // and reference it inline in the email body via cid:signature.
  const sig = signatureBlob(d.signature);
  if (sig) {
    attachments.push(sig.copyBlob().setName('signature.png'));
    inlineImages.signature = sig;
  }

  const subject     = 'New Tiger Merchant Application – ' + (d.dba || d.legalName || 'Unknown');
  const opts        = { name: FROM_NAME, attachments: attachments, inlineImages: inlineImages };

  // onboarding@icans.ai — no banking, SSN masked to last 4
  GmailApp.sendEmail(
    ONBOARDING_EMAIL, subject, '',
    Object.assign({}, opts, { htmlBody: buildEmail(d, false) })
  );

  // zach@tigerprocessing.com — everything including banking + full SSN
  GmailApp.sendEmail(
    ZACH_EMAIL, subject, '',
    Object.assign({}, opts, { htmlBody: buildEmail(d, true) })
  );
}

function buildAttachments(files) {
  return files.map(function(f) {
    const bytes = Utilities.base64Decode(f.data);
    return Utilities.newBlob(bytes, f.type || 'application/octet-stream', f.name);
  });
}

// Decode a data: URL (e.g. "data:image/png;base64,....") into a Blob.
function signatureBlob(dataUrl) {
  if (!dataUrl || String(dataUrl).indexOf('data:image') !== 0) return null;
  var parts = String(dataUrl).split(',');
  var mime  = (parts[0].match(/data:(.*?);/) || [])[1] || 'image/png';
  var b64   = parts[1] || '';
  if (!b64) return null;
  var bytes = Utilities.base64Decode(b64);
  return Utilities.newBlob(bytes, mime, 'signature.png');
}

// ── HTML email builder ────────────────────────────────────────────────────────
function buildEmail(d, includeBanking) {
  const ownerName  = join([d.owner1FirstName, d.owner1LastName]);
  const owner2Name = d.hasSecondOwner ? join([d.owner2FirstName, d.owner2LastName]) : '';
  const owner3Name = d.hasThirdOwner  ? join([d.owner3FirstName, d.owner3LastName]) : '';
  const address    = join([d.address1, d.city, d.state, d.zip], ', ');
  const bizType    = BUSINESS_TYPE_LABELS[d.businessType] || d.businessType || '';
  const dateStr    = new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });

  // SSN display: full for Zach, last-4 masked for onboarding
  const ssn1Display = includeBanking ? (d.owner1Ssn || '') : maskSsn(d.owner1Ssn);
  const ssn2Display = includeBanking ? (d.owner2Ssn || '') : maskSsn(d.owner2Ssn);
  const ssn3Display = includeBanking ? (d.owner3Ssn || '') : maskSsn(d.owner3Ssn);

  var h = '';
  h += '<div style="font-family:Inter,Arial,sans-serif;max-width:680px;margin:0 auto;background:#F4F3F8;padding:24px;">';

  // Header
  h += '<div style="background:#494A7D;border-radius:12px;padding:24px 28px;margin-bottom:16px;">';
  h += '<p style="color:#B6F3BF;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.1em;margin:0 0 6px;">Tiger Payments · New Application</p>';
  h += '<h1 style="color:#fff;font-size:22px;font-weight:700;letter-spacing:-0.02em;margin:0 0 4px;">' + esc(d.dba || d.legalName) + '</h1>';
  h += '<p style="color:rgba(255,255,255,0.65);font-size:13px;margin:0;">Submitted ' + (d.submittedAt || dateStr) + (includeBanking ? ' · Full application including banking' : ' · Excluding banking information') + '</p>';
  h += '</div>';

  // Business
  h += section('Business Profile', [
    row('Business Name (DBA)', d.dba),
    row('Legal Name',          d.legalName),
    row('Business Type',       bizType),
    row('Federal EIN',         d.ein),
    row('Industry',            d.industry),
    row('Website',             d.website),
    row('Business Start Date', d.businessStartDate),
    row('Business Address',    address),
    row('Billing Address',     d.billingAddress),
    row('Where to Send Statements', d.whereToSendStatements),
    row('Contact Phone',       d.contactPhone),
  ]);

  // Owners
  var ownerRows = [
    row('Primary Owner',   ownerName),
    row('Title',           d.owner1Title),
    row('Ownership',       d.owner1Ownership ? d.owner1Ownership + '%' : ''),
    row('Date of Birth',   d.owner1Dob),
    row('SSN',             ssn1Display),
    row('Email',           d.owner1Email),
    row('Phone',           d.owner1Phone),
    row('Address',         d.owner1Address || address),
  ];
  if (d.hasSecondOwner && owner2Name) {
    ownerRows.push(divider());
    ownerRows = ownerRows.concat([
      row('Additional Owner', owner2Name),
      row('Title',            d.owner2Title),
      row('Ownership',        d.owner2Ownership ? d.owner2Ownership + '%' : ''),
      row('Date of Birth',    d.owner2Dob),
      row('SSN',              ssn2Display),
      row('Email',            d.owner2Email),
      row('Phone',            d.owner2Phone),
      row('Address',          d.owner2Address),
    ]);
  }
  if (d.hasThirdOwner && owner3Name) {
    ownerRows.push(divider());
    ownerRows = ownerRows.concat([
      row('Third Owner',  owner3Name),
      row('Title',        d.owner3Title),
      row('Ownership',    d.owner3Ownership ? d.owner3Ownership + '%' : ''),
      row('Date of Birth',d.owner3Dob),
      row('SSN',          ssn3Display),
      row('Email',        d.owner3Email),
      row('Phone',        d.owner3Phone),
      row('Address',      d.owner3Address),
    ]);
  }
  h += section('Ownership', ownerRows);

  // Processing
  h += section('Processing Volume', [
    row('Monthly Volume',      d.monthlyVolume),
    row('Gross Yearly Sales',  d.grossYearlySales),
    row('Average Ticket',      d.avgTicket),
    row('Highest Ticket',      d.highestTicket),
    row('Current Processor',   d.currentProcessor),
    row('Number of Trucks',    d.numTrucks),
    row('Number of Dumpsters', d.numDumpsters),
  ]);

  // Banking (Zach only)
  if (includeBanking) {
    h += section('Banking', [
      row('Bank Name',       d.bankName),
      row('Routing Number',  d.bankRoutingNumber),
      row('Account Number',  d.bankAccountNumber),
      row('Account Type',    d.bankAccountType),
    ]);
  }

  // Signature
  if (d.signature) {
    h += '<div style="background:#fff;border-radius:10px;padding:18px 20px;margin-bottom:14px;box-shadow:0 1px 3px rgba(0,0,0,0.06);">';
    h += '<p style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.1em;color:#2E9039;margin:0 0 12px;">Signature</p>';
    h += '<img src="cid:signature" alt="Customer signature" style="max-width:320px;height:auto;border:1px solid #E8E7EE;border-radius:8px;background:#fff;"/>';
    h += '<p style="font-size:12px;color:#6E6A93;margin:10px 0 0;">Electronically signed by ' + esc(ownerName) + ' · ' + esc(d.submittedAt || dateStr) + '</p>';
    h += '</div>';
  }

  h += '<p style="font-size:11px;color:#9B9DBC;text-align:center;margin-top:20px;">Tiger Payment Solutions · tigerapplication@icans.ai</p>';
  h += '</div>';
  return h;
}

function section(title, rows) {
  var h = '<div style="background:#fff;border-radius:10px;padding:18px 20px;margin-bottom:14px;box-shadow:0 1px 3px rgba(0,0,0,0.06);">';
  h += '<p style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.1em;color:#2E9039;margin:0 0 12px;">' + title + '</p>';
  rows.forEach(function(r) { h += r; });
  h += '</div>';
  return h;
}

function row(label, value) {
  var v = value ? esc(String(value)) : '<span style="color:#B5B2C9;font-style:italic;">—</span>';
  return '<div style="display:flex;gap:10px;padding:6px 0;border-bottom:1px solid #F4F3F8;font-size:13px;line-height:1.4;">'
    + '<span style="color:#6E6A93;min-width:170px;flex-shrink:0;font-size:12px;">' + label + '</span>'
    + '<span style="color:#1A1830;font-weight:500;">' + v + '</span>'
    + '</div>';
}

function divider() {
  return '<div style="border-top:1px dashed #E8E7EE;margin:10px 0 6px;"></div>';
}

function esc(s) {
  if (!s) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function join(arr, sep) {
  return (arr || []).filter(Boolean).join(sep || ' ');
}
