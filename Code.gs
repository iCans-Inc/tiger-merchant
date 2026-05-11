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

// ── Entry point ───────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    saveToSheet(data);
    sendEmails(data);
    return ok({ success: true });
  } catch (err) {
    Logger.log(err.toString());
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
  const address    = join([d.address1, d.city, d.state, d.zip], ', ');
  const bizType    = BUSINESS_TYPE_LABELS[d.businessType] || d.businessType || '';

  sheet.appendRow([
    d.dba || d.legalName,                          // Business Name
    d.legalName        || '',                       // Legal Company Name
    ownerName,                                      // Full Name
    d.contactPhone     || d.owner1Phone || '',      // Phone
    d.owner1Email      || '',                       // Email
    address,                                        // Business Address
    d.billingAddress   || '',                       // Billing Address
    d.website          || '',                       // Website
    bizType,                                        // Business Type
    d.businessStartDate || '',                      // Business Start Date
    d.whereToSendStatements || '',                  // Where to send Statements
    ownerName,                                      // Owner Name
    d.owner1Dob        || '',                       // Owner DOB
    d.owner1Ownership  ? d.owner1Ownership + '%' : '', // Ownership %
    d.owner1Title      || '',                       // Company Title
    d.owner1Phone      || d.contactPhone || '',     // Owner Phone
    d.owner1Address    || d.address1 || '',         // Owner Address
    owner2Name,                                     // Second Owner Name
    d.owner2Dob        || '',                       // Second Owner DOB
    d.owner2Ownership  ? d.owner2Ownership + '%' : '', // Second Owner %
    d.owner2Title      || '',                       // Second Owner Title
    d.owner2Phone      || '',                       // Second Owner Phone
    d.owner2Address    || '',                       // Second Owner Address
    d.grossYearlySales || '',                       // Gross Yearly Sales
    d.monthlyVolume    || '',                       // Yearly Credit Card Sales
    d.avgTicket        || '',                       // AVG Transaction Amount
    d.highestTicket    || '',                       // Highest Transaction Amount
    d.numTrucks        || '',                       // Number of Trucks
    d.numDumpsters     || '',                       // Number of Dumpsters
  ]);
}

// ── Emails ────────────────────────────────────────────────────────────────────
function sendEmails(d) {
  const attachments = buildAttachments(d.files || []);
  const subject     = 'New Tiger Merchant Application – ' + (d.dba || d.legalName || 'Unknown');
  const opts        = { name: FROM_NAME, attachments: attachments };

  // onboarding@icans.ai — no banking
  GmailApp.sendEmail(
    ONBOARDING_EMAIL, subject, '',
    Object.assign({}, opts, { htmlBody: buildEmail(d, false) })
  );

  // zach@tigerprocessing.com — everything
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

// ── HTML email builder ────────────────────────────────────────────────────────
function buildEmail(d, includeBanking) {
  const ownerName  = join([d.owner1FirstName, d.owner1LastName]);
  const owner2Name = d.hasSecondOwner ? join([d.owner2FirstName, d.owner2LastName]) : '';
  const address    = join([d.address1, d.city, d.state, d.zip], ', ');
  const bizType    = BUSINESS_TYPE_LABELS[d.businessType] || d.businessType || '';
  const dateStr    = new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });

  var h = '';
  h += '<div style="font-family:Inter,Arial,sans-serif;max-width:680px;margin:0 auto;background:#F4F3F8;padding:24px;">';

  // Header
  h += '<div style="background:#494A7D;border-radius:12px;padding:24px 28px;margin-bottom:16px;">';
  h += '<p style="color:#B6F3BF;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.1em;margin:0 0 6px;">Tiger Payments · New Application</p>';
  h += '<h1 style="color:#fff;font-size:22px;font-weight:700;letter-spacing:-0.02em;margin:0 0 4px;">' + esc(d.dba || d.legalName) + '</h1>';
  h += '<p style="color:rgba(255,255,255,0.65);font-size:13px;margin:0;">Submitted ' + dateStr + (includeBanking ? ' · Full application including banking' : ' · Excluding banking information') + '</p>';
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
    row('SSN (Last 4)',    d.owner1Ssn),
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
      row('SSN (Last 4)',     d.owner2Ssn),
      row('Email',            d.owner2Email),
      row('Phone',            d.owner2Phone),
      row('Address',          d.owner2Address),
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
