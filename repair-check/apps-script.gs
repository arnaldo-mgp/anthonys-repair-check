/**
 * Anthony's Auto Craft: Collision Repair Check lead receiver.
 *
 * Bind this script to the Google Sheet that should hold the leads
 * (Extensions > Apps Script inside the Sheet), then deploy it as a Web App.
 * Each POST from the quiz page becomes one row on the "Leads" tab and an email alert.
 */

// ---- Config -----------------------------------------------------------------
const NOTIFY_EMAILS = [
  'owner@example.com',
  // 'service@example.com',
];
const SHEET_NAME = 'Leads';

// Column order follows the quiz payload; answers are split into their own columns.
const HEADERS = [
  'name', 'phone', 'email', 'best_time',
  'situation', 'make', 'adas', 'insurance', 'referral', 'timing',
  'utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'utm_term',
  'gclid', 'fbclid', 'ndclid',
  'landing_url', 'referrer', 'submitted_at',
];
const ANSWER_KEYS = ['situation', 'make', 'adas', 'insurance', 'referral', 'timing'];

// Human-readable answers for the email.
const LABELS = {
  situation: {
    _q: 'What happened',
    collision: 'Accident or collision',
    minor: 'Dent, scratch or minor damage',
    redo: 'A previous repair they are not happy with',
    unsure: 'Not sure yet',
  },
  make: { _q: 'Vehicle make', Other: 'Other make' },
  adas: { _q: 'Driver-assist features', yes: 'Yes', no: 'No', unsure: 'Not sure' },
  insurance: {
    _q: 'Paying with',
    own: 'Claim with their own insurance',
    other: "Claim with the other driver's insurance",
    self: 'Paying out of pocket',
    unsure: "Haven't decided yet",
  },
  referral: {
    _q: 'Insurer recommended a shop?',
    yes: 'Yes, insurer suggested a shop / network',
    no: 'No',
    notyet: "Not yet, hasn't filed",
    '': '(skipped: paying out of pocket)',
  },
  timing: { _q: 'Needs car back', asap: 'As soon as possible', weeks: 'In the next few weeks', flexible: 'Flexible' },
};

// ---- Web App ------------------------------------------------------------------
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const answers = data.answers || {};

    const name = clean(data.name, 80);
    const phone = String(data.phone || '').replace(/\D/g, '').slice(0, 15);
    if (!name || phone.length < 10) return json({ ok: false, error: 'invalid' });

    const flat = {};
    HEADERS.forEach(function (h) {
      flat[h] = clean(ANSWER_KEYS.indexOf(h) > -1 ? answers[h] : data[h], h.indexOf('url') > -1 || h === 'referrer' ? 2000 : 200);
    });
    flat.name = name;
    flat.phone = phone;

    lock.waitLock(10000);
    sheet().appendRow(HEADERS.map(function (h) { return asCell(flat[h]); }));
    lock.releaseLock();

    sendAlert(flat);
    return json({ ok: true });
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: 'server' });
  }
}

// Visiting the /exec URL in a browser just confirms the deployment is live.
function doGet() {
  return json({ ok: true, service: 'repair-check' });
}

// ---- Helpers -------------------------------------------------------------------
function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function clean(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max || 200);
}

// Keep values as plain text: stops formula injection (=, +, -, @) and keeps phone digits intact.
function asCell(v) {
  return v === '' ? '' : "'" + v;
}

function label(key, value) {
  const map = LABELS[key] || {};
  return map[value] != null ? map[value] : (value || '-');
}

function formatPhone(d) {
  return d.length === 10 ? '(' + d.slice(0, 3) + ') ' + d.slice(3, 6) + '-' + d.slice(6) : d;
}

function html(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function sendAlert(f) {
  const recipients = NOTIFY_EMAILS.filter(String).join(',');
  if (!recipients) return;

  const make = f.make ? label('make', f.make) : 'no make';
  const subject = 'New repair check lead: ' + f.name + ' (' + make + ')';
  const phonePretty = formatPhone(f.phone);
  const bestTime = f.best_time || 'No preference';

  const rows = [
    ['Name', html(f.name)],
    ['Phone', '<a href="tel:+1' + f.phone + '">' + html(phonePretty) + '</a>'],
    ['Email', f.email ? '<a href="mailto:' + html(f.email) + '">' + html(f.email) + '</a>' : '-'],
    ['Best time to call', html(bestTime)],
  ];
  ANSWER_KEYS.forEach(function (k) { rows.push([LABELS[k]._q, html(label(k, f[k]))]); });
  const source = ['utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'utm_term', 'gclid', 'fbclid', 'ndclid']
    .filter(function (k) { return f[k]; })
    .map(function (k) { return k + ': ' + html(f[k]); })
    .join('<br>') || 'Direct / no tracking';
  rows.push(['Source', source]);
  rows.push(['Landing page', html(f.landing_url || '-')]);
  rows.push(['Submitted', html(f.submitted_at || new Date().toISOString())]);

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;font-size:15px;color:#111">' +
    '<h2 style="margin:0 0 6px">New repair check lead</h2>' +
    '<p style="margin:0 0 16px">Call <a href="tel:+1' + f.phone + '"><b>' + html(phonePretty) + '</b></a> ' +
    '(' + html(bestTime) + ') to book an in-person estimate.</p>' +
    '<table cellpadding="6" style="border-collapse:collapse">' +
    rows.map(function (r) {
      return '<tr><td style="color:#666;vertical-align:top;white-space:nowrap">' + r[0] + '</td><td>' + r[1] + '</td></tr>';
    }).join('') +
    '</table></div>';

  const textBody = rows.map(function (r) { return r[0] + ': ' + r[1].replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(); }).join('\n');

  MailApp.sendEmail({ to: recipients, subject: subject, body: textBody, htmlBody: htmlBody, name: 'Repair Check' });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Run once from the editor to authorize the script and check the sheet + email.
function testLead() {
  const res = doPost({ postData: { contents: JSON.stringify({
    name: 'Test', phone: '4155551234', email: '', best_time: 'Morning',
    answers: { situation: 'collision', make: 'Tesla', adas: 'yes', insurance: 'own', referral: 'yes', timing: 'asap' },
    landing_url: 'https://example.com/repair-check', referrer: '', submitted_at: new Date().toISOString(),
  }) } });
  console.log(res.getContent());
}
