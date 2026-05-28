// ─────────────────────────────────────────────────────────────────────────────
//  UMass Athletics — Daily Wellness Reminder + Compliance Check
//
//  SETUP:
//  1. Open the Google Sheet linked to your wellness form
//  2. Extensions → Apps Script → paste this entire file
//  3. Fill in FORM_URL and COMPLIANCE_EMAIL below
//  4. Add athlete emails to the ROSTER sheet (see createRosterSheet below)
//  5. Run installTriggers() once from the editor to schedule daily reminders
// ─────────────────────────────────────────────────────────────────────────────

// ── Config ───────────────────────────────────────────────────────────────────

const FORM_URL         = 'https://YOUR_GITHUB_USERNAME.github.io/umass-wellness/';
const COMPLIANCE_EMAIL = 'your@email.com';   // who receives the daily compliance report
const REMINDER_HOUR    = 7;                   // 7 AM — morning reminder to athletes
const COMPLIANCE_HOUR  = 12;                  // Noon — compliance report to you
const WELLNESS_SHEET   = 'Wellness';          // sheet where form data lands
const ROSTER_SHEET     = 'Roster';            // sheet with athlete name + email list

// ─────────────────────────────────────────────────────────────────────────────
//  1. MORNING REMINDER — sent to all athletes on the roster
// ─────────────────────────────────────────────────────────────────────────────

function sendDailyReminder() {
  const athletes = getRoster();
  if (!athletes.length) return;

  const day    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEEE, MMMM d');
  const subject = `✅ Daily Wellness Check-In — ${day}`;
  const body    = `Hi {NAME},

Don't forget to complete your daily wellness check-in for today (${day}).

It takes less than 60 seconds:
${FORM_URL}

Save this link to your phone's home screen so it's always one tap away.

— UMass Athletics Sports Performance`;

  athletes.forEach(({ name, email }) => {
    if (!email) return;
    const firstName    = name.split(' ')[0];
    const personalBody = body.replace('{NAME}', firstName);
    GmailApp.sendEmail(email, subject, personalBody);
  });

  Logger.log(`Reminder sent to ${athletes.length} athletes.`);
}

// ─────────────────────────────────────────────────────────────────────────────
//  2. COMPLIANCE REPORT — sent to staff at noon showing who hasn't submitted
// ─────────────────────────────────────────────────────────────────────────────

function sendComplianceReport() {
  const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const athletes = getRoster();
  const submitted = getSubmittedToday(today);

  const missing = athletes.filter(a => !submitted.has(a.name.toLowerCase().trim()));
  const done    = athletes.filter(a =>  submitted.has(a.name.toLowerCase().trim()));

  const day     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEEE, MMMM d');
  const subject = `📋 Wellness Compliance — ${day} (${done.length}/${athletes.length} submitted)`;

  let body = `Wellness Check-In Compliance Report — ${day}\n\n`;
  body += `✅ Submitted (${done.length}): ${done.map(a => a.name).join(', ') || 'None'}\n\n`;
  body += `❌ Missing  (${missing.length}): ${missing.map(a => a.name).join(', ') || 'None'}\n\n`;
  body += `View full data: https://docs.google.com/spreadsheets/d/${SpreadsheetApp.getActiveSpreadsheet().getId()}`;

  GmailApp.sendEmail(COMPLIANCE_EMAIL, subject, body);
  Logger.log(`Compliance report sent. ${done.length}/${athletes.length} submitted.`);
}

// ─────────────────────────────────────────────────────────────────────────────
//  3. FOLLOW-UP NUDGE — re-send to athletes who haven't submitted by noon
// ─────────────────────────────────────────────────────────────────────────────

function sendFollowUpNudge() {
  const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const athletes = getRoster();
  const submitted = getSubmittedToday(today);

  const missing = athletes.filter(a => !submitted.has(a.name.toLowerCase().trim()) && a.email);
  if (!missing.length) return;

  const day     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEEE, MMMM d');
  const subject = `⏰ Reminder: Wellness Check-In still needed — ${day}`;
  const body    = `Hi {NAME},

We haven't received your wellness check-in yet today (${day}).

Please take 60 seconds to fill it out now:
${FORM_URL}

Thanks,
UMass Athletics Sports Performance`;

  missing.forEach(({ name, email }) => {
    const firstName    = name.split(' ')[0];
    const personalBody = body.replace('{NAME}', firstName);
    GmailApp.sendEmail(email, subject, personalBody);
  });

  Logger.log(`Follow-up nudge sent to ${missing.length} athletes.`);
}

// ─────────────────────────────────────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────────────────────────────────────

function getRoster() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ROSTER_SHEET);
  if (!sheet) { Logger.log('No Roster sheet found. Run createRosterSheet() first.'); return []; }
  const rows  = sheet.getDataRange().getValues().slice(1); // skip header
  return rows
    .filter(r => r[0])
    .map(r => ({ name: r[0].toString().trim(), email: r[1].toString().trim() }));
}

function getSubmittedToday(dateStr) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(WELLNESS_SHEET);
  if (!sheet) return new Set();
  const rows  = sheet.getDataRange().getValues().slice(1);
  const names = new Set();
  rows.forEach(r => {
    const rowDate = Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (rowDate === dateStr) names.add(r[2].toString().toLowerCase().trim()); // col C = Name
  });
  return names;
}

// ─────────────────────────────────────────────────────────────────────────────
//  ONE-TIME SETUP — run these functions once from the Apps Script editor
// ─────────────────────────────────────────────────────────────────────────────

// Run this once to create a blank Roster sheet with headers
function createRosterSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(ROSTER_SHEET);
  if (sheet) { Logger.log('Roster sheet already exists.'); return; }
  sheet = ss.insertSheet(ROSTER_SHEET);
  sheet.appendRow(['Name', 'Email', 'Sport']);
  sheet.getRange(1, 1, 1, 3).setBackground('#5C0000').setFontColor('white').setFontWeight('bold');
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 150);
  Logger.log('Roster sheet created. Add athlete names and emails.');
}

// Run this once to install all daily triggers
function installTriggers() {
  // Clear any existing triggers first
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // 7 AM — morning reminder to all athletes
  ScriptApp.newTrigger('sendDailyReminder')
    .timeBased().everyDays(1).atHour(REMINDER_HOUR).create();

  // 12 PM — compliance report to staff
  ScriptApp.newTrigger('sendComplianceReport')
    .timeBased().everyDays(1).atHour(COMPLIANCE_HOUR).create();

  // 1 PM — follow-up nudge to athletes who still haven't submitted
  ScriptApp.newTrigger('sendFollowUpNudge')
    .timeBased().everyDays(1).atHour(COMPLIANCE_HOUR + 1).create();

  Logger.log('✅ Triggers installed: reminders at 7am, compliance at noon, nudge at 1pm.');
}

// Run this to remove all triggers (e.g. off-season)
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('All triggers removed.');
}
