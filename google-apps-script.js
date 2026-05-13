// ============================================================
// NISC FORM HANDLER — Google Apps Script (UPDATED)
// Handles BOTH:
//   1. Member signups (from join.html) → "NISC Members" tab
//   2. Event registrations (from events.html) → "Event Registrations" tab
//
// SETUP STEPS:
// 1. Go to https://script.google.com → New Project
// 2. Paste this entire file
// 3. Click Deploy → New Deployment → Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Copy the Web App URL → paste into BOTH join.html and events.html
// ============================================================

const SHEET_ID     = '1VfbaqCL3Aca76jpgSocEtaalxIAeLPpKwyuL4CkxywE'; // ✅ Your Sheet ID
const NOTIFY_EMAIL = 'niscnpj@gmail.com';

// ─── ROUTER ─────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.formType === 'event_registration') {
      saveEventRegistration(data);
      sendEventNotification(data);
    } else {
      saveMemberSignup(data);
      sendMemberNotification(data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'NISC Form Handler Active' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── MEMBER SIGNUP (join.html) ───────────────────────────────────────────────
function saveMemberSignup(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('NISC Members');

  if (!sheet) {
    sheet = ss.insertSheet('NISC Members');
    const headers = [
      'Timestamp', 'First Name', 'Last Name', 'Email', 'Phone',
      'Age Range', 'City', 'Role', 'Institution', 'Field of Study',
      'Skill Level', 'Interests', 'Goal', 'How They Heard'
    ];
    sheet.appendRow(headers);
    styleHeader(sheet, headers.length, '#5b6eff');
  }

  sheet.appendRow([
    timestamp(),
    data.firstName   || '',
    data.lastName    || '',
    data.email       || '',
    data.phone       || '',
    data.age         || '',
    data.city        || '',
    data.role        || '',
    data.college     || '',
    data.field       || '',
    data.level       || '',
    data.interests   || '',
    data.goal        || '',
    data.source      || ''
  ]);
}

function sendMemberNotification(data) {
  const subject = `New NISC Member: ${data.firstName} ${data.lastName}`;
  const body = `
New member joined NISC!

NAME:          ${data.firstName} ${data.lastName}
EMAIL:         ${data.email}
PHONE:         ${data.phone}
AGE RANGE:     ${data.age}
CITY:          ${data.city}
ROLE:          ${data.role}
INSTITUTION:   ${data.college}
FIELD:         ${data.field || 'Not specified'}
SKILL LEVEL:   ${data.level}
INTERESTS:     ${data.interests}
GOAL:          ${data.goal}
HEARD FROM:    ${data.source || 'Not specified'}

View all members:
https://docs.google.com/spreadsheets/d/${SHEET_ID}
`;
  MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
}

// ─── EVENT REGISTRATION (events.html) ───────────────────────────────────────
function saveEventRegistration(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Event Registrations');

  if (!sheet) {
    sheet = ss.insertSheet('Event Registrations');
    const headers = [
      'Timestamp', 'Type', 'Event Title', 'Event Date', 'Event Type',
      'Event Location', 'Name', 'Email', 'Phone', 'City',
      'Level', 'Institution', 'Goal / Expectations',
      'Attendance Preference', 'How They Heard'
    ];
    sheet.appendRow(headers);
    styleHeader(sheet, headers.length, '#22d3ee');
  }

  sheet.appendRow([
    timestamp(),
    data.registrationType || 'Registration',
    data.eventTitle       || '',
    data.eventDate        || '',
    data.eventType        || '',
    data.eventLocation    || '',
    data.name             || '',
    data.email            || '',
    data.phone            || '',
    data.city             || '',
    data.level            || '',
    data.college          || '',
    data.goal             || '',
    data.preference       || '',
    data.source           || ''
  ]);
}

function sendEventNotification(data) {
  const icon = data.registrationType === 'Waitlist' ? '📋'
    : data.registrationType === 'Notification Request' ? '🔔' : '🎟';
  const subject = `${icon} ${data.registrationType}: ${data.name} — ${data.eventTitle}`;
  const body = `
${data.registrationType} received!

EVENT:         ${data.eventTitle}
DATE:          ${data.eventDate}
TYPE:          ${data.eventType}
LOCATION:      ${data.eventLocation}

NAME:          ${data.name}
EMAIL:         ${data.email}
PHONE:         ${data.phone}
CITY:          ${data.city}
LEVEL:         ${data.level}
INSTITUTION:   ${data.college || 'Not specified'}
PREFERENCE:    ${data.preference}
EXPECTATIONS:  ${data.goal}
HEARD FROM:    ${data.source || 'Not specified'}

View all registrations:
https://docs.google.com/spreadsheets/d/${SHEET_ID}
`;
  MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
}

// ─── HELPERS ────────────────────────────────────────────────────────────────
function timestamp() {
  return new Date().toLocaleString('en-NP', { timeZone: 'Asia/Kathmandu' });
}

function styleHeader(sheet, colCount, color) {
  const range = sheet.getRange(1, 1, 1, colCount);
  range.setBackground(color);
  range.setFontColor('#ffffff');
  range.setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, colCount);
}
