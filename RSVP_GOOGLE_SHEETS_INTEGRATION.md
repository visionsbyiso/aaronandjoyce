# RSVP Google Sheets Integration Guide

This guide connects your RSVP form to this Google Sheet:

`1m0NbG85E69YSyl64UR9DGsn_cEAG5kHacDC7-LGS-40`

## 1. Create the Apps Script Web App

1. Open the Google Sheet.
2. Go to `Extensions` > `Apps Script`.
3. Replace the default code with this:

```javascript
const SHEET_ID = '1m0NbG85E69YSyl64UR9DGsn_cEAG5kHacDC7-LGS-40';
const SHEET_NAME = 'RSVP';
const HEADERS = [
  'Timestamp',
  'Name',
  'Attendance',
  'Guests',
  'Plus One Name',
  'Message',
  'Event Date',
  'Event Time',
  'Event Timezone',
  'Submitted At (ISO)',
  'Page'
];

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }
  ensureHeaders_(sh);
  return sh;
}

function ensureHeaders_(sh) {
  const width = Math.max(sh.getLastColumn(), HEADERS.length);
  const firstRow = sh.getRange(1, 1, 1, width).getValues()[0].map(v => String(v).trim());
  const isEmpty = firstRow.every(v => !v);
  if (isEmpty) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }

  // Backward-compatibility: older setup had no "Plus One Name" column.
  if (!firstRow.includes('Plus One Name')) {
    sh.insertColumnAfter(4);
    sh.getRange(1, 5).setValue('Plus One Name');
  }
}

function doPost(e) {
  const p = e.parameter || {};
  const sh = getSheet_();
  const plusOneName = p.plus_one_name || p.guests || '';

  sh.appendRow([
    new Date(),
    p.name || '',
    p.attending || '',
    p.guests || plusOneName,
    plusOneName,
    p.message || '',
    p.event_date || '',
    p.event_time || '',
    p.event_timezone || '',
    p.submitted_at || '',
    p.page || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

## 2. Deploy

1. Click `Deploy` > `New deployment`.
2. Select type `Web app`.
3. `Execute as`: `Me`.
4. `Who has access`: `Anyone` (or `Anyone with Google account` if all guests are signed in).
5. Click `Deploy`.
6. Copy the Web App URL ending with `/exec`.

## 3. Connect to Your HTML

In `/Users/coleen/Downloads/Asuncion-Estacio/wedding-evite.html`, find:

```javascript
const RSVP_GOOGLE_WEB_APP_URL='...';
```

Replace it with your deployed Web App URL.

## 4. Test

1. Open the invitation page.
2. Submit RSVP with sample name and optional plus-one name.
3. Check the `RSVP` tab in your sheet for a new row.

## Notes

- Current form now sends these fields: `name`, `attending`, `guests`, `plus_one_name`, `message`, `event_date`, `event_time`, `event_timezone`, `submitted_at`, `page`.
- If the Web App URL is not set, the RSVP UI still works locally and does not post to Google Sheets.
