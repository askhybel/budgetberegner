/**
 * Tilmeldinger – Digital budgetformidling (online seminar 1. december 2027)
 *
 * Opsætning (ca. 3 min):
 * 1. Opret et nyt Google Sheet, fx "Tilmeldinger – Digital budgetformidling".
 * 2. Udvidelser → Apps Script. Slet indholdet og indsæt hele denne fil. Gem.
 * 3. Implementer → Ny implementering → Type: Webapp
 *      Udfør som: Mig
 *      Hvem har adgang: Alle
 *    → Implementer → godkend adgang → kopier webapp-URL'en (slutter på /exec).
 * 4. Indsæt URL'en i index.html i konstanten SHEET_ENDPOINT.
 *
 * Ændrer du scriptet senere: Implementer → Administrer implementeringer → rediger → Ny version
 * (så beholder URL'en sig den samme).
 */

const SHEET_NAME = 'Tilmeldinger';
const HEADERS = ['Tidspunkt', 'Navn', 'Titel', 'Virksomhed', 'Email', 'Kilde'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};

    // Honeypot: bots udfylder det skjulte felt "website"
    if (p.website) return json_({ ok: true });

    const email = String(p.email || '').trim();
    if (!p.navn || !p.titel || !p.virksomhed || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return json_({ ok: false, error: 'Manglende eller ugyldige felter' });
    }

    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      clean_(p.navn),
      clean_(p.titel),
      clean_(p.virksomhed),
      clean_(email),
      clean_(p.kilde)
    ]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, info: 'Tilmeldings-endpoint kører' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Forhindrer formel-injektion i arket (værdier der starter med = + - @)
function clean_(v) {
  const s = String(v == null ? '' : v).trim().slice(0, 300);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
