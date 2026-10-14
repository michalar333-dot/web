/**
 * mám capacitu – ukládání rezervací z webu do Google tabulky
 *
 * Návod: viz NAVOD.txt (krok 1). Celý tento kód vlož do
 * Rozšíření → Apps Script ve své tabulce „Rezervace mám capacitu“.
 */

// Chceš e-mail při každé nové rezervaci? Nech true. Nechceš? Přepiš na false.
const POSILAT_EMAIL = true;

const SLOUPCE = ['Čas', 'Jméno', 'Kontakt', 'Položky', 'Kusů', 'Cena (Kč)', 'Poznámka', 'Stav'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(SLOUPCE);
      sheet.getRange(1, 1, 1, SLOUPCE.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    const ciste = v => String(v || '').slice(0, 2000).replace(/^[=+\-@]/, "'$&");
    sheet.appendRow([
      new Date(),
      ciste(p.jmeno),
      ciste(p.kontakt),
      ciste(p.polozky),
      Number(p.kusy) || 0,
      Number(p.cena) || 0,
      ciste(p.poznamka),
      'nová'
    ]);
    if (POSILAT_EMAIL) {
      MailApp.sendEmail(
        Session.getEffectiveUser().getEmail(),
        'Nová rezervace: ' + (p.jmeno || '') + ' (' + (p.kusy || 0) + ' ks)',
        'Jméno: ' + (p.jmeno || '') + '\nKontakt: ' + (p.kontakt || '') +
        '\n\n' + (p.polozky || '') + '\n\nCelkem: ' + (p.kusy || 0) + ' ks / ' + (p.cena || 0) + ' Kč' +
        (p.poznamka ? '\nPoznámka: ' + p.poznamka : '') +
        '\n\nTabulka: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
      );
    }
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// Kontrola v prohlížeči: když otevřeš adresu skriptu, uvidíš „běží“.
function doGet() {
  return ContentService.createTextOutput('Rezervace mám capacitu: běží');
}
