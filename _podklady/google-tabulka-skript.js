/** @OnlyCurrentDoc */

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
    // Ochrana proti robotům a nesmyslům: povinná pole, rozumné délky,
    // formulář vyplněný aspoň za 3 vteřiny, max. 50 kusů.
    const kusy = Number(p.kusy) || 0;
    if (!p.jmeno || !p.kontakt || !p.polozky) return ContentService.createTextOutput('chybí údaje');
    if (String(p.jmeno).length > 100 || String(p.kontakt).length > 150) return ContentService.createTextOutput('příliš dlouhé');
    if (kusy < 1 || kusy > 50) return ContentService.createTextOutput('špatný počet');
    if ((Number(p.cas) || 0) < 3000) return ContentService.createTextOutput('příliš rychlé');
    // Max. 200 rezervací za den, ať tabulku nikdo nezahltí.
    const cache = CacheService.getScriptCache();
    const den = 'pocet-' + Utilities.formatDate(new Date(), 'Europe/Prague', 'yyyy-MM-dd');
    const dnes = Number(cache.get(den)) || 0;
    if (dnes >= 200) return ContentService.createTextOutput('limit');
    cache.put(den, String(dnes + 1), 21600);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(SLOUPCE);
      sheet.getRange(1, 1, 1, SLOUPCE.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    const ciste = v => String(v || '').slice(0, 1000).replace(/^[=+\-@]/, "'$&");
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
