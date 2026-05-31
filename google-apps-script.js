// ══════════════════════════════════════════════════════════════
//  BRYLLUP RSVP — Google Apps Script
//  Lagrer svar fra bryllupsnettsiden direkte i dette arket.
//
//  OPPSETT (gjøres én gang):
//  1. Gå til https://script.google.com og opprett et nytt prosjekt
//  2. Slett alt i editoren og lim inn hele denne filen
//  3. Klikk "Kjør" → "doPost" én gang for å gi tillatelser
//  4. Klikk "Distribuer" → "Ny distribusjon"
//       - Type: Nettapp
//       - Kjør som: Meg
//       - Hvem har tilgang: Alle
//  5. Klikk "Distribuer" og kopier URL-en du får
//  6. Lim inn URL-en i bryllup.html der det står:
//       const APPS_SCRIPT_URL = 'DIN_APPS_SCRIPT_URL_HER';
// ══════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Legg til kolonneoverskrifter første gang
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Tidspunkt', 'Navn', 'E-post', 'Kommer', 'Allergier / matpreferanser', 'Melding']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    }

    const params = e.parameter;

    sheet.appendRow([
      new Date(),
      params.name     || '',
      params.email    || '',
      params.attending === 'ja' ? 'Ja' : 'Nei',
      params.dietary  || '',
      params.message  || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Trengs for at Google skal godta distribusjonen
function doGet() {
  return ContentService
    .createTextOutput('RSVP endpoint er aktivt.')
    .setMimeType(ContentService.MimeType.TEXT);
}
