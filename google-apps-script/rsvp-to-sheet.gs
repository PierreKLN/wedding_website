/**
 * Script à coller dans l'éditeur Apps Script d'un Google Sheet
 * (Extensions > Apps Script) pour que chaque soumission du formulaire
 * RSVP du site ajoute automatiquement une ligne dans la feuille.
 *
 * Installation :
 * 1. Crée un Google Sheet (ou ouvre celui qui doit recevoir les réponses).
 * 2. Extensions > Apps Script, colle ce code, enregistre.
 * 3. Déployer > Nouveau déploiement > type "Application Web".
 *    - Exécuter en tant que : Moi
 *    - Qui a accès : Tout le monde
 * 4. Copie l'URL "/exec" obtenue et colle-la dans SHEET_WEBHOOK_URL
 *    au début du script du site (index.html).
 * 5. À la première exécution, Google demandera d'autoriser le script
 *    (accès à ce Sheet) : c'est normal, à accepter une seule fois.
 */

const SHEET_NAME = 'Réponses RSVP';

function doPost(e) {
  const sheet = getOrCreateSheet();
  const p = e.parameter;

  sheet.appendRow([
    new Date(),
    p.name || '',
    p.address || '',
    p.presence || '',
    p.diet || '',
    p.car || '',
    p.sleepover || '',
    p.brunch || '',
    p.song || '',
    p.message || ''
  ]);

  return ContentService.createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'Horodatage', 'Nom et prénom', 'Adresse', 'Présent·e ?',
      'Régime particulier', 'Voiture ?', 'Dort sur place ?',
      'Présent·e au brunch ?', 'Chanson', 'Message'
    ]);
  }
  return sheet;
}
