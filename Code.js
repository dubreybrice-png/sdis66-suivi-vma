/**
 * SDIS 66 — Suivi VMA
 * Point d'entrée principal — doGet, menu, fonctions serveur
 */

/* ═══════════════════════════════════════════════════════
   WEBAPP
   ═══════════════════════════════════════════════════════ */

function doGet(e) {
  var token = (e && e.parameter && e.parameter.followToken) ? e.parameter.followToken : '';
  var template = HtmlService.createTemplateFromFile(token ? 'Followup' : 'Index');
  template.baseUrl = ScriptApp.getService().getUrl();
  template.followToken = token;

  return template.evaluate()
    .setTitle('SDIS 66 — Suivi VMA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ═══════════════════════════════════════════════════════
   MENU GOOGLE SHEETS
   ═══════════════════════════════════════════════════════ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏥 Suivi VMA')
    .addItem('📋 Mettre à jour la liste des CIS', 'populateCisMailingSheet')
    .addSeparator()
    .addItem('🧰 Initialiser Drive + programmes sport', 'initializeSportProgramResources')
    .addItem('🌐 Ouvrir l\'application web', 'openWebApp')
    .addToUi();
}

function openWebApp() {
  var url = ScriptApp.getService().getUrl();
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("' + url + '");google.script.host.close();</script>'
  );
  SpreadsheetApp.getUi().showModalDialog(html, 'Ouverture de l\'application…');
}

function initializeSportProgramResources() {
  var result = getSportProgramCatalog();
  var count = Object.keys(result || {}).length;
  SpreadsheetApp.getUi().alert(
    'Initialisation terminée',
    count + ' programme(s) PDF disponible(s) dans Drive.\n\nLa webapp peut maintenant afficher les aperçus PDF.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
