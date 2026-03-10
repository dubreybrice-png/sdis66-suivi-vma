/**
 * SDIS 66 — Suivi VMA
 * Point d'entrée principal — doGet, menu, fonctions serveur
 */

/* ═══════════════════════════════════════════════════════
   WEBAPP
   ═══════════════════════════════════════════════════════ */

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  template.cisParam = (e && e.parameter && e.parameter.cis) ? e.parameter.cis : '';
  template.baseUrl  = ScriptApp.getService().getUrl();

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
