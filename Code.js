/**
 * SDIS 66 — Suivi VMA
 * Point d'entrée principal — doGet, menu, fonctions serveur
 */

/* ═══════════════════════════════════════════════════════
   WEBAPP
   ═══════════════════════════════════════════════════════ */

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var followToken = params.followToken || '';
  var cisToken    = params.cisToken || '';

  var templateName = 'Index';
  if (followToken) templateName = 'Followup';
  else if (cisToken) templateName = 'CisView';

  var template = HtmlService.createTemplateFromFile(templateName);
  template.baseUrl = ScriptApp.getService().getUrl();
  template.followToken = followToken;
  template.cisToken = cisToken;

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
    .addItem('🔗 Afficher les URLs chefs de centre', 'showCisViewLinks')
    .addSeparator()
    .addItem('🧰 Initialiser Drive + programmes sport', 'initializeSportProgramResources')
    .addItem('📋 Installer triggers contrôle qualité', 'setupControleTriggers')
    .addItem('🌐 Ouvrir l\'application web', 'openWebApp')
    .addToUi();
}

function showCisViewLinks() {
  var links = getCisViewLinks();
  if (!links || links.length === 0) {
    SpreadsheetApp.getUi().alert('Aucun CIS trouvé', 'Lance d\'abord "Mettre à jour la liste des CIS".', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  var html = '<div style="font-family:monospace;font-size:12px;max-height:500px;overflow:auto;padding:10px;">';
  html += '<p style="margin-bottom:10px;font-family:sans-serif;"><strong>' + links.length + ' CIS</strong> — Copiez l\'URL et envoyez-la au chef de centre correspondant.</p>';
  for (var i = 0; i < links.length; i++) {
    var l = links[i];
    html += '<div style="margin-bottom:12px;padding:8px;background:#f5f5f5;border-radius:6px;border-left:3px solid #1D2951;">';
    html += '<strong>' + l.cis + '</strong>';
    if (l.email) html += ' <span style="color:#888;">('+l.email+')</span>';
    html += '<br><input type="text" value="'+l.url+'" style="width:100%;margin-top:4px;padding:4px;font-size:11px;border:1px solid #ccc;border-radius:3px;" onclick="this.select()" readonly>';
    html += '</div>';
  }
  html += '</div>';
  var output = HtmlService.createHtmlOutput(html).setWidth(650).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(output, '🔗 URLs Chefs de Centre');
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
