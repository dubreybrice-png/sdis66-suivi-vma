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
    .addItem('🏛️ Afficher CIS Divers', 'openCisDivers')
    .addItem('📬 Installer envoi hebdo CIS groupés (Brice)', 'setupWeeklyGroupedCisEmailTrigger')
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
  var html = '<div style="font-family:Arial,sans-serif;font-size:12px;max-height:540px;overflow:auto;padding:12px;">';
  html += '<div style="display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:12px;">';
  html += '<p style="margin:0;"><strong>' + links.length + ' CIS</strong> — liens prêts à être envoyés.</p>';
  html += '<button id="sendAllBtn" onclick="sendAllCisEmails()" style="background:#1D2951;color:#fff;border:none;border-radius:6px;padding:8px 12px;cursor:pointer;font-weight:600;">Envoyer à tous</button>';
  html += '</div>';
  html += '<div id="status" style="display:none;margin-bottom:10px;padding:8px 10px;border-radius:6px;background:#eef4ff;color:#1D2951;"></div>';
  for (var i = 0; i < links.length; i++) {
    var l = links[i];
    var cisEncoded = encodeURIComponent(l.cis || '');
    html += '<div style="margin-bottom:12px;padding:10px;background:#f5f5f5;border-radius:6px;border-left:3px solid #1D2951;">';
    html += '<div style="display:flex;justify-content:space-between;align-items:center;gap:10px;">';
    html += '<div><strong>' + l.cis + '</strong>';
    if (l.email) html += ' <span style="color:#888;">(' + l.email + ')</span>';
    html += '</div>';
    html += '<button data-cis="' + cisEncoded + '" onclick="sendCisEmail(this)" ' + (l.email ? '' : 'disabled ') + 'style="background:' + (l.email ? '#2e7d32' : '#bdbdbd') + ';color:#fff;border:none;border-radius:6px;padding:7px 10px;cursor:' + (l.email ? 'pointer' : 'not-allowed') + ';font-weight:600;">Envoyer</button>';
    html += '</div>';
    html += '<br><input type="text" value="'+l.url+'" style="width:100%;margin-top:4px;padding:4px;font-size:11px;border:1px solid #ccc;border-radius:3px;" onclick="this.select()" readonly>';
    html += '</div>';
  }
  html += '<script>'
    + 'function setStatus(message,isError){'
    + 'var box=document.getElementById("status");'
    + 'box.style.display="block";'
    + 'box.style.background=isError?"#fdecea":"#eef4ff";'
    + 'box.style.color=isError?"#b3261e":"#1D2951";'
    + 'box.textContent=message;'
    + '}'
    + 'function sendCisEmail(button){'
    + 'var cisName=decodeURIComponent(button.getAttribute("data-cis")||"");'
    + 'if(!cisName){setStatus("Erreur : CIS manquant",true);return;}'
    + 'button.disabled=true;'
    + 'var label=button.textContent;'
    + 'button.textContent="Envoi...";'
    + 'google.script.run.withSuccessHandler(function(result){'
      + 'button.disabled=false;button.textContent=label;'
      + 'setStatus("Lien envoyé pour " + result.cis + " à " + result.count + " destinataire(s).",false);'
    + '}).withFailureHandler(function(err){'
      + 'button.disabled=false;button.textContent=label;'
      + 'setStatus("Erreur : " + (err.message || err),true);'
    + '}).sendCisViewLinkEmail(cisName);'
    + '}'
    + 'function sendAllCisEmails(){'
    + 'var button=document.getElementById("sendAllBtn");'
    + 'button.disabled=true;'
    + 'var label=button.textContent;'
    + 'button.textContent="Envoi...";'
    + 'google.script.run.withSuccessHandler(function(result){'
      + 'button.disabled=false;button.textContent=label;'
      + 'var msg="Envoi terminé : " + result.sentCount + " CIS envoyé(s)";'
      + 'if(result.skippedCount){msg += ", " + result.skippedCount + " ignoré(s) sans email";}'
      + 'setStatus(msg + ".",false);'
    + '}).withFailureHandler(function(err){'
      + 'button.disabled=false;button.textContent=label;'
      + 'setStatus("Erreur : " + (err.message || err),true);'
    + '}).sendAllCisViewLinkEmails();'
    + '}'
  + '</script>';
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

function openCisDivers() {
  var url = getWeeklyGroupedCisLink_();
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("' + url + '");google.script.host.close();</script>'
  );
  SpreadsheetApp.getUi().showModalDialog(html, 'Ouverture CIS Divers…');
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
