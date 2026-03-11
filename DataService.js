/**
 * SDIS 66 — Suivi VMA
 * Service de données : lecture spreadsheet, logique métier, agrégation
 */

/* ═══════════════════════════════════════════════════════
   HELPERS PRIVÉS
   ═══════════════════════════════════════════════════════ */

function getSpreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

/**
 * Lit un onglet et retourne les lignes (sans l'en-tête).
 * Retourne [] si l'onglet n'existe pas ou est vide.
 */
function getSheetData_(sheetName) {
  var sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1);
}

/** Ajoute n mois à une date */
function addMonths_(date, months) {
  var d = new Date(date);
  d.setMonth(d.getMonth() + months);
  return d;
}

/** Formate une date en JJ/MM/AAAA */
function formatDate_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  var dd = ('0' + date.getDate()).slice(-2);
  var mm = ('0' + (date.getMonth() + 1)).slice(-2);
  return dd + '/' + mm + '/' + date.getFullYear();
}

/* ═══════════════════════════════════════════════════════
   SPÉCIALITÉS
   ═══════════════════════════════════════════════════════ */

/**
 * Construit un dictionnaire { "NOM Prénom": ["Bruleur", "Grimp", …] }
 */
function getSpecialties_() {
  var data = getSheetData_(CONFIG.SHEETS.SPECIALITE);
  var map = {};
  data.forEach(function (row) {
    var nom  = (row[CONFIG.COLS_SPE.NOM]  || '').toString().trim();
    var type = (row[CONFIG.COLS_SPE.TYPE] || '').toString().trim();
    if (nom && type) {
      if (!map[nom]) map[nom] = [];
      map[nom].push(type);
    }
  });
  return map;
}

/* ═══════════════════════════════════════════════════════
   SPORT
   ═══════════════════════════════════════════════════════ */

/**
 * Construit { matricule: [{date, testName, resultat}] }
 * Ne garde que les épreuves reconnues dans CONFIG.SPORT_TESTS
 */
function getSportData_() {
  var data = getSheetData_(CONFIG.SHEETS.SPORT);
  var map = {};

  data.forEach(function (row) {
    var matricule = (row[CONFIG.COLS_SPORT.MATRICULE] || '').toString().trim();
    var testName  = (row[CONFIG.COLS_SPORT.TEST_NAME] || '').toString().replace(/\s+/g, ' ').trim();

    if (!matricule || CONFIG.SPORT_TESTS.indexOf(testName) === -1) return;

    var dateVal  = row[CONFIG.COLS_SPORT.DATE];
    var resultat = (row[CONFIG.COLS_SPORT.RESULTAT] || '').toString().trim();

    if (!map[matricule]) map[matricule] = [];
    map[matricule].push({
      date:     formatDate_(dateVal),
      testName: testName,
      resultat: resultat
    });
  });

  return map;
}

/* ═══════════════════════════════════════════════════════
   EXAMENS COMPLÉMENTAIRES (CRUD)
   ═══════════════════════════════════════════════════════ */

/** Retourne (et crée si besoin) l'onglet Examens */
function getExamensSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.EXAMENS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.EXAMENS);
    sheet.getRange(1, 1, 1, 8).setValues([[
      'ID', 'Matricule', 'Type', 'Détail examen',
      'Date demande', 'Date résultat attendu', 'Commentaire', 'Statut'
    ]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Tous les examens ouverts */
function getAllExamens() {
  var sheet = getExamensSheet_();
  if (sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getValues().slice(1);
  var examens = [];

  data.forEach(function (row) {
    var statut = (row[CONFIG.COLS_EXAMENS.STATUT] || '').toString().trim();
    if (statut === 'cloture') return;

    var dateRes = row[CONFIG.COLS_EXAMENS.DATE_RESULTAT];
    examens.push({
      id:             (row[CONFIG.COLS_EXAMENS.ID] || '').toString(),
      matricule:      (row[CONFIG.COLS_EXAMENS.MATRICULE] || '').toString().trim(),
      type:           (row[CONFIG.COLS_EXAMENS.TYPE] || '').toString().trim(),
      detail:         (row[CONFIG.COLS_EXAMENS.DETAIL] || '').toString().trim(),
      dateDemande:    formatDate_(row[CONFIG.COLS_EXAMENS.DATE_DEMANDE]),
      dateResultat:   formatDate_(dateRes),
      dateResultatRaw:(dateRes instanceof Date && !isNaN(dateRes.getTime())) ? dateRes.getTime() : null,
      commentaire:    (row[CONFIG.COLS_EXAMENS.COMMENTAIRE] || '').toString().trim(),
      statut:         statut || 'ouvert'
    });
  });

  return examens;
}

/** Enregistre un nouvel examen et retourne l'objet complet */
function saveExamen(examenData) {
  var sheet = getExamensSheet_();
  var id = Utilities.getUuid();
  var dateDemande  = examenData.dateDemande  ? new Date(examenData.dateDemande)  : new Date();
  var dateResultat = examenData.dateResultat ? new Date(examenData.dateResultat) : null;

  sheet.appendRow([
    id,
    examenData.matricule || '',
    examenData.type || '',
    examenData.detail || '',
    dateDemande,
    dateResultat,
    examenData.commentaire || '',
    'ouvert'
  ]);

  return {
    id:             id,
    matricule:      (examenData.matricule || '').toString().trim(),
    type:           (examenData.type || '').toString().trim(),
    detail:         (examenData.detail || '').toString().trim(),
    dateDemande:    formatDate_(dateDemande),
    dateResultat:   formatDate_(dateResultat),
    dateResultatRaw:dateResultat ? dateResultat.getTime() : null,
    commentaire:    (examenData.commentaire || '').toString().trim(),
    statut:         'ouvert'
  };
}

/** Clôture un examen par son ID */
function closeExamen(examenId) {
  var sheet = getExamensSheet_();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === examenId) {
      sheet.getRange(i + 1, CONFIG.COLS_EXAMENS.STATUT + 1).setValue('cloture');
      return true;
    }
  }
  return false;
}

/* ═══════════════════════════════════════════════════════
   DÉTERMINATION DU TYPE DE VISITE
   ═══════════════════════════════════════════════════════ */

/**
 * Retourne le type de visite selon les règles métier :
 *
 * 1. Spécialité Bruleur / SAV / SAL / caisson          → VMA tous les ans
 * 2. Spécialité Grimp  ET âge ≥ 43                     → VMA tous les ans
 * 3. Spécialité diabétique                              → VMA tous les ans
 * ─── Hors spécialité (ou spécialité sans match VMA) ───
 * 4. ≥ 39 ans, né en année paire                       → Visite médicale biennale
 * 5. ≥ 39 ans, né en année impaire                     → Visite prévention
 * 6. < 39 ans, né en année paire                       → Visite médicale biennale
 * 7. < 39 ans, né en année impaire                     → Visite médicale 2027
 */
function determineVisitType_(agent, specialties) {
  var agentSpe = specialties[agent.nomPrenom];

  /* ── Règles spécialité ── */
  if (agentSpe && agentSpe.length > 0) {
    for (var i = 0; i < agentSpe.length; i++) {
      var spe = agentSpe[i];
      if (CONFIG.VMA_SPECIALTIES.indexOf(spe) !== -1) return 'VMA tous les ans';
      if (spe === 'Grimp' && agent.age >= CONFIG.VMA_GRIMP_AGE) return 'VMA tous les ans';
      if (spe.toLowerCase() === 'diabétique') return 'VMA tous les ans';
    }
    // Dans données spécialité mais aucune règle VMA matchée
    // (ex. Grimp < 43) → on tombe dans les règles d'âge ci-dessous
  }

  /* ── Règles d'âge (tous les agents non-VMA) ── */
  var birthYear = agent.birthYear;
  if (!birthYear) return 'Non déterminé';

  var isBirthEven = birthYear % 2 === 0;

  if (agent.age >= CONFIG.AGE_THRESHOLD) {
    // ≥ 39 ans
    return isBirthEven ? 'Visite médicale biennale' : 'Visite prévention';
  } else {
    // < 39 ans
    return isBirthEven ? 'Visite médicale biennale' : 'Visite médicale 2027';
  }
}

/* ═══════════════════════════════════════════════════════
   CHARGEMENT DES AGENTS
   ═══════════════════════════════════════════════════════ */

/**
 * Charge, dédoublonne et enrichit la liste complète des agents
 */
function getAllAgents() {
  var retardData  = getSheetData_(CONFIG.SHEETS.RETARD);
  var aVenirData  = getSheetData_(CONFIG.SHEETS.A_VENIR);
  var allData     = retardData.concat(aVenirData);
  var specialties = getSpecialties_();

  var agentsMap = {};

  allData.forEach(function (row) {
    var matricule = row[CONFIG.COLS.MATRICULE];
    if (!matricule || matricule.toString().trim() === '') return;

    var key = matricule.toString().trim();
    if (agentsMap[key]) return;

    var age              = parseInt(row[CONFIG.COLS.AGE]) || 0;
    var centreSecondaire = (row[CONFIG.COLS.CENTRE_SECONDAIRE] || '').toString().trim();
    var centrePrincipal  = (row[CONFIG.COLS.CENTRE_PRINCIPAL]  || '').toString().trim();
    var dateNaissance    = row[CONFIG.COLS.DATE_NAISSANCE];
    var dateVisite       = row[CONFIG.COLS.DATE_VISITE];
    var nomPrenom        = (row[CONFIG.COLS.NOM_PRENOM] || '').toString().trim();
    var objetVisite      = (row[CONFIG.COLS.OBJET_VISITE] || '').toString().trim();

    var datePerteCompetence = null;
    if (dateVisite instanceof Date && !isNaN(dateVisite.getTime())) {
      datePerteCompetence = addMonths_(dateVisite, CONFIG.MONTHS_TO_ADD);
    }

    var birthYear = (dateNaissance instanceof Date && !isNaN(dateNaissance.getTime()))
      ? dateNaissance.getFullYear() : null;
    var perteYear = datePerteCompetence ? datePerteCompetence.getFullYear() : null;
    var visitYear = (dateVisite instanceof Date && !isNaN(dateVisite.getTime()))
      ? dateVisite.getFullYear() : null;

    var agent = {
      age:                    age,
      centreSecondaire:       centreSecondaire,
      centrePrincipal:        centrePrincipal,
      dateNaissance:          formatDate_(dateNaissance),
      datePerteCompetence:    formatDate_(datePerteCompetence),
      datePerteCompetenceRaw: datePerteCompetence ? datePerteCompetence.getTime() : null,
      nomPrenom:              nomPrenom,
      matricule:              key,
      objetVisite:            objetVisite,
      birthYear:              birthYear,
      perteYear:              perteYear,
      visitYear:              visitYear,
      typeVisite:             '',
      sport:                  []
    };

    agent.typeVisite = determineVisitType_(agent, specialties);
    agentsMap[key]   = agent;
  });

  // Tri chronologique par date de perte de compétence
  var agents = Object.keys(agentsMap).map(function (k) { return agentsMap[k]; });
  agents.sort(function (a, b) {
    if (!a.datePerteCompetenceRaw && !b.datePerteCompetenceRaw) return 0;
    if (!a.datePerteCompetenceRaw) return 1;
    if (!b.datePerteCompetenceRaw) return -1;
    return a.datePerteCompetenceRaw - b.datePerteCompetenceRaw;
  });

  return agents;
}

/* ═══════════════════════════════════════════════════════
   FONCTIONS PUBLIQUES
   ═══════════════════════════════════════════════════════ */

/** Liste triée de tous les CIS */
function getCisList() {
  var allAgents = getAllAgents();
  var cisSet = {};
  allAgents.forEach(function (a) {
    if (a.centrePrincipal)  cisSet[a.centrePrincipal]  = true;
    if (a.centreSecondaire) cisSet[a.centreSecondaire] = true;
  });
  return Object.keys(cisSet).sort();
}

/** Agents d'un CIS donné (principal OU secondaire) */
function getAgentsByCis(cisName) {
  return getAllAgents().filter(function (a) {
    return a.centrePrincipal === cisName || a.centreSecondaire === cisName;
  });
}

/**
 * Données complètes pour la page web :
 *  - agents avec sport
 *  - examens ouverts
 *  - totalAgents
 */
function getPageData() {
  var agents    = getAllAgents();
  var sportData = getSportData_();
  var examens   = getAllExamens();

  // Rattacher les données sport aux agents
  agents.forEach(function (a) {
    a.sport = sportData[a.matricule] || [];
  });

  return {
    agents:      agents,
    examens:     examens,
    totalAgents: agents.length
  };
}

/** Remplit la colonne A de l'onglet "cis / mailing" avec tous les CIS */
function populateCisMailingSheet() {
  var ss    = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CIS_MAILING);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.CIS_MAILING);

  var cisList = getCisList();
  sheet.getRange(1, 1).setValue('CIS');

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 1).clearContent();

  if (cisList.length > 0) {
    var values = cisList.map(function (c) { return [c]; });
    sheet.getRange(2, 1, values.length, 1).setValues(values);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    cisList.length + ' CIS ajoutés dans l\'onglet "cis / mailing"',
    'Mise à jour terminée'
  );
  return cisList;
}
