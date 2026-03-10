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
 * Lit un onglet et retourne les lignes (sans l'en-tête)
 */
function getSheetData_(sheetName) {
  var sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1);
}

/**
 * Ajoute n mois à une date
 */
function addMonths_(date, months) {
  var d = new Date(date);
  d.setMonth(d.getMonth() + months);
  return d;
}

/**
 * Formate une date en JJ/MM/AAAA
 */
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
   DÉTERMINATION DU TYPE DE VISITE
   ═══════════════════════════════════════════════════════ */

/**
 * Retourne le type de visite selon les règles métier :
 *
 * 1. Spécialité Bruleur / SAV / SAL / caisson          → VMA tous les ans
 * 2. Spécialité Grimp  ET âge ≥ 43                     → VMA tous les ans
 * 3. Spécialité diabétique                              → VMA tous les ans
 * 4. (Hors spécialité) > 39 ans, né pair, perte pair   → Visite médicale biennale
 * 5. (Hors spécialité) > 39 ans, né impair, perte pair → Visite prévention
 * 6. (Hors spécialité) < 39 ans, visite en 2026        → Visite médicale biennale
 */
function determineVisitType_(agent, specialties) {
  var agentSpe = specialties[agent.nomPrenom];

  /* ── Règles spécialité ── */
  if (agentSpe && agentSpe.length > 0) {
    for (var i = 0; i < agentSpe.length; i++) {
      var spe = agentSpe[i];

      // Bruleur, SAV, SAL, caisson → VMA tous les ans
      if (CONFIG.VMA_SPECIALTIES.indexOf(spe) !== -1) {
        return 'VMA tous les ans';
      }
      // Grimp + âge ≥ 43 → VMA tous les ans
      if (spe === 'Grimp' && agent.age >= CONFIG.VMA_GRIMP_AGE) {
        return 'VMA tous les ans';
      }
      // diabétique → VMA tous les ans
      if (spe.toLowerCase() === 'diabétique') {
        return 'VMA tous les ans';
      }
    }
    // Dans données spécialité mais aucune règle VMA matchée
    // (ex. Grimp < 43 ans, autre spécialité non listée)
    // → on ne renvoie rien ; on ne tombe PAS dans les règles d'âge
    return '';
  }

  /* ── Règles d'âge (hors données spécialité) ── */
  var birthYear = agent.birthYear;
  var perteYear = agent.perteYear;
  if (!birthYear || !perteYear) return '';

  var isBirthEven = birthYear % 2 === 0;
  var isPerteEven = perteYear % 2 === 0;

  if (agent.age > CONFIG.AGE_THRESHOLD) {
    if (isBirthEven && isPerteEven) {
      return 'Visite médicale biennale';
    }
    if (!isBirthEven && isPerteEven) {
      return 'Visite prévention';
    }
  }

  if (agent.age < CONFIG.AGE_THRESHOLD && agent.visitYear === CONFIG.REFERENCE_YEAR) {
    return 'Visite médicale biennale';
  }

  return '';
}

/* ═══════════════════════════════════════════════════════
   CHARGEMENT DES AGENTS
   ═══════════════════════════════════════════════════════ */

/**
 * Charge, dédoublonne et enrichit la liste complète des agents
 * (Copie retard puis Copie a venir, premier vu par matricule gardé)
 */
function getAllAgents() {
  var retardData = getSheetData_(CONFIG.SHEETS.RETARD);
  var aVenirData = getSheetData_(CONFIG.SHEETS.A_VENIR);
  var allData    = retardData.concat(aVenirData);
  var specialties = getSpecialties_();

  var agentsMap = {};

  allData.forEach(function (row) {
    var matricule = row[CONFIG.COLS.MATRICULE];
    if (!matricule || matricule.toString().trim() === '') return;

    var key = matricule.toString().trim();
    if (agentsMap[key]) return; // dédoublonnage par matricule

    var age              = parseInt(row[CONFIG.COLS.AGE]) || 0;
    var centreSecondaire = (row[CONFIG.COLS.CENTRE_SECONDAIRE] || '').toString().trim();
    var centrePrincipal  = (row[CONFIG.COLS.CENTRE_PRINCIPAL]  || '').toString().trim();
    var dateNaissance    = row[CONFIG.COLS.DATE_NAISSANCE];
    var dateVisite       = row[CONFIG.COLS.DATE_VISITE];
    var nomPrenom        = (row[CONFIG.COLS.NOM_PRENOM] || '').toString().trim();
    var objetVisite      = (row[CONFIG.COLS.OBJET_VISITE] || '').toString().trim();

    // Date de perte de compétence = date visite + 3 mois
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
      age:                      age,
      centreSecondaire:         centreSecondaire,
      centrePrincipal:          centrePrincipal,
      dateNaissance:            formatDate_(dateNaissance),
      datePerteCompetence:      formatDate_(datePerteCompetence),
      datePerteCompetenceRaw:   datePerteCompetence ? datePerteCompetence.getTime() : null,
      nomPrenom:                nomPrenom,
      matricule:                key,
      objetVisite:              objetVisite,
      birthYear:                birthYear,
      perteYear:                perteYear,
      visitYear:                visitYear,
      typeVisite:               ''
    };

    agent.typeVisite = determineVisitType_(agent, specialties);
    agentsMap[key] = agent;
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
   FONCTIONS PUBLIQUES (appelées par le client / menu)
   ═══════════════════════════════════════════════════════ */

/**
 * Liste triée de tous les CIS présents dans les données
 */
function getCisList() {
  var allAgents = getAllAgents();
  var cisSet = {};
  allAgents.forEach(function (a) {
    if (a.centrePrincipal)  cisSet[a.centrePrincipal]  = true;
    if (a.centreSecondaire) cisSet[a.centreSecondaire] = true;
  });
  return Object.keys(cisSet).sort();
}

/**
 * Agents d'un CIS donné (principal OU secondaire)
 */
function getAgentsByCis(cisName) {
  return getAllAgents().filter(function (a) {
    return a.centrePrincipal === cisName || a.centreSecondaire === cisName;
  });
}

/**
 * Données complètes pour la page web (home ou détail CIS)
 */
function getPageData(cisName) {
  var cisList = getCisList();

  /* ── Page d'accueil ── */
  if (!cisName || cisName === '' || cisName === 'home') {
    var allAgents = getAllAgents();
    var cisCounts = {};
    cisList.forEach(function (c) { cisCounts[c] = 0; });

    allAgents.forEach(function (a) {
      if (a.centrePrincipal  && cisCounts[a.centrePrincipal]  !== undefined) cisCounts[a.centrePrincipal]++;
      if (a.centreSecondaire && cisCounts[a.centreSecondaire] !== undefined) cisCounts[a.centreSecondaire]++;
    });

    return {
      agents:      [],
      cisName:     '',
      cisList:     cisList,
      cisCounts:   cisCounts,
      totalAgents: allAgents.length,
      page:        'home'
    };
  }

  /* ── Page détail ── */
  var agents, title;
  if (cisName === 'all') {
    agents = getAllAgents();
    title  = 'Vue complète';
  } else {
    agents = getAgentsByCis(cisName);
    title  = cisName;
  }

  return {
    agents:  agents,
    cisName: title,
    cisList: cisList,
    page:    'detail'
  };
}

/**
 * Remplit la colonne A de l'onglet "cis / mailing" avec tous les CIS
 */
function populateCisMailingSheet() {
  var ss    = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CIS_MAILING);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.CIS_MAILING);
  }

  var cisList = getCisList();

  // En-tête
  sheet.getRange(1, 1).setValue('CIS');

  // Nettoyer colonne A (garder les autres colonnes intactes)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
  }

  // Écrire la liste
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
