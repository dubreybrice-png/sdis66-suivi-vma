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

function formatIsoDate_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  var dd = ('0' + date.getDate()).slice(-2);
  var mm = ('0' + (date.getMonth() + 1)).slice(-2);
  return date.getFullYear() + '-' + mm + '-' + dd;
}

function parseAnyDateToIso_(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) return formatIsoDate_(value);
  var s = value.toString().trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return m[3] + '-' + m[2] + '-' + m[1];
  var d = new Date(s);
  if (!isNaN(d.getTime())) return formatIsoDate_(d);
  return '';
}

function normalizeSessions_(sessions) {
  if (!Array.isArray(sessions)) return [];
  var out = [];
  for (var i = 0; i < sessions.length; i++) {
    var iso = parseAnyDateToIso_(sessions[i].dateIso || sessions[i].date);
    if (!iso) continue;
    out.push({
      dateIso: iso,
      date: formatDate_(new Date(iso + 'T12:00:00')),
      commentaire: (sessions[i].commentaire || '').toString().trim(),
      createdAt: sessions[i].createdAt || new Date().toISOString()
    });
  }
  out.sort(function (a, b) { return a.dateIso.localeCompare(b.dateIso); });
  return out;
}

/** Retourne/crée le dossier Drive des documents de suivi sport */
function getSportDocsFolder_() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('SPORT_DOCS_FOLDER_ID');
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      // dossier supprimé/inaccessible: on recrée plus bas
    }
  }

  var name = 'SDIS66 - Suivi VMA - Documents Sport';
  var it = DriveApp.getFoldersByName(name);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(name);
  props.setProperty('SPORT_DOCS_FOLDER_ID', folder.getId());
  return folder;
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

/**
 * Charge l'onglet "derniere visite" et retourne un dictionnaire
 * { matricule (string) → Date de dernière visite effectuée }
 */
function getLastVisitDates_() {
  var data = getSheetData_(CONFIG.SHEETS.DERNIERE_VISITE);
  var map = {};
  data.forEach(function (row) {
    var matricule = (row[CONFIG.COLS_DERNIERE_VISITE.MATRICULE] || '').toString().trim();
    var dateVal   = row[CONFIG.COLS_DERNIERE_VISITE.DATE];
    if (!matricule) return;
    if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
      map[matricule] = dateVal;
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
    var rawRes    = row[CONFIG.COLS_SPORT.RESULTAT];
    var resultat  = (rawRes !== undefined && rawRes !== null && rawRes !== '') ? rawRes.toString().trim() : '';

    // Ignorer si pas de matricule, test non reconnu, ou pas de résultat
    if (!matricule || CONFIG.SPORT_TESTS.indexOf(testName) === -1 || resultat === '') return;

    var dateVal = row[CONFIG.COLS_SPORT.DATE];
    var dateRaw = (dateVal instanceof Date && !isNaN(dateVal.getTime())) ? dateVal.getTime() : null;

    if (!map[matricule]) map[matricule] = [];
    map[matricule].push({
      date:     formatDate_(dateVal),
      dateRaw:  dateRaw,
      testName: testName,
      resultat: resultat
    });
  });

  // Trier par date décroissante (plus récent en premier)
  Object.keys(map).forEach(function(key) {
    map[key].sort(function(a, b) { return (b.dateRaw || 0) - (a.dateRaw || 0); });
  });

  return map;
}

/* ═══════════════════════════════════════════════════════
   SPORT META (SUIVI + EAP + DOCUMENTS)
   ═══════════════════════════════════════════════════════ */

function getProgramDefinitions_() {
  return {
    perte_poids: {
      label: 'Perte de poids',
      content:
        'Programme de perte de poids (progressif, sécurisé)\n\n' +
        'Objectif général\n' +
        '- Réduction progressive de la masse grasse\n' +
        '- Préservation de la masse musculaire\n' +
        '- Amélioration du souffle et de la tolérance à l\'effort\n\n' +
        'Semaine type (4 séances)\n' +
        '1) Cardio zone modérée (45 min)\n' +
        '2) Renforcement global (45 min)\n' +
        '3) Cardio fractionné doux (30 min)\n' +
        '4) Mobilité + gainage (30 min)\n\n' +
        'Repères nutrition\n' +
        '- Assiette équilibrée 80% du temps\n' +
        '- Hydratation 2L/jour\n' +
        '- Protéines à chaque repas\n\n' +
        'Précautions\n' +
        '- Progressivité des charges\n' +
        '- Arrêt si douleur articulaire anormale\n' +
        '- Point mensuel EAP'
    },
    reathletisation: {
      label: 'Réathlétisation',
      content:
        'Programme de réathlétisation\n\n' +
        'Phase 1 (2 semaines)\n' +
        '- Mobilité active et contrôle moteur\n' +
        '- Endurance fondamentale 20-30 min\n\n' +
        'Phase 2 (3 semaines)\n' +
        '- Renforcement fonctionnel\n' +
        '- Travail unilatéral et gainage\n\n' +
        'Phase 3 (3 semaines)\n' +
        '- Reprise d\'intensité\n' +
        '- Intervalles courts\n' +
        '- Ateliers opérationnels\n\n' +
        'Critères de passage\n' +
        '- Douleur < 2/10\n' +
        '- Symétrie motrice satisfaisante\n' +
        '- Validation EAP'
    },
    reprise_genou: {
      label: 'Reprise post blessure genou',
      content:
        'Programme de reprise post blessure genou\n\n' +
        '1) Mobilité\n' +
        '- Flexion/extension contrôlée\n' +
        '- Mobilité hanche-cheville\n\n' +
        '2) Renforcement\n' +
        '- Chaîne antérieure/postérieure\n' +
        '- Squat partiel, fente statique\n' +
        '- Ischio-jambiers et fessiers\n\n' +
        '3) Proprioception\n' +
        '- Appuis unipodaux\n' +
        '- Variations visuelles et surfaces\n\n' +
        '4) Retour terrain\n' +
        '- Course progressive\n' +
        '- Changements de direction\n' +
        '- Port de charge progressif'
    },
    reprise_accouchement: {
      label: 'Reprise post accouchement',
      content:
        'Programme de reprise post accouchement\n\n' +
        'Pré-requis\n' +
        '- Feu vert médical\n' +
        '- Travail respiratoire et périnéal intégré\n\n' +
        'Étape 1 (2-4 semaines)\n' +
        '- Marche active\n' +
        '- Renforcement doux du tronc\n\n' +
        'Étape 2 (4-6 semaines)\n' +
        '- Renforcement global progressif\n' +
        '- Endurance modérée\n\n' +
        'Étape 3\n' +
        '- Reprise des impacts progressifs\n' +
        '- Ateliers métiers adaptés\n\n' +
        'Surveillance\n' +
        '- Fatigue\n' +
        '- Douleurs pelviennes/lombaires\n' +
        '- Qualité de récupération'
    },
    prevention_epaule: {
      label: 'Prévention blessure épaule',
      content:
        'Programme de prévention des blessures de l\'épaule\n\n' +
        'Échauffement\n' +
        '- Mobilité scapulo-thoracique\n' +
        '- Rotations contrôlées\n\n' +
        'Renforcement\n' +
        '- Coiffe des rotateurs avec élastique\n' +
        '- Stabilisateurs omoplates\n' +
        '- Gainage anti-rotation\n\n' +
        'Fonctionnel pompier\n' +
        '- Tirage/port de charge progressif\n' +
        '- Travail au-dessus de la tête\n\n' +
        'Fréquence\n' +
        '- 2 à 3 séances / semaine\n' +
        '- 20 à 30 minutes'
    },
    prevention_cheville: {
      label: 'Prévention blessure cheville',
      content:
        'Programme de prévention des blessures de cheville\n' +
        '(suite à entorse – reprise progressive)\n\n' +
        '1. Mobilité de la cheville (échauffement)\n' +
        'Objectif : récupérer l\'amplitude articulaire et lubrifier l\'articulation.\n\n' +
        '- Cercles de cheville : 10 cercles dans chaque sens, 2 séries\n' +
        '- Alphabet avec le pied : 1 alphabet complet, 1 à 2 séries\n' +
        '- Étirement mollet / tendon d\'Achille au mur : 20 secondes, 3 répétitions\n\n' +
        '2. Renforcement musculaire\n' +
        'Objectif : renforcer les muscles stabilisateurs de la cheville.\n\n' +
        '- Éversion / inversion avec élastique : 12 répétitions, 3 séries\n' +
        '- Flexion plantaire (montées sur pointe) : 15 répétitions, 3 séries\n' +
        '- Progression : sur un seul pied\n\n' +
        '3. Proprioception (clé pour les pompiers)\n' +
        'Objectif : éviter les récidives sur terrains instables.\n\n' +
        '- Équilibre sur un pied : 30 secondes, 3 répétitions\n' +
        '- Progression : yeux fermés puis surface instable\n' +
        '- Variante pompier : rotation du tronc / lancer de balle\n\n' +
        '4. Reprise fonctionnelle\n' +
        '- Petits sauts : 10 sauts, 3 séries\n' +
        '- Sauts latéraux : 10 répétitions, 3 séries\n' +
        '- Course légère : 10 à 15 minutes\n\n' +
        'Conseils spécifiques SP\n' +
        '- Échauffement systématique\n' +
        '- Proprioception au moins 2 mois\n' +
        '- Reprise progressive terrains instables\n' +
        '- Chevillière si instabilité ressentie\n\n' +
        'Risque de récidive majoré lors de la course en terrain irrégulier, du port de charge et de la fatigue en intervention.'
    }
  };
}

function buildDemoSessionsForProgram_(programKey) {
  var now = new Date();
  var base = now.getTime();
  var comments = {
    perte_poids: ['Cardio modéré 40 min', 'Renforcement complet OK', 'Bonne récupération et hydratation'],
    reathletisation: ['Travail mobilité + gainage', 'Atelier unilatéral validé', 'Progression sans douleur'],
    reprise_genou: ['Proprioception unipodale', 'Course légère 12 min', 'Aucun gonflement post séance'],
    reprise_accouchement: ['Marche active + respiration', 'Renforcement tronc doux', 'Fatigue correcte, RAS'],
    prevention_epaule: ['Coiffe des rotateurs élastique', 'Stabilité scapulaire', 'Mobilité overhead améliorée'],
    prevention_cheville: ['Équilibre sur coussin', 'Sauts latéraux contrôlés', 'Cheville stable en fin de séance']
  };
  var c = comments[programKey] || ['Séance réalisée'];
  return normalizeSessions_([
    { dateIso: formatIsoDate_(new Date(base - 28 * 24 * 3600000)), commentaire: c[0] || 'Séance 1' },
    { dateIso: formatIsoDate_(new Date(base - 11 * 24 * 3600000)), commentaire: c[1] || 'Séance 2' },
    { dateIso: formatIsoDate_(new Date(base - 3 * 24 * 3600000)), commentaire: c[2] || 'Séance 3' }
  ]);
}

function generateFollowToken_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 20);
}

function makeFollowLink_(token) {
  var base = ScriptApp.getService().getUrl();
  return base + '?followToken=' + encodeURIComponent(token);
}

function getProgramFolder_() {
  var root = getSportDocsFolder_();
  var it = root.getFoldersByName('Programmes Démo');
  return it.hasNext() ? it.next() : root.createFolder('Programmes Démo');
}

function ensureProgramCatalog_() {
  var defs = getProgramDefinitions_();
  var folder = getProgramFolder_();
  var props = PropertiesService.getScriptProperties();
  var map;

  try {
    map = JSON.parse(props.getProperty('SPORT_PROGRAM_PDF_MAP') || '{}');
  } catch (e) {
    map = {};
  }

  Object.keys(defs).forEach(function (key) {
    var existingId = map[key];
    if (existingId) {
      try {
        DriveApp.getFileById(existingId);
        return;
      } catch (e) {
        // recréation
      }
    }

    var def = defs[key];
    var doc = DocumentApp.create('Programme - ' + def.label);
    doc.getBody().setText(def.content);
    doc.saveAndClose();

    var docFile = DriveApp.getFileById(doc.getId());
    var pdfBlob = docFile.getAs(MimeType.PDF).setName('Programme - ' + def.label + '.pdf');
    var pdfFile = folder.createFile(pdfBlob);
    map[key] = pdfFile.getId();

    // On nettoie le Google Doc temporaire
    docFile.setTrashed(true);
  });

  props.setProperty('SPORT_PROGRAM_PDF_MAP', JSON.stringify(map));

  var out = {};
  Object.keys(defs).forEach(function (key) {
    var f = DriveApp.getFileById(map[key]);
    out[key] = {
      key: key,
      label: defs[key].label,
      fileId: f.getId(),
      url: f.getUrl(),
      previewUrl: 'https://drive.google.com/file/d/' + f.getId() + '/preview'
    };
  });
  return out;
}

/** Retourne (et crée si besoin) l'onglet de métadonnées sport */
function getSportMetaSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SPORT_META);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.SPORT_META);
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Matricule',
      'Suivi en place',
      'Nom EAP',
      'Programme',
      'Séances JSON',
      'Documents JSON',
      'Dernière mise à jour',
      'Token suivi'
    ]]);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < 8) {
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Matricule',
      'Suivi en place',
      'Nom EAP',
      'Programme',
      'Séances JSON',
      'Documents JSON',
      'Dernière mise à jour',
      'Token suivi'
    ]]);
  }
  return sheet;
}

/** Retourne { matricule: { suiviEnPlace, nomEap, programmeKey, sessions[], documents[] } } */
function getSportMetaMap_() {
  var sheet = getSportMetaSheet_();
  if (sheet.getLastRow() < 2) return {};

  var data = sheet.getDataRange().getValues().slice(1);
  var map = {};

  data.forEach(function (row) {
    var matricule = (row[0] || '').toString().trim();
    if (!matricule) return;

    var suiviRaw = (row[1] || '').toString().trim().toLowerCase();
    var suiviEnPlace = (suiviRaw === 'oui' || suiviRaw === 'true' || suiviRaw === '1');
    var nomEap = (row[2] || '').toString().trim();
    var programmeKey = (row[3] || '').toString().trim();
    var followToken = (row[7] || '').toString().trim();
    var sessions = [];
    var documents = [];

    // Compatibilité ancienne structure (5 colonnes): row[3] = docs
    var looksLikeOldDocs = false;
    if (row[3] && !row[4] && !row[5]) {
      var s = (row[3] || '').toString().trim();
      looksLikeOldDocs = s.indexOf('[') === 0;
    }

    if (!looksLikeOldDocs && row[4]) {
      try {
        sessions = JSON.parse(row[4]);
        sessions = normalizeSessions_(sessions);
      } catch (e) {
        sessions = [];
      }
    }

    var docsRaw = looksLikeOldDocs ? row[3] : row[5];
    if (docsRaw) {
      try {
        documents = JSON.parse(docsRaw);
        if (!Array.isArray(documents)) documents = [];
      } catch (e) {
        documents = [];
      }
    }

    if (programmeKey && sessions.length === 0) {
      sessions = buildDemoSessionsForProgram_(programmeKey);
    }

    if (!followToken && suiviEnPlace && programmeKey) {
      followToken = generateFollowToken_();
    }

    map[matricule] = {
      suiviEnPlace: suiviEnPlace,
      nomEap: nomEap,
      programmeKey: programmeKey,
      sessions: sessions,
      documents: documents,
      followToken: followToken,
      followLink: followToken ? makeFollowLink_(followToken) : ''
    };
  });

  return map;
}

/** Enregistre le suivi sport (checkbox + nom EAP + programme) */
function saveSportFollowUp(matricule, suiviEnPlace, nomEap, programmeKey) {
  matricule = (matricule || '').toString().trim();
  if (!matricule) throw new Error('Matricule manquant');

  var defs = getProgramDefinitions_();
  if (programmeKey && !defs[programmeKey]) {
    throw new Error('Programme inconnu');
  }

  var sheet = getSportMetaSheet_();
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  var existingSessions = '[]';
  var existingDocs = '[]';
  var existingProgram = '';
  var existingToken = '';

  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === matricule) {
      rowIndex = i + 1;
      existingProgram = (data[i][3] || '').toString().trim();
      existingSessions = data[i][4] || '[]';
      existingDocs = data[i][5] || '[]';
      existingToken = (data[i][7] || '').toString().trim();
      // Compat old 5 cols
      if (!data[i][5] && data[i][3] && (data[i][3] || '').toString().trim().indexOf('[') === 0) {
        existingProgram = '';
        existingSessions = '[]';
        existingDocs = data[i][3] || '[]';
      }
      break;
    }
  }

  var finalProgram = (programmeKey !== undefined && programmeKey !== null)
    ? (programmeKey || '')
    : existingProgram;

  if (finalProgram && (!existingSessions || existingSessions === '[]')) {
    existingSessions = JSON.stringify(buildDemoSessionsForProgram_(finalProgram));
  }

  var finalToken = existingToken;
  if (suiviEnPlace && finalProgram && !finalToken) finalToken = generateFollowToken_();
  if (!suiviEnPlace) finalToken = '';

  var values = [
    matricule,
    suiviEnPlace ? 'oui' : 'non',
    (nomEap || '').toString().trim(),
    finalProgram,
    existingSessions,
    existingDocs,
    new Date(),
    finalToken
  ];

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, 8).setValues([values]);
  } else {
    sheet.appendRow(values);
  }

  return true;
}

/** Upload un document sport (PDF/Word) sur Drive et rattache à l'agent */
function uploadSportDocument(matricule, fileName, mimeType, base64Data) {
  matricule = (matricule || '').toString().trim();
  fileName = (fileName || '').toString().trim();
  mimeType = (mimeType || '').toString().trim();
  base64Data = (base64Data || '').toString().trim();

  if (!matricule) throw new Error('Matricule manquant');
  if (!fileName || !base64Data) throw new Error('Fichier invalide');

  var allowed = {
    'application/pdf': true,
    'application/msword': true,
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': true,
    'application/vnd.oasis.opendocument.text': true
  };
  var extOk = /\.(pdf|doc|docx|odt)$/i.test(fileName);
  if (!allowed[mimeType] && !extOk) {
    throw new Error('Seuls les fichiers PDF/Word sont autorisés');
  }

  var bytes = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(bytes, mimeType, fileName);
  var folder = getSportDocsFolder_();
  var file = folder.createFile(blob);

  var doc = {
    id: file.getId(),
    name: file.getName(),
    mimeType: file.getMimeType(),
    url: file.getUrl(),
    previewUrl: 'https://drive.google.com/file/d/' + file.getId() + '/preview',
    createdAt: new Date().toISOString()
  };

  var sheet = getSportMetaSheet_();
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  var suivi = 'non';
  var nomEap = '';
  var programmeKey = '';
  var sessionsJson = '[]';
  var docs = [];
  var followToken = '';

  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === matricule) {
      rowIndex = i + 1;
      suivi = (data[i][1] || '').toString().trim().toLowerCase() === 'oui' ? 'oui' : 'non';
      nomEap = (data[i][2] || '').toString().trim();
      programmeKey = (data[i][3] || '').toString().trim();
      sessionsJson = data[i][4] || '[]';
      followToken = (data[i][7] || '').toString().trim();
      var docsRaw = data[i][5];
      if (!docsRaw && data[i][3] && (data[i][3] || '').toString().trim().indexOf('[') === 0) {
        // compat ancienne structure
        docsRaw = data[i][3];
        programmeKey = '';
        sessionsJson = '[]';
      }
      if (docsRaw) {
        try {
          docs = JSON.parse(docsRaw);
          if (!Array.isArray(docs)) docs = [];
        } catch (e) {
          docs = [];
        }
      }
      break;
    }
  }

  docs.push(doc);

  if (programmeKey && (!sessionsJson || sessionsJson === '[]')) {
    sessionsJson = JSON.stringify(buildDemoSessionsForProgram_(programmeKey));
  }

  if (suivi === 'oui' && programmeKey && !followToken) followToken = generateFollowToken_();

  var values = [matricule, suivi, nomEap, programmeKey, sessionsJson, JSON.stringify(docs), new Date(), followToken];
  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, 8).setValues([values]);
  } else {
    sheet.appendRow(values);
  }

  return {
    document: doc,
    count: docs.length
  };
}

/** Supprime un document sport de Drive et de la feuille méta */
function deleteSportDocument(matricule, fileId) {
  matricule = (matricule || '').toString().trim();
  fileId = (fileId || '').toString().trim();
  if (!matricule) throw new Error('Matricule manquant');
  if (!fileId) throw new Error('ID fichier manquant');

  // Supprimer de Drive
  try {
    var file = DriveApp.getFileById(fileId);
    file.setTrashed(true);
  } catch (e) {
    // fichier déjà supprimé ou inaccessible – on continue pour nettoyer la méta
  }

  // Retirer de la feuille méta
  var sheet = getSportMetaSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === matricule) {
      var docsRaw = data[i][5] || '[]';
      var docs = [];
      try { docs = JSON.parse(docsRaw); if (!Array.isArray(docs)) docs = []; } catch (e) { docs = []; }
      docs = docs.filter(function (d) { return d.id !== fileId; });
      sheet.getRange(i + 1, 6).setValue(JSON.stringify(docs));
      break;
    }
  }

  return true;
}

/** Retourne le catalogue des programmes avec liens Drive/preview */
function getSportProgramCatalog() {
  return ensureProgramCatalog_();
}

function getSportProgramCatalogSafe_() {
  try {
    return {
      catalog: ensureProgramCatalog_(),
      authRequired: false
    };
  } catch (e) {
    var msg = (e && e.message) ? e.message : e;
    if (msg && msg.toString().indexOf('DriveApp') !== -1) {
      return {
        catalog: {},
        authRequired: true
      };
    }
    throw e;
  }
}

function getOrCreateSportFollowLink(matricule) {
  matricule = (matricule || '').toString().trim();
  if (!matricule) throw new Error('Matricule manquant');

  var sheet = getSportMetaSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === matricule) {
      var suivi = (data[i][1] || '').toString().trim().toLowerCase() === 'oui';
      var programKey = (data[i][3] || '').toString().trim();
      if (!suivi || !programKey) throw new Error('Le suivi et le programme doivent être renseignés');
      var token = (data[i][7] || '').toString().trim();
      if (!token) {
        token = generateFollowToken_();
        sheet.getRange(i + 1, 8).setValue(token);
      }
      return {
        token: token,
        url: makeFollowLink_(token)
      };
    }
  }
  throw new Error('Agent introuvable dans Suivi sport meta');
}

function getFollowupAgentPageData(followToken) {
  followToken = (followToken || '').toString().trim();
  if (!followToken) throw new Error('Lien de suivi invalide');

  var sheet = getSportMetaSheet_();
  var data = sheet.getDataRange().getValues();
  var found = null;
  for (var i = 1; i < data.length; i++) {
    if ((data[i][7] || '').toString().trim() === followToken) {
      found = {
        row: i + 1,
        matricule: (data[i][0] || '').toString().trim(),
        suivi: (data[i][1] || '').toString().trim().toLowerCase() === 'oui',
        nomEap: (data[i][2] || '').toString().trim(),
        programKey: (data[i][3] || '').toString().trim(),
        sessionsJson: data[i][4] || '[]'
      };
      break;
    }
  }
  if (!found || !found.suivi || !found.programKey) throw new Error('Suivi non disponible pour ce lien');

  var sessions;
  try {
    sessions = normalizeSessions_(JSON.parse(found.sessionsJson || '[]'));
  } catch (e) {
    sessions = [];
  }

  var agents = getAllAgents();
  var agent = null;
  for (var j = 0; j < agents.length; j++) {
    if (agents[j].matricule === found.matricule) { agent = agents[j]; break; }
  }

  var catalog = ensureProgramCatalog_();
  var prog = catalog[found.programKey];
  if (!prog) throw new Error('Programme introuvable');

  return {
    token: followToken,
    matricule: found.matricule,
    agentName: agent ? agent.nomPrenom : found.matricule,
    centre: agent ? (agent.centrePrincipal || '') : '',
    nomEap: found.nomEap,
    programKey: found.programKey,
    programLabel: prog.label,
    programUrl: prog.url,
    programPreviewUrl: prog.previewUrl,
    sessions: sessions
  };
}

function logFollowupSession(followToken, dateIso, commentaire) {
  followToken = (followToken || '').toString().trim();
  dateIso = (dateIso || '').toString().trim();
  commentaire = (commentaire || '').toString().trim();
  if (!followToken) throw new Error('Lien invalide');
  if (!dateIso || !/^\d{4}-\d{2}-\d{2}$/.test(dateIso)) throw new Error('Date invalide');

  var sheet = getSportMetaSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][7] || '').toString().trim() === followToken) {
      var sessions = [];
      try {
        sessions = normalizeSessions_(JSON.parse(data[i][4] || '[]'));
      } catch (e) {
        sessions = [];
      }
      sessions.push({
        dateIso: dateIso,
        date: formatDate_(new Date(dateIso + 'T12:00:00')),
        commentaire: commentaire,
        createdAt: new Date().toISOString()
      });
      sessions = normalizeSessions_(sessions);

      sheet.getRange(i + 1, 5).setValue(JSON.stringify(sessions));
      sheet.getRange(i + 1, 7).setValue(new Date());

      return {
        ok: true,
        sessions: sessions
      };
    }
  }
  throw new Error('Lien de suivi introuvable');
}

/* ═══════════════════════════════════════════════════════
   VACCINS
   ═══════════════════════════════════════════════════════ */

/**
 * Construit { matricule: { hb: [{date, dateRaw, nom}], dtp: [{date, dateRaw, nom}], immunise: bool, nonRepondeur: bool } }
 * HB = nom contient "Hépatite B" (insensible casse)
 * DTP = nom contient Boostrix / DTP / DTPC / Revaxis
 */
function getVaccinData_() {
  var data = getSheetData_(CONFIG.SHEETS.VACCINS);
  var map = {};

  data.forEach(function (row) {
    var matricule = (row[CONFIG.COLS_VACCINS.MATRICULE] || '').toString().trim();
    if (!matricule) return;

    var nomVaccin = (row[CONFIG.COLS_VACCINS.NOM_VACCIN] || '').toString().trim();
    if (!nomVaccin) return;

    var dateVal = row[CONFIG.COLS_VACCINS.DATE];
    var dateRaw = (dateVal instanceof Date && !isNaN(dateVal.getTime())) ? dateVal.getTime() : null;

    var immunise = (row[CONFIG.COLS_VACCINS.IMMUNISE] || '').toString().trim().toLowerCase() === 'oui';
    var nonRepondeur = (row[CONFIG.COLS_VACCINS.NON_REPONDEUR] || '').toString().trim().toLowerCase() === 'oui';

    if (!map[matricule]) map[matricule] = { hb: [], dtp: [], immunise: false, nonRepondeur: false };

    var entry = { date: formatDate_(dateVal), dateRaw: dateRaw, nom: nomVaccin };
    var nomLower = nomVaccin.toLowerCase();

    // Classification HB / DTP
    if (nomLower.indexOf('hépatite b') !== -1 || nomLower.indexOf('hepatite b') !== -1) {
      map[matricule].hb.push(entry);
    }
    if (nomLower.indexOf('boostrix') !== -1 || nomLower.indexOf('dtp') !== -1 || nomLower.indexOf('revaxis') !== -1) {
      map[matricule].dtp.push(entry);
    }

    if (immunise) map[matricule].immunise = true;
    if (nonRepondeur) map[matricule].nonRepondeur = true;
  });

  // Trier chronologiquement (plus ancien d'abord)
  Object.keys(map).forEach(function (key) {
    map[key].hb.sort(function (a, b) { return (a.dateRaw || 0) - (b.dateRaw || 0); });
    map[key].dtp.sort(function (a, b) { return (a.dateRaw || 0) - (b.dateRaw || 0); });
  });

  return map;
}

/* ═══════════════════════════════════════════════════════
   SÉROLOGIES
   ═══════════════════════════════════════════════════════ */

/**
 * Construit { matricule: [{type, resultat}] }
 * Types HB intéressants : "Ac anti HBc", "Ac anti HBs", "Ag HBs"
 */
function getSeroData_() {
  var data = getSheetData_(CONFIG.SHEETS.SERO);
  var map = {};

  data.forEach(function (row) {
    var matricule = (row[CONFIG.COLS_SERO.MATRICULE] || '').toString().trim();
    if (!matricule) return;

    var type = (row[CONFIG.COLS_SERO.TYPE] || '').toString().trim();
    var resultat = (row[CONFIG.COLS_SERO.RESULTAT] || '').toString().trim();
    if (!type) return;

    var dateVal = row[CONFIG.COLS_SERO.DATE];
    var dateRaw = (dateVal instanceof Date && !isNaN(dateVal.getTime())) ? dateVal.getTime() : null;

    if (!map[matricule]) map[matricule] = [];
    map[matricule].push({ type: type, resultat: resultat, date: formatDate_(dateVal), dateRaw: dateRaw });
  });

  // Trier par date décroissante (plus récent en premier)
  Object.keys(map).forEach(function (key) {
    map[key].sort(function (a, b) { return (b.dateRaw || 0) - (a.dateRaw || 0); });
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
    sheet.getRange(1, 1, 1, 12).setValues([[
      'ID', 'Matricule', 'Type', 'Détail examen',
      'Date demande', 'Date résultat attendu', 'Commentaire', 'Statut', 'Géré par',
      'Relance 1', 'Relance 2', 'Relance 3'
    ]]);
    sheet.setFrozenRows(1);
  } else {
    // Garantir la présence des colonnes Relance 1/2/3
    if (sheet.getLastColumn() < 12) {
      sheet.getRange(1, 1, 1, 12).setValues([[
        'ID', 'Matricule', 'Type', 'Détail examen',
        'Date demande', 'Date résultat attendu', 'Commentaire', 'Statut', 'Géré par',
        'Relance 1', 'Relance 2', 'Relance 3'
      ]]);
    }
  }
  return sheet;
}

/* ═══════════════════════════════════════════════════════
   AGENTS INACTIFS (ARCHIVAGE)
   ═══════════════════════════════════════════════════════ */

/** Retourne (et crée si besoin) l'onglet Inactifs */
function getInactifsSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.INACTIFS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.INACTIFS);
    sheet.getRange(1, 1, 1, 3).setValues([['Matricule', 'NOM Prénom', 'Date archivage']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Retourne un Set des matricules inactifs */
function getInactiveMatricules_() {
  var sheet = getInactifsSheet_();
  if (sheet.getLastRow() < 2) return {};
  var data = sheet.getDataRange().getValues().slice(1);
  var map = {};
  data.forEach(function (row) {
    var m = (row[0] || '').toString().trim();
    if (m) map[m] = true;
  });
  return map;
}

/** Archive un agent (ajoute à l'onglet Inactifs) */
function archiveAgent(matricule, nomPrenom) {
  var sheet = getInactifsSheet_();
  // Vérifier s'il est déjà archivé
  if (sheet.getLastRow() >= 2) {
    var data = sheet.getDataRange().getValues().slice(1);
    for (var i = 0; i < data.length; i++) {
      if (data[i][0].toString().trim() === matricule) return true; // déjà archivé
    }
  }
  sheet.appendRow([matricule, nomPrenom || '', new Date()]);
  return true;
}

/** Restaure un agent (supprime de l'onglet Inactifs) */
function restoreAgent(matricule) {
  var sheet = getInactifsSheet_();
  if (sheet.getLastRow() < 2) return false;
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0].toString().trim() === matricule) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
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

    var dateDem = row[CONFIG.COLS_EXAMENS.DATE_DEMANDE];
    var dateRes = row[CONFIG.COLS_EXAMENS.DATE_RESULTAT];
    var rel1 = row[CONFIG.COLS_EXAMENS.RELANCE_1];
    var rel2 = row[CONFIG.COLS_EXAMENS.RELANCE_2];
    var rel3 = row[CONFIG.COLS_EXAMENS.RELANCE_3];
    examens.push({
      id:              (row[CONFIG.COLS_EXAMENS.ID] || '').toString(),
      matricule:       (row[CONFIG.COLS_EXAMENS.MATRICULE] || '').toString().trim(),
      type:            (row[CONFIG.COLS_EXAMENS.TYPE] || '').toString().trim(),
      detail:          (row[CONFIG.COLS_EXAMENS.DETAIL] || '').toString().trim(),
      dateDemande:     formatDate_(dateDem),
      dateDemandeRaw:  (dateDem instanceof Date && !isNaN(dateDem.getTime())) ? dateDem.getTime() : null,
      dateResultat:    formatDate_(dateRes),
      dateResultatRaw: (dateRes instanceof Date && !isNaN(dateRes.getTime())) ? dateRes.getTime() : null,
      commentaire:     (row[CONFIG.COLS_EXAMENS.COMMENTAIRE] || '').toString().trim(),
      statut:          statut || 'ouvert',
      gerePar:         (row[CONFIG.COLS_EXAMENS.GERE_PAR] || '').toString().trim(),
      relance1Date:    formatDate_(rel1),
      relance1Raw:     (rel1 instanceof Date && !isNaN(rel1.getTime())) ? rel1.getTime() : null,
      relance2Date:    formatDate_(rel2),
      relance2Raw:     (rel2 instanceof Date && !isNaN(rel2.getTime())) ? rel2.getTime() : null,
      relance3Date:    formatDate_(rel3),
      relance3Raw:     (rel3 instanceof Date && !isNaN(rel3.getTime())) ? rel3.getTime() : null,
      acquitte:        (row[CONFIG.COLS_EXAMENS.ACQUITTE] || '').toString().trim().toLowerCase()
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
  /* Auto-calculer la date limite si non fournie */
  if (!dateResultat) {
    var dlDays = { biologie: 15, automesure: 7, imagerie: 15 };
    var typ = (examenData.type || '').toString().trim();
    dateResultat = new Date(dateDemande);
    if (typ === 'avis_specialise') { dateResultat.setMonth(dateResultat.getMonth() + 2); }
    else if (dlDays[typ]) { dateResultat.setDate(dateResultat.getDate() + dlDays[typ]); }
    else { dateResultat = null; }
  }
  var commentaire  = (examenData.commentaire || '').toString().trim();
  if (commentaire) {
    commentaire = formatDate_(new Date()) + ' \u2014 ' + commentaire;
  }

  sheet.appendRow([
    id,
    examenData.matricule || '',
    examenData.type || '',
    examenData.detail || '',
    dateDemande,
    dateResultat,
    commentaire,
    'ouvert',
    examenData.gerePar || '',
    '',
    '',
    '',
    ''
  ]);

  return {
    id:              id,
    matricule:       (examenData.matricule || '').toString().trim(),
    type:            (examenData.type || '').toString().trim(),
    detail:          (examenData.detail || '').toString().trim(),
    dateDemande:     formatDate_(dateDemande),
    dateDemandeRaw:  dateDemande ? dateDemande.getTime() : null,
    dateResultat:    formatDate_(dateResultat),
    dateResultatRaw: dateResultat ? dateResultat.getTime() : null,
    commentaire:     commentaire,
    statut:          'ouvert',
    gerePar:         (examenData.gerePar || '').toString().trim(),
    relance1Date:    '',
    relance1Raw:     null,
    relance2Date:    '',
    relance2Raw:     null,
    relance3Date:    '',
    relance3Raw:     null,
    acquitte:        ''
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

/** Met à jour le champ 'géré par' d'un examen */
function updateExamGerePar(examId, gerePar) {
  var sheet = getExamensSheet_();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === examId) {
      sheet.getRange(i + 1, CONFIG.COLS_EXAMENS.GERE_PAR + 1).setValue(gerePar || '');
      return true;
    }
  }
  return false;
}

function updateExamDateResultat(examId, dateResultat) {
  var sheet = getExamensSheet_();
  var data = sheet.getDataRange().getValues();
  var parsedDate = dateResultat ? new Date(dateResultat) : null;

  if (dateResultat && (!(parsedDate instanceof Date) || isNaN(parsedDate.getTime()))) {
    throw new Error('Date de résultat invalide');
  }

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === examId) {
      sheet.getRange(i + 1, CONFIG.COLS_EXAMENS.DATE_RESULTAT + 1).setValue(parsedDate || '');
      return {
        ok: true,
        dateResultat: formatDate_(parsedDate),
        dateResultatRaw: parsedDate ? parsedDate.getTime() : null
      };
    }
  }
  throw new Error('Examen introuvable');
}

/** Met à jour la date de demande et gère l'alerte prescription */
function updateExamDateDemande(examId, dateDemande) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEETS.EXAMENS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][CONFIG.COLS_EXAMENS.ID] === examId) {
      var d = new Date(dateDemande);
      sheet.getRange(i + 1, CONFIG.COLS_EXAMENS.DATE_DEMANDE + 1).setValue(d);
      /* Si la date est dans le futur, planifier l'alerte */
      var today = new Date();
      today.setHours(0, 0, 0, 0);
      var dDay = new Date(d); dDay.setHours(0, 0, 0, 0);
      var newAcquitte = dDay > today ? 'planifie' : '';
      if (sheet.getLastColumn() < 13) sheet.getRange(1, 13).setValue('Acquitté');
      sheet.getRange(i + 1, 13).setValue(newAcquitte);
      return {
        dateDemande: Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        dateDemandeRaw: d.getTime(),
        acquitte: newAcquitte
      };
    }
  }
  throw new Error('Examen introuvable : ' + examId);
}

function acquitterExamen(examId) {
  var sheet = getExamensSheet_();
  var data = sheet.getDataRange().getValues();
  // Ensure column M exists
  if (sheet.getLastColumn() < 13) {
    sheet.getRange(1, 13).setValue('Acquitté');
  }
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === examId) {
      sheet.getRange(i + 1, 13).setValue('oui');
      return true;
    }
  }
  throw new Error('Examen introuvable');
}

function updateExamCommentaire(examId, commentaire) {
  var sheet = getExamensSheet_();
  var data = sheet.getDataRange().getValues();
  var value = (commentaire || '').toString().trim();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === examId) {
      sheet.getRange(i + 1, CONFIG.COLS_EXAMENS.COMMENTAIRE + 1).setValue(value);
      return {
        ok: true,
        commentaire: value
      };
    }
  }
  throw new Error('Examen introuvable');
}

/** Coche/décoche une relance (1/2/3) pour un examen */
function setExamRelance(examId, relanceIndex, checked) {
  var idxMap = {
    1: CONFIG.COLS_EXAMENS.RELANCE_1,
    2: CONFIG.COLS_EXAMENS.RELANCE_2,
    3: CONFIG.COLS_EXAMENS.RELANCE_3
  };
  var colIdx = idxMap[relanceIndex];
  if (colIdx === undefined) throw new Error('Relance invalide');

  var sheet = getExamensSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === examId) {
      var value = checked ? new Date() : '';
      sheet.getRange(i + 1, colIdx + 1).setValue(value);
      return {
        ok: true,
        date: checked ? formatDate_(value) : ''
      };
    }
  }
  throw new Error('Examen introuvable');
}

/* ═══════════════════════════════════════════════════════
   DÉTERMINATION DU TYPE DE VISITE
   ═══════════════════════════════════════════════════════ */

/**
 * Retourne le type de visite selon les règles métier unifiées
 * (retard et à venir = même traitement).
 *
 * La date de dernière visite effective vient de l'onglet "derniere visite".
 * On en extrait l'année (visitYear) et on applique :
 *
 * 1. Spécialité VMA → VMA tous les ans
 * 2. visitYear ≤ 2024 (ou inconnue) → Visite médicale 2026
 * 3. visitYear = 2025 :
 *    a. Né année paire           → Visite médicale biennale (2026)
 *    b. Né année impaire, ≥ 39   → Visite prévention (2026)
 *    c. Né année impaire, < 39   → Visite médicale 2027 (rien en 2026)
 * 4. visitYear ≥ 2026            → Déjà vu
 */
function determineVisitType_(agent, specialties) {
  var agentSpe = specialties[agent.nomPrenom];
  var refYear  = CONFIG.REFERENCE_YEAR; // 2026

  /* ── 1. Spécialités → VMA tous les ans ── */
  if (agentSpe && agentSpe.length > 0) {
    for (var i = 0; i < agentSpe.length; i++) {
      var spe = agentSpe[i];
      if (CONFIG.VMA_SPECIALTIES.indexOf(spe) !== -1) {
        return { type: 'VMA tous les ans', raison: 'Spécialité : ' + spe };
      }
      if (spe === 'Grimp' && agent.age >= CONFIG.VMA_GRIMP_AGE) {
        return { type: 'VMA tous les ans', raison: 'Spécialité : Grimp (≥ ' + CONFIG.VMA_GRIMP_AGE + ' ans)' };
      }
      if (spe.toLowerCase() === 'diabétique') {
        return { type: 'VMA tous les ans', raison: 'Spécialité : Diabétique' };
      }
    }
  }

  /* ── 2. Dernière visite ≤ 2024 ou inconnue → visite médicale obligatoire ── */
  var visitYear = agent.visitYear;
  if (!visitYear || visitYear <= refYear - 2) {
    return {
      type: 'Visite médicale ' + refYear,
      raison: 'Dernière visite en ' + (visitYear || '?') + ' (≤ ' + (refYear - 2) + ') → visite médicale tous les 2 ans max'
    };
  }

  /* ── 3. Dernière visite = 2025 → parité + âge ── */
  if (visitYear === refYear - 1) {
    var birthYear = agent.birthYear;
    if (!birthYear) {
      return { type: 'Non déterminé', raison: 'Visite en ' + (refYear - 1) + ' mais date de naissance inconnue' };
    }

    var isEvenBirth = (birthYear % 2 === 0);

    /* 3a. Né année paire → visite médicale biennale */
    if (isEvenBirth) {
      return {
        type: 'Visite médicale biennale',
        raison: 'Né en ' + birthYear + ' (paire), visite en ' + (refYear - 1) + ' → visite médicale en ' + refYear
      };
    }

    /* 3c. Né année impaire + < 39 ans → visite médicale 2027 (pas de prévention en 2026) */
    if (agent.age < CONFIG.AGE_THRESHOLD) {
      return {
        type: 'Visite médicale ' + (refYear + 1),
        raison: 'Né en ' + birthYear + ' (impaire), < ' + CONFIG.AGE_THRESHOLD + ' ans, visite en ' + (refYear - 1) + ' → prochaine visite médicale en ' + (refYear + 1)
      };
    }

    /* 3b. Né année impaire + ≥ 39 ans → visite prévention */
    return {
      type: 'Visite prévention',
      raison: 'Né en ' + birthYear + ' (impaire), visite en ' + (refYear - 1) + ' → prévention en ' + refYear
    };
  }

  /* ── 4. Dernière visite ≥ 2026 → déjà à jour ── */
  if (visitYear >= refYear) {
    return { type: 'Déjà vu en ' + visitYear, raison: 'Visite effectuée en ' + visitYear };
  }

  return { type: 'Non déterminé', raison: 'Cas non couvert (visite ' + visitYear + ')' };
}

/* ═══════════════════════════════════════════════════════
   CHARGEMENT DES AGENTS
   ═══════════════════════════════════════════════════════ */

/**
 * Charge, dédoublonne et enrichit la liste complète des agents
 */
function getAllAgents() {
  var retardData      = getSheetData_(CONFIG.SHEETS.RETARD);
  var aVenirData      = getSheetData_(CONFIG.SHEETS.A_VENIR);
  var specialties     = getSpecialties_();
  var lastVisitDates  = getLastVisitDates_();

  /* Marquer chaque ligne avec sa source */
  var allData = [];
  retardData.forEach(function (row) { allData.push({ row: row, source: 'retard' }); });
  aVenirData.forEach(function (row) { allData.push({ row: row, source: 'a_venir' }); });

  var agentsMap = {};

  allData.forEach(function (item) {
    var row = item.row;
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

    /* La colonne E = date prévue (deadline prochaine visite), PAS la date effective.
       La vraie date de dernière visite vient de l'onglet "derniere visite". */
    var dateDerniereVisite = lastVisitDates[key] || null;
    var visitYear = dateDerniereVisite ? dateDerniereVisite.getFullYear() : null;

    /* dateProchVisite = colonne E brute (date avant laquelle l'agent doit passer) */
    var dateProchVisite = null;
    if (dateVisite instanceof Date && !isNaN(dateVisite.getTime())) {
      dateProchVisite = dateVisite;
    }

    var agent = {
      age:                    age,
      centreSecondaire:       centreSecondaire,
      centrePrincipal:        centrePrincipal,
      dateNaissance:          formatDate_(dateNaissance),
      dateProchVisite:        formatDate_(dateProchVisite),
      dateProchVisiteRaw:     dateProchVisite ? dateProchVisite.getTime() : null,
      prochVisiteYear:        dateProchVisite ? dateProchVisite.getFullYear() : null,
      datePerteCompetence:    formatDate_(datePerteCompetence),
      datePerteCompetenceRaw: datePerteCompetence ? datePerteCompetence.getTime() : null,
      nomPrenom:              nomPrenom,
      matricule:              key,
      objetVisite:            objetVisite,
      birthYear:              birthYear,
      perteYear:              perteYear,
      visitYear:              visitYear,
      isRetard:               item.source === 'retard',
      specialites:            specialties[nomPrenom] || [],
      typeVisite:             '',
      typeVisiteRaison:       '',
      sport:                  []
    };

    var visitResult = determineVisitType_(agent, specialties);
    agent.typeVisite       = visitResult.type;
    agent.typeVisiteRaison = visitResult.raison;
    agentsMap[key]         = agent;
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
 *  - agents actifs (avec sport)
 *  - agents inactifs (avec sport)
 *  - examens ouverts
 *  - totalAgents
 */
function getPageData() {
  var agents    = getAllAgents();
  var sportData = getSportData_();
  var examens   = getAllExamens();
  var inactifs  = getInactiveMatricules_();
  var vaccinData = getVaccinData_();
  var seroData   = getSeroData_();
  var sportMeta  = getSportMetaMap_();
  var programInfo = getSportProgramCatalogSafe_();

  var bruleurData = getBruleurCaissonData_();

  // Séparer actifs / inactifs et rattacher sport
  var activeAgents   = [];
  var inactiveAgents = [];

  agents.forEach(function (a) {
    a.sport = sportData[a.matricule] || [];
    if (inactifs[a.matricule]) {
      inactiveAgents.push(a);
    } else {
      activeAgents.push(a);
    }
  });

  return {
    agents:         activeAgents,
    inactiveAgents: inactiveAgents,
    examens:        examens,
    totalAgents:    activeAgents.length,
    vaccins:        vaccinData,
    seros:          seroData,
    sportMeta:      sportMeta,
    sportPrograms:  programInfo.catalog,
    sportProgramAuthRequired: programInfo.authRequired,
    bruleurCaisson: bruleurData
  };
}

/** Remplit la colonne A de l'onglet "cis / mailing" avec tous les CIS */
function populateCisMailingSheet() {
  var ss    = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CIS_MAILING);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.CIS_MAILING);

  // Set headers
  sheet.getRange(1, 1, 1, 3).setValues([['CIS', 'Email', 'Token']]);

  // Read existing data to preserve emails & tokens
  var existingMap = {}; // { cisName: { email, token } }
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var oldData = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = 0; i < oldData.length; i++) {
      var name = (oldData[i][0] || '').toString().trim();
      if (!name) continue;
      existingMap[name] = {
        email: (oldData[i][1] || '').toString().trim(),
        token: (oldData[i][2] || '').toString().trim()
      };
    }
    sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }

  var cisList = getCisList();
  if (cisList.length > 0) {
    var values = cisList.map(function (c) {
      var existing = existingMap[c] || {};
      var token = existing.token || Utilities.getUuid().replace(/-/g, '').slice(0, 16);
      return [c, existing.email || '', token];
    });
    sheet.getRange(2, 1, values.length, 3).setValues(values);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    cisList.length + ' CIS mis à jour avec tokens dans l\'onglet "cis / mailing"',
    'Mise à jour terminée'
  );
  return cisList;
}

/* ═══════════════════════════════════════════════════════
   SUIVI SPÉCIALITÉ BRÛLEUR / CAISSON
   ═══════════════════════════════════════════════════════ */

/**
 * Retourne (et crée si besoin) l'onglet "Suivi Bruleur Caisson"
 * Colonnes : Matricule | Exposition>20ans | Scanner Statut | Scanner Date |
 *            ECBU JSON (tableau d'entrées)
 */
function getBruleurSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.BRULEUR);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.BRULEUR);
    sheet.getRange(1, 1, 1, 6).setValues([[
      'Matricule', 'Exposition>20ans', 'Scanner Statut', 'Scanner Date', 'ECBU JSON', 'Consentement'
    ]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Lit toutes les données brûleur/caisson
 * Retourne { matricule: { exposition: bool, scanner: {statut, date}, ecbus: [{date, statut, dateVu}] } }
 */
function getBruleurCaissonData_() {
  var sheet = getBruleurSheet_();
  if (sheet.getLastRow() < 2) return {};
  var data = sheet.getDataRange().getValues().slice(1);
  var map = {};

  data.forEach(function (row) {
    var matricule = (row[0] || '').toString().trim();
    if (!matricule) return;

    var exposition = (row[1] || '').toString().trim().toLowerCase();
    var scannerStatut = (row[2] || '').toString().trim();
    var scannerDate = row[3];
    var ecbuJson = (row[4] || '').toString().trim();
    var consentement = (row[5] || '').toString().trim();

    var ecbus = [];
    if (ecbuJson) {
      try { ecbus = JSON.parse(ecbuJson); } catch (e) { ecbus = []; }
    }

    map[matricule] = {
      exposition: exposition === 'oui' || exposition === 'true' || exposition === '1',
      scanner: {
        statut: scannerStatut || '',
        date: (scannerDate instanceof Date && !isNaN(scannerDate.getTime())) ? formatDate_(scannerDate) : (scannerDate || '').toString().trim()
      },
      consentement: consentement || '',
      ecbus: ecbus
    };
  });

  return map;
}

/** Sauvegarde le flag exposition > 20 ans */
function saveBruleurExposition(matricule, exposition) {
  matricule = (matricule || '').toString().trim();
  if (!matricule) throw new Error('Matricule manquant');

  var sheet = getBruleurSheet_();
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === matricule) { rowIndex = i + 1; break; }
  }

  if (rowIndex === -1) {
    sheet.appendRow([matricule, exposition ? 'oui' : '', '', '', '[]', '']);
  } else {
    sheet.getRange(rowIndex, 2).setValue(exposition ? 'oui' : '');
    // Si on décoche, on ne supprime pas les données existantes
  }
  return true;
}

/** Sauvegarde le statut scanner (prescrit / recu / vu) */
function saveScannerStatus(matricule, statut) {
  matricule = (matricule || '').toString().trim();
  if (!matricule) throw new Error('Matricule manquant');
  var validStatuts = ['', 'prescrit', 'recu', 'vu'];
  if (validStatuts.indexOf(statut) === -1) throw new Error('Statut invalide');

  var sheet = getBruleurSheet_();
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === matricule) { rowIndex = i + 1; break; }
  }

  var now = statut ? formatDate_(new Date()) : '';
  if (rowIndex === -1) {
    sheet.appendRow([matricule, '', statut, now, '[]', '']);
  } else {
    sheet.getRange(rowIndex, 3).setValue(statut);
    if (statut) sheet.getRange(rowIndex, 4).setValue(now);
  }
  return true;
}

/** Ajoute ou met à jour une entrée ECBU */
function saveEcbuEntry(matricule, ecbuData) {
  matricule = (matricule || '').toString().trim();
  if (!matricule) throw new Error('Matricule manquant');

  var sheet = getBruleurSheet_();
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  var existingEcbus = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === matricule) {
      rowIndex = i + 1;
      try { existingEcbus = JSON.parse(data[i][4] || '[]'); } catch (e) { existingEcbus = []; }
      break;
    }
  }

  // ecbuData = { id?, date, statut }
  if (ecbuData.id) {
    // Mise à jour
    for (var j = 0; j < existingEcbus.length; j++) {
      if (existingEcbus[j].id === ecbuData.id) {
        existingEcbus[j].date = ecbuData.date || existingEcbus[j].date;
        existingEcbus[j].statut = ecbuData.statut || existingEcbus[j].statut;
        if (ecbuData.statut === 'vu') {
          existingEcbus[j].dateVu = formatDate_(new Date());
        }
        break;
      }
    }
  } else {
    // Nouvel ECBU
    var newId = 'ecbu_' + new Date().getTime();
    existingEcbus.push({
      id: newId,
      date: ecbuData.date || formatDate_(new Date()),
      statut: ecbuData.statut || 'prescrit',
      dateVu: ''
    });
  }

  var jsonStr = JSON.stringify(existingEcbus);
  if (rowIndex === -1) {
    sheet.appendRow([matricule, '', '', '', jsonStr, '']);
  } else {
    sheet.getRange(rowIndex, 5).setValue(jsonStr);
  }
  return existingEcbus;
}

/** Supprime une entrée ECBU */
function deleteEcbuEntry(matricule, ecbuId) {
  matricule = (matricule || '').toString().trim();
  if (!matricule) throw new Error('Matricule manquant');

  var sheet = getBruleurSheet_();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === matricule) {
      var ecbus = [];
      try { ecbus = JSON.parse(data[i][4] || '[]'); } catch (e) { ecbus = []; }
      ecbus = ecbus.filter(function (e) { return e.id !== ecbuId; });
      sheet.getRange(i + 1, 5).setValue(JSON.stringify(ecbus));
      return ecbus;
    }
  }
  return [];
}

/** Sauvegarde le statut consentement (envoye / recu / '') */
function saveConsentementStatus(matricule, statut) {
  matricule = (matricule || '').toString().trim();
  if (!matricule) throw new Error('Matricule manquant');
  var validStatuts = ['', 'envoye', 'recu'];
  if (validStatuts.indexOf(statut) === -1) throw new Error('Statut invalide');

  var sheet = getBruleurSheet_();
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === matricule) { rowIndex = i + 1; break; }
  }

  // Ensure column 6 exists
  if (sheet.getLastColumn() < 6) {
    sheet.getRange(1, 6).setValue('Consentement');
  }

  if (rowIndex === -1) {
    sheet.appendRow([matricule, '', '', '', '[]', statut]);
  } else {
    sheet.getRange(rowIndex, 6).setValue(statut);
  }
  return true;
}

/* ═══════════════════════════════════════════════════════
   EMAIL AUTOMATIQUE — RÉSUMÉ HEBDOMADAIRE (v55)
   ═══════════════════════════════════════════════════════ */

var EMAIL_CONTACTS = {
  'Cécile': 'cecile.verges@sdis66.fr',
  'Célia':  'celia.bertoncelo@sdis66.fr'
};

/**
 * Installe les triggers pour l'envoi automatique :
 * lundi 8h + vendredi 8h.
 * Supprime les anciens triggers avant d'en créer de nouveaux.
 */
function setupEmailTriggers() {
  removeEmailTriggers();
  // Lundi 8h
  ScriptApp.newTrigger('sendAllWeeklySummaries')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  // Vendredi 8h
  ScriptApp.newTrigger('sendAllWeeklySummaries')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(8)
    .create();
  return 'Triggers installés : lundi 8h + vendredi 8h';
}

/** Supprime tous les triggers d'envoi email */
function removeEmailTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendAllWeeklySummaries') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  return 'Triggers supprimés';
}

/** Vérifie si les triggers sont actifs */
function getEmailTriggersStatus() {
  var triggers = ScriptApp.getProjectTriggers();
  var count = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendAllWeeklySummaries') count++;
  }
  return { active: count > 0, count: count };
}

/** Point d'entrée du trigger : envoie un email à chaque personne */
function sendAllWeeklySummaries() {
  var persons = Object.keys(EMAIL_CONTACTS);
  for (var i = 0; i < persons.length; i++) {
    try {
      sendWeeklySummaryFor_(persons[i]);
    } catch (e) {
      Logger.log('Erreur envoi email ' + persons[i] + ': ' + e.message);
    }
  }
}

/** Envoi du résumé pour une personne donnée */
function sendWeeklySummaryFor_(person) {
  var email = EMAIL_CONTACTS[person];
  if (!email) return;

  var allExams = getAllExamens();
  var allAgts  = getAllAgents();
  var bruleur  = getBruleurCaissonData_();
  var now = new Date(); now.setHours(0, 0, 0, 0);
  var todayTs = now.getTime();
  var SIX_MONTHS = 6 * 30.44 * 24 * 3600000;

  var labels = { biologie: 'Biologie', automesure: 'Automesure tensionnelle', imagerie: 'Imagerie', avis_specialise: 'Avis spécialisé', autre: 'Autre' };
  var icons  = { biologie: '🧪', automesure: '🩻', imagerie: '☢️', avis_specialise: '🩺', autre: '🅰️' };

  /* Map matricule → nom */
  var nameMap = {};
  for (var i = 0; i < allAgts.length; i++) {
    nameMap[allAgts[i].matricule] = allAgts[i].nomPrenom;
  }

  var personExams = allExams.filter(function (e) { return e.gerePar === person; });

  /* Collecter événements */
  var overdue = [];
  var upcoming = [];

  for (var i = 0; i < personExams.length; i++) {
    var ex = personExams[i];
    var agentName = nameMap[ex.matricule] || ex.matricule;

    /* Résultats à récupérer */
    if (ex.dateResultatRaw) {
      var item = {
        date: ex.dateResultatRaw,
        dateStr: ex.dateResultat,
        agent: agentName,
        type: labels[ex.type] || ex.type,
        icon: icons[ex.type] || '📋',
        detail: ex.detail,
        action: 'Récupérer résultat'
      };
      if (ex.dateResultatRaw < todayTs) { overdue.push(item); } else { upcoming.push(item); }
    }

    /* Prescriptions planifiées */
    if (ex.acquitte === 'planifie' && ex.dateDemandeRaw) {
      var item2 = {
        date: ex.dateDemandeRaw,
        dateStr: ex.dateDemande,
        agent: agentName,
        type: labels[ex.type] || ex.type,
        icon: icons[ex.type] || '📋',
        detail: ex.detail,
        action: 'Prescrire examen'
      };
      if (ex.dateDemandeRaw < todayTs) { overdue.push(item2); } else { upcoming.push(item2); }
    }

    /* Relances */
    for (var r = 1; r <= 3; r++) {
      if (ex['relance' + r + 'Raw']) {
        var rDate = ex['relance' + r + 'Raw'];
        var item3 = {
          date: rDate,
          dateStr: ex['relance' + r + 'Date'],
          agent: agentName,
          type: labels[ex.type] || ex.type,
          icon: '🔔',
          detail: ex.detail,
          action: 'Relance ' + r
        };
        if (rDate < todayTs) { overdue.push(item3); } else { upcoming.push(item3); }
      }
    }
  }

  /* ECBU prochains */
  for (var i = 0; i < allAgts.length; i++) {
    var a = allAgts[i];
    var bd = bruleur[a.matricule];
    if (bd && bd.exposition && bd.ecbus && bd.ecbus.length > 0) {
      var lastEcbuDate = null;
      for (var j = bd.ecbus.length - 1; j >= 0; j--) {
        if (bd.ecbus[j].date) {
          var parts = bd.ecbus[j].date.match(/(\d{2})\/(\d{2})\/(\d{4})/);
          if (parts) lastEcbuDate = new Date(parseInt(parts[3]), parseInt(parts[2]) - 1, parseInt(parts[1]));
          break;
        }
      }
      if (lastEcbuDate) {
        var nextEcbu = new Date(lastEcbuDate.getTime() + SIX_MONTHS);
        var ecbuItem = {
          date: nextEcbu.getTime(),
          dateStr: formatDate_(nextEcbu),
          agent: a.nomPrenom,
          type: 'ECBU (brûleur)',
          icon: '🧪',
          detail: 'Prochain contrôle 6 mois',
          action: 'ECBU à faire'
        };
        if (nextEcbu.getTime() < todayTs) { overdue.push(ecbuItem); } else { upcoming.push(ecbuItem); }
      }
    }
  }

  /* Tri */
  overdue.sort(function (a, b) { return a.date - b.date; });
  upcoming.sort(function (a, b) { return a.date - b.date; });

  /* Limiter */
  var overdueSlice  = overdue.slice(0, 20);
  var upcomingSlice = upcoming.slice(0, 20);

  if (overdueSlice.length === 0 && upcomingSlice.length === 0) return; // rien à envoyer

  /* Construire le HTML */
  var dayNames = ['dimanche','lundi','mardi','mercredi','jeudi','vendredi','samedi'];
  var monthNames = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];
  var todayLabel = dayNames[now.getDay()] + ' ' + now.getDate() + ' ' + monthNames[now.getMonth()] + ' ' + now.getFullYear();

  var html = '';
  html += '<div style="font-family:\'Segoe UI\',Arial,sans-serif;max-width:650px;margin:0 auto;background:#ffffff;">';

  /* Header */
  html += '<div style="background:linear-gradient(135deg,#1e3a5f 0%,#2980b9 100%);color:white;padding:30px 24px;border-radius:12px 12px 0 0;">';
  html += '<h1 style="margin:0;font-size:22px;">🏥 SDIS 66 — Suivi VMA</h1>';
  html += '<p style="margin:8px 0 0;opacity:0.9;font-size:14px;">Résumé pour <strong>' + person + '</strong> — ' + todayLabel + '</p>';
  html += '</div>';

  /* Stats banner */
  html += '<div style="display:flex;background:#f8f9fa;border-bottom:1px solid #e0e0e0;">';
  html += '<div style="flex:1;text-align:center;padding:14px;border-right:1px solid #e0e0e0;">';
  html += '<div style="font-size:28px;font-weight:700;color:#c62828;">' + overdue.length + '</div>';
  html += '<div style="font-size:12px;color:#666;text-transform:uppercase;letter-spacing:0.5px;">En retard</div></div>';
  html += '<div style="flex:1;text-align:center;padding:14px;">';
  html += '<div style="font-size:28px;font-weight:700;color:#2e7d32;">' + upcoming.length + '</div>';
  html += '<div style="font-size:12px;color:#666;text-transform:uppercase;letter-spacing:0.5px;">À venir</div></div>';
  html += '</div>';

  /* Section EN RETARD */
  if (overdueSlice.length > 0) {
    html += '<div style="padding:20px 24px;">';
    html += '<h2 style="margin:0 0 14px;font-size:16px;color:#c62828;border-bottom:2px solid #ffcdd2;padding-bottom:8px;">⏰ Dossiers en retard (' + overdue.length + ')</h2>';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<tr style="background:#fafafa;"><th style="text-align:left;padding:8px 10px;color:#666;font-weight:600;border-bottom:1px solid #eee;">Date</th>';
    html += '<th style="text-align:left;padding:8px 10px;color:#666;font-weight:600;border-bottom:1px solid #eee;">Agent</th>';
    html += '<th style="text-align:left;padding:8px 10px;color:#666;font-weight:600;border-bottom:1px solid #eee;">Examen</th>';
    html += '<th style="text-align:left;padding:8px 10px;color:#666;font-weight:600;border-bottom:1px solid #eee;">Action</th></tr>';
    for (var i = 0; i < overdueSlice.length; i++) {
      var ev = overdueSlice[i];
      var bg = i % 2 === 0 ? '#fff' : '#fafafa';
      html += '<tr style="background:' + bg + ';">';
      html += '<td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;color:#c62828;font-weight:600;white-space:nowrap;">' + ev.dateStr + '</td>';
      html += '<td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;font-weight:600;">' + ev.agent + '</td>';
      html += '<td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;">' + ev.icon + ' ' + ev.type + (ev.detail ? ' — ' + ev.detail : '') + '</td>';
      html += '<td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;"><span style="background:#ffebee;color:#c62828;padding:2px 8px;border-radius:10px;font-size:12px;font-weight:600;">' + ev.action + '</span></td>';
      html += '</tr>';
    }
    html += '</table>';
    if (overdue.length > 20) {
      html += '<p style="margin:10px 0 0;font-size:12px;color:#999;">… et ' + (overdue.length - 20) + ' autre(s) dossier(s) en retard.</p>';
    }
    html += '</div>';
  }

  /* Section À VENIR */
  if (upcomingSlice.length > 0) {
    html += '<div style="padding:20px 24px;">';
    html += '<h2 style="margin:0 0 14px;font-size:16px;color:#1565c0;border-bottom:2px solid #bbdefb;padding-bottom:8px;">📆 Prochains dossiers à gérer (' + upcoming.length + ')</h2>';
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<tr style="background:#fafafa;"><th style="text-align:left;padding:8px 10px;color:#666;font-weight:600;border-bottom:1px solid #eee;">Date</th>';
    html += '<th style="text-align:left;padding:8px 10px;color:#666;font-weight:600;border-bottom:1px solid #eee;">Agent</th>';
    html += '<th style="text-align:left;padding:8px 10px;color:#666;font-weight:600;border-bottom:1px solid #eee;">Examen</th>';
    html += '<th style="text-align:left;padding:8px 10px;color:#666;font-weight:600;border-bottom:1px solid #eee;">Action</th></tr>';
    for (var i = 0; i < upcomingSlice.length; i++) {
      var ev2 = upcomingSlice[i];
      var bg2 = i % 2 === 0 ? '#fff' : '#fafafa';
      html += '<tr style="background:' + bg2 + ';">';
      html += '<td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;color:#1565c0;font-weight:600;white-space:nowrap;">' + ev2.dateStr + '</td>';
      html += '<td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;font-weight:600;">' + ev2.agent + '</td>';
      html += '<td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;">' + ev2.icon + ' ' + ev2.type + (ev2.detail ? ' — ' + ev2.detail : '') + '</td>';
      html += '<td style="padding:8px 10px;border-bottom:1px solid #f0f0f0;"><span style="background:#e3f2fd;color:#1565c0;padding:2px 8px;border-radius:10px;font-size:12px;font-weight:600;">' + ev2.action + '</span></td>';
      html += '</tr>';
    }
    html += '</table>';
    if (upcoming.length > 20) {
      html += '<p style="margin:10px 0 0;font-size:12px;color:#999;">… et ' + (upcoming.length - 20) + ' autre(s) dossier(s) à venir.</p>';
    }
    html += '</div>';
  }

  /* Footer */
  html += '<div style="background:#f8f9fa;padding:16px 24px;border-radius:0 0 12px 12px;border-top:1px solid #e0e0e0;text-align:center;">';
  html += '<p style="margin:0;font-size:12px;color:#999;">📧 Email automatique — SDIS 66 Suivi VMA</p>';
  html += '<p style="margin:4px 0 0;font-size:11px;color:#bbb;">Envoyé le ' + todayLabel + ' à 8h00</p>';
  html += '</div>';

  html += '</div>';

  /* Envoi */
  MailApp.sendEmail({
    to: email,
    subject: '🏥 SDIS 66 — Résumé VMA ' + person + ' (' + todayLabel + ')',
    htmlBody: html
  });
}

/* ═══════════════════════════════════════════════════════
   PERMIS C
   ═══════════════════════════════════════════════════════ */

/**
 * Construit { matricule: { dateLimite: "JJ/MM/AAAA", dateLimiteRaw: timestamp } }
 */
function getPermisData_() {
  var data = getSheetData_(CONFIG.SHEETS.PERMIS_C);
  var map = {};

  data.forEach(function (row) {
    var matricule = (row[CONFIG.COLS_PERMIS_C.MATRICULE] || '').toString().trim();
    if (!matricule) return;

    var dateVal = row[CONFIG.COLS_PERMIS_C.DATE_LIMITE];
    var dateRaw = (dateVal instanceof Date && !isNaN(dateVal.getTime())) ? dateVal.getTime() : null;

    // Garder la date la plus ancienne s'il y a des doublons
    if (!map[matricule] || (dateRaw && (!map[matricule].dateLimiteRaw || dateRaw < map[matricule].dateLimiteRaw))) {
      map[matricule] = {
        dateLimite: formatDate_(dateVal),
        dateLimiteRaw: dateRaw
      };
    }
  });

  return map;
}

/* ═══════════════════════════════════════════════════════
   VUE CHEF DE CENTRE (CIS VIEW)
   ═══════════════════════════════════════════════════════ */

/**
 * Lit l'onglet "cis / mailing" et retourne { token: cisName, ... }
 * Génère automatiquement un token pour chaque CIS qui n'en a pas.
 */
function getCisTokenMap_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CIS_MAILING);
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  // Ensure column C header
  var headerC = sheet.getRange(1, 3).getValue();
  if (!headerC || headerC.toString().trim() !== 'Token') {
    sheet.getRange(1, 3).setValue('Token');
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var map = {};
  var needWrite = false;

  for (var i = 0; i < data.length; i++) {
    var cisName = (data[i][0] || '').toString().trim();
    if (!cisName) continue;

    var token = (data[i][2] || '').toString().trim();
    if (!token) {
      token = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
      data[i][2] = token;
      needWrite = true;
    }
    map[token] = cisName;
  }

  // Write back auto-generated tokens
  if (needWrite) {
    sheet.getRange(2, 1, data.length, 3).setValues(data);
  }

  return map;
}

/**
 * Retourne la liste des CIS avec leurs URLs pour la page admin.
 * Utilisé pour générer/consulter les liens à donner aux chefs de centre.
 */
function getCisViewLinks() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CIS_MAILING);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Ensure tokens exist
  getCisTokenMap_();

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var baseUrl = ScriptApp.getService().getUrl();
  var result = [];

  for (var i = 0; i < data.length; i++) {
    var cisName = (data[i][0] || '').toString().trim();
    var email   = (data[i][1] || '').toString().trim();
    var token   = (data[i][2] || '').toString().trim();
    if (!cisName || !token) continue;

    result.push({
      cis: cisName,
      email: email,
      token: token,
      url: baseUrl + '?cisToken=' + encodeURIComponent(token)
    });
  }

  return result;
}

/**
 * Données pour la vue chef de centre.
 * Retourne { cisName, agents: [...], error: null } ou { error: "..." }
 */
function getCisViewData(token) {
  token = (token || '').toString().trim();
  if (!token) return { error: 'Token manquant' };

  var tokenMap = getCisTokenMap_();
  var cisName = tokenMap[token];
  if (!cisName) return { error: 'Lien invalide ou expiré' };

  var allAgents   = getAllAgents();
  var sportData   = getSportData_();
  var permisData  = getPermisData_();
  var inactifs    = getInactiveMatricules_();

  // Filtrer les agents du CIS (principal OU secondaire), exclure inactifs
  var cisAgents = allAgents.filter(function (a) {
    if (inactifs[a.matricule]) return false;
    return a.centrePrincipal === cisName || a.centreSecondaire === cisName;
  });

  // Enrichir avec ICP et permis C
  var now = new Date().getTime();
  var agents = cisAgents.map(function (a) {
    // Dernière ICP = date la plus récente parmi toutes les épreuves sportives
    var sport = sportData[a.matricule] || [];
    var latestIcpRaw = null;
    for (var i = 0; i < sport.length; i++) {
      if (sport[i].dateRaw && (!latestIcpRaw || sport[i].dateRaw > latestIcpRaw)) {
        latestIcpRaw = sport[i].dateRaw;
      }
    }

    var permis = permisData[a.matricule];

    return {
      nomPrenom:            a.nomPrenom,
      centrePrincipal:      a.centrePrincipal,
      centreSecondaire:     a.centreSecondaire,
      datePerteCompetence:  a.datePerteCompetence,
      datePerteCompetenceRaw: a.datePerteCompetenceRaw,
      typeVisite:           a.typeVisite,
      dateIcp:              latestIcpRaw ? formatDate_(new Date(latestIcpRaw)) : '',
      dateIcpRaw:           latestIcpRaw,
      datePermisC:          permis ? permis.dateLimite : '',
      datePermisCRaw:       permis ? permis.dateLimiteRaw : null
    };
  });

  // Tri par date de perte de compétence (plus proche en premier)
  agents.sort(function (a, b) {
    if (!a.datePerteCompetenceRaw && !b.datePerteCompetenceRaw) return 0;
    if (!a.datePerteCompetenceRaw) return 1;
    if (!b.datePerteCompetenceRaw) return -1;
    return a.datePerteCompetenceRaw - b.datePerteCompetenceRaw;
  });

  return {
    cisName: cisName,
    agents: agents,
    error: null
  };
}