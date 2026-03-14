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
      relance3Raw:     (rel3 instanceof Date && !isNaN(rel3.getTime())) ? rel3.getTime() : null
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
    relance3Raw:     null
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
    // Dans données spécialité mais aucune règle VMA matchée
    // (ex. Grimp < 43) → on tombe dans les règles d'âge ci-dessous
  }

  /* ── Règles d'âge (tous les agents non-VMA) ── */
  var birthYear = agent.birthYear;
  if (!birthYear) return { type: 'Non déterminé', raison: 'Date de naissance inconnue' };

  var isBirthEven = birthYear % 2 === 0;
  var pariteLabel = isBirthEven ? 'paire' : 'impaire';

  if (agent.age >= CONFIG.AGE_THRESHOLD) {
    if (isBirthEven) {
      return { type: 'Visite médicale biennale', raison: 'Maintien activité ≥ ' + CONFIG.AGE_THRESHOLD + ' ans, né en année ' + pariteLabel + ' (' + birthYear + ')' };
    } else {
      return { type: 'Visite prévention', raison: 'Maintien activité ≥ ' + CONFIG.AGE_THRESHOLD + ' ans, né en année ' + pariteLabel + ' (' + birthYear + ')' };
    }
  } else {
    if (isBirthEven) {
      return { type: 'Visite médicale biennale', raison: 'Volontaire de -' + CONFIG.AGE_THRESHOLD + ' ans, né en année ' + pariteLabel + ' (' + birthYear + ')' };
    } else {
      return { type: 'Visite médicale 2027', raison: 'Volontaire de -' + CONFIG.AGE_THRESHOLD + ' ans, né en année ' + pariteLabel + ' (' + birthYear + ')' };
    }
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
    sportProgramAuthRequired: programInfo.authRequired
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
