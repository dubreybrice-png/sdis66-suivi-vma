/**
 * SDIS 66 — Suivi VMA
 * Configuration globale
 */

var CONFIG = {
  SPREADSHEET_ID: '1-6759nuMIn7A_ouAoALG-oHgQiJXI15EVezLPELdpUg',

  /* ── Noms des onglets ── */
  SHEETS: {
    RETARD:       'Copie retard',
    A_VENIR:      'Copie a venir',
    CIS_MAILING:  'cis / mailing',
    SPECIALITE:   'données spécialité'
  },

  /* ── Colonnes onglets Copie retard / Copie a venir (0-based) ── */
  COLS: {
    AGE:                0,  // A
    CENTRE_SECONDAIRE:  1,  // B
    CENTRE_PRINCIPAL:   2,  // C
    DATE_NAISSANCE:     3,  // D
    DATE_VISITE:        4,  // E  (date prochaine visite)
    EMAIL:              5,  // F
    MATRICULE:          6,  // G
    NOM_PRENOM:         7,  // H
    OBJET_VISITE:       8   // I
  },

  /* ── Colonnes onglet données spécialité (0-based) ── */
  COLS_SPE: {
    TYPE: 1,  // B — type de spécialité
    NOM:  2   // C — NOM Prénom
  },

  /* ── Règles métier ── */
  VMA_SPECIALTIES: ['Bruleur', 'SAV', 'SAL', 'caisson'],
  VMA_GRIMP_AGE:   43,
  AGE_THRESHOLD:   39,
  MONTHS_TO_ADD:   3,
  REFERENCE_YEAR:  2026
};
