# SDIS 66 — Suivi VMA

Application Google Apps Script de suivi des Visites Médicales d'Aptitude (VMA) par centre d'incendie et de secours.

## 📋 Fonctionnalités

- **Page par CIS** : chaque centre a sa propre URL avec le tableau de ses agents
- **Vue complète** : tous les agents triés par date de perte de compétence, avec recherche en temps réel
- **Type de visite automatique** : calcul du type de visite selon les spécialités et l'âge
- **Export PDF** : téléchargement du tableau pour impression
- **Gestion CIS / Mailing** : remplissage automatique de la liste des CIS dans l'onglet dédié

## 📊 Source de données

| Onglet | Description |
|--------|-------------|
| `Copie retard` | Agents dont la visite est en retard |
| `Copie a venir` | Agents dont la visite est à venir |
| `données spécialité` | Spécialités des agents (Bruleur, SAV, Grimp…) |
| `cis / mailing` | Liste des CIS + emails (col A auto-remplie) |

### Colonnes (Copie retard / Copie a venir)

| Col | Champ |
|-----|-------|
| A | Âge |
| B | Centre secondaire |
| C | Centre principal |
| D | Date de naissance |
| E | Date prochaine visite |
| F | Email |
| G | Matricule |
| H | NOM Prénom |
| I | Objet de visite |

> **Date de perte de compétence** = Date prochaine visite (col E) **+ 3 mois**

## 🏥 Règles de type de visite

### Agents dans « données spécialité »
| Condition | Type |
|-----------|------|
| Spécialité Bruleur, SAV, SAL ou caisson | VMA tous les ans |
| Spécialité Grimp ET âge ≥ 43 ans | VMA tous les ans |
| Spécialité diabétique | VMA tous les ans |

### Agents hors « données spécialité »
| Condition | Type |
|-----------|------|
| > 39 ans, né année paire, perte compétence année paire | Visite médicale biennale |
| > 39 ans, né année impaire, perte compétence année paire | Visite prévention |
| < 39 ans, visite prévue en 2026 | Visite médicale biennale |

## 🚀 Déploiement

### Prérequis
- [Node.js](https://nodejs.org/)
- [clasp](https://github.com/google/clasp) : `npm install -g @google/clasp`
- Être connecté : `clasp login`

### Push vers Google Apps Script
```bash
clasp push --force
```

### Déployer en webapp
```bash
clasp deploy
```

### Raccourci Windows
```bash
clasp-helper.bat push-deploy
```

## 🔗 IDs

- **Spreadsheet** : `1-6759nuMIn7A_ouAoALG-oHgQiJXI15EVezLPELdpUg`
- **Script** : `1-sst5KyL9hrxAi6GJApvhKgYGxXm-fwpT8N6qbM2EXo89k2Dz5yutiJ5`

## 📁 Structure

```
├── .clasp.json          # Configuration clasp (script ID)
├── appsscript.json      # Manifeste Google Apps Script
├── Config.js            # Configuration (IDs, colonnes, règles)
├── DataService.js       # Lecture données, logique métier
├── Code.js              # Point d'entrée (doGet, menu)
├── Index.html           # Interface web (SPA)
├── clasp-helper.bat     # Script helper Windows
└── README.md
```

## 🌐 URLs

Une fois déployé, l'application est accessible via :
- **Accueil** : `{URL_WEBAPP}`
- **CIS spécifique** : `{URL_WEBAPP}?cis=NomDuCIS`
- **Vue complète** : `{URL_WEBAPP}?cis=all`

## 📝 Menu Google Sheets

Le menu **🏥 Suivi VMA** apparaît dans Google Sheets avec :
- **Mettre à jour la liste des CIS** : remplit la colonne A de l'onglet `cis / mailing`
- **Ouvrir l'application web** : ouvre la webapp dans un nouvel onglet
