// ═══════════════════════════════════════════════════════════════════
// config.js — Constantes du projet Lutecia Care Deal Flow
// ───────────────────────────────────────────────────────────────────
// Ce fichier ne contient QUE des constantes.
// Aucune logique, aucun appel API, aucun DOM.
// ═══════════════════════════════════════════════════════════════════

// ─── Identifiants Azure AD ──────────────────────────────────────────
const AZURE = {
  clientId: '8e1ee67e-a7d8-4362-bff0-e08f806e5404',
  tenantId: '968cbaab-6a7a-4e30-a00f-af1805e561b6',
  redirectUri: window.location.origin
};

// ─── Scopes Microsoft Graph nécessaires ─────────────────────────────
const GRAPH_SCOPES = [
  'Files.ReadWrite.All',
  'Sites.ReadWrite.All',
  'User.Read'
];

// ─── SharePoint ─────────────────────────────────────────────────────
const SHAREPOINT = {
  hostname: 'lutetiacare.sharepoint.com',
  sitePath: '/sites/MALutecia'
};

// ─── Identifiants des fichiers Excel sur SharePoint ─────────────────
const FILE_IDS = {
  MASTER:   '47BBEE2C-785E-4822-B90D-B4E50447B395',
  ACTIONS:  '3C19EF7D-EB87-41E7-B053-DC389A200BD0',
  ARCHIVES: '53C9D41F-84A8-465F-BAC3-573330E7594C'
};

// ─── Onglets attendus dans les fichiers ─────────────────────────────
const SHEET_NAMES = {
  MASTER:   'Master Screening',
  ACTIONS:  'Actions',
  ARCHIVES: 'Archives'
};

// ─── Schéma attendu du MASTER (validation au démarrage) ─────────────
// Headers attendus dans l'ordre. Le code accepte des colonnes en plus
// (warning console) mais refuse si une attendue manque (erreur bloquante).
const MASTER_SCHEMA = {
  // Position de la ligne des headers dans l'onglet (1-indexé)
  headerRow: 4,
  // Position de la première ligne de données (1-indexé)
  firstDataRow: 5,
  // Headers obligatoires
  requiredColumns: [
    'id_centre',
    'nom_centre',
    'mapping',
    'adresse',
    'pays',
    'nb_centres',
    'appartenance_reseau',
    'id_structure',
    'nom_reseau',
    'label_deal',
    'statut_deal',
    'responsable_LC',
    'contact_deal',
    'responsabilite',
    'coordonnes_contact',
    'tep',
    'scintigraphie',
    'nb_tep',
    'nb_gamma',
    'contact_url',
    // KPIs
    'CA_2022','CA_2023','CA_2024','CA_2025','CA_2026','CA_2027','CA_2028','CA_2029','CA_2030',
    'Rlnt_2022','Rlnt_2023','Rlnt_2024','Rlnt_2025','Rlnt_2026','Rlnt_2027','Rlnt_2028','Rlnt_2029','Rlnt_2030',
    'Actes_2022','Actes_2023','Actes_2024','Actes_2025','Actes_2026','Actes_2027','Actes_2028','Actes_2029','Actes_2030',
    'EBIDTA_2022','EBIDTA_2023','EBIDTA_2024','EBIDTA_2025','EBIDTA_2026','EBIDTA_2027','EBIDTA_2028','EBIDTA_2029','EBIDTA_2030',
    // Métadonnées
    'archived',
    'modified_by',
    'modified_at'
  ]
};

const ACTIONS_SCHEMA = {
  headerRow: 1,
  firstDataRow: 2,
  requiredColumns: [
    'id_action',
    'id_centre',
    'id_structure',
    'type',
    'date',
    'echeance',
    'note',
    'realise',
    'auteur',
    'created_at',
    'modified_by',
    'modified_at',
    'archived'
  ]
};

const ARCHIVES_SCHEMA = {
  headerRow: 1,
  firstDataRow: 2,
  // Toutes les colonnes du master + 2 colonnes archive
  // (validation seulement sur la présence de ces 2 dernières)
  requiredArchiveColumns: ['archived_by', 'archived_at']
};

// ─── Valeurs autorisées pour les champs énumérés ────────────────────
const ENUMS = {
  statut_deal: ['', 'contact', 'discussion', 'advanced', 'loi'],
  tep: ['Oui', 'Non', 'N/D'],
  scintigraphie: ['Oui', 'Non', 'N/D'],
  appartenance_reseau: ['Oui', 'Non', 'N/D'],
  archived: ['TRUE', 'FALSE']
};

// ─── KPIs : indicateurs et années (utilisé par les calculs) ─────────
const KPI_INDICATORS = ['CA', 'Rlnt', 'Actes', 'EBIDTA'];
const KPI_YEARS = [2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030];

// ─── Types d'actions autorisées ─────────────────────────────────────
const ACTION_TYPES = ['Appel', 'Email', 'Visio', 'Réunion présentiel', 'Alerte', 'Autre'];

// ─── Paramètres divers ──────────────────────────────────────────────
const APP_CONFIG = {
  // Nombre de lignes affichées par "page" dans le tableau front
  pageSize: 100,
  // Délai avant disparition de la notification "Enregistré"
  savedToastMs: 2000,
  // Format de date pour affichage UI (ISO côté Excel, FR côté UI)
  dateLocale: 'fr-FR'
};

// ─── Logs et debug ──────────────────────────────────────────────────
const DEBUG = {
  enabled: true,        // logs détaillés en console (à passer à false en prod stabilisée)
  logGraphCalls: true   // logger chaque appel Graph
};
