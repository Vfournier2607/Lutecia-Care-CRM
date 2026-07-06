// ═══════════════════════════════════════════════════════════════════
// config.js — Constantes du projet Lutecia Care Deal Flow
// Source unique de vérité pour toute la configuration.
// ═══════════════════════════════════════════════════════════════════

const AZURE = {
  clientId:    '8e1ee67e-a7d8-4362-bff0-e08f806e5404',
  tenantId:    '968cbaab-6a7a-4e30-a00f-af1805e561b6',
  redirectUri: window.location.origin
};

const GRAPH_SCOPES = [
  'Files.ReadWrite.All',
  'Sites.ReadWrite.All',
  'User.Read'
];

const SHAREPOINT = {
  hostname: 'lutetiacare.sharepoint.com',
  sitePath: '/sites/MALutecia'
};

// Fichier Excel unique (output.xlsx) contenant tous les onglets
const OUTPUT_FILE_ID = '67B2C3FB-FFE6-442B-825D-C784B3223DF6';

// En-têtes des onglets
const JOURNAL_HEADERS = [
  'timestamp', 'auteur', 'type_entite', 'id_entite',
  'nom_entite', 'champ_modifie', 'ancienne_valeur', 'nouvelle_valeur', 'pays'
];

const ETAT_HEADERS = [
  'id_entite', 'type_entite', 'nom_entite', 'pays',
  'statut', 'label', 'responsable', 'contact', 'source', 'note', 'lien',
  'tep', 'gamma_cam', 'radiotherapie', 'nb_tep', 'nb_gamma',
  'notes_db', 'region', 'ville', 'adresse', 'nb_centres',
  'nb_machines', 'respcontact',
  'modifiedBy', 'modifiedAt'
];

const KPI_HEADERS     = ['id_entite', 'lignes', 'valeurs', 'annees', 'modifiedBy', 'modifiedAt'];

// Onglet Screening : données brutes importées, partagées entre les membres.
// Une ligne par centre, taguée par pays + métadonnées d'import.
const SCREENING_HEADERS = [
  'pays', 'filename', 'importedAt',
  'nom', 'type', 'reseau', 'nom_reseau',
  'tep', 'gamma_cam', 'nb_tep', 'nb_gamma',
  'adresse', 'cp', 'region', 'ville', 'mapping', 'contact', 'notes'
];
const ACTIONS_HEADERS = ['id_entite', 'type', 'date', 'note', 'auteur', 'createdAt', 'echeance', 'realise', 'modifiedBy', 'modifiedAt'];

const GEO_HEADERS = ['cle', 'lat', 'lng'];

// Screenings intégrés à l'application (data/screenings.json).
// Incrémenter la version pour remplacer les données chez tous les utilisateurs
// (local + SharePoint) lors de l'ajout d'un pays ou d'une mise à jour.
const SEED_VERSION = '2026-07-06-v3';

const DEBUG = {
  enabled:      false,
  logGraphCalls: false
};
