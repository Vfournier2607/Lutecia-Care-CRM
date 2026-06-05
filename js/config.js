// config.js — Constantes du projet Lutecia Care Deal Flow
// Source unique de vérité pour toute la configuration.

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

const OUTPUT_FILE_ID = '67B2C3FB-FFE6-442B-825D-C784B3223DF6';

const JOURNAL_HEADERS = [
  'timestamp', 'auteur', 'type_entite', 'id_entite',
  'nom_entite', 'champ_modifie', 'ancienne_valeur', 'nouvelle_valeur', 'pays'
];

const ETAT_HEADERS = [
  'id_entite', 'nom_entite', 'pays',
  'statut', 'label', 'responsable', 'contact', 'source', 'note',
  'modifiedBy', 'modifiedAt'
];

const KPI_HEADERS     = ['id_entite', 'lignes', 'valeurs', 'annees', 'modifiedBy', 'modifiedAt'];
const ACTIONS_HEADERS = ['id_entite', 'type', 'date', 'note', 'auteur', 'createdAt'];

const DEBUG = {
  enabled:      false,
  logGraphCalls: false
};
