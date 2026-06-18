// ═══════════════════════════════════════════════════════════════════
// graph.js — Couche réseau Microsoft Graph + initialisation MSAL
// Dépend de : config.js, auth.js (chargés avant)
// ═══════════════════════════════════════════════════════════════════

// Indicateur lisible par le code applicatif
let GRAPH_READY = false;

const Graph = (function() {
  const BASE = 'https://graph.microsoft.com/v1.0';
  let _siteId = null;

  // ── Appel HTTP avec token Bearer ────────────────────────────────
  async function _fetch(method, endpoint, body) {
    const token = await Auth.getToken();
    if (DEBUG.logGraphCalls) console.log('[graph]', method, endpoint);
    const res = await fetch(BASE + endpoint, {
      method,
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type':  'application/json'
      },
      body: body !== undefined ? JSON.stringify(body) : undefined
    });
    if (res.status === 204) return null;
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error ? err.error.message : 'Graph ' + res.status);
    }
    return res.json();
  }

  // ── siteId SharePoint (mis en cache) ────────────────────────────
  async function _siteIdGet() {
    if (_siteId) return _siteId;
    const data = await _fetch('GET', '/sites/' + SHAREPOINT.hostname + ':' + SHAREPOINT.sitePath);
    _siteId = data.id;
    return _siteId;
  }

  // ── URL de base du workbook ─────────────────────────────────────
  async function _wbUrl() {
    const siteId = await _siteIdGet();
    return '/sites/' + siteId + '/drive/items/' + OUTPUT_FILE_ID + '/workbook';
  }

  // ── Lire un onglet ──────────────────────────────────────────────
  // Retourne un tableau d'objets {header: valeur}
  async function readSheet(sheetName) {
    try {
      const wb  = await _wbUrl();
      const data = await _fetch('GET', wb + '/worksheets/' + encodeURIComponent(sheetName) + '/usedRange');
      if (!data || !data.values || data.values.length < 2) return [];
      const headers = data.values[0];
      return data.values.slice(1)
        .map(row => {
          const obj = {};
          headers.forEach((h, i) => { obj[h] = row[i] !== null ? String(row[i]) : ''; });
          return obj;
        })
        .filter(r => Object.values(r).some(v => v !== ''));
    } catch(e) {
      if (DEBUG.enabled) console.warn('[graph] readSheet', sheetName, e.message);
      return [];
    }
  }

  // ── Écrire un onglet en entier ──────────────────────────────────
  // Vide d'abord le contenu existant pour éviter les lignes orphelines.
  async function writeSheet(sheetName, headers, rows) {
    const wb      = await _wbUrl();
    const wsBase  = wb + '/worksheets/' + encodeURIComponent(sheetName);
    // 1. Vider le contenu de la plage utilisée (sans supprimer les cellules)
    try {
      await _fetch('POST', wsBase + '/usedRange/clear', { applyTo: 'Contents' });
    } catch(_) { /* feuille vide : normal */ }
    // 2. Écrire les nouvelles données
    const values  = [headers].concat(rows.map(r => headers.map(h => r[h] !== undefined ? String(r[h]) : '')));
    const lastCol = String.fromCharCode(64 + headers.length);
    const lastRow = values.length;
    await _fetch('PATCH',
      wsBase + '/range(address=\'A1:' + lastCol + lastRow + '\')',
      { values }
    );
  }

  // ── Créer un onglet si absent et poser les en-têtes ─────────────
  async function _ensureSheet(sheetName, headers) {
    const wb     = await _wbUrl();
    const sheets = await _fetch('GET', wb + '/worksheets');
    const exists = sheets.value.some(s => s.name === sheetName);
    if (!exists) {
      await _fetch('POST', wb + '/worksheets', { name: sheetName });
    }
    // Vérifier si la ligne 1 est vide
    try {
      const data = await _fetch('GET', wb + '/worksheets/' + encodeURIComponent(sheetName) + '/usedRange');
      if (!data || !data.values || data.values.length === 0) {
        const lastCol = String.fromCharCode(64 + headers.length);
        await _fetch('PATCH',
          wb + '/worksheets/' + encodeURIComponent(sheetName) + '/range(address=\'A1:' + lastCol + '1\')',
          { values: [headers] }
        );
      }
    } catch(_) { /* onglet vide, en-têtes déjà posés ci-dessus */ }
  }

  // ── Initialiser les onglets nécessaires ─────────────────────────
  async function initSheets() {
    await _ensureSheet('Journal', JOURNAL_HEADERS);
    await _ensureSheet('Etat',    ETAT_HEADERS);
    await _ensureSheet('KPIs',    KPI_HEADERS);
    await _ensureSheet('Actions', ACTIONS_HEADERS);
    GRAPH_READY = true;
  }

  // ════════════════════════════════════════════════════════════════
  // API publique
  // ════════════════════════════════════════════════════════════════

  async function loadEtat() {
    const rows = await readSheet('Etat');
    const out  = {};
    rows.forEach(r => { if (r.id_entite) out[r.id_entite] = r; });
    return out;
  }

  async function updateEtat(id, data) {
    const rows = await readSheet('Etat');
    const idx  = rows.findIndex(r => r.id_entite === id);
    const user = Auth.getUser();
    const row  = Object.assign({}, idx >= 0 ? rows[idx] : {}, data, {
      id_entite:  id,
      modifiedBy: user ? user.name : '',
      modifiedAt: new Date().toISOString()
    });
    if (idx >= 0) rows[idx] = row;
    else          rows.push(row);
    await writeSheet('Etat', ETAT_HEADERS, rows);
  }

  async function loadActions(id) {
    const rows = await readSheet('Actions');
    return rows.filter(r => r.id_entite === id);
  }

  async function saveAction(id, action) {
    const rows = await readSheet('Actions');
    rows.push(Object.assign({ id_entite: id }, action));
    await writeSheet('Actions', ACTIONS_HEADERS, rows);
  }

  async function loadKpi(id) {
    const rows = await readSheet('KPIs');
    const row  = rows.find(r => r.id_entite === id);
    if (!row) return null;
    try {
      return {
        lignes:  JSON.parse(row.lignes  || '[]'),
        valeurs: JSON.parse(row.valeurs || '{}'),
        annees:  JSON.parse(row.annees  || '[]')
      };
    } catch(_) { return null; }
  }

  async function saveKpi(id, kpiData) {
    const rows = await readSheet('KPIs');
    const idx  = rows.findIndex(r => r.id_entite === id);
    const user = Auth.getUser();
    const row  = {
      id_entite:  id,
      lignes:     JSON.stringify(kpiData.lignes),
      valeurs:    JSON.stringify(kpiData.valeurs),
      annees:     JSON.stringify(kpiData.annees || []),
      modifiedBy: user ? user.name : '',
      modifiedAt: new Date().toISOString()
    };
    if (idx >= 0) rows[idx] = row;
    else          rows.push(row);
    await writeSheet('KPIs', KPI_HEADERS, rows);
  }

  async function appendJournal(ligne) {
    const rows = await readSheet('Journal');
    rows.push(ligne);
    await writeSheet('Journal', JOURNAL_HEADERS, rows);
  }

  return { readSheet, writeSheet, initSheets, loadEtat, updateEtat, loadActions, saveAction, loadKpi, saveKpi, appendJournal };
})();

// ── Connexion principale appelée au démarrage de chaque page ────────
// Lance le login MSAL (silent puis popup si nécessaire) et initialise les onglets.
async function graphConnect() {
  try {
    await Auth.login();
    await Graph.initSheets();
    return true;
  } catch(e) {
    if (DEBUG.enabled) console.warn('[graph] connect:', e.message);
    GRAPH_READY = false;
    return false;
  }
}

// Raccourcis globaux (compatibilité avec le code applicatif existant)
function graphGetUserName() {
  const u = Auth.getUser();
  return u ? u.name : '';
}
async function graphLoadEtat()         { return Graph.loadEtat(); }
async function graphUpdateEtat(id, d)  { return Graph.updateEtat(id, d); }
async function graphLoadActions(id)    { return Graph.loadActions(id); }
async function graphSaveAction(id, a)  { return Graph.saveAction(id, a); }
async function graphLoadKpi(id)        { return Graph.loadKpi(id); }
async function graphSaveKpi(id, k)     { return Graph.saveKpi(id, k); }
async function graphAppendJournal(l)   { return Graph.appendJournal(l); }
