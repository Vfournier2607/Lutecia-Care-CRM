// graph.js — Couche réseau Microsoft Graph
// Dépend de : config.js, auth.js

let GRAPH_READY = false;

const Graph = (function() {
  const BASE = 'https://graph.microsoft.com/v1.0';
  let _siteId = null;

  async function _fetch(method, endpoint, body) {
    const token = await Auth.getToken();
    if (DEBUG.logGraphCalls) console.log('[graph]', method, endpoint);
    const res = await fetch(BASE + endpoint, {
      method,
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: body !== undefined ? JSON.stringify(body) : undefined
    });
    if (res.status === 204) return null;
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error ? err.error.message : 'Graph ' + res.status);
    }
    return res.json();
  }

  async function _siteIdGet() {
    if (_siteId) return _siteId;
    const data = await _fetch('GET', '/sites/' + SHAREPOINT.hostname + ':' + SHAREPOINT.sitePath);
    _siteId = data.id;
    return _siteId;
  }

  async function _wbUrl() {
    return '/sites/' + (await _siteIdGet()) + '/drive/items/' + OUTPUT_FILE_ID + '/workbook';
  }

  async function readSheet(sheetName) {
    try {
      const data = await _fetch('GET', (await _wbUrl()) + '/worksheets/' + encodeURIComponent(sheetName) + '/usedRange');
      if (!data || !data.values || data.values.length < 2) return [];
      const h = data.values[0];
      return data.values.slice(1)
        .map(row => { const o = {}; h.forEach((k,i) => { o[k] = row[i] !== null ? String(row[i]) : ''; }); return o; })
        .filter(r => Object.values(r).some(v => v !== ''));
    } catch(e) { if (DEBUG.enabled) console.warn('[graph] readSheet', sheetName, e.message); return []; }
  }

  async function writeSheet(sheetName, headers, rows) {
    const vals = [headers].concat(rows.map(r => headers.map(h => r[h] !== undefined ? String(r[h]) : '')));
    const lc   = String.fromCharCode(64 + headers.length);
    await _fetch('PATCH',
      (await _wbUrl()) + '/worksheets/' + encodeURIComponent(sheetName) + "/range(address='A1:" + lc + vals.length + "')",
      { values: vals });
  }

  async function _ensureSheet(sheetName, headers) {
    const wb = await _wbUrl();
    const sheets = await _fetch('GET', wb + '/worksheets');
    if (!sheets.value.some(s => s.name === sheetName)) await _fetch('POST', wb + '/worksheets', { name: sheetName });
    try {
      const data = await _fetch('GET', wb + '/worksheets/' + encodeURIComponent(sheetName) + '/usedRange');
      if (!data || !data.values || !data.values.length) {
        const lc = String.fromCharCode(64 + headers.length);
        await _fetch('PATCH', wb + '/worksheets/' + encodeURIComponent(sheetName) + "/range(address='A1:" + lc + "1')", { values: [headers] });
      }
    } catch(_) {}
  }

  async function initSheets() {
    await _ensureSheet('Journal', JOURNAL_HEADERS);
    await _ensureSheet('Etat',    ETAT_HEADERS);
    await _ensureSheet('KPIs',    KPI_HEADERS);
    await _ensureSheet('Actions', ACTIONS_HEADERS);
    GRAPH_READY = true;
  }

  async function loadEtat() {
    const out = {}; (await readSheet('Etat')).forEach(r => { if (r.id_entite) out[r.id_entite] = r; }); return out;
  }

  async function updateEtat(id, data) {
    const rows = await readSheet('Etat'), idx = rows.findIndex(r => r.id_entite === id), user = Auth.getUser();
    const row  = Object.assign({}, idx >= 0 ? rows[idx] : {}, data, { id_entite: id, modifiedBy: user?user.name:'', modifiedAt: new Date().toISOString() });
    if (idx >= 0) rows[idx] = row; else rows.push(row);
    await writeSheet('Etat', ETAT_HEADERS, rows);
  }

  async function loadActions(id) { return (await readSheet('Actions')).filter(r => r.id_entite === id); }

  async function saveAction(id, action) {
    const rows = await readSheet('Actions'); rows.push(Object.assign({ id_entite: id }, action));
    await writeSheet('Actions', ACTIONS_HEADERS, rows);
  }

  async function loadKpi(id) {
    const row = (await readSheet('KPIs')).find(r => r.id_entite === id);
    if (!row) return null;
    try { return { lignes: JSON.parse(row.lignes||'[]'), valeurs: JSON.parse(row.valeurs||'{}'), annees: JSON.parse(row.annees||'[]') }; }
    catch(_) { return null; }
  }

  async function saveKpi(id, k) {
    const rows = await readSheet('KPIs'), idx = rows.findIndex(r => r.id_entite === id), user = Auth.getUser();
    const row = { id_entite: id, lignes: JSON.stringify(k.lignes), valeurs: JSON.stringify(k.valeurs), annees: JSON.stringify(k.annees||[]), modifiedBy: user?user.name:'', modifiedAt: new Date().toISOString() };
    if (idx >= 0) rows[idx] = row; else rows.push(row);
    await writeSheet('KPIs', KPI_HEADERS, rows);
  }

  async function appendJournal(ligne) {
    const rows = await readSheet('Journal'); rows.push(ligne); await writeSheet('Journal', JOURNAL_HEADERS, rows);
  }

  return { readSheet, writeSheet, initSheets, loadEtat, updateEtat, loadActions, saveAction, loadKpi, saveKpi, appendJournal };
})();

async function graphConnect() {
  try { await Auth.login(); await Graph.initSheets(); return true; }
  catch(e) { if (DEBUG.enabled) console.warn('[graph]', e.message); GRAPH_READY = false; return false; }
}

function graphGetUserName()            { const u = Auth.getUser(); return u ? u.name : ''; }
async function graphLoadEtat()         { return Graph.loadEtat(); }
async function graphUpdateEtat(id, d)  { return Graph.updateEtat(id, d); }
async function graphLoadActions(id)    { return Graph.loadActions(id); }
async function graphSaveAction(id, a)  { return Graph.saveAction(id, a); }
async function graphLoadKpi(id)        { return Graph.loadKpi(id); }
async function graphSaveKpi(id, k)     { return Graph.saveKpi(id, k); }
async function graphAppendJournal(l)   { return Graph.appendJournal(l); }
