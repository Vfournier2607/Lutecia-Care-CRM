// ═══════════════════════════════════════════════════════════════════
// graph.js — Couche réseau Microsoft Graph API
// ───────────────────────────────────────────────────────────────────
// Responsabilités :
//   - Résoudre le siteId SharePoint à la volée
//   - Lire un onglet Excel complet (usedRange)
//   - Écrire une cellule
//   - Écrire une ligne complète (append)
//   - Mettre à jour des cellules en batch
//
// Dépend de : config.js, auth.js
// ═══════════════════════════════════════════════════════════════════

const Graph = (function() {
  const BASE = 'https://graph.microsoft.com/v1.0';
  let _siteId = null;

  // ─── Fetch wrapper avec auth automatique ──────────────────────────
  async function _fetch(url, options = {}) {
    const token = await Auth.getToken();
    const headers = Object.assign({
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json'
    }, options.headers || {});

    if (DEBUG.logGraphCalls) {
      console.log('[graph] ' + (options.method || 'GET') + ' ' + url);
    }

    const response = await fetch(url, { ...options, headers });
    if (!response.ok) {
      const text = await response.text();
      const err = new Error('Graph ' + response.status + ': ' + text.slice(0, 300));
      err.status = response.status;
      err.url = url;
      console.error('[graph] ERREUR', err);
      throw err;
    }

    // Certaines réponses (PATCH cellule) renvoient du contenu, d'autres pas
    const contentType = response.headers.get('content-type') || '';
    if (contentType.includes('application/json')) {
      return await response.json();
    }
    return null;
  }

  // ─── Résolution du siteId (cache après 1er appel) ─────────────────
  async function getSiteId() {
    if (_siteId) return _siteId;
    const url = BASE + '/sites/' + SHAREPOINT.hostname + ':' + SHAREPOINT.sitePath;
    const data = await _fetch(url);
    _siteId = data.id;
    if (DEBUG.enabled) console.log('[graph] siteId résolu:', _siteId);
    return _siteId;
  }

  // ─── Construction d'URL workbook ──────────────────────────────────
  // siteId est de la forme "hostname,siteCollectionId,siteId"
  // On utilise /drive/items/{fileId} pour accéder à un fichier
  async function _workbookUrl(fileId) {
    const siteId = await getSiteId();
    return BASE + '/sites/' + siteId + '/drive/items/' + fileId + '/workbook';
  }

  // ─── Lecture d'un onglet complet (usedRange) ──────────────────────
  // Retourne un tableau 2D : data[row][col]
  async function readSheet(fileId, sheetName) {
    const wbUrl = await _workbookUrl(fileId);
    const url = wbUrl + "/worksheets('" + encodeURIComponent(sheetName) + "')/usedRange?$select=values";
    const data = await _fetch(url);
    return data.values || [];
  }

  // ─── Lecture du dimensions (sans charger les valeurs) ─────────────
  // Utile pour connaître la taille d'un onglet avant lecture
  async function getSheetInfo(fileId, sheetName) {
    const wbUrl = await _workbookUrl(fileId);
    const url = wbUrl + "/worksheets('" + encodeURIComponent(sheetName) + "')/usedRange?$select=address,rowCount,columnCount";
    return await _fetch(url);
  }

  // ─── Écrire une cellule (1 cellule) ───────────────────────────────
  // address en notation Excel : "A5", "Z42", etc.
  async function writeCell(fileId, sheetName, address, value) {
    const wbUrl = await _workbookUrl(fileId);
    const url = wbUrl + "/worksheets('" + encodeURIComponent(sheetName) + "')/range(address='" + address + "')";
    return await _fetch(url, {
      method: 'PATCH',
      body: JSON.stringify({ values: [[value]] })
    });
  }

  // ─── Écrire plusieurs cellules d'une ligne (batch sur une range) ──
  // address = "A5:Z5" (par exemple), values = ["val1", "val2", ...]
  async function writeRange(fileId, sheetName, address, values) {
    const wbUrl = await _workbookUrl(fileId);
    const url = wbUrl + "/worksheets('" + encodeURIComponent(sheetName) + "')/range(address='" + address + "')";
    // values doit être un tableau 2D
    const values2D = Array.isArray(values[0]) ? values : [values];
    return await _fetch(url, {
      method: 'PATCH',
      body: JSON.stringify({ values: values2D })
    });
  }

  // ─── Ajouter une ligne à la fin (append) ──────────────────────────
  // values = tableau de cellules (1D) qui sera ajouté en fin
  async function appendRow(fileId, sheetName, values) {
    const wbUrl = await _workbookUrl(fileId);
    // Récupérer le tableau ou usedRange pour connaître la prochaine ligne
    // Méthode : on lit l'address de usedRange, on calcule la ligne suivante
    const info = await getSheetInfo(fileId, sheetName);
    // info.address = "Master Screening!A1:BG1571"
    const addressOnly = info.address.split('!')[1];
    const match = addressOnly.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) throw new Error('Impossible de parser l\'address: ' + info.address);
    const startCol = match[1];
    const endCol = match[3];
    const endRow = parseInt(match[4], 10);
    const newRow = endRow + 1;
    const range = startCol + newRow + ':' + endCol + newRow;
    return await writeRange(fileId, sheetName, range, [values]);
  }

  // ─── Supprimer une ligne (range entière) ──────────────────────────
  // Utilisé pour move-to-archive : suppression de la ligne du master
  async function deleteRow(fileId, sheetName, rowNumber) {
    const wbUrl = await _workbookUrl(fileId);
    const url = wbUrl + "/worksheets('" + encodeURIComponent(sheetName) + "')/range(address='" + rowNumber + ':' + rowNumber + "')/delete";
    return await _fetch(url, {
      method: 'POST',
      body: JSON.stringify({ shift: 'Up' })
    });
  }

  // ─── Helper : convertir un index de colonne (0-based) en lettre Excel ─
  // 0 → A, 25 → Z, 26 → AA, etc.
  function colIndexToLetter(idx) {
    let letter = '';
    let n = idx;
    while (n >= 0) {
      letter = String.fromCharCode(65 + (n % 26)) + letter;
      n = Math.floor(n / 26) - 1;
    }
    return letter;
  }

  // ─── Helper : convertir une lettre Excel en index (0-based) ───────
  function letterToColIndex(letter) {
    let idx = 0;
    for (let i = 0; i < letter.length; i++) {
      idx = idx * 26 + (letter.charCodeAt(i) - 64);
    }
    return idx - 1;
  }

  // ─── API publique ─────────────────────────────────────────────────
  return {
    getSiteId: getSiteId,
    readSheet: readSheet,
    getSheetInfo: getSheetInfo,
    writeCell: writeCell,
    writeRange: writeRange,
    appendRow: appendRow,
    deleteRow: deleteRow,
    colIndexToLetter: colIndexToLetter,
    letterToColIndex: letterToColIndex
  };
})();
