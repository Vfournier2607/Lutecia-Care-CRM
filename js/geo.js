// ═══════════════════════════════════════════════════════════════════
// geo.js — Géocodage des centres (adresse → lat/lng)
// Nominatim (OpenStreetMap), 1 requête/seconde maximum.
// Cache : localStorage 'lutecia_geo' + onglet Geo de output.xlsx
// (un centre n'est géocodé qu'une seule fois pour toute l'équipe).
// Dépend de : config.js, graph.js (facultatif — marche sans sync)
// ═══════════════════════════════════════════════════════════════════

var Geo = (function() {
  var STORAGE_GEO = 'lutecia_geo';
  var NOMINATIM   = 'https://nominatim.openstreetmap.org/search';
  var _dirty      = false; // vrai si de nouvelles coordonnées doivent partir vers Graph

  function _load() {
    try { return JSON.parse(localStorage.getItem(STORAGE_GEO) || '{}'); }
    catch(e) { return {}; }
  }
  function _save(cache) {
    try { localStorage.setItem(STORAGE_GEO, JSON.stringify(cache)); } catch(e) {}
  }

  // Clé stable d'un centre pour le cache géo
  function cle(c) {
    return (c.pays + '|' + c.nom).toLowerCase().replace(/[^a-z0-9|]/g, '_');
  }

  // Géocode une adresse via Nominatim. Renvoie {lat, lng} ou null.
  async function _nominatim(query, pays) {
    var url = NOMINATIM + '?format=json&limit=1&q=' + encodeURIComponent(query)
            + (pays ? '&countrycodes=' + (PAYS_CODES[pays] || '') : '');
    var res = await fetch(url, { headers: { 'Accept': 'application/json' } });
    if (!res.ok) return null;
    var data = await res.json();
    if (!data.length) return null;
    return { lat: parseFloat(data[0].lat), lng: parseFloat(data[0].lon) };
  }

  // Codes ISO pour aider Nominatim à rester dans le bon pays
  var PAYS_CODES = {
    'France': 'fr', 'Italie': 'it', 'Italy': 'it', 'Allemagne': 'de', 'Germany': 'de',
    'Espagne': 'es', 'Spain': 'es', 'Portugal': 'pt', 'Belgique': 'be', 'Belgium': 'be',
    'Suisse': 'ch', 'Switzerland': 'ch', 'Autriche': 'at', 'Austria': 'at',
    'Pays-bas': 'nl', 'Netherlands': 'nl', 'Royaume-uni': 'gb', 'Uk': 'gb',
    'Pologne': 'pl', 'Poland': 'pl', 'Grece': 'gr', 'Grèce': 'gr', 'Greece': 'gr'
  };

  // Géocode un centre : adresse complète, sinon ville + CP, sinon ville seule.
  async function _geocodeCentre(c) {
    var essais = [];
    if (c.adresse) essais.push(c.adresse + ', ' + (c.ville || '') + ', ' + c.pays);
    if (c.cp || c.ville) essais.push((c.cp || '') + ' ' + (c.ville || '') + ', ' + c.pays);
    if (c.ville) essais.push(c.ville + ', ' + c.pays);
    for (var i = 0; i < essais.length; i++) {
      try {
        var pos = await _nominatim(essais[i], c.pays);
        if (pos) return pos;
      } catch(e) { /* réseau : essai suivant */ }
    }
    return null;
  }

  // Récupère les coordonnées connues (cache local). Renvoie {cle: {lat,lng}}
  function connues() { return _load(); }

  // Coordonnées pré-calculées livrées avec l'app (niveau ville).
  // Fusionnées dans le cache local (sans écraser un géocodage précis existant).
  var _seedPromise = null;
  function chargerSeed() {
    if (_seedPromise) return _seedPromise;
    _seedPromise = fetch('data/geo.json')
      .then(function(r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
      .then(function(seed) {
        var cache = _load();
        var nouveau = false;
        Object.keys(seed).forEach(function(k) {
          if (!cache[k] || cache[k].lat === null) { cache[k] = seed[k]; nouveau = true; }
        });
        if (nouveau) _save(cache);
        return cache;
      })
      .catch(function(e) { console.warn('geo seed:', e.message); return _load(); });
    return _seedPromise;
  }

  // Géocode progressivement une liste de centres (1/s), avec callback de progression.
  // onProgress(fait, total) — onDone(cache)
  async function geocoder(centres, onProgress, onDone) {
    var cache = _load();
    var manquants = centres.filter(function(c) { return !cache[cle(c)]; });
    var total = manquants.length;
    if (!total) { if (onDone) onDone(cache); return cache; }
    for (var i = 0; i < manquants.length; i++) {
      var c = manquants[i];
      var pos = await _geocodeCentre(c);
      cache[cle(c)] = pos || { lat: null, lng: null }; // null mémorisé : ne pas retenter en boucle
      _dirty = true;
      _save(cache);
      if (onProgress) onProgress(i + 1, total);
      // Politesse Nominatim : 1 requête par seconde
      if (i < manquants.length - 1) await new Promise(function(r) { setTimeout(r, 1100); });
    }
    // Partager les nouvelles coordonnées avec l'équipe
    if (_dirty && window.GRAPH_READY) {
      try { await pousserVersGraph(); } catch(e) { console.warn('Geo sync:', e.message); }
    }
    if (onDone) onDone(cache);
    return cache;
  }

  // ── Sync SharePoint (onglet Geo : cle, lat, lng) ──────────────────
  async function pousserVersGraph() {
    var cache = _load();
    var rows = Object.keys(cache).map(function(k) {
      return { cle: k, lat: cache[k].lat !== null ? cache[k].lat : '', lng: cache[k].lng !== null ? cache[k].lng : '' };
    });
    await Graph.writeSheet('Geo', GEO_HEADERS, rows);
    _dirty = false;
  }

  async function chargerDepuisGraph() {
    var rows = await Graph.readSheet('Geo');
    if (!rows.length) return _load();
    var cache = _load();
    var nouveau = false;
    rows.forEach(function(r) {
      if (!r.cle || cache[r.cle]) return;
      cache[r.cle] = {
        lat: r.lat !== '' ? parseFloat(r.lat) : null,
        lng: r.lng !== '' ? parseFloat(r.lng) : null
      };
      nouveau = true;
    });
    if (nouveau) _save(cache);
    return cache;
  }

  return { cle: cle, connues: connues, chargerSeed: chargerSeed, geocoder: geocoder, pousserVersGraph: pousserVersGraph, chargerDepuisGraph: chargerDepuisGraph };
})();
