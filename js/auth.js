// ═══════════════════════════════════════════════════════════════════
// auth.js — Authentification MSAL et identité utilisateur
// ───────────────────────────────────────────────────────────────────
// Responsabilités :
//   - Initialiser MSAL
//   - Connexion (popup ou silent)
//   - Récupération du token Graph
//   - Exposer l'identité réelle de l'utilisateur (email Microsoft)
//
// Dépend de : config.js
// ═══════════════════════════════════════════════════════════════════

const Auth = (function() {
  let _msal = null;
  let _account = null;

  // ─── Initialisation MSAL ──────────────────────────────────────────
  function init() {
    if (_msal) return _msal;

    _msal = new msal.PublicClientApplication({
      auth: {
        clientId: AZURE.clientId,
        authority: 'https://login.microsoftonline.com/' + AZURE.tenantId,
        redirectUri: AZURE.redirectUri
      },
      cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false
      }
    });

    // Réutiliser un compte déjà connecté si présent
    const accounts = _msal.getAllAccounts();
    if (accounts.length > 0) {
      _account = accounts[0];
      if (DEBUG.enabled) console.log('[auth] Compte existant détecté:', _account.username);
    }

    return _msal;
  }

  // ─── Connexion (silent + popup en fallback) ───────────────────────
  // Retourne une promesse qui résout avec l'objet account
  async function login() {
    init();

    // 1. Tentative silencieuse si un compte existe
    if (_account) {
      try {
        await _msal.acquireTokenSilent({
          scopes: GRAPH_SCOPES,
          account: _account
        });
        if (DEBUG.enabled) console.log('[auth] Login silent OK');
        return _account;
      } catch (e) {
        if (DEBUG.enabled) console.warn('[auth] Login silent failed, popup needed:', e.errorCode);
      }
    }

    // 2. Popup login
    try {
      const result = await _msal.loginPopup({
        scopes: GRAPH_SCOPES,
        prompt: 'select_account'
      });
      _account = result.account;
      if (DEBUG.enabled) console.log('[auth] Login popup OK:', _account.username);
      return _account;
    } catch (e) {
      console.error('[auth] Login popup failed:', e);
      throw new Error('Authentification échouée: ' + (e.errorCode || e.message));
    }
  }

  // ─── Récupérer un token Graph (silent en priorité, popup si besoin) ─
  async function getToken() {
    if (!_account) {
      throw new Error('Pas de compte connecté. Appeler login() d\'abord.');
    }

    try {
      const result = await _msal.acquireTokenSilent({
        scopes: GRAPH_SCOPES,
        account: _account
      });
      return result.accessToken;
    } catch (e) {
      if (DEBUG.enabled) console.warn('[auth] Token silent failed, popup:', e.errorCode);
      try {
        const result = await _msal.acquireTokenPopup({ scopes: GRAPH_SCOPES });
        return result.accessToken;
      } catch (e2) {
        console.error('[auth] Token popup failed:', e2);
        throw new Error('Impossible d\'obtenir un token Graph: ' + (e2.errorCode || e2.message));
      }
    }
  }

  // ─── Identité utilisateur ──────────────────────────────────────────
  function getUser() {
    if (!_account) return null;
    return {
      email: _account.username,                      // ex: vincent.fournier@lutecia.care
      name:  _account.name || _account.username,    // ex: Vincent Fournier
      id:    _account.localAccountId
    };
  }

  // ─── Déconnexion ──────────────────────────────────────────────────
  async function logout() {
    if (!_msal || !_account) return;
    try {
      await _msal.logoutPopup({ account: _account });
      _account = null;
      if (DEBUG.enabled) console.log('[auth] Logout OK');
    } catch (e) {
      console.error('[auth] Logout failed:', e);
    }
  }

  // ─── Accès interne pour debug ─────────────────────────────────────
  function _getAccount() {
    return _account;
  }

  // ─── API publique ─────────────────────────────────────────────────
  return {
    init: init,
    login: login,
    getToken: getToken,
    getUser: getUser,
    logout: logout,
    _getAccount: _getAccount
  };
})();
