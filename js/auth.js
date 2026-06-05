// ═══════════════════════════════════════════════════════════════════
// auth.js — Authentification MSAL
// Dépend de : config.js (doit être chargé avant)
// ═══════════════════════════════════════════════════════════════════

const Auth = (function() {
  let _msal    = null;
  let _account = null;

  async function init() {
    if (_msal) return;
    _msal = new msal.PublicClientApplication({
      auth: {
        clientId:    AZURE.clientId,
        authority:   'https://login.microsoftonline.com/' + AZURE.tenantId,
        redirectUri: AZURE.redirectUri
      },
      cache: {
        cacheLocation:       'sessionStorage',
        storeAuthStateInCookie: false
      }
    });
    await _msal.initialize();
    const accounts = _msal.getAllAccounts();
    if (accounts.length > 0) _account = accounts[0];
  }

  // Connexion : silent en priorité, popup en fallback.
  // Toujours retourner l'account Microsoft réel — jamais un nom saisi manuellement.
  async function login() {
    await init();
    if (_account) {
      try {
        await _msal.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: _account });
        return _account;
      } catch (_) { /* token expiré → popup */ }
    }
    const result = await _msal.loginPopup({ scopes: GRAPH_SCOPES, prompt: 'select_account' });
    _account = result.account;
    return _account;
  }

  async function getToken() {
    if (!_account) throw new Error('Non authentifié — appeler login() d\'abord.');
    try {
      const r = await _msal.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: _account });
      return r.accessToken;
    } catch (_) {
      const r = await _msal.acquireTokenPopup({ scopes: GRAPH_SCOPES });
      _account = r.account;
      return r.accessToken;
    }
  }

  // Retourne le nom d'affichage tiré du compte Microsoft (jamais une saisie manuelle).
  function getUser() {
    if (!_account) return null;
    return {
      name:  _account.name || _account.username,
      email: _account.username,
      id:    _account.localAccountId
    };
  }

  async function logout() {
    if (!_msal || !_account) return;
    await _msal.logoutPopup({ account: _account });
    _account = null;
  }

  return { login, getToken, getUser, logout };
})();
