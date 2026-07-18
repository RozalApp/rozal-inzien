// ══════════════════════════════════════════════════════════════
// AUTH — inloggen met Microsoft account via MSAL.js (browser-veilig)
// Gebruikt de "Authorization Code Flow met PKCE", de standaard-
// methode voor Single-Page Applications. Er is geen client secret
// nodig en er wordt nooit een wachtwoord opgeslagen — alleen een
// tijdelijk toegangsbewijs (token) dat automatisch ververst wordt.
// ══════════════════════════════════════════════════════════════

const GRAPH_SCOPES = ['User.Read', 'Sites.Read.All', 'Files.Read.All'];

const msalConfig = {
  auth: {
    clientId:    CONFIG.clientId,
    authority:   `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: CONFIG.redirectUri,
  },
  cache: {
    cacheLocation: 'localStorage', // zodat je ingelogd blijft na herstart
    storeAuthStateInCookie: false,
  },
};

let msalInstance = null;
let _huidigeAccount = null;

async function initAuth() {
  msalInstance = new msal.PublicClientApplication(msalConfig);
  await msalInstance.initialize();

  // Afhandelen van redirect terugkomst (na inloggen bij Microsoft)
  const response = await msalInstance.handleRedirectPromise();
  if (response && response.account) {
    _huidigeAccount = response.account;
    msalInstance.setActiveAccount(response.account);
    return true;
  }

  // Al eerder ingelogd? Dan staat er een account in de cache
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    _huidigeAccount = accounts[0];
    msalInstance.setActiveAccount(accounts[0]);
    return true;
  }
  return false;
}

function inloggen() {
  msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
}

function uitloggen() {
  msalInstance.logoutRedirect();
}

function huidigeGebruikerNaam() {
  return _huidigeAccount ? (_huidigeAccount.name || _huidigeAccount.username) : '';
}

// Haalt een geldig toegangstoken op — ververst automatisch op de achtergrond
async function getAccessToken() {
  if (!_huidigeAccount) throw new Error('Niet ingelogd');
  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes:  GRAPH_SCOPES,
      account: _huidigeAccount,
    });
    return result.accessToken;
  } catch (e) {
    // Stil verversen lukte niet (bv. token te oud) — opnieuw laten inloggen
    console.warn('Token verversen mislukt, opnieuw inloggen nodig:', e);
    msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
    throw e;
  }
}
