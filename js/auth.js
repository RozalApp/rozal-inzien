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
    storeAuthStateInCookie: true,  // Safari-fix: bewaart inlogstatus ook in
                                    // een cookie, nodig omdat Safari tijdens
                                    // de heen-en-terug naar Microsoft soms
                                    // browseropslag kwijtraakt (vooral als
                                    // app op het beginscherm).
  },
};

let msalInstance = null;
let _huidigeAccount = null;

async function initAuth() {
  msalInstance = new msal.PublicClientApplication(msalConfig);
  await msalInstance.initialize();

  // Afhandelen van redirect terugkomst (na inloggen bij Microsoft)
  let response = null;
  try {
    response = await msalInstance.handleRedirectPromise();
  } catch (e) {
    console.warn('Redirect afhandelen mislukt, oude inlogstatus wordt opgeschoond:', e);
    ruimVastgelopenInlogstatusOp();
    // Belangrijk: een mislukte poging laat soms een kapot #-restje in de
    // URL staan. Zonder dit weg te halen blijft de app dat restje bij elke
    // volgende poging opnieuw proberen te verwerken en steeds opnieuw falen.
    if (window.location.hash) {
      history.replaceState(null, '', window.location.pathname + window.location.search);
    }
    response = null;
    // Ga gewoon door naar het inlogscherm — niet opnieuw proberen te parsen.
  }

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

// Verwijdert MSAL's tijdelijke "bezig met inloggen"-vlaggetjes. Die kunnen
// blijven staan als een eerdere inlogpoging halverwege misging (bv. door
// een verkeerd ingestelde redirect-URL) — zonder dit zou de app daarna
// blijven denken dat er al een inlogpoging loopt.
function ruimVastgelopenInlogstatusOp() {
  try {
    [sessionStorage, localStorage].forEach(opslag => {
      Object.keys(opslag)
        .filter(k => k.startsWith('msal.') || k.toLowerCase().includes('interaction.status'))
        .forEach(k => opslag.removeItem(k));
    });
  } catch (e) { console.warn('Opschonen mislukt:', e); }
}

function inloggen() {
  try {
    msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
  } catch (e) {
    console.warn('Inloggen mislukt, opnieuw proberen na opschonen:', e);
    ruimVastgelopenInlogstatusOp();
    msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
  }
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
