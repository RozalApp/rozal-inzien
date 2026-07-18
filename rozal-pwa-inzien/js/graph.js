// ══════════════════════════════════════════════════════════════
// GRAPH API — alleen-lezen aanroepen met het MSAL-token
// ══════════════════════════════════════════════════════════════

function vertaalFout(fout) {
  const status = fout && fout.status ? fout.status : 0;
  if (status === 404) return 'Bestand niet gevonden op SharePoint.';
  if (status === 423) return 'Het Excel-bestand is nu open op kantoor. Probeer over een minuutje opnieuw.';
  if (status === 429) return 'Even te veel verzoeken tegelijk. Wacht een moment.';
  if (status === 401 || status === 403) return 'Geen toegang. Log opnieuw in.';
  if (status >= 500) return 'SharePoint is tijdelijk niet bereikbaar. Probeer het zo weer.';
  return 'Er ging iets mis bij het ophalen van de gegevens.';
}

const _RETRY_STATUSSEN = [404, 429, 502, 503, 504];
const _RETRY_WACHTTIJD = [1200, 2500, 5000];

async function _graphFetch(aanroep) {
  let lastErr;
  for (let poging = 0; poging <= 2; poging++) {
    try { return await aanroep(); }
    catch (e) {
      lastErr = e;
      if (!_RETRY_STATUSSEN.includes(e.status) || poging === 2) throw e;
      await new Promise(r => setTimeout(r, _RETRY_WACHTTIJD[poging]));
    }
  }
  throw lastErr;
}

async function graphGet(url) {
  return _graphFetch(async () => {
    const token = await getAccessToken();
    const resp = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
      headers: { Authorization: 'Bearer ' + token },
    });
    if (!resp.ok) {
      const err = new Error(`Graph ${resp.status}: ${url}`);
      err.status = resp.status;
      throw err;
    }
    return resp.json();
  });
}
