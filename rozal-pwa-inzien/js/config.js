// ══════════════════════════════════════════════════════════════
// CONFIGURATIE — alleen publieke gegevens
// Er staat GEEN client secret in dit bestand. Dat mag ook niet:
// deze code komt op GitHub Pages te staan en is voor iedereen
// zichtbaar. Inloggen gaat via MSAL.js (jouw eigen Microsoft-
// account), niet via een geheime sleutel.
// ══════════════════════════════════════════════════════════════
const CONFIG = {
  tenantId:         '9adf6406-9d51-428d-9c5d-34a46087b7a8',
  clientId:         '81f2417d-ac40-457b-8f28-302fd9952be4',
  spHost:           'rozal.sharepoint.com',
  klantentabelGuid: '5eaf4ac0-5c08-48ac-b934-63f218df30c8',
  docPad:           'Rozal/Documenten ROZAL',
  docMappen: {
    'Offertes':        'OFFERTES',
    'Facturen':        'ADMINISTRATIE',
    'Reparatiebonnen': 'RozalApp/Reparatiebonnen',
  },
  // Redirect URI moet exact overeenkomen met wat in Azure staat
  // bij App registraties → Rozal Kantoor → Authentication → SPA
  redirectUri: window.location.origin + window.location.pathname,
};
