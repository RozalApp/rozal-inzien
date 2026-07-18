// ══════════════════════════════════════════════════════════════
// DATA — laadt planning, klanten en reparatiebonnen (alleen lezen)
// Kolomstructuur is exact hetzelfde als in de kantoor-app, zodat
// we altijd dezelfde gegevens zien.
// ══════════════════════════════════════════════════════════════

let alleAfspraken   = [];
let alleKlanten     = [];
let alleReparaties  = [];

// ── Datum/tijd helpers (Excel serienummer ↔ leesbaar) ─────────
function excelDatumNaarIso(v) {
  const n = Number(v);
  if (!v || isNaN(n) || n < 1) return '';
  const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
  if (isNaN(d.getTime())) return '';
  return d.toISOString().substring(0, 10);
}
function excelTijdNaarUurMin(v) {
  const n = Number(v);
  if (isNaN(n)) return { uur: 0, minuut: 0 };
  const mins = Math.round(n * 24 * 60);
  return { uur: Math.floor(mins / 60) % 24, minuut: mins % 60 };
}
function pad2(n) { return String(n).padStart(2, '0'); }

// ── PLANNING ────────────────────────────────────────────────
async function laadPlanning() {
  let file = null;
  for (const term of ['rozal_planning', 'planning']) {
    const zoek = await graphGet(`/sites/${CONFIG.spHost}/drive/root/search(q='${term}.xlsx')`);
    file = (zoek.value || []).find(i => i.name?.toLowerCase().endsWith('.xlsx') && i.name?.toLowerCase().includes('planning'));
    if (file) break;
  }
  if (!file) throw new Error('Planning Excel niet gevonden');

  const driveId = file.parentReference?.driveId;
  const werkblad = await graphGet(`/drives/${driveId}/items/${file.id}/workbook/worksheets`);
  const wsNaam = werkblad.value?.[0]?.name;
  if (!wsNaam) throw new Error('Geen werkblad gevonden');

  const bereik = await graphGet(`/drives/${driveId}/items/${file.id}/workbook/worksheets/${encodeURIComponent(wsNaam)}/usedRange`);
  const rijen = bereik.values || [];

  alleAfspraken = [];
  for (let r = 3; r < rijen.length; r++) {
    const rij = rijen[r];
    if (!rij[1] || typeof rij[1] !== 'number') continue;
    const datum = excelDatumNaarIso(rij[1]);
    if (!datum) continue;
    const { uur, minuut } = excelTijdNaarUurMin(rij[2]);
    alleAfspraken.push({
      id: String(rij[0] || ''), datum, uur, minuut,
      duurMinuten: parseInt(rij[3]) || 60,
      type: rij[4] || '', naam: rij[5] || '', adres: rij[6] || '',
      monteurs: rij[7] || '', opmerking: rij[8] || '',
      definitief: String(rij[12] || '').toLowerCase() === 'ja',
    });
  }
  return alleAfspraken;
}

// ── KLANTEN ─────────────────────────────────────────────────
async function laadKlanten() {
  let url = `/sites/${CONFIG.spHost}/lists/${CONFIG.klantentabelGuid}/items?expand=fields&$select=id,fields&$top=500`;
  let resultaten = [];
  while (url) {
    const data = await graphGet(url);
    resultaten = resultaten.concat(data.value || []);
    url = data['@odata.nextLink'] ? data['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : null;
  }
  alleKlanten = resultaten.map(item => {
    const f = item.fields || {};
    return {
      id: item.id,
      voornaam:   f.Voornaam || '',
      achternaam: f.Achternaam || '',
      bedrijf:    f.Bedrijfsnaam || '',
      adres:      f.Adres || '',
      postcode:   f.Postcode || '',
      plaats:     f.Plaats || '',
      telefoon:   f.Telefoonnummer || f.Telefoon || '',
      email:      f.E_x002d_mail_x0020_adress || f['E-mail adress'] || f.Email || f.Emailadres || '',
    };
  }).filter(k => k.achternaam || k.bedrijf)
    .sort((a, b) => (a.achternaam || a.bedrijf).localeCompare(b.achternaam || b.bedrijf, 'nl'));
  return alleKlanten;
}
function klantVolledigeNaam(k) {
  return k.bedrijf ? k.bedrijf : `${k.voornaam} ${k.achternaam}`.trim();
}

// ── REPARATIEBONNEN ─────────────────────────────────────────
// Kolommen: A=bonnr B=datum C=voornaam D=achternaam E=adres
//           F=postcode G=plaats H=email I=telefoon J=omschrijving
//           K=lat L=lon M=status
let _repDriveId = null, _repItemId = null;

async function laadReparatiebonnen() {
  const zoek = await graphGet(`/sites/${CONFIG.spHost}/drive/root/search(q='Reparatiebonnen plus locatie')`);
  const file = (zoek.value || []).find(f => f.name?.toLowerCase().includes('reparatiebonnen') && f.name?.toLowerCase().endsWith('.xlsx'));
  if (!file) throw new Error('Reparatiebonnen Excel niet gevonden');

  _repDriveId = file.parentReference?.driveId;
  _repItemId  = file.id;

  const bereik = await graphGet(`/drives/${_repDriveId}/items/${_repItemId}/workbook/worksheets/Blad1/usedRange`);
  const rijen = (bereik.values || []).slice(1).filter(r => r[0]);

  alleReparaties = rijen.map(r => ({
    bonnr:        String(r[0]),
    datum:        excelDatumNaarIso(r[1]),
    voornaam:     String(r[2] || ''),
    achternaam:   String(r[3] || ''),
    adres:        String(r[4] || ''),
    postcode:     String(r[5] || ''),
    plaats:       String(r[6] || ''),
    email:        String(r[7] || ''),
    telefoon:     String(r[8] || ''),
    omschrijving: String(r[9] || ''),
    status:       String(r[12] || 'Open') || 'Open',
  })).sort((a, b) => Number(b.bonnr) - Number(a.bonnr));
  return alleReparaties;
}

// ── PDF ZOEKEN EN OPENEN ────────────────────────────────────
// We openen het bestand via de gewone SharePoint-link (webUrl) —
// dat werkt overal betrouwbaar in Safari, ook in de PWA.
async function zoekEnOpenPdf(mapKey, bonnrOfZoekterm) {
  const mapPad = CONFIG.docPad + '/' + CONFIG.docMappen[mapKey];
  const term = String(bonnrOfZoekterm).split('_')[0];
  const data = await graphGet(`/sites/${CONFIG.spHost}/drive/root:/${encodeURIComponent(mapPad).replace(/%2F/g, '/')}:/search(q='${encodeURIComponent(term)}')?$top=20&$select=name,webUrl`);
  const items = data.value || [];
  const gevonden = items.find(f => f.name?.startsWith(term + '_') && f.name?.toLowerCase().endsWith('.pdf')) || items.find(f => f.name?.toLowerCase().endsWith('.pdf'));
  if (!gevonden) throw new Error('Bon-bestand niet gevonden op SharePoint');
  window.open(gevonden.webUrl, '_blank');
}

// Zoekt alle bestanden (offertes/facturen/reparatiebonnen/etc.) die bij een
// klant horen — exact dezelfde aanpak als de Windows-app.
//
// Bestandsnamen volgen bij Rozal het patroon "Achternaam tussenvoegsel,
// Woonplaats.pdf", bv. klant "Jan de Bruin" uit Hedel → "Bruin de, Hedel.pdf".
// Het tussenvoegsel staat dus NA de achternaam in de bestandsnaam, niet
// ervoor zoals in de spreektaal. We splitsen de achternaam daarom in losse
// woorden en eisen dat ELK woord ergens in de bestandsnaam voorkomt —
// ongeacht de volgorde — plus de woonplaats.
const TUSSENVOEGSELS = ['van','de','den','der','ten','ter','te','het','op','aan','bij','tot','uit','voor','over','in'];

function achternaamDelenVan(achternaam) {
  return (achternaam || '').trim().split(/\s+/).filter(Boolean);
}

async function zoekBestandenVoorKlant(klant) {
  const achternaamVeld = (klant.achternaam || klant.bedrijf || '').trim();
  const plaats = (klant.plaats || '').trim();
  if (!achternaamVeld) return [];

  const achternaamDelen = achternaamDelenVan(achternaamVeld);
  // Het "hoofdwoord" (zonder tussenvoegsel) gebruiken we als zoekterm richting
  // Graph — dat geeft de beste treffers; de fijnfiltering gebeurt hierna lokaal.
  const hoofdwoord = achternaamDelen.find(w => !TUSSENVOEGSELS.includes(w.toLowerCase())) || achternaamDelen[achternaamDelen.length - 1];
  const zoekterm = plaats ? `${hoofdwoord} ${plaats}` : hoofdwoord;
  const plaatsNorm = plaats.toLowerCase().replace(/[^a-z0-9]/g, '');

  const resultaten = [];
  for (const key of Object.keys(CONFIG.docMappen)) {
    try {
      const mapPad = CONFIG.docPad + '/' + CONFIG.docMappen[key];
      const data = await graphGet(`/sites/${CONFIG.spHost}/drive/root:/${encodeURIComponent(mapPad).replace(/%2F/g, '/')}:/search(q='${encodeURIComponent(zoekterm)}')?$top=25&$select=name,webUrl,lastModifiedDateTime`);
      (data.value || [])
        .filter(f => f.name?.toLowerCase().endsWith('.pdf'))
        .filter(f => bestandHoortBijKlant(f.name, achternaamDelen, plaatsNorm, key))
        .forEach(f => resultaten.push({ type: key, naam: f.name, webUrl: f.webUrl, datum: f.lastModifiedDateTime }));
    } catch (e) { console.warn('Bestanden zoeken mislukt voor', key, e); }
  }
  resultaten.sort((a, b) => new Date(b.datum) - new Date(a.datum));
  return resultaten;
}

function bestandHoortBijKlant(bestandsnaam, achternaamDelen, plaatsNorm, mapKey) {
  const basisNaam = bestandsnaam.toLowerCase().replace(/\.[^.]+$/, '');
  const basisNaamNorm = basisNaam.replace(/[^a-z0-9]/g, '');

  // Elk woord van de achternaam (incl. tussenvoegsel) moet ergens voorkomen
  const heeftAchternaam = achternaamDelen.every(deel => basisNaam.includes(deel.toLowerCase()));
  if (!heeftAchternaam) return false;

  if (plaatsNorm && !basisNaamNorm.includes(plaatsNorm)) return false;

  // Facturen/Bestelorders: extra herkenningswoord vereist (net als Windows-app)
  const marker = CONFIG.docMappenFilter?.[mapKey];
  if (marker && !basisNaam.includes(marker)) return false;

  return true;
}
