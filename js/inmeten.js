// ══════════════════════════════════════════════════════════════
// INMETEN — inmeetformulier op de iPhone, hetzelfde als op de tablet.
//  • Klantgegevens + meerdere locaties met product en bijbehorende velden
//  • Concepten worden automatisch bewaard op de telefoon
//  • Versturen: Excel-bestand (zelfde opmaak als de tablet) naar
//    SharePoint → RozalApp/Inmeetdocumenten, foto's in een eigen submap
// Dit is het enige onderdeel van de app dat schrijft naar SharePoint.
// ══════════════════════════════════════════════════════════════

// ── Producten en velden (overgenomen uit Product.java van de tablet) ──
const INMEET_PRODUCTEN = [
  { naam: 'Rolluik', velden: ['Breedte','Hoogte','Type lamellen','Soort bediening','Type bediening','Bedieningszijde (vanaf buiten gezien)','Lamelkleur','Kastkleur','Onderlatkleur','Geleiderkleur','Geleider links','Geleider rechts','Kasttype','Kapmaat','Opmerkingen'] },
  { naam: 'RS Screen', velden: ['Breedte','Hoogte','Soort/kleur doek','Soort bediening','Type bediening','Bedieningszijde (vanaf buiten gezien)','Onderlat kleur','Kast kleur','Geleider kleur','Geleider links','Geleider rechts','Kasttype','Opmerkingen'] },
  { naam: 'Zonnescherm', velden: ['Breedte','Uitval','Knikarmscherm','Soort bediening','Type bediening','Bedieningszijde (vanaf buiten gezien)','Zender','Kleur doek','Kleur cassette','Bevestiging','Opmerkingen'] },
  { naam: 'Rolpoort Alu', velden: ['Breedte','Hoogte','Type lamellen','Soort bediening','Type bediening','Bedieningszijde (vanaf buiten gezien)','Lamelkleur','Kastkleur','Onderlatkleur','Geleiderkleur','Geleider links','Geleider rechts','Kasttype','Kapmaat','Opmerkingen'] },
  { naam: 'Aluminium Kozijn', velden: ['Breedte','Hoogte','Vulling','Kleur kozijn','Opmerkingen'] },
  { naam: 'Tahoma Switch Box', velden: ['Aantal','Opmerkingen'] },
  { naam: 'Terrasoverkapping RZ-Line', velden: ['Aantal vlakken','Kleur','Dakvulling','Fundatie','Staanders','Opmerkingen'] },
  { naam: 'Raamhorren', velden: ['Breedte','Hoogte','Type','Kleur','Opmerkingen'] },
  { naam: 'Glasschuifwand RZ-Aluminium', velden: ['Aantal','Uitvoering','Rails','Glaspanelen','Wandaansluiting','Kleur','Greep','Tochtprofielen','Schuifrichting','Opmerkingen'] },
  { naam: 'Alu Markies', velden: ['Breedte','Uitval','Markies','Bovenkap','Zijkappen','Volant','Soort/Kleur doek','Soort bediening','Type bediening','Bedieningszijde (vanaf buiten gezien)','Omkasting kleur','Opmerkingen'] },
  { naam: 'Uitvalscherm', velden: ['Breedte','Uitval','Uitvalscherm','Soort bediening','Type bediening','Bedieningszijde (vanaf buiten gezien)','Soort/Kleur doek','Kleur cassette','Opmerkingen'] },
  { naam: 'Hordeur', velden: ['Breedte totaal','Hoogte','Uitvoering','Kleur','Kleur gaas','Onderstrip','Opmerkingen'] },
  { naam: 'Jaloezie', velden: ['Breedte','Hoogte','Uitvoering','Soort lamel','Lamelbreedte','Kleur lamel','Type bediening','Afwerking','Opmerkingen'] },
  { naam: 'Duo Rolgordijn model 3011', velden: ['Breedte','Hoogte','Soort','Stofkleur','Type bediening','Onderlat','Opmerkingen'] },
  { naam: 'Garagedeur AW 40 UniPro', velden: ['Breedte','Hoogte','Uitvoering','Paneel','Kleur binnenzijde','Kleur buitenzijde','Extra zender','Bediening','Materialen','Opmerkingen'] },
  { naam: 'Zonweringsdoek vervangen', velden: ['Breedte','Uitval','Soort/kleur doek','Opmerkingen'] },
  { naam: 'Rolluik t.b.v. Velux dakraam', velden: ['Aantal','Type dakraam','Maatcode','Bediening','Afwerking buiten','Opmerkingen'] },
  { naam: 'Rolgordijn t.b.v. Velux dakraam', velden: ['Aantal','Code raam','Kleur doek','Kleur cassette','Opmerkingen'] },
  { naam: 'Handzender Situo io 1-kanaals', velden: ['Aantal','Opmerkingen'] },
  { naam: 'Handzender Ysia Patio io 16-kanaals', velden: ['Aantal','Opmerkingen'] },
  { naam: 'Handzender Situo io 5-kanaals', velden: ['Aantal','Opmerkingen'] },
  { naam: 'Verandazonwering Vinci 150', velden: ['Uitvoering','Kleur Frame','Kleur doek','Bediening','Positie'] },
  { naam: 'Verandazonwering Vinci 250', velden: ['Uitvoering','Kleur Frame','Kleur doek','Bediening','Positie'] },
];

// Keuzelijsten (exact zoals op de tablet / in de kantoor-app). Geen vermelding = vrij tekstveld.
const INMEET_KEUZES = {
  'Rolluik': {
    'Type lamellen': ['RZ 40','RZH 40','RZ 41','Ecoroll-L','RZ 52'],
    'Soort bediening': ['Handbediend','Elektrisch','Elektrisch en Band'],
    'Type bediening': ['Band met bandopwinder','Band met bandopwinder (incl. bandvertrager)','Somfy motor LT met draaischakelaar','Somfy motor Ilmo','Somfy motor Oximo io','Somfy motor io','Somfy motor RS 100 io Solar','Somfy motor RS 100 io','Faher Solar Bi','Solar motor','Elektrisch en Band'],
    'Geleider links': ['HTF','LHTF','LHF','HF','F-HTF12','F-HTF20'],
    'Geleider rechts': ['HTF','LHTF','LHF','HF','F-HTF12','F-HTF20'],
    'Kasttype': ['Schuin','Halfrond','Rond'],
    'Kapmaat': ['150','165','180','205','250'],
  },
  'RS Screen': {
    'Soort bediening': ['Handbediend','Elektrisch'],
    'Type bediening': ['Oogwindwerk met slingerstang','Somfy motor LT','Somfy motor Altus RTS','Somfy motor Sunea io','Somfy motor Maestria io','Somfy motor io','Somfy motor RS 100 io Solar','Somfy motor Sunilus io','Solar motor'],
    'Geleider links': ['HTF','MHTF','LHTF','CTG','UTG','UTG-S','MUTG','UG','STG-S','MSTG-S','MSTG-S50','STG','MSTG','STG-Click','STG-0','Zip','Zip+lip'],
    'Geleider rechts': ['HTF','MHTF','LHTF','CTG','UTG','UTG-S','MUTG','UG','STG-S','MSTG-S','MSTG-S50','STG','MSTG','STG-Click','STG-0','Zip','Zip+lip'],
    'Kasttype': ['Classic 75','Classic 85','Classic 95','Classic 95 Recht','Classic 105','Classic 105 Recht','Classic 120','Classic 125','Uni 85','Uni 95','Uni 105','Uni 105 Recht','Uni 120','Uni 125','Nova 83','Nova 103','Nova Solar','Nova 83 Zip Recht','Nova 103 Zip Afgerond','Flowzz','Nova 123 Zip','Classic 125 Recht','Uni 125 Recht','Smart 130'],
  },
  'Zonnescherm': {
    'Knikarmscherm': ['Gaudi 300','Gaudi 400','Moneo 400','Pisano 230','Tucan','Capri','Formosa'],
  },
};

function inmeetProduct(naam) { return INMEET_PRODUCTEN.find(p => p.naam === naam) || null; }
function inmeetIsBedieningszijde(v) { return v.toLowerCase().startsWith('bedieningszijde'); }
function inmeetOpties(product, veld) {
  if (inmeetIsBedieningszijde(veld)) return ['Links', 'Rechts'];
  return (INMEET_KEUZES[product] && INMEET_KEUZES[product][veld]) || null;
}
function inmeetIsNumeriek(veld) {
  const l = veld.toLowerCase().trim();
  return ['breedte','hoogte','uitval','uitvallengte','kapmaat','lamelbreedte','aantal','aantal vlakken','maatcode']
    .some(n => l === n || l.startsWith(n + ' ') || l.endsWith(' ' + n));
}
// Kleurvelden voor "alles dezelfde kleur?" — kleur van doek/stof blijft altijd apart
function inmeetIsKleurSync(veld) {
  const l = veld.toLowerCase();
  return l.includes('kleur') && !l.includes('doek') && !l.includes('stof');
}

// ── Staat ─────────────────────────────────────────────────────
const INMEET_LS = 'rozal-inmeet-concepten-v1';
let inmeetConcepten = [];     // opgeslagen concepten (zonder foto's)
let inmeetConcept   = null;   // het formulier dat nu open staat
let inmeetRecent    = [];     // laatst verstuurde inmetingen (uit SharePoint)
let inmeetBezig     = false;  // bezig met versturen
let _inmeetOpslagTimer = null;

function inmeetNieuwId() { return 'c' + Date.now() + Math.random().toString(36).slice(2, 6); }
function inmeetNieuweLocatie() { return { id: inmeetNieuwId(), omschrijving: '', product: '', velden: {}, zelfdeKleur: false, fotos: [] }; }
function inmeetNieuwConcept() {
  return { id: inmeetNieuwId(), naam: '', adres: '', telefoon: '', email: '', gewijzigd: Date.now(), locaties: [inmeetNieuweLocatie()] };
}

function inmeetLeeg(c) {
  if ((c.naam || '').trim() || (c.adres || '').trim() || (c.telefoon || '').trim() || (c.email || '').trim()) return false;
  return !c.locaties.some(l =>
    (l.omschrijving || '').trim() || l.product ||
    Object.values(l.velden || {}).some(v => String(v || '').trim()) ||
    (l.fotos && l.fotos.length));
}

// ── Opslag van concepten (localStorage) ──────────────────────
function inmeetLaadConcepten() {
  try { inmeetConcepten = JSON.parse(localStorage.getItem(INMEET_LS) || '[]') || []; }
  catch (e) { inmeetConcepten = []; }
}
function inmeetSchrijfConcepten() {
  try {
    const schoon = inmeetConcepten.map(c => ({ ...c, locaties: c.locaties.map(l => ({ ...l, fotos: [] })) }));
    localStorage.setItem(INMEET_LS, JSON.stringify(schoon));
  } catch (e) { console.warn('Concepten opslaan mislukt:', e); }
}

// Het open formulier opslaan (of verwijderen als het volledig leeg is)
function inmeetBewaarNu() {
  clearTimeout(_inmeetOpslagTimer);
  if (!inmeetConcept) return;
  const c = inmeetConcept;
  c.gewijzigd = Date.now();
  const idx = inmeetConcepten.findIndex(x => x.id === c.id);
  if (inmeetLeeg(c)) {
    if (idx >= 0) inmeetConcepten.splice(idx, 1);
  } else if (idx >= 0) {
    inmeetConcepten[idx] = c;
  } else {
    inmeetConcepten.unshift(c);
  }
  inmeetSchrijfConcepten();
}
function inmeetPlanOpslag() {
  clearTimeout(_inmeetOpslagTimer);
  _inmeetOpslagTimer = setTimeout(inmeetBewaarNu, 500);
}
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') inmeetBewaarNu(); });
window.addEventListener('pagehide', inmeetBewaarNu);

// ── Tabblad ───────────────────────────────────────────────────
function inmeetInit() {
  inmeetLaadConcepten();
  inmeetRenderTab();
  inmeetLaadRecent().then(inmeetRenderTab).catch(e => console.warn('Recente inmetingen laden mislukt:', e));
}
function inmeetTabOpen() {
  inmeetLaadConcepten();
  inmeetRenderTab();
}

function inmeetDatumTekst(ms) {
  const d = new Date(ms);
  return `${d.getDate()} ${MAAND_NAMEN[d.getMonth()].slice(0, 3)} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
}

function inmeetRenderTab() {
  const sub = document.getElementById('im-sub');
  if (sub) sub.textContent = inmeetConcepten.length ? inmeetConcepten.length + ' niet afgerond' : 'Nieuwe inmeting starten';

  const cEl = document.getElementById('im-concepten');
  if (cEl) {
    cEl.innerHTML = inmeetConcepten.length
      ? inmeetConcepten.map(c => {
          const nr = c.locaties.filter(l => l.product || l.omschrijving).length;
          return `<div class="card" onclick="inmeetHervat('${c.id}')">
            <div class="bestand-icon" style="background:var(--oranje-licht)">📐</div>
            <div class="plan-info">
              <div class="plan-naam">${escapeHtml(c.naam || '(naam nog niet ingevuld)')}</div>
              <div class="plan-detail">${nr} locatie${nr === 1 ? '' : 's'} · ${inmeetDatumTekst(c.gewijzigd)}</div>
            </div>
            <div class="im-wis" onclick="event.stopPropagation(); inmeetVerwijderConcept('${c.id}')" title="Verwijderen">✕</div>
          </div>`;
        }).join('')
      : '<div class="leeg-melding" style="padding:18px">Geen formulieren in behandeling</div>';
  }

  const rEl = document.getElementById('im-recent');
  if (rEl) {
    rEl.innerHTML = inmeetRecent.length
      ? inmeetRecent.map(f => `<div class="card" onclick="window.open('${escapeHtml(f.webUrl)}', '_blank')">
          <div class="bestand-icon" style="background:var(--groen-licht)">✅</div>
          <div class="plan-info"><div class="plan-naam">${escapeHtml(f.klant)}</div><div class="plan-detail">${escapeHtml(f.datum)}</div></div>
          <div class="klant-pijl">›</div>
        </div>`).join('')
      : '<div class="leeg-melding" style="padding:18px">Nog geen verstuurde inmetingen gevonden</div>';
  }
}

function inmeetVerwijderConcept(id) {
  const c = inmeetConcepten.find(x => x.id === id);
  if (!c) return;
  if (!confirm('Formulier van ' + (c.naam || 'deze klant') + ' verwijderen?')) return;
  inmeetConcepten = inmeetConcepten.filter(x => x.id !== id);
  inmeetSchrijfConcepten();
  inmeetRenderTab();
}

// Laatste verstuurde inmetingen uit de SharePoint-map
async function inmeetLaadRecent() {
  const mapPad = CONFIG.docPad + '/' + CONFIG.docMappen['Inmeetdocumenten'];
  const url = `/sites/${CONFIG.spHost}/drive/root:/${inmeetPadEnc(mapPad)}:/children?$select=name,webUrl,lastModifiedDateTime&$orderby=lastModifiedDateTime desc&$top=25`;
  const data = await graphGet(url);
  inmeetRecent = (data.value || [])
    .filter(f => f.name && f.name.toLowerCase().endsWith('.xlsx') && !f.name.toLowerCase().includes('overzicht'))
    .slice(0, 15)
    .map(f => {
      const m = f.name.match(/^Opmeting_(.+)_(\d{8})_(\d{4,6})\.xlsx$/i);
      const klant = (m ? m[1] : f.name.replace(/\.xlsx$/i, '')).replace(/_/g, ' ').trim();
      const d = new Date(f.lastModifiedDateTime);
      return { klant, webUrl: f.webUrl, datum: `${d.getDate()} ${MAAND_NAMEN[d.getMonth()]} ${d.getFullYear()} · ${pad2(d.getHours())}:${pad2(d.getMinutes())}` };
    });
}

function inmeetPadEnc(p) { return p.split('/').map(encodeURIComponent).join('/'); }

// ── Formulier openen / sluiten ────────────────────────────────
function inmeetNieuw() {
  inmeetConcept = inmeetNieuwConcept();
  inmeetOpenFormulier();
}
function inmeetHervat(id) {
  const c = inmeetConcepten.find(x => x.id === id);
  if (!c) return;
  inmeetConcept = JSON.parse(JSON.stringify(c));
  inmeetConcept.locaties.forEach(l => { l.fotos = []; });
  if (!inmeetConcept.locaties.length) inmeetConcept.locaties.push(inmeetNieuweLocatie());
  inmeetOpenFormulier();
}
function inmeetOpenFormulier() {
  document.getElementById('im-titel').textContent = inmeetConcept.naam || 'Nieuwe inmeting';
  inmeetRenderFormulier();
  document.getElementById('im-content').scrollTop = 0;
  document.getElementById('overlay-inmeet').classList.add('open');
}
function inmeetSluitFormulier() {
  if (inmeetBezig) return;
  inmeetBewaarNu();
  inmeetConcept = null;
  document.getElementById('overlay-inmeet').classList.remove('open');
  inmeetRenderTab();
}

// ── Formulier tekenen ─────────────────────────────────────────
function inmeetRenderFormulier() {
  const c = inmeetConcept;
  if (!c) return;
  const content = document.getElementById('im-content');
  const scroll = content.scrollTop;

  let html = `
    <div class="im-kaart">
      <div class="im-kop">Klantgegevens</div>
      <label class="im-lbl">Bestaande klant zoeken</label>
      <input class="im-inp" id="im-klantzoek" placeholder="Zoek op naam of plaats..." oninput="inmeetKlantZoek(this.value)" autocomplete="off">
      <div id="im-klantresultaten"></div>
      <label class="im-lbl">Naam *</label>
      <input class="im-inp" value="${escapeHtml(c.naam)}" oninput="inmeetKlantVeld('naam', this.value)" autocomplete="off">
      <label class="im-lbl">Adres *</label>
      <input class="im-inp" value="${escapeHtml(c.adres)}" oninput="inmeetKlantVeld('adres', this.value)" autocomplete="off">
      <label class="im-lbl">Telefoon</label>
      <input class="im-inp" type="tel" value="${escapeHtml(c.telefoon)}" oninput="inmeetKlantVeld('telefoon', this.value)" autocomplete="off">
      <label class="im-lbl">E-mail</label>
      <input class="im-inp" type="email" value="${escapeHtml(c.email)}" oninput="inmeetKlantVeld('email', this.value)" autocomplete="off" autocapitalize="none">
    </div>`;

  c.locaties.forEach((loc, li) => { html += inmeetLocatieHtml(loc, li, c.locaties.length > 1); });

  html += `<div class="im-knop-sec" onclick="inmeetVoegLocatieToe()">＋ Locatie toevoegen</div>`;
  content.innerHTML = html;
  content.scrollTop = scroll;
  const titel = document.getElementById('im-titel');
  if (titel) titel.textContent = c.naam || 'Nieuwe inmeting';
}

function inmeetLocatieHtml(loc, li, kanVerwijderen) {
  const prod = inmeetProduct(loc.product);
  let velden = '';
  if (prod) {
    const syncVelden = prod.velden.filter(inmeetIsKleurSync);
    if (syncVelden.length >= 2) {
      velden += `<label class="im-check"><input type="checkbox" ${loc.zelfdeKleur ? 'checked' : ''} onchange="inmeetZelfdeKleur(${li}, this.checked)"> Alles dezelfde kleur?</label>`;
    }
    prod.velden.forEach((veld, vi) => { velden += inmeetVeldHtml(loc, li, prod, veld, vi); });
  }

  const fotos = loc.fotos.map((f, fi) =>
    `<div class="im-foto"><img src="${f.url}" alt=""><div class="im-foto-x" onclick="inmeetVerwijderFoto(${li}, ${fi})">✕</div></div>`).join('');

  return `<div class="im-kaart">
    <div class="im-kop-rij">
      <div class="im-kop" style="color:var(--oranje)">Locatie ${li + 1}</div>
      <div class="im-acties">
        <span class="im-mini" onclick="inmeetDupliceerLocatie(${li})" title="Kopie maken">⧉</span>
        ${kanVerwijderen ? `<span class="im-mini" onclick="inmeetVerwijderLocatie(${li})" title="Verwijderen">✕</span>` : ''}
      </div>
    </div>
    <label class="im-lbl">Omschrijving locatie *</label>
    <input class="im-inp" value="${escapeHtml(loc.omschrijving)}" placeholder="Bijv. Slaapkamer links" oninput="inmeetLocVeld(${li}, this.value)" autocomplete="off">
    <label class="im-lbl">Product *</label>
    <select class="im-inp" onchange="inmeetKiesProduct(${li}, this.value)">
      <option value="">— kies een product —</option>
      ${INMEET_PRODUCTEN.map(p => `<option value="${escapeHtml(p.naam)}" ${p.naam === loc.product ? 'selected' : ''}>${escapeHtml(p.naam)}</option>`).join('')}
    </select>
    ${velden}
    <div class="im-fotos">${fotos}</div>
    <label class="im-foto-knop">📷 Foto toevoegen
      <input type="file" accept="image/*" multiple style="display:none" onchange="inmeetFotoToevoegen(${li}, this)">
    </label>
  </div>`;
}

function inmeetVeldHtml(loc, li, prod, veld, vi) {
  const waarde = (loc.velden[veld] != null) ? String(loc.velden[veld]) : '';
  const id = `imv-${li}-${vi}`;
  const kop = `<label class="im-lbl">${escapeHtml(veld)}</label>`;
  if (veld === 'Opmerkingen') {
    return `${kop}<textarea class="im-inp" id="${id}" rows="2" oninput="inmeetVeldInput(${li}, ${vi}, this.value)">${escapeHtml(waarde)}</textarea>`;
  }
  const opties = inmeetOpties(prod.naam, veld);
  if (opties) {
    const lijst = (waarde && !opties.includes(waarde)) ? [waarde].concat(opties) : opties;
    return `${kop}<select class="im-inp" id="${id}" onchange="inmeetVeldInput(${li}, ${vi}, this.value)">
      <option value="">—</option>
      ${lijst.map(o => `<option value="${escapeHtml(o)}" ${o === waarde ? 'selected' : ''}>${escapeHtml(o)}</option>`).join('')}
    </select>`;
  }
  const modus = inmeetIsNumeriek(veld) ? ' inputmode="decimal"' : '';
  return `${kop}<input class="im-inp" id="${id}" value="${escapeHtml(waarde)}"${modus} oninput="inmeetVeldInput(${li}, ${vi}, this.value)" autocomplete="off">`;
}

// ── Invoer verwerken ──────────────────────────────────────────
function inmeetKlantVeld(veld, waarde) {
  inmeetConcept[veld] = waarde;
  if (veld === 'naam') document.getElementById('im-titel').textContent = waarde || 'Nieuwe inmeting';
  inmeetPlanOpslag();
}
function inmeetLocVeld(li, waarde) {
  inmeetConcept.locaties[li].omschrijving = waarde;
  inmeetPlanOpslag();
}
function inmeetVeldInput(li, vi, waarde) {
  const loc = inmeetConcept.locaties[li];
  const prod = inmeetProduct(loc.product);
  if (!prod) return;
  const veld = prod.velden[vi];
  loc.velden[veld] = waarde;

  // "Alles dezelfde kleur?": de andere kleurvelden meenemen
  if (loc.zelfdeKleur && inmeetIsKleurSync(veld)) {
    prod.velden.forEach((andere, ai) => {
      if (ai === vi || !inmeetIsKleurSync(andere)) return;
      loc.velden[andere] = waarde;
      const el = document.getElementById(`imv-${li}-${ai}`);
      if (el) el.value = waarde;
    });
  }
  inmeetPlanOpslag();
}
function inmeetZelfdeKleur(li, aan) {
  inmeetConcept.locaties[li].zelfdeKleur = !!aan;
  inmeetPlanOpslag();
}
function inmeetKiesProduct(li, naam) {
  const loc = inmeetConcept.locaties[li];
  if (loc.product === naam) return;
  loc.product = naam;
  loc.velden = {};
  inmeetRenderFormulier();
  inmeetPlanOpslag();
}
function inmeetVoegLocatieToe() {
  inmeetConcept.locaties.push(inmeetNieuweLocatie());
  inmeetRenderFormulier();
  const content = document.getElementById('im-content');
  content.scrollTop = content.scrollHeight;
  inmeetPlanOpslag();
}
function inmeetDupliceerLocatie(li) {
  const kopie = JSON.parse(JSON.stringify({ ...inmeetConcept.locaties[li], fotos: [] }));
  kopie.id = inmeetNieuwId();
  inmeetConcept.locaties.splice(li + 1, 0, kopie);
  inmeetRenderFormulier();
  inmeetPlanOpslag();
}
function inmeetVerwijderLocatie(li) {
  if (inmeetConcept.locaties.length <= 1) return;
  if (!confirm('Locatie ' + (li + 1) + ' verwijderen?')) return;
  inmeetConcept.locaties[li].fotos.forEach(f => URL.revokeObjectURL(f.url));
  inmeetConcept.locaties.splice(li, 1);
  inmeetRenderFormulier();
  inmeetPlanOpslag();
}

// ── Bestaande klant zoeken ────────────────────────────────────
function inmeetKlantZoek(term) {
  const el = document.getElementById('im-klantresultaten');
  const t = (term || '').toLowerCase().trim();
  if (t.length < 2) { el.innerHTML = ''; return; }
  const treffers = alleKlanten
    .map((k, idx) => ({ k, idx }))
    .filter(({ k }) => (klantVolledigeNaam(k) + ' ' + (k.plaats || '') + ' ' + (k.adres || '')).toLowerCase().includes(t))
    .slice(0, 6);
  el.innerHTML = treffers.length
    ? '<div class="im-resultaten">' + treffers.map(({ k, idx }) =>
        `<div class="im-resultaat" onclick="inmeetKiesKlant(${idx})"><b>${escapeHtml(klantVolledigeNaam(k))}</b><br><span>${escapeHtml([k.adres, k.plaats].filter(Boolean).join(', '))}</span></div>`).join('') + '</div>'
    : '<div class="im-resultaten"><div class="im-resultaat"><span>Geen klanten gevonden</span></div></div>';
}
function inmeetKiesKlant(idx) {
  const k = alleKlanten[idx];
  if (!k) return;
  const c = inmeetConcept;
  c.naam = klantVolledigeNaam(k);
  c.adres = [k.adres, k.postcode, k.plaats].filter(Boolean).join(', ');
  c.telefoon = k.telefoon || '';
  c.email = k.email || '';
  inmeetRenderFormulier();
  inmeetPlanOpslag();
}

// ── Foto's ────────────────────────────────────────────────────
// Foto's worden verkleind (max 1600 px) zodat uploaden snel gaat.
function inmeetVerkleinFoto(bestand) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(bestand);
    const img = new Image();
    img.onload = () => {
      const max = 1600;
      const schaal = Math.min(1, max / Math.max(img.width, img.height));
      const cv = document.createElement('canvas');
      cv.width = Math.round(img.width * schaal);
      cv.height = Math.round(img.height * schaal);
      cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height);
      URL.revokeObjectURL(url);
      cv.toBlob(b => b ? resolve(b) : reject(new Error('Foto verkleinen mislukt')), 'image/jpeg', 0.82);
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('Foto kon niet gelezen worden')); };
    img.src = url;
  });
}
async function inmeetFotoToevoegen(li, input) {
  const bestanden = Array.from(input.files || []);
  input.value = '';
  if (!bestanden.length) return;
  for (const b of bestanden) {
    try {
      const blob = await inmeetVerkleinFoto(b);
      inmeetConcept.locaties[li].fotos.push({ id: inmeetNieuwId(), blob, url: URL.createObjectURL(blob) });
    } catch (e) { alert('Foto toevoegen mislukt: ' + e.message); }
  }
  inmeetRenderFormulier();
}
function inmeetVerwijderFoto(li, fi) {
  const f = inmeetConcept.locaties[li].fotos[fi];
  if (f) URL.revokeObjectURL(f.url);
  inmeetConcept.locaties[li].fotos.splice(fi, 1);
  inmeetRenderFormulier();
}

// ── Versturen ─────────────────────────────────────────────────
function inmeetToast(tekst) {
  let t = document.getElementById('im-toast');
  if (!t) { t = document.createElement('div'); t.id = 'im-toast'; document.body.appendChild(t); }
  t.textContent = tekst;
  t.classList.add('toon');
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.classList.remove('toon'), 3500);
}

function inmeetValideer() {
  const c = inmeetConcept;
  if (!(c.naam || '').trim())  return 'Vul de naam van de klant in.';
  if (!(c.adres || '').trim()) return 'Vul het adres in.';
  for (let i = 0; i < c.locaties.length; i++) {
    const l = c.locaties[i];
    if (!(l.omschrijving || '').trim()) return `Locatie ${i + 1}: vul een omschrijving in.`;
    if (!l.product) return `Locatie ${i + 1}: kies een product.`;
  }
  return null;
}

async function inmeetVerstuur() {
  if (inmeetBezig || !inmeetConcept) return;
  const fout = inmeetValideer();
  if (fout) { alert(fout); return; }

  const c = inmeetConcept;
  const aantalFotos = c.locaties.reduce((s, l) => s + l.fotos.length, 0);
  if (!confirm(`Inmeting van ${c.naam.trim()} versturen naar kantoor?\n\n${c.locaties.length} locatie(s)${aantalFotos ? ' en ' + aantalFotos + ' foto(s)' : ''} worden geüpload.`)) return;

  const knop = document.getElementById('im-verstuur');
  inmeetBezig = true;
  knop.classList.add('laden');
  knop.textContent = '⏳ Versturen...';
  try {
    const resultaat = await inmeetUpload(c);
    // Gelukt: concept weghalen en terug naar het tabblad
    inmeetConcepten = inmeetConcepten.filter(x => x.id !== c.id);
    inmeetSchrijfConcepten();
    c.locaties.forEach(l => l.fotos.forEach(f => URL.revokeObjectURL(f.url)));
    inmeetConcept = null;
    document.getElementById('overlay-inmeet').classList.remove('open');
    inmeetToast(resultaat.fotosMislukt
      ? `✅ Verstuurd, maar ${resultaat.fotosMislukt} foto(s) niet geüpload`
      : '✅ Inmeting verstuurd naar kantoor');
    inmeetRenderTab();
    inmeetLaadRecent().then(inmeetRenderTab).catch(() => {});
  } catch (e) {
    console.error('Inmeting versturen mislukt:', e);
    alert('Versturen mislukt: ' + inmeetFoutTekst(e) + '\n\nJe formulier is bewaard, je kunt het later opnieuw proberen.');
    inmeetBewaarNu();
  } finally {
    inmeetBezig = false;
    knop.classList.remove('laden');
    knop.textContent = '📤 Versturen naar kantoor';
  }
}

function inmeetFoutTekst(e) {
  const s = e && e.status;
  if (s === 401 || s === 403) return 'Geen schrijfrechten. Controleer in Azure of "Files.ReadWrite.All" is toegevoegd en log opnieuw in.';
  if (s === 404) return 'De map Inmeetdocumenten is niet gevonden op SharePoint.';
  if (s === 423) return 'Het bestand is geblokkeerd. Probeer het zo opnieuw.';
  if (s === 507) return 'Geen opslagruimte meer op SharePoint.';
  if (s >= 500) return 'SharePoint is tijdelijk niet bereikbaar.';
  if (!navigator.onLine) return 'Geen internetverbinding.';
  return (e && e.message) || 'Onbekende fout';
}

// Excel (+ foto's) naar SharePoint. De Excel komt direct in Inmeetdocumenten
// (de Windows-app zoekt daar), de foto's in een eigen submap.
async function inmeetUpload(c) {
  const nu = new Date();
  const stempel = `${nu.getFullYear()}${pad2(nu.getMonth() + 1)}${pad2(nu.getDate())}_${pad2(nu.getHours())}${pad2(nu.getMinutes())}${pad2(nu.getSeconds())}`;
  const veilig = c.naam.trim().replace(/[^a-zA-Z0-9_\-]/g, '_');
  const bestand = `Opmeting_${veilig}_${stempel}.xlsx`;
  const mapPad = CONFIG.docPad + '/' + CONFIG.docMappen['Inmeetdocumenten'];
  const mapUrl = `/sites/${CONFIG.spHost}/drive/root:/${inmeetPadEnc(mapPad)}`;

  const data = {
    naam: c.naam.trim(), adres: c.adres.trim(), telefoon: (c.telefoon || '').trim(), email: (c.email || '').trim(),
    locaties: c.locaties.map(l => {
      const prod = inmeetProduct(l.product);
      return {
        omschrijving: l.omschrijving.trim(),
        product: l.product,
        velden: prod.velden.map(v => [v, String(l.velden[v] == null ? '' : l.velden[v]).trim()]),
      };
    }),
  };
  const xlsx = maakInmeetXlsx(data, nu);
  await graphSchrijf('PUT', `${mapUrl}/${encodeURIComponent(bestand)}:/content`, xlsx,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

  // Foto's (als de upload mislukt blijft de Excel gewoon staan)
  let fotosMislukt = 0;
  const heeftFotos = c.locaties.some(l => l.fotos.length);
  if (heeftFotos) {
    let mapId = null;
    try {
      const sub = await graphSchrijf('POST', `${mapUrl}:/children`, JSON.stringify({
        name: `Opmeting_${veilig}_${stempel}_fotos`, folder: {}, '@microsoft.graph.conflictBehavior': 'rename',
      }), 'application/json');
      mapId = sub && sub.id;
    } catch (e) { console.warn('Foto-submap maken mislukt:', e); }

    for (let li = 0; li < c.locaties.length; li++) {
      const l = c.locaties[li];
      const omschr = (l.omschrijving || '').trim().replace(/[^a-zA-Z0-9_\-]/g, '_') || ('Locatie' + (li + 1));
      for (let fi = 0; fi < l.fotos.length; fi++) {
        if (!mapId) { fotosMislukt++; continue; }
        try {
          const bytes = new Uint8Array(await l.fotos[fi].blob.arrayBuffer());
          await graphSchrijf('PUT', `/sites/${CONFIG.spHost}/drive/items/${mapId}:/${encodeURIComponent(`Foto_${omschr}_${fi + 1}.jpg`)}:/content`, bytes, 'image/jpeg');
        } catch (e) { console.warn('Foto uploaden mislukt:', e); fotosMislukt++; }
      }
    }
  }
  return { bestand, fotosMislukt };
}
