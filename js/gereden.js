// ══════════════════════════════════════════════════════════════
// GEREDEN BONNEN — kaart met pins van alle gereden montage- en reparatiebonnen
//
// • Gereden = taken die de monteur op de tablet op "Klaar" heeft gezet
//   (TaakRegistratie.xlsx, blad "Taken").
// • Type, adres en monteurs komen uit de planning (zelfde klant + datum).
// • Reparatie: de locatie komt uit "Reparatiebonnen plus locatie" (kolom K/L) via het
//   bonnummer (#44) in de planning. Montage: het adres wordt opgezocht via OpenStreetMap
//   (Nominatim) en onthouden op dit toestel, zodat het de volgende keer meteen staat.
// • Zelfstandig: dit bestand voegt zelf de tab, de pagina en de opmaak toe.
// ══════════════════════════════════════════════════════════════

const GER_THUIS   = [51.754438, 5.264027];          // Rozal, Hedel
const GER_GEO_KEY = 'rozal_ger_geo_v1';
const GER_TYPE_KLEUR = { Montage: '#2D5016', Reparatie: '#C0392B', Onbekend: '#7f8c8d' };

let gerRijen = [];
let gerBezig = false;
let gerGeladen = false;
let gerType = 'alle';          // 'alle' | 'Montage' | 'Reparatie'
let gerPeriode = '30';         // 'week' | '30' | 'alles'
let gerActief = null;          // id van de gekozen bon
let gerKaart = null;
let gerMarkers = {};           // id -> Leaflet-marker
let gerThuisMarker = null;
let gerFout = '';
let gerGeoCache = {};
let gerGeoWachtrij = [];
let gerGeoBezig = false;
let gerPasPassen = true;       // kaart opnieuw passend zoomen na wijziging van filter/data
try { gerGeoCache = JSON.parse(localStorage.getItem(GER_GEO_KEY) || '{}') || {}; } catch (e) { gerGeoCache = {}; }

function gerEsc(s) {
  return String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function gerP2(n) { return String(n).padStart(2, '0'); }
function gerIso(d) { return d.getFullYear() + '-' + gerP2(d.getMonth() + 1) + '-' + gerP2(d.getDate()); }
function gerDatumKort(iso) {
  const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? (parseInt(m[3], 10) + ' ' + MAAND_NAMEN[parseInt(m[2], 10) - 1].slice(0, 3)) : (iso || '');
}
function gerDagNaam(iso) {
  const m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return iso || '';
  const dt = new Date(+m[1], +m[2] - 1, +m[3]);
  return DAG_VOLNAAM[(dt.getDay() + 6) % 7] + ' ' + (+m[3]) + ' ' + MAAND_NAMEN[+m[2] - 1];
}

// ── Datum/tijd uit Excel ──────────────────────────────────────
function gerDatumNorm(v) {
  if (v === null || v === undefined || v === '') return '';
  const n = Number(v);
  if (!isNaN(n) && n > 40000 && n < 60000) return new Date(Date.UTC(1899, 11, 30) + Math.floor(n) * 86400000).toISOString().slice(0, 10);
  let s = String(v).trim();
  if (s.indexOf('T') > -1) s = s.split('T')[0];
  if (s.indexOf(' ') > -1) s = s.split(' ')[0];
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return m[1] + '-' + gerP2(+m[2]) + '-' + gerP2(+m[3]);
  m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (m) return m[3] + '-' + gerP2(+m[2]) + '-' + gerP2(+m[1]);
  return '';
}
function gerTijd(v) {
  if (v === null || v === undefined || v === '' || v === 0) return '';
  if (typeof v === 'number') {
    const mins = Math.round((v - Math.floor(v)) * 1440) % 1440;
    return gerP2(Math.floor(mins / 60)) + ':' + gerP2(mins % 60);
  }
  const m = String(v).match(/(\d{1,2}):(\d{2})/);
  return m ? gerP2(+m[1]) + ':' + m[2] : '';
}
function gerMin(t) { const m = String(t || '').match(/^(\d{1,2}):(\d{2})$/); return m ? (+m[1]) * 60 + (+m[2]) : null; }
function gerDuur(van, tot) {
  const a = gerMin(van), b = gerMin(tot);
  if (a === null || b === null) return '';
  let d = b - a; if (d < 0) d += 1440;
  return d >= 60 ? (Math.floor(d / 60) + ' u ' + (d % 60) + ' min') : (d + ' min');
}

// ── Gegevens ophalen ──────────────────────────────────────────
async function gerLeesRegistratie() {
  const zoek = await graphGet(`/sites/${CONFIG.spHost}/drive/root/search(q='TaakRegistratie')`);
  const file = (zoek.value || []).find(f => f.name && f.name.toLowerCase().includes('taakregistratie') && f.name.toLowerCase().endsWith('.xlsx'));
  if (!file) throw new Error('TaakRegistratie.xlsx niet gevonden');
  const driveId = file.parentReference && file.parentReference.driveId;
  const ws = await graphGet(`/drives/${driveId}/items/${file.id}/workbook/worksheets`);
  const lijst = ws.value || [];
  const blad = lijst.find(s => s.name === 'Taken') || lijst[0];
  if (!blad) throw new Error('Geen werkblad gevonden in TaakRegistratie');
  const b = await graphGet(`/drives/${driveId}/items/${file.id}/workbook/worksheets/${encodeURIComponent(blad.name)}/usedRange`);
  const uit = [];
  (b.values || []).slice(1).forEach((r, i) => {
    const klant = String(r[1] || '').trim();
    if (!klant) return;
    const klaar = gerTijd(r[5]);
    const kol = String(r[10] || '').trim().toLowerCase();
    if (!(kol === 'klaar' || (!kol && klaar))) return;          // alleen gereden (afgeronde) taken
    uit.push({
      id: 'g' + i,
      klant, datum: gerDatumNorm(r[2]),
      onderweg: gerTijd(r[3]), bijKlant: gerTijd(r[4]), klaar,
      opm: String(r[9] === null || r[9] === undefined ? '' : r[9]).trim(),
    });
  });
  return uit;
}

// Bonnummer -> { lat, lon, klacht } uit "Reparatiebonnen plus locatie"
async function gerLeesReparatieCoords() {
  const map = {};
  try {
    if (typeof _repDriveId === 'undefined' || !_repDriveId || !_repItemId) await laadReparatiebonnen();
    const b = await graphGet(`/drives/${_repDriveId}/items/${_repItemId}/workbook/worksheets/Blad1/usedRange`);
    (b.values || []).slice(1).forEach(r => {
      if (!r[0]) return;
      const lat = parseFloat(String(r[10]).replace(',', '.'));
      const lon = parseFloat(String(r[11]).replace(',', '.'));
      map[String(r[0])] = { lat: isNaN(lat) ? null : lat, lon: isNaN(lon) ? null : lon, klacht: String(r[9] || '') };
    });
  } catch (e) { console.warn('Reparatielocaties niet geladen:', e); }
  return map;
}

function gerSamenstellen(reg, repMap) {
  const perSleutel = {};
  alleAfspraken.forEach(a => {
    const s = String(a.naam || '').toLowerCase().trim() + '|' + a.datum;
    (perSleutel[s] = perSleutel[s] || []).push(a);
  });
  return reg.map(r => {
    const kand = perSleutel[r.klant.toLowerCase().trim() + '|' + r.datum] || [];
    const afs = kand.find(a => /montage|reparatie/i.test(a.type || '')) || kand[0] || null;
    let type = 'Onbekend';
    if (afs && /reparatie/i.test(afs.type || '')) type = 'Reparatie';
    else if (afs && /montage/i.test(afs.type || '')) type = 'Montage';

    let bonnr = '', klacht = '', lat = null, lon = null;
    if (type === 'Reparatie' && afs) {
      const m = String(afs.opmerking || '').match(/#(\d{1,6})/);
      if (m && repMap[m[1]]) {
        bonnr = m[1];
        klacht = repMap[m[1]].klacht || '';
        lat = repMap[m[1]].lat; lon = repMap[m[1]].lon;
      } else if (m) bonnr = m[1];
    }
    return Object.assign({}, r, {
      type, adres: afs ? String(afs.adres || '') : '', monteurs: afs ? String(afs.monteurs || '') : '',
      bonnr, klacht, lat, lon,
    });
  }).sort((a, b) => b.datum.localeCompare(a.datum) || (b.klaar || '').localeCompare(a.klaar || ''));
}

async function gerLaad(vers) {
  if (gerBezig) return;
  gerBezig = true;
  const lijst = document.getElementById('ger-lijst');
  if (lijst && !gerRijen.length) lijst.innerHTML = '<div class="laad-scherm"><div class="spinner"></div><div>Gereden bonnen laden...</div></div>';
  try {
    if (!alleAfspraken.length) { try { await laadPlanning(); } catch (e) { console.warn('Planning niet geladen:', e); } }
    const [reg, repMap] = await Promise.all([gerLeesRegistratie(), gerLeesReparatieCoords()]);
    gerRijen = gerSamenstellen(reg, repMap);
    gerFout = '';
    gerGeladen = true;
  } catch (e) {
    gerFout = (typeof vertaalFout === 'function') ? vertaalFout(e) : (e.message || 'Laden mislukt');
  } finally {
    gerBezig = false;
  }
  gerPasPassen = true;
  gerRender();
}

// ── Filters ───────────────────────────────────────────────────
function gerZichtbaar() {
  const vandaag = new Date();
  let vanaf = '';
  if (gerPeriode === 'week') vanaf = gerIso(maandagVan(vandaag));
  else if (gerPeriode === '30') { const d = new Date(vandaag); d.setDate(d.getDate() - 30); vanaf = gerIso(d); }
  return gerRijen.filter(r => {
    if (gerType !== 'alle' && r.type !== gerType) return false;
    if (vanaf && r.datum < vanaf) return false;
    return true;
  });
}
function gerZetType(t, el) {
  gerType = t;
  document.querySelectorAll('#ger-type-chips .chip').forEach(c => c.classList.remove('actief'));
  if (el) el.classList.add('actief');
  gerActief = null; gerPasPassen = true; gerRender();
}
function gerZetPeriode(p, el) {
  gerPeriode = p;
  document.querySelectorAll('#ger-periode-chips .chip').forEach(c => c.classList.remove('actief'));
  if (el) el.classList.add('actief');
  gerActief = null; gerPasPassen = true; gerRender();
}

// ── Lijst ─────────────────────────────────────────────────────
function gerKaartHtml(r) {
  const kleur = GER_TYPE_KLEUR[r.type] || GER_TYPE_KLEUR.Onbekend;
  const rij = gerDuur(r.onderweg, r.bijKlant), werk = gerDuur(r.bijKlant, r.klaar);
  const tijden = [r.onderweg ? '🚐 ' + r.onderweg : '', r.bijKlant ? '🔧 ' + r.bijKlant : '', r.klaar ? '✅ ' + r.klaar : ''].filter(Boolean).join(' → ');
  const duren = [rij ? 'rijtijd ' + rij : '', werk ? 'werktijd ' + werk : ''].filter(Boolean).join(' · ');
  let loc;
  if (r.lat !== null && r.lon !== null) loc = '📍 ' + gerEsc(r.adres || 'Locatie bekend');
  else if (r.adres) loc = '📍 ' + gerEsc(r.adres) + ' <span style="color:#aaa">(locatie zoeken…)</span>';
  else loc = '<span style="color:#aaa">Geen adres gevonden in de planning</span>';
  const maps = r.adres ? '<a class="ger-route" href="https://maps.google.com/?q=' + encodeURIComponent(r.adres) + '" target="_blank" onclick="event.stopPropagation()">Route</a>' : '';
  return `<div class="card ger-kaart${gerActief === r.id ? ' actief' : ''}" id="ger-k-${r.id}" onclick="gerKies('${r.id}', false)">
    <div class="ger-kop"><span class="ger-stip" style="background:${kleur}"></span><div class="ger-naam">${gerEsc(r.klant)}</div>
      <span class="ger-type" style="background:${kleur}1A;color:${kleur}">${gerEsc(r.type)}</span></div>
    <div class="ger-sub">📅 ${gerEsc(gerDatumKort(r.datum))}${r.monteurs ? ' · 👷 ' + gerEsc(r.monteurs) : ''}</div>
    ${tijden ? `<div class="ger-sub">${tijden}${duren ? '<br>' + gerEsc(duren) : ''}</div>` : ''}
    ${r.bonnr ? `<div class="ger-sub">🧾 Bon #${gerEsc(r.bonnr)}${r.klacht ? ' — ' + gerEsc(r.klacht) : ''}</div>` : ''}
    ${r.opm ? `<div class="ger-opm">💬 ${gerEsc(r.opm)}</div>` : ''}
    <div class="ger-sub ger-loc">${loc} ${maps}</div>
  </div>`;
}

function gerRender() {
  const lijst = document.getElementById('ger-lijst');
  const sub = document.getElementById('ger-sub');
  if (!lijst) return;
  const items = gerZichtbaar();
  const opKaart = items.filter(r => r.lat !== null && r.lon !== null).length;
  if (sub) sub.textContent = `${items.length} gereden bon${items.length === 1 ? '' : 'nen'}` + (items.length ? ` · ${opKaart} op de kaart` : '');

  let html = '';
  if (gerFout) html += `<div class="leeg-dag" style="color:#c0392b;padding:14px 18px">⚠️ ${gerEsc(gerFout)}</div>`;
  if (!items.length) {
    html += gerFout ? '' : '<div class="leeg-dag">Geen gereden bonnen in deze periode</div>';
  } else {
    let laatste = '';
    items.forEach(r => {
      if (r.datum !== laatste) { laatste = r.datum; html += `<div class="dag-label" style="margin-top:12px">${gerEsc(gerDagNaam(r.datum))}</div>`; }
      html += gerKaartHtml(r);
    });
  }
  lijst.innerHTML = html;

  gerZetMarkers(items);
  gerGeocodeWachtrij(items);
}

// ── Kaart (Leaflet) ───────────────────────────────────────────
function gerLaadLeaflet(klaar) {
  if (window.L) { klaar(); return; }
  if (!document.getElementById('leaflet-css')) {
    const l = document.createElement('link');
    l.id = 'leaflet-css'; l.rel = 'stylesheet'; l.href = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css';
    document.head.appendChild(l);
  }
  let s = document.getElementById('leaflet-js');
  if (!s) {
    s = document.createElement('script');
    s.id = 'leaflet-js'; s.src = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js';
    s.onerror = () => { const k = document.getElementById('ger-kaart'); if (k) k.innerHTML = '<div class="leeg-dag">Kaart kon niet worden geladen (geen internet?)</div>'; };
    document.head.appendChild(s);
  }
  s.addEventListener('load', klaar);
}

function gerKaartInit() {
  if (gerKaart || !window.L) return;
  const el = document.getElementById('ger-kaart');
  if (!el) return;
  gerKaart = L.map('ger-kaart', { zoomControl: true, attributionControl: true }).setView(GER_THUIS, 10);
  L.tileLayer('https://basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png', {
    attribution: '© OpenStreetMap © CARTO', maxZoom: 19,
  }).addTo(gerKaart);
  const thuis = L.divIcon({
    className: '',
    html: '<div style="width:0;height:0;border-left:10px solid transparent;border-right:10px solid transparent;border-bottom:18px solid #27ae60;filter:drop-shadow(0 2px 4px rgba(0,0,0,.5))"></div>',
    iconSize: [20, 18], iconAnchor: [10, 18],
  });
  gerThuisMarker = L.marker(GER_THUIS, { icon: thuis, zIndexOffset: 9999 }).bindTooltip('Rozal, Hedel').addTo(gerKaart);
  gerKaart.on('click', () => { gerActief = null; gerRender(); });
}

function gerZetMarkers(items) {
  if (!window.L || !gerKaart) return;
  Object.keys(gerMarkers).forEach(id => { try { gerKaart.removeLayer(gerMarkers[id]); } catch (e) { /* weg */ } });
  gerMarkers = {};
  const punten = [];
  items.forEach(r => {
    if (r.lat === null || r.lon === null) return;
    const actief = gerActief === r.id;
    const kleur = GER_TYPE_KLEUR[r.type] || GER_TYPE_KLEUR.Onbekend;
    const maat = actief ? 22 : 15;
    const icon = L.divIcon({
      className: '',
      html: `<div style="width:${maat}px;height:${maat}px;background:${kleur};border:${actief ? 3 : 2}px solid #fff;border-radius:50%;box-shadow:0 2px 6px rgba(0,0,0,.5)"></div>`,
      iconSize: [maat, maat], iconAnchor: [maat / 2, maat / 2],
    });
    const m = L.marker([r.lat, r.lon], { icon });
    m.bindTooltip(gerEsc(r.klant) + ' · ' + gerEsc(gerDatumKort(r.datum)), { direction: 'top', offset: [0, -maat / 2] });
    m.on('click', e => { L.DomEvent.stopPropagation(e); gerKies(r.id, true); });
    m.addTo(gerKaart);
    gerMarkers[r.id] = m;
    punten.push([r.lat, r.lon]);
  });
  if (gerPasPassen && punten.length) {
    gerPasPassen = false;
    try { gerKaart.fitBounds(L.latLngBounds(punten.concat([GER_THUIS])), { padding: [30, 30], maxZoom: 13 }); } catch (e) { /* negeren */ }
  }
}

function gerKies(id, vanKaart) {
  gerActief = id;
  const r = gerRijen.find(x => x.id === id);
  gerRender();
  if (!r) return;
  if (!vanKaart && gerKaart && r.lat !== null && r.lon !== null) {
    gerKaart.setView([r.lat, r.lon], Math.max(gerKaart.getZoom(), 13), { animate: true });
    const kaartEl = document.getElementById('ger-kaart');
    if (kaartEl && kaartEl.scrollIntoView) kaartEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
  if (vanKaart) {
    const el = document.getElementById('ger-k-' + id);
    if (el && el.scrollIntoView) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }
}

// ── Adressen opzoeken (Nominatim), één voor één, met onthouden resultaat ──
function gerGeoSleutel(adres) { return String(adres || '').toLowerCase().replace(/\s+/g, ' ').trim(); }

function gerGeocodeWachtrij(items) {
  items.forEach(r => {
    if (r.lat !== null && r.lon !== null) return;
    const s = gerGeoSleutel(r.adres);
    if (!s) return;
    const c = gerGeoCache[s];
    if (c && c.length === 2) { r.lat = c[0]; r.lon = c[1]; return; }
    if (c === 0) return;                                   // eerder niet gevonden: niet opnieuw proberen
    if (gerGeoWachtrij.indexOf(s) === -1) gerGeoWachtrij.push(s);
  });
  // rijen die net uit de cache kwamen: pins tekenen
  gerZetMarkers(items);
  gerGeoStart();
}

async function gerGeoStart() {
  if (gerGeoBezig) return;
  gerGeoBezig = true;
  try {
    while (gerGeoWachtrij.length) {
      const s = gerGeoWachtrij.shift();
      let uitkomst = 0;
      try {
        const url = 'https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=nl&q=' + encodeURIComponent(s + ', Nederland');
        const resp = await fetch(url, { headers: { 'Accept-Language': 'nl' } });
        if (resp.ok) {
          const d = await resp.json();
          if (d && d.length) uitkomst = [parseFloat(d[0].lat), parseFloat(d[0].lon)];
        } else if (resp.status === 429) { gerGeoWachtrij.unshift(s); await new Promise(r => setTimeout(r, 5000)); continue; }
      } catch (e) { uitkomst = null; }
      if (uitkomst !== null) {                              // bij netwerkfout niets onthouden
        gerGeoCache[s] = uitkomst;
        try { localStorage.setItem(GER_GEO_KEY, JSON.stringify(gerGeoCache)); } catch (e) { /* vol: negeren */ }
        gerRijen.forEach(r => { if (gerGeoSleutel(r.adres) === s && uitkomst) { r.lat = uitkomst[0]; r.lon = uitkomst[1]; } });
        gerRender();
      }
      await new Promise(r => setTimeout(r, 1100));          // Nominatim: max. 1 verzoek per seconde
    }
  } finally {
    gerGeoBezig = false;
  }
}

// ── Tab openen / verversen ────────────────────────────────────
function gerOpenTab() {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-item').forEach(t => t.classList.remove('active'));
  document.getElementById('page-gereden').classList.add('active');
  document.getElementById('tab-gereden').classList.add('active');
  const sc = document.querySelector('.screen-content'); if (sc) sc.scrollTop = 0;
  const av = document.getElementById('gebruiker-avatar'), mijnAv = document.querySelector('#page-gereden .avatar');
  if (av && mijnAv) mijnAv.textContent = av.textContent;
  gerLaadLeaflet(() => {
    gerKaartInit();
    setTimeout(() => { if (gerKaart) { gerKaart.invalidateSize(); gerPasPassen = true; gerRender(); } }, 150);
  });
  if (!gerGeladen) gerLaad(false); else gerRender();
}

async function gerVerversKnop(el) {
  el.classList.add('laden');
  try { await gerLaad(true); }
  finally { el.classList.remove('laden'); }
}

// ── Zelfstandig: eigen tab, pagina en opmaak ──────────────────
(function() {
  const css = `
  #ger-kaart { height:300px; border-radius:12px; border:1px solid #E0E8D0; margin:0 0 10px; overflow:hidden; background:#eef2e8; position:relative; z-index:0; }
  .ger-kaart { display:block; cursor:pointer; }
  .ger-kaart.actief { border-color:#2D5016; box-shadow:0 0 0 2px rgba(45,80,22,.25); }
  .ger-kop { display:flex; align-items:center; gap:8px; }
  .ger-stip { width:9px; height:9px; border-radius:50%; flex-shrink:0; }
  .ger-naam { font-size:14px; font-weight:700; flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .ger-type { font-size:10.5px; font-weight:700; padding:3px 9px; border-radius:20px; flex-shrink:0; }
  .ger-sub { font-size:12px; color:#777; margin-top:4px; line-height:1.4; }
  .ger-opm { font-size:12px; margin-top:6px; background:#f6f8f2; border-radius:8px; padding:6px 9px; }
  .ger-route { color:#2D5016; font-weight:700; text-decoration:none; margin-left:6px; }
  .tab-label { white-space:nowrap; }
  .tab-bar .tab-label { font-size:10px; }`;
  const st = document.createElement('style');
  st.textContent = css;
  document.head.appendChild(st);

  function bouw() {
    const sc = document.querySelector('.screen-content');
    const bar = document.querySelector('.tab-bar');
    if (!sc || !bar || document.getElementById('page-gereden')) return;

    const avatar = document.getElementById('gebruiker-avatar');
    const page = document.createElement('div');
    page.className = 'page';
    page.id = 'page-gereden';
    page.innerHTML = `
      <div class="app-topbar">
        <div><div class="app-titel">Gereden bonnen</div><div class="app-sub" id="ger-sub">—</div></div>
        <div class="topbar-rechts">
          <div class="refresh-btn" id="refresh-gereden" onclick="gerVerversKnop(this)" title="Vernieuwen">⟳</div>
          <div class="avatar" onclick="uitloggen()" title="Uitloggen">${avatar ? gerEsc(avatar.textContent) : ''}</div>
        </div>
      </div>
      <div class="chip-filter" id="ger-type-chips">
        <div class="chip actief" onclick="gerZetType('alle', this)">Alle</div>
        <div class="chip" onclick="gerZetType('Montage', this)">Montage</div>
        <div class="chip" onclick="gerZetType('Reparatie', this)">Reparatie</div>
      </div>
      <div class="chip-filter" id="ger-periode-chips">
        <div class="chip" onclick="gerZetPeriode('week', this)">Deze week</div>
        <div class="chip actief" onclick="gerZetPeriode('30', this)">30 dagen</div>
        <div class="chip" onclick="gerZetPeriode('alles', this)">Alles</div>
      </div>
      <div class="dag-blok" style="padding-bottom:0"><div id="ger-kaart"><div class="leeg-dag">Kaart laden…</div></div></div>
      <div class="dag-blok" style="padding-top:0" id="ger-lijst"><div class="leeg-dag">Tik op de tab om de gereden bonnen te laden</div></div>`;
    sc.appendChild(page);

    const tab = document.createElement('div');
    tab.className = 'tab-item';
    tab.id = 'tab-gereden';
    tab.onclick = gerOpenTab;
    tab.innerHTML = '<div class="tab-icon">🚐</div><div class="tab-label">Gereden</div>';
    const na = document.getElementById('tab-reparaties');
    if (na && na.nextSibling) bar.insertBefore(tab, na.nextSibling); else bar.appendChild(tab);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bouw);
  else bouw();
})();
