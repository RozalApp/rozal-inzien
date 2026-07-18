// ══════════════════════════════════════════════════════════════
// APP — UI-weergave voor Planning, Klanten en Reparatiebonnen
// ══════════════════════════════════════════════════════════════

const KLEUR  = { Inmeten:'#E8820C', Montage:'#2D5016', Reparatie:'#C0392B', Overig:'#2F6FAE' };
const CHIPBG = { Inmeten:['#FEF3E2','#E8820C'], Montage:['#EAF0E1','#2D5016'], Reparatie:['#fde8e6','#C0392B'], Overig:['#E7F0F8','#2F6FAE'] };
const STATUS_KLEUR = {
  open:        ['#fde8e6', '#C0392B'],
  ingepland:   ['#FEF3E2', '#E8820C'],
  klaar:       ['#E7F0F8', '#2F6FAE'],
  gefactureerd:['#EAF0E1', '#2D5016'],
};
const DAG_NAMEN = ['Ma','Di','Wo','Do','Vr','Za','Zo'];
const DAG_VOLNAAM = ['Maandag','Dinsdag','Woensdag','Donderdag','Vrijdag','Zaterdag','Zondag'];
const MAAND_NAMEN = ['januari','februari','maart','april','mei','juni','juli','augustus','september','oktober','november','december'];

let huidigeMaandag = maandagVan(new Date());
let geselecteerdeDag = new Date();
let repStatusFilter = 'alle';
let repZoekterm = '';

function maandagVan(d) {
  const dt = new Date(d); const dag = (dt.getDay() + 6) % 7;
  dt.setDate(dt.getDate() - dag); dt.setHours(0,0,0,0); return dt;
}
function isoDatum(d) { return d.toISOString().substring(0,10); }
function zelfdeD(a, b) { return isoDatum(a) === isoDatum(b); }

// ── Opstarten na login ──────────────────────────────────────
async function startApp() {
  document.getElementById('app-shell').classList.add('active');
  document.getElementById('login-shell').style.display = 'none';

  await Promise.allSettled([laadPlanning(), laadKlanten(), laadReparatiebonnen()]);
  renderWeekStrip();
  renderKlanten();
  renderReparatiebonnen();
}

// ── Tabs ──────────────────────────────────────────────────
function switchTab(naam) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-item').forEach(t => t.classList.remove('active'));
  document.getElementById('page-' + naam).classList.add('active');
  document.getElementById('tab-' + naam).classList.add('active');
  document.querySelector('.screen-content').scrollTop = 0;
}

// ── PLANNING: weekstrip + dagoverzicht ─────────────────────
function renderWeekStrip() {
  const maandLabel = MAAND_NAMEN[huidigeMaandag.getMonth()] + ' ' + huidigeMaandag.getFullYear();
  document.getElementById('pl-maand-label').textContent = maandLabel.charAt(0).toUpperCase() + maandLabel.slice(1);

  const vandaag = new Date();
  const strip = document.getElementById('week-strip');
  strip.innerHTML = '';
  for (let i = 0; i < 7; i++) {
    const d = new Date(huidigeMaandag); d.setDate(d.getDate() + i);
    const heeftItems = alleAfspraken.some(a => a.datum === isoDatum(d));
    const actief = zelfdeD(d, geselecteerdeDag);
    const isVandaag = zelfdeD(d, vandaag);
    const el = document.createElement('div');
    el.className = 'week-dag' + (isVandaag ? ' vandaag' : '') + (actief ? ' actief' : '');
    el.innerHTML = `<div class="week-dag-naam">${DAG_NAMEN[i]}</div><div class="week-dag-num">${d.getDate()}</div>${heeftItems ? '<div class="week-dag-stip"></div>' : ''}`;
    el.onclick = () => { geselecteerdeDag = d; renderWeekStrip(); };
    strip.appendChild(el);
  }
  renderDagOverzicht();
}
function verschuifWeek(richting) {
  huidigeMaandag.setDate(huidigeMaandag.getDate() + richting * 7);
  huidigeMaandag = new Date(huidigeMaandag);
  renderWeekStrip();
}
function gaNaarVandaag() {
  huidigeMaandag = maandagVan(new Date());
  geselecteerdeDag = new Date();
  renderWeekStrip();
}

function renderDagOverzicht() {
  const dagIdx = (geselecteerdeDag.getDay() + 6) % 7;
  document.getElementById('dag-label').textContent = `${DAG_VOLNAAM[dagIdx]} ${geselecteerdeDag.getDate()} ${MAAND_NAMEN[geselecteerdeDag.getMonth()]}`;

  const datumIso = isoDatum(geselecteerdeDag);
  const items = alleAfspraken.filter(a => a.datum === datumIso).sort((a,b) => (a.uur*60+a.minuut) - (b.uur*60+b.minuut));
  const el = document.getElementById('dag-items');
  if (!items.length) { el.innerHTML = '<div class="leeg-dag">Geen afspraken op deze dag</div>'; return; }

  el.innerHTML = items.map((a, i) => {
    const kleur = KLEUR[a.type] || KLEUR.Overig;
    const chip = CHIPBG[a.type] || CHIPBG.Overig;
    const plaats = (a.adres || '').split(',').pop().trim();
    return `<div class="card" data-idx="${i}" onclick="toonAfspraakDetail(${i})">
      <div class="plan-tijd">${pad2(a.uur)}:${pad2(a.minuut)}</div>
      <div class="plan-stip" style="background:${kleur}"></div>
      <div class="plan-info"><div class="plan-naam">${escapeHtml(a.naam || a.type)}</div><div class="plan-detail">${escapeHtml(a.opmerking || a.type)}${plaats ? ' · ' + escapeHtml(plaats) : ''}</div></div>
      <div class="plan-monteur" style="background:${chip[0]};color:${chip[1]}">${escapeHtml(a.monteurs || '—')}</div>
    </div>`;
  }).join('');
  el._items = items;
  window._huidigeDagItems = items;
}

function toonAfspraakDetail(i) {
  const a = window._huidigeDagItems[i];
  const kleur = KLEUR[a.type] || KLEUR.Overig;
  document.getElementById('sh-header').style.background = kleur;
  document.getElementById('sh-naam').textContent = a.naam || a.type;
  document.getElementById('sh-sub').textContent = `${pad2(a.uur)}:${pad2(a.minuut)} · ${a.type}`;

  let body = `
    <div class="detail-veld"><div class="detail-label">Adres</div><div class="detail-waarde">${a.adres ? `<a href="https://maps.google.com/?q=${encodeURIComponent(a.adres)}" target="_blank">${escapeHtml(a.adres)}</a>` : '—'}</div></div>
    <div class="detail-veld"><div class="detail-label">Monteur</div><div class="detail-waarde">${escapeHtml(a.monteurs || '—')}</div></div>
    <div class="detail-veld"><div class="detail-label">Opmerking</div><div class="detail-waarde">${escapeHtml((a.opmerking || '').replace(/\s*\[pdf:[^\]]+\]/, '')) || '—'}</div></div>`;

  const pdfMatch = (a.opmerking || '').match(/\[pdf:([^\]]+)\]/);
  if (pdfMatch) {
    body += `<div class="pdf-btn" id="sh-pdf-btn" onclick="openBonPdfKnop('${pdfMatch[1].replace(/'/g,"\\'")}', 'sh-pdf-btn')">📄 Reparatiebon openen</div>`;
  }
  document.getElementById('sh-body').innerHTML = body;
  document.getElementById('afspraak-sheet').classList.add('open');
  document.getElementById('sheet-backdrop').classList.add('open');
}
function sluitSheet() {
  document.getElementById('afspraak-sheet').classList.remove('open');
  document.getElementById('sheet-backdrop').classList.remove('open');
}

// ── KLANTEN ─────────────────────────────────────────────────
function renderKlanten() {
  document.getElementById('kl-aantal').textContent = alleKlanten.length + ' klanten';
  const term = (document.getElementById('kl-zoek')?.value || '').toLowerCase().trim();
  const lijst = term
    ? alleKlanten.filter(k => klantVolledigeNaam(k).toLowerCase().includes(term) || (k.plaats||'').toLowerCase().includes(term) || (k.telefoon||'').includes(term))
    : alleKlanten;

  const el = document.getElementById('klanten-lijst');
  if (!lijst.length) { el.innerHTML = '<div class="leeg-melding">Geen klanten gevonden</div>'; return; }
  el.innerHTML = lijst.slice(0, 200).map((k, i) => {
    const naam = klantVolledigeNaam(k);
    const init = naam.split(' ').filter(Boolean).slice(0,2).map(w => w[0]).join('').toUpperCase();
    const idx = alleKlanten.indexOf(k);
    return `<div class="card" onclick="toonKlantDetail(${idx})">
      <div class="klant-avatar">${init || '?'}</div>
      <div class="plan-info"><div class="klant-naam">${escapeHtml(naam)}</div><div class="klant-adres">${escapeHtml([k.adres, k.plaats].filter(Boolean).join(', ') || '—')}</div></div>
      <div class="klant-pijl">›</div>
    </div>`;
  }).join('');
}
function klZoekOnInput() { renderKlanten(); }

async function toonKlantDetail(idx) {
  const k = alleKlanten[idx];
  document.getElementById('ok-naam').textContent = klantVolledigeNaam(k);
  document.getElementById('ok-adres').textContent = [k.adres, k.postcode, k.plaats].filter(Boolean).join(', ') || '—';
  document.getElementById('ok-tel').textContent = k.telefoon || '—';
  document.getElementById('ok-email').textContent = k.email || '—';
  document.getElementById('ok-bestanden').innerHTML = '<div class="laad-scherm" style="padding:20px 0"><div class="spinner"></div><div>Bestanden zoeken...</div></div>';
  document.getElementById('overlay-klant').classList.add('open');

  try {
    const bestanden = await zoekBestandenVoorKlant(k);
    const iconMap = { Reparatiebonnen:['🔧','#fde8e6'], Offertes:['📄','#FEF3E2'], Facturen:['🧾','#EAF0E1'] };
    if (!bestanden.length) {
      document.getElementById('ok-bestanden').innerHTML = '<div class="leeg-melding">Geen bestanden gevonden</div>';
      return;
    }
    document.getElementById('ok-bestanden').innerHTML = bestanden.map(b => {
      const icon = iconMap[b.type] || ['📄', '#EAF0E1'];
      const datumStr = b.datum ? new Date(b.datum).toLocaleDateString('nl-NL') : '';
      return `<div class="card" onclick="window.open('${b.webUrl}', '_blank')">
        <div class="bestand-icon" style="background:${icon[1]}">${icon[0]}</div>
        <div class="plan-info"><div class="bestand-naam">${escapeHtml(b.naam)}</div><div class="bestand-sub">${b.type}${datumStr ? ' · ' + datumStr : ''}</div></div>
        <div class="klant-pijl">›</div>
      </div>`;
    }).join('');
  } catch (e) {
    document.getElementById('ok-bestanden').innerHTML = '<div class="leeg-melding">Zoeken mislukt: ' + escapeHtml(vertaalFout(e)) + '</div>';
  }
}

// ── REPARATIEBONNEN ─────────────────────────────────────────
function renderReparatiebonnen() {
  const openCount = alleReparaties.filter(r => !['klaar','gefactureerd'].includes((r.status||'').toLowerCase())).length;
  document.getElementById('rep-aantal').textContent = openCount + ' openstaand';

  let lijst = alleReparaties;
  if (repStatusFilter !== 'alle') lijst = lijst.filter(r => (r.status||'open').toLowerCase() === repStatusFilter);
  if (repZoekterm) {
    const t = repZoekterm.toLowerCase();
    lijst = lijst.filter(r => `${r.voornaam} ${r.achternaam} ${r.plaats} ${r.bonnr}`.toLowerCase().includes(t));
  }

  const el = document.getElementById('bon-lijst');
  if (!lijst.length) { el.innerHTML = '<div class="leeg-melding">Geen reparatiebonnen gevonden</div>'; return; }
  el.innerHTML = lijst.map(r => {
    const statusKey = (r.status || 'open').toLowerCase();
    const kleur = STATUS_KLEUR[statusKey] || STATUS_KLEUR.open;
    const naam = `${r.voornaam} ${r.achternaam}`.trim() || 'Onbekende klant';
    return `<div class="card rep-card" onclick="toonBonDetail('${r.bonnr}')">
      <div class="rep-top"><div class="rep-bonnr">#${r.bonnr}</div><div class="rep-status" style="background:${kleur[0]};color:${kleur[1]}">${escapeHtml(r.status || 'Open')}</div></div>
      <div class="rep-klant">${escapeHtml(naam)}${r.plaats ? ' · ' + escapeHtml(r.plaats) : ''}</div>
      <div class="rep-klacht">${escapeHtml(r.omschrijving || '—')}</div>
    </div>`;
  }).join('');
}
function repZetFilter(status, el) {
  repStatusFilter = status;
  document.querySelectorAll('#rep-chips .chip').forEach(c => c.classList.remove('actief'));
  el.classList.add('actief');
  renderReparatiebonnen();
}
function repZoekOnInput(v) { repZoekterm = v; renderReparatiebonnen(); }

function toonBonDetail(bonnr) {
  const r = alleReparaties.find(x => x.bonnr === bonnr);
  if (!r) return;
  const naam = `${r.voornaam} ${r.achternaam}`.trim() || 'Onbekende klant';
  document.getElementById('ob-titel').textContent = 'Bon #' + bonnr;
  document.getElementById('ob-klant').textContent = naam;
  document.getElementById('ob-adres').textContent = [r.adres, r.postcode, r.plaats].filter(Boolean).join(', ') || '—';
  document.getElementById('ob-klacht').textContent = r.omschrijving || '—';
  const statusKey = (r.status || 'open').toLowerCase();
  const kleur = STATUS_KLEUR[statusKey] || STATUS_KLEUR.open;
  document.getElementById('ob-status').innerHTML = `<span style="background:${kleur[0]};color:${kleur[1]};padding:3px 10px;border-radius:20px;font-size:11.5px;font-weight:700">${escapeHtml(r.status || 'Open')}</span>`;
  document.getElementById('ob-telefoon').textContent = r.telefoon || '—';

  const btn = document.getElementById('ob-pdf-btn');
  btn.textContent = '📄 Reparatiebon PDF openen';
  btn.classList.remove('laden');
  btn.onclick = () => openBonPdfKnop(bonnr, 'ob-pdf-btn');
  document.getElementById('overlay-bon').classList.add('open');
}

async function openBonPdfKnop(bonnrOfBestand, btnId) {
  const btn = document.getElementById(btnId);
  const origineel = btn.textContent;
  btn.textContent = '⏳ Bon zoeken...'; btn.classList.add('laden');
  try {
    await zoekEnOpenPdf('Reparatiebonnen', bonnrOfBestand);
  } catch (e) {
    alert('Kon de bon niet openen: ' + vertaalFout(e));
  } finally {
    btn.textContent = origineel; btn.classList.remove('laden');
  }
}

function sluitOverlay(id) { document.getElementById(id).classList.remove('open'); }
function escapeHtml(s) { return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
