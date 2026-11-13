// ══════════════════════════════════════════════════════════════
// BOEKINGEN — online boekingen / showroomafspraken bekijken en
// "as zaterdag" dichtzetten voor online boeken.
//
// • De boekingen komen uit de planning (online boekingen staan daar als
//   concept-afspraak met "ONLINE BOEKING" in de opmerking).
// • Dichtzetten gaat via dezelfde Azure Function als de Windows-app
//   (Instellingen → Boekingbeheer → geblokkeerde periodes).
// ══════════════════════════════════════════════════════════════

const BOEKING_URL = 'https://rozal-booking-fn-ehbcbmdeb8h4e7dj.westeurope-01.azurewebsites.net/api/rozal-booking';
let boekFilter = 'komend';          // 'komend' | 'eerder'
let boekZaterdagBezig = false;

// ── Hulpfuncties ──────────────────────────────────────────────
function boekDagNaam(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  const dt = new Date(y, m - 1, d);
  return `${DAG_VOLNAAM[(dt.getDay() + 6) % 7]} ${d} ${MAAND_NAMEN[m - 1]}`;
}
function boekKort(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  return `${d} ${MAAND_NAMEN[m - 1].slice(0, 3)}`;
}
function boekIsBoeking(a) {
  return (a.opmerking || '').toUpperCase().includes('ONLINE BOEKING') || /showroom/i.test(a.type || '');
}

// Zaterdag van de week (ma t/m zo) waarin we nu zitten, als 'yyyy-mm-dd'
function boekZaterdagDezeWeek() {
  const nu = new Date();
  const maandag = maandagVan(nu);
  const za = new Date(maandag);
  za.setDate(maandag.getDate() + 5);
  return isoDatum(za);
}

// Datum uit de geblokkeerd-lijst (Excel-serienummer, dd-mm-yyyy of ISO) -> 'yyyy-mm-dd'
function boekBlokIso(v) {
  if (v === null || v === undefined || v === '') return '';
  const s = String(v).trim();
  if (/^\d{5}$/.test(s)) return new Date(Math.round((parseInt(s, 10) - 25569) * 86400000)).toISOString().slice(0, 10);
  const m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (m) return `${m[3]}-${pad2(m[2])}-${pad2(m[1])}`;
  return s.slice(0, 10);
}

// ── Tab openen / verversen ────────────────────────────────────
function boekingenTabOpen() {
  boekingenRender();
  boekingenControleerZaterdag();
}

async function boekingenVervers() {
  await laadPlanning();
  renderWeekStrip();
  boekingenRender();
  boekingenControleerZaterdag();
}

function boekingenZetFilter(f, el) {
  boekFilter = f;
  document.querySelectorAll('#boek-chips .chip').forEach(c => c.classList.remove('actief'));
  if (el) el.classList.add('actief');
  boekingenRender();
}

// ── Lijst met boekingen ───────────────────────────────────────
function boekingenRender() {
  const lijst = document.getElementById('boek-lijst');
  const sub = document.getElementById('boek-sub');
  if (!lijst) return;
  const vandaag = isoDatum(new Date());

  const alle = alleAfspraken.filter(boekIsBoeking);
  const komend = alle.filter(a => a.datum >= vandaag)
    .sort((a, b) => a.datum.localeCompare(b.datum) || (a.uur * 60 + a.minuut) - (b.uur * 60 + b.minuut));
  const eerder = alle.filter(a => a.datum < vandaag)
    .sort((a, b) => b.datum.localeCompare(a.datum) || (b.uur * 60 + b.minuut) - (a.uur * 60 + a.minuut))
    .slice(0, 40);

  sub.textContent = `${komend.length} komend${komend.length === 1 ? 'e boeking' : 'e boekingen'}`;
  const items = boekFilter === 'komend' ? komend : eerder;
  if (!items.length) {
    lijst.innerHTML = `<div class="leeg-dag">${boekFilter === 'komend' ? 'Geen komende boekingen' : 'Geen eerdere boekingen'}</div>`;
    return;
  }

  let html = '', laatsteDag = '';
  items.forEach(a => {
    if (a.datum !== laatsteDag) {
      laatsteDag = a.datum;
      html += `<div class="dag-label" style="margin-top:14px">${escapeHtml(boekDagNaam(a.datum))}${a.datum === vandaag ? ' · vandaag' : ''}</div>`;
    }
    html += boekingKaartHtml(a);
  });
  lijst.innerHTML = html;
}

// Opmerking opdelen in regels (splitsen op | en nieuwe regels), "ONLINE BOEKING" zelf weglaten
function boekingRegels(opm) {
  return String(opm || '').split(/\||\n/).map(s => s.trim()).filter(s => s && !/^online boeking$/i.test(s));
}

function boekingKaartHtml(a) {
  const online = (a.opmerking || '').toUpperCase().includes('ONLINE BOEKING');
  const kleur = a.type === 'Inmeten' ? KLEUR.Inmeten : '#2F6FAE';
  const regels = boekingRegels(a.opmerking).map(r => {
    const mail = r.match(/^(e-?mail)\s*:\s*(\S+@\S+)/i);
    const tel = r.match(/^(tel(?:efoon)?(?:nummer)?)\s*:\s*([+\d][\d\s\-]{6,})/i);
    if (mail) return `<div class="boek-regel"><span>E-mail</span><a href="mailto:${escapeHtml(mail[2])}">${escapeHtml(mail[2])}</a></div>`;
    if (tel) return `<div class="boek-regel"><span>Telefoon</span><a href="tel:${escapeHtml(tel[2].replace(/\s/g, ''))}">${escapeHtml(tel[2].trim())}</a></div>`;
    const kv = r.match(/^([^:]{2,24}):\s*(.+)$/);
    if (kv) return `<div class="boek-regel"><span>${escapeHtml(kv[1])}</span><div>${escapeHtml(kv[2])}</div></div>`;
    return `<div class="boek-regel"><div>${escapeHtml(r)}</div></div>`;
  }).join('');
  const detailAdres = a.adres ? `<div class="boek-regel"><span>Adres</span><div>${escapeHtml(a.adres)}</div></div>` : '';

  return `<div class="card boek-kaart">
    <div class="boek-kop">
      <div class="plan-tijd" style="width:auto;margin-right:2px">${pad2(a.uur)}:${pad2(a.minuut)}</div>
      <div class="plan-info"><div class="plan-naam">${escapeHtml(a.naam || 'Onbekend')}</div>
      <div class="plan-detail" style="white-space:normal">${escapeHtml(a.type || 'Afspraak')}${online ? ' · online geboekt' : ''}${a.definitief ? ' · definitief' : ' · concept'}</div></div>
      <div class="plan-monteur" style="background:${online ? '#E7F0F8' : '#EAF0E1'};color:${kleur}">${online ? 'online' : 'showroom'}</div>
    </div>
    ${(regels || detailAdres) ? `<div class="boek-details">${detailAdres}${regels}</div>` : ''}
  </div>`;
}

// ── "As zaterdag" dichtzetten ─────────────────────────────────
async function boekHaalGeblokkeerd() {
  const resp = await fetch(BOEKING_URL + '?actie=haal-instellingen');
  if (!resp.ok) throw new Error('Boekingsdienst gaf fout ' + resp.status);
  const data = await resp.json();
  if (!data || !Array.isArray(data.geblokkeerd)) throw new Error('Onverwacht antwoord van de boekingsdienst');
  return data.geblokkeerd;
}

async function boekingenControleerZaterdag() {
  const iso = boekZaterdagDezeWeek();
  const knop = document.getElementById('boek-za-knop');
  const status = document.getElementById('boek-za-status');
  if (!knop || !status) return;
  knop.textContent = `🚫 As zaterdag (${boekKort(iso)}) dichtzetten`;
  const vandaag = isoDatum(new Date());
  if (iso < vandaag) { status.textContent = 'Deze zaterdag is al geweest.'; knop.classList.add('uit'); return; }
  knop.classList.remove('uit');
  status.textContent = 'Status controleren…';
  try {
    const lijst = await boekHaalGeblokkeerd();
    const dicht = lijst.some(g => g.type === 'dag' && boekBlokIso(g.datum) === iso);
    status.textContent = dicht ? `✅ Zaterdag ${boekKort(iso)} staat dicht voor online boeken.` : `Zaterdag ${boekKort(iso)} staat open voor online boeken.`;
    knop.classList.toggle('uit', dicht);
  } catch (e) {
    status.textContent = '';
  }
}

async function boekingenZaterdagDicht() {
  if (boekZaterdagBezig) return;
  const iso = boekZaterdagDezeWeek();
  const vandaag = isoDatum(new Date());
  if (iso < vandaag) { alert('Deze zaterdag is al geweest, daar valt niets meer dicht te zetten.'); return; }

  // Staan er al boekingen op die dag? Die blijven staan — dat zeggen we erbij.
  const reedsGeboekt = alleAfspraken.filter(a => a.datum === iso && boekIsBoeking(a)).length;
  let vraag = `Weet je het zeker?\n\nZaterdag ${boekDagNaam(iso)} wordt de hele dag dichtgezet voor online boeken.`;
  if (reedsGeboekt) vraag += `\n\nLet op: er staan al ${reedsGeboekt} boeking${reedsGeboekt === 1 ? '' : 'en'} op die dag. Die blijven staan.`;
  if (!confirm(vraag)) return;

  boekZaterdagBezig = true;
  const knop = document.getElementById('boek-za-knop');
  const status = document.getElementById('boek-za-status');
  knop.classList.add('uit');
  status.textContent = 'Bezig met dichtzetten…';
  try {
    // 1. Huidige lijst ophalen (de dienst vervangt bij opslaan de hele lijst)
    const lijst = await boekHaalGeblokkeerd();
    if (lijst.some(g => g.type === 'dag' && boekBlokIso(g.datum) === iso)) {
      status.textContent = `✅ Zaterdag ${boekKort(iso)} stond al dicht.`;
      return;
    }
    // 2. Zaterdag toevoegen en de hele lijst terugsturen
    const nieuw = lijst.concat([{ type: 'dag', datum: iso, tijd: '', omschrijving: 'Zaterdag dicht' }]);
    const resp = await fetch(BOEKING_URL + '?actie=sla-instellingen', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ actie: 'sla-instellingen', type: 'geblokkeerd', data: nieuw }),
    });
    if (!resp.ok) throw new Error('Opslaan mislukt (' + resp.status + ')');
    // 3. Controleren of het echt is opgeslagen
    const controle = await boekHaalGeblokkeerd();
    if (!controle.some(g => g.type === 'dag' && boekBlokIso(g.datum) === iso)) throw new Error('Zaterdag stond na opslaan niet in de lijst');
    status.textContent = `✅ Zaterdag ${boekKort(iso)} is dichtgezet voor online boeken.`;
    alert(`Zaterdag ${boekDagNaam(iso)} is dichtgezet.`);
  } catch (e) {
    console.error('Zaterdag dichtzetten mislukt:', e);
    knop.classList.remove('uit');
    status.textContent = '⚠️ Dichtzetten mislukt.';
    alert('Dichtzetten is niet gelukt.\n\n' + (e && e.message ? e.message : e) +
      '\n\nAls dit blijft gebeuren: controleer of https://rozalapp.github.io is toegestaan in de CORS-instellingen van de Azure Function.');
  } finally {
    boekZaterdagBezig = false;
  }
}
