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
  boekBlokInit();
  boekBlokLaad();
}

async function boekingenVervers() {
  await laadPlanning();
  renderWeekStrip();
  boekingenRender();
  boekingenControleerZaterdag();
  boekBlokLaad();
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
    const hint = alleAfspraken.length
      ? `<div style="font-size:11px;margin-top:6px">(planning geladen: ${alleAfspraken.length} afspraken, daarvan ${alle.length} online/showroom)</div>`
      : `<div style="font-size:11px;margin-top:6px">De planning is nog niet geladen — tik op ⟳ rechtsboven.</div>`;
    lijst.innerHTML = `<div class="leeg-dag">${boekFilter === 'komend' ? 'Geen komende boekingen' : 'Geen eerdere boekingen'}${hint}</div>`;
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
  let vraag = `Weet je het zeker?\n\n${boekDagNaam(iso)} wordt de hele dag dichtgezet voor online boeken.`;
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
    alert(`${boekDagNaam(iso)} is dichtgezet.`);
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


// ── Dag / periode / tijdstip / tijdvak dichtzetten ────────────
// Zelfde lijst + zelfde soorten als in de Windows-app (Boekingbeheer):
//   dag = hele dag, slot = één tijdstip (een tijdvak wordt per half uur opgeslagen).
let boekBlokBezig = false;

function boekTijdStr(v) {
  if (typeof v === 'number') { const m = Math.round(v * 1440); return pad2(Math.floor(m / 60) % 24) + ':' + pad2(m % 60); }
  return String(v || '').trim().slice(0, 5);
}
function boekTijdMin(t) { const p = boekTijdStr(t).split(':'); return (parseInt(p[0], 10) || 0) * 60 + (parseInt(p[1], 10) || 0); }
function boekMinStr(m) { return pad2(Math.floor(m / 60)) + ':' + pad2(m % 60); }
function boekBlokSleutel(g) { return g.type + '|' + boekBlokIso(g.datum) + '|' + (g.type === 'slot' ? boekTijdStr(g.tijd) : ''); }

function boekBlokInit() {
  if (document.getElementById('boek-blok-kaart')) return;
  const pagina = document.getElementById('page-boekingen');
  const eersteBlok = pagina && pagina.querySelector('.dag-blok');
  if (!eersteBlok) return;
  const wrap = document.createElement('div');
  wrap.className = 'dag-blok';
  wrap.style.paddingBottom = '0';
  wrap.innerHTML = `
    <div class="boek-za-kaart" id="boek-blok-kaart">
      <div style="font-weight:700;font-size:14px;margin-bottom:10px">🚫 Dag of tijd dichtzetten</div>
      <label class="boek-bl-label">Wat wil je dichtzetten?</label>
      <select id="boek-bl-type" class="boek-bl-input" onchange="boekBlokType()">
        <option value="dag">Één dag</option>
        <option value="periode">Meerdere dagen (van → t/m)</option>
        <option value="slot">Eén tijdstip</option>
        <option value="tijdvak">Tijdvak (van → tot tijd)</option>
      </select>
      <label class="boek-bl-label" id="boek-bl-datum-label">Datum</label>
      <input type="date" id="boek-bl-datum" class="boek-bl-input">
      <div id="boek-bl-tot-wrap" style="display:none">
        <label class="boek-bl-label">Tot en met datum</label>
        <input type="date" id="boek-bl-datum-tot" class="boek-bl-input">
      </div>
      <div id="boek-bl-tijd-wrap" style="display:none">
        <label class="boek-bl-label" id="boek-bl-tijd-label">Tijdstip</label>
        <input type="time" id="boek-bl-tijd" class="boek-bl-input" step="1800">
      </div>
      <div id="boek-bl-tijdtot-wrap" style="display:none">
        <label class="boek-bl-label">Tot tijdstip</label>
        <input type="time" id="boek-bl-tijd-tot" class="boek-bl-input" step="1800">
      </div>
      <label class="boek-bl-label">Omschrijving (optioneel)</label>
      <input type="text" id="boek-bl-omschr" class="boek-bl-input" placeholder="bijv. Vakantie, Vergadering">
      <div class="boek-za-knop" id="boek-bl-knop" style="margin-top:12px" onclick="boekBlokToevoegen()">🚫 Dichtzetten</div>
      <div class="boek-za-status" id="boek-bl-status"></div>
      <div style="font-weight:700;font-size:13px;margin:12px 0 6px">Nu dichtgezet</div>
      <div id="boek-bl-lijst" style="font-size:13px;color:#888">Laden…</div>
    </div>`;
  eersteBlok.parentNode.insertBefore(wrap, eersteBlok.nextSibling);
  boekBlokType();
}

function boekBlokType() {
  const t = document.getElementById('boek-bl-type').value;
  const toon = (id, aan) => { const e = document.getElementById(id); if (e) e.style.display = aan ? '' : 'none'; };
  toon('boek-bl-tot-wrap', t === 'periode');
  toon('boek-bl-tijd-wrap', t === 'slot' || t === 'tijdvak');
  toon('boek-bl-tijdtot-wrap', t === 'tijdvak');
  document.getElementById('boek-bl-datum-label').textContent = t === 'periode' ? 'Van datum' : 'Datum';
  document.getElementById('boek-bl-tijd-label').textContent = t === 'tijdvak' ? 'Van tijdstip' : 'Tijdstip';
}

async function boekBlokVerstuur(lijst) {
  const resp = await fetch(BOEKING_URL + '?actie=sla-instellingen', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ actie: 'sla-instellingen', type: 'geblokkeerd', data: lijst }),
  });
  if (!resp.ok) throw new Error('Opslaan mislukt (' + resp.status + ')');
}

async function boekBlokToevoegen() {
  if (boekBlokBezig) return;
  const type = document.getElementById('boek-bl-type').value;
  const datum = document.getElementById('boek-bl-datum').value;
  const datumTot = document.getElementById('boek-bl-datum-tot').value;
  const tijd = document.getElementById('boek-bl-tijd').value;
  const tijdTot = document.getElementById('boek-bl-tijd-tot').value;
  const omschr = document.getElementById('boek-bl-omschr').value.trim();
  const vandaag = isoDatum(new Date());

  if (!datum) { alert('Kies een datum.'); return; }
  if (datum < vandaag) { alert('Die datum is al geweest.'); return; }
  if (type === 'periode') {
    if (!datumTot) { alert('Kies de einddatum.'); return; }
    if (datumTot < datum) { alert('De einddatum ligt vóór de begindatum.'); return; }
  }
  if (type === 'slot' && !tijd) { alert('Kies een tijdstip.'); return; }
  if (type === 'tijdvak') {
    if (!tijd || !tijdTot) { alert('Kies een begin- en eindtijd.'); return; }
    if (boekTijdMin(tijdTot) <= boekTijdMin(tijd)) { alert('De eindtijd moet na de begintijd liggen.'); return; }
  }

  // Wat gaan we toevoegen?
  const nieuwe = [];
  if (type === 'dag') {
    nieuwe.push({ type: 'dag', datum: datum, tijd: '', omschrijving: omschr });
  } else if (type === 'periode') {
    const [y, m, d] = datum.split('-').map(Number);
    const dt = new Date(y, m - 1, d);
    let iso = datum, teller = 0;
    while (iso <= datumTot && teller < 400) {
      nieuwe.push({ type: 'dag', datum: iso, tijd: '', omschrijving: omschr || 'Periode geblokkeerd' });
      dt.setDate(dt.getDate() + 1); iso = isoDatum(dt); teller++;
    }
  } else if (type === 'slot') {
    nieuwe.push({ type: 'slot', datum: datum, tijd: boekTijdStr(tijd), omschrijving: omschr });
  } else {
    for (let mnt = boekTijdMin(tijd); mnt < boekTijdMin(tijdTot); mnt += 30) {
      nieuwe.push({ type: 'slot', datum: datum, tijd: boekMinStr(mnt), omschrijving: omschr });
    }
  }

  boekBlokBezig = true;
  const knop = document.getElementById('boek-bl-knop');
  const status = document.getElementById('boek-bl-status');
  knop.classList.add('uit');
  status.textContent = 'Bezig met dichtzetten…';
  try {
    const lijst = await boekHaalGeblokkeerd();   // verse lijst: de dienst vervangt de hele lijst
    const bestaand = new Set(lijst.map(boekBlokSleutel));
    const echtNieuw = nieuwe.filter(g => !bestaand.has(boekBlokSleutel(g)));
    if (!echtNieuw.length) { status.textContent = '✅ Dit stond al dicht.'; return; }
    await boekBlokVerstuur(lijst.concat(echtNieuw));
    const controle = await boekHaalGeblokkeerd();
    const na = new Set(controle.map(boekBlokSleutel));
    if (!echtNieuw.every(g => na.has(boekBlokSleutel(g)))) throw new Error('Na opslaan stond het niet volledig in de lijst');
    status.textContent = '✅ Dichtgezet voor online boeken.';
    document.getElementById('boek-bl-datum').value = '';
    document.getElementById('boek-bl-datum-tot').value = '';
    document.getElementById('boek-bl-tijd').value = '';
    document.getElementById('boek-bl-tijd-tot').value = '';
    document.getElementById('boek-bl-omschr').value = '';
    boekBlokRender(controle);
    boekingenControleerZaterdag();
  } catch (e) {
    console.error('Dichtzetten mislukt:', e);
    status.textContent = '⚠️ Dichtzetten mislukt.';
    alert('Dichtzetten is niet gelukt.\n\n' + (e && e.message ? e.message : e));
  } finally {
    boekBlokBezig = false;
    knop.classList.remove('uit');
  }
}

async function boekBlokLaad() {
  const el = document.getElementById('boek-bl-lijst');
  if (!el) return;
  try { boekBlokRender(await boekHaalGeblokkeerd()); }
  catch (e) { el.textContent = 'Kon de lijst niet laden.'; }
}

// Lijst tonen: dagen achter elkaar worden "van t/m", tijdslots achter elkaar worden één tijdvak
function boekBlokRender(lijst) {
  const el = document.getElementById('boek-bl-lijst');
  if (!el) return;
  const vandaag = isoDatum(new Date());
  const items = lijst.map(g => ({ g, iso: boekBlokIso(g.datum), t: g.type === 'slot' ? boekTijdMin(g.tijd) : -1 }))
    .filter(x => x.iso && x.iso >= vandaag)
    .sort((a, b) => a.iso.localeCompare(b.iso) || a.t - b.t);
  if (!items.length) { el.textContent = 'Niets dichtgezet.'; return; }

  const groepen = [];   // { label, sleutels: [..] }
  let i = 0;
  while (i < items.length) {
    const s = items[i], om = s.g.omschrijving || '';
    const sleutels = [boekBlokSleutel(s.g)];
    let eind = s, j = i + 1;
    while (j < items.length) {
      const n = items[j];
      if (n.g.type !== s.g.type || (n.g.omschrijving || '') !== om) break;
      if (s.g.type === 'dag') {
        const [y, m, d] = eind.iso.split('-').map(Number);
        const volgende = new Date(y, m - 1, d + 1);
        if (isoDatum(volgende) !== n.iso) break;
      } else {
        if (n.iso !== eind.iso || n.t !== eind.t + 30) break;
      }
      sleutels.push(boekBlokSleutel(n.g)); eind = n; j++;
    }
    let label;
    if (s.g.type === 'dag') {
      label = '📅 ' + (sleutels.length === 1 ? boekDagNaam(s.iso) : boekKort(s.iso) + ' t/m ' + boekKort(eind.iso));
    } else {
      label = '🕐 ' + boekDagNaam(s.iso) + ' ' + boekMinStr(s.t) + (sleutels.length === 1 ? '' : '–' + boekMinStr(eind.t + 30));
    }
    if (om) label += ' — <em>' + escapeHtml(om) + '</em>';
    groepen.push({ label, sleutels });
    i = j;
  }
  el.innerHTML = groepen.map((gr, idx) =>
    `<div class="boek-bl-rij"><span>${gr.label}</span><button onclick='boekBlokVerwijder(${JSON.stringify(gr.sleutels)})'>✕</button></div>`
  ).join('');
}

async function boekBlokVerwijder(sleutels) {
  if (boekBlokBezig) return;
  if (!confirm('Weer openzetten voor online boeken?')) return;
  boekBlokBezig = true;
  const status = document.getElementById('boek-bl-status');
  status.textContent = 'Bezig…';
  try {
    const lijst = await boekHaalGeblokkeerd();
    const weg = new Set(sleutels);
    await boekBlokVerstuur(lijst.filter(g => !weg.has(boekBlokSleutel(g))));
    const controle = await boekHaalGeblokkeerd();
    if (controle.some(g => weg.has(boekBlokSleutel(g)))) throw new Error('Verwijderen is niet gelukt');
    status.textContent = '✅ Weer open voor online boeken.';
    boekBlokRender(controle);
    boekingenControleerZaterdag();
  } catch (e) {
    status.textContent = '⚠️ Mislukt.';
    alert('Weer openzetten is niet gelukt.\n\n' + (e && e.message ? e.message : e));
  } finally {
    boekBlokBezig = false;
  }
}

// ── Zelfstandig: eigen opmaak + koppeling aan de tab ───────────
// (zodat het ook werkt als style.css of app.js nog in een oude versie in de cache staan)
(function() {
  const css = `
  .boek-za-kaart { background:#fff; border:1px solid #E0E8D0; border-radius:12px; padding:12px 14px; }
  .boek-za-knop { background:#C0392B; color:#fff; font-weight:700; font-size:14px; text-align:center; padding:13px; border-radius:10px; cursor:pointer; -webkit-tap-highlight-color:transparent; user-select:none; }
  .boek-za-knop:active { opacity:.85; }
  .boek-za-knop.uit { background:#bbb; pointer-events:none; }
  .boek-za-status { font-size:12px; color:#888; margin-top:8px; text-align:center; min-height:15px; }
  .boek-kaart { display:block; cursor:default; }
  .boek-kop { display:flex; align-items:center; gap:10px; }
  .boek-details { margin-top:10px; padding-top:8px; border-top:1px solid #E0E8D0; }
  .boek-regel { display:flex; gap:8px; font-size:12px; padding:2px 0; }
  .boek-regel > span { color:#888; width:74px; flex-shrink:0; }
  .boek-regel a { color:#2D5016; font-weight:600; text-decoration:none; word-break:break-all; }
  .tab-label { white-space:nowrap; }
  .boek-bl-label { display:block; font-size:12px; color:#888; margin:10px 0 4px; }
  .boek-bl-input { width:100%; box-sizing:border-box; font-size:16px; padding:10px; border:1px solid #D5DEC5; border-radius:8px; background:#fff; color:#222; min-height:42px; -webkit-appearance:none; appearance:none; }
  .boek-bl-rij { display:flex; align-items:center; gap:8px; padding:8px 10px; background:#F6F9F0; border:1px solid #E0E8D0; border-radius:8px; margin-bottom:6px; color:#222; }
  .boek-bl-rij span { flex:1; }
  .boek-bl-rij button { background:none; border:none; color:#C0392B; font-size:18px; padding:2px 8px; cursor:pointer; }`;
  const st = document.createElement('style');
  st.textContent = css;
  document.head.appendChild(st);

  document.addEventListener('click', function(e) {
    const tab = e.target.closest && e.target.closest('#tab-boekingen');
    if (tab) setTimeout(function() {
      document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
      document.querySelectorAll('.tab-item').forEach(t => t.classList.remove('active'));
      document.getElementById('page-boekingen').classList.add('active');
      tab.classList.add('active');
      boekingenTabOpen();
    }, 0);
  });
})();

async function boekingenVerversKnop(el) {
  el.classList.add('laden');
  try { await boekingenVervers(); }
  catch (e) { alert('Vernieuwen mislukt: ' + (typeof vertaalFout === 'function' ? vertaalFout(e) : e)); }
  finally { el.classList.remove('laden'); }
}
