// ══════════════════════════════════════════════════════════════
// FOTO'S — opleverfoto's per klant (alleen kijken)
// Bron: SharePoint  RozalApp/Opleverfoto's
// Bestandsnaam: <klant>_<yyyy-mm-dd>_<HHmm>_<nr>.jpg
// ══════════════════════════════════════════════════════════════

let _fotoLijst   = [];       // alle foto's
let _fotoGroepen = [];       // zichtbare groepen (per klant)
let _fotoOpen    = {};       // klant-sleutel -> uitgeklapt
let _fotoGeladen = false;
let _fotoBezig   = false;
let _fotoMapPad  = null;
const FOTO_MAPNAMEN = ["Opleverfoto's", 'Opleverfoto’s', 'Opleverfotos'];

function fotoEncPad(pad) {
  return pad.split('/').map(s => encodeURIComponent(s).replace(/'/g, '%27')).join('/');
}
function fotoLeesbaar(veilig) { return String(veilig).replace(/_+/g, ' ').trim(); }
function fotoDatumNL(d) {
  const m = String(d || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? `${m[3]}-${m[2]}-${m[1]}` : (d || '');
}

async function fotoHaalLijst(pad) {
  let url = `/sites/${CONFIG.spHost}/drive/root:/${fotoEncPad(pad)}:/children?$select=id,name,parentReference&$expand=thumbnails&$top=999`;
  let items = [];
  let data = await graphGet(url);
  items = items.concat(data.value || []);
  let guard = 0;
  while (data['@odata.nextLink'] && guard++ < 20) {
    data = await graphGet(String(data['@odata.nextLink']).replace('https://graph.microsoft.com/v1.0', ''));
    items = items.concat(data.value || []);
  }
  return items;
}

async function fotosLaad() {
  const basis = CONFIG.docPad + '/RozalApp/';
  const kandidaten = _fotoMapPad ? [_fotoMapPad] : FOTO_MAPNAMEN.map(n => basis + n);
  let items = null, laatste = null;
  for (const pad of kandidaten) {
    try { items = await fotoHaalLijst(pad); _fotoMapPad = pad; break; }
    catch (e) { laatste = e; }
  }
  if (!items) {
    if (laatste && laatste.status && laatste.status !== 404) throw laatste;
    _fotoLijst = [];
    return;
  }
  const fotos = [];
  items.forEach(f => {
    const m = String(f.name || '').match(/^(.+)_(\d{4}-\d{2}-\d{2})_(\d{3,6})_(\d+)\.jpe?g$/i);
    if (!m) return;
    const th = f.thumbnails && f.thumbnails[0];
    const tu = th ? ((th.medium && th.medium.url) || (th.small && th.small.url) || (th.large && th.large.url) || null) : null;
    fotos.push({
      id: f.id, driveId: f.parentReference && f.parentReference.driveId,
      naam: f.name, veilig: m[1], datum: m[2], tijd: m[3], nr: parseInt(m[4], 10),
      thumb: tu, dl: null, dlTijd: 0,
    });
  });
  fotos.sort((a, b) => (a.datum + a.tijd).localeCompare(b.datum + b.tijd) || a.nr - b.nr);
  _fotoLijst = fotos;
}

// ── Tab openen / verversen ────────────────────────────────────
async function fotosTabOpen() {
  if (_fotoGeladen) { fotosRender(); return; }
  await fotosVervers();
}

async function fotosVervers() {
  if (_fotoBezig) return;
  _fotoBezig = true;
  const lijst = document.getElementById('fotos-lijst');
  if (lijst && !_fotoLijst.length) lijst.innerHTML = '<div class="laad-scherm"><div class="spinner"></div><div>Foto\'s laden...</div></div>';
  try {
    await fotosLaad();
    _fotoGeladen = true;
    fotosRender();
  } catch (e) {
    if (lijst) lijst.innerHTML = `<div class="leeg-dag">⚠️ ${escapeHtml(vertaalFout(e))}</div>`;
  } finally {
    _fotoBezig = false;
  }
}

// ── Lijst per klant ───────────────────────────────────────────
function fotosRender() {
  const lijst = document.getElementById('fotos-lijst');
  const sub = document.getElementById('fotos-sub');
  if (!lijst) return;

  const zoek = (document.getElementById('fotos-zoek').value || '').toLowerCase().trim().replace(/\s+/g, ' ');
  const van = document.getElementById('fotos-van').value || '';
  const tot = document.getElementById('fotos-tot').value || '';

  const fotos = _fotoLijst.filter(f => {
    if (zoek && fotoLeesbaar(f.veilig).toLowerCase().indexOf(zoek) === -1) return false;
    if (van && f.datum < van) return false;
    if (tot && f.datum > tot) return false;
    return true;
  });

  const perKlant = {};
  fotos.forEach(f => {
    const s = String(f.veilig).toLowerCase();
    if (!perKlant[s]) perKlant[s] = { sleutel: s, naam: fotoLeesbaar(f.veilig), fotos: [], dagen: {} };
    perKlant[s].fotos.push(f);
    (perKlant[s].dagen[f.datum] = perKlant[s].dagen[f.datum] || []).push(f);
  });
  const groepen = Object.values(perKlant);
  groepen.forEach(g => { g.laatste = g.fotos.reduce((m, f) => f.datum > m ? f.datum : m, ''); });
  groepen.sort((a, b) => b.laatste.localeCompare(a.laatste) || a.naam.localeCompare(b.naam));
  _fotoGroepen = groepen;

  sub.textContent = `${groepen.length} klant${groepen.length === 1 ? '' : 'en'} · ${fotos.length} foto's`;

  if (!groepen.length) {
    lijst.innerHTML = `<div class="leeg-dag">${_fotoLijst.length ? 'Geen foto\'s met deze filters' : 'Er zijn nog geen opleverfoto\'s'}</div>`;
    return;
  }

  const direct = groepen.length === 1 || zoek !== '';
  lijst.innerHTML = groepen.map((g, gi) => {
    const open = _fotoOpen[g.sleutel] !== undefined ? _fotoOpen[g.sleutel] : direct;
    let html = `<div class="card foto-kaart"><div class="foto-kop" onclick="fotosToggle(${gi})">
      <div class="klant-avatar">📷</div>
      <div style="flex:1;min-width:0"><div class="klant-naam">${escapeHtml(g.naam)}</div>
      <div class="klant-adres">${g.fotos.length} foto's · laatst ${escapeHtml(fotoDatumNL(g.laatste))}</div></div>
      <div class="klant-pijl">${open ? '⌄' : '›'}</div></div>`;
    if (open) {
      Object.keys(g.dagen).sort().reverse().forEach(d => {
        html += `<div class="foto-dag">${escapeHtml(fotoDatumNL(d))}</div><div class="foto-thumbs">`;
        g.dagen[d].forEach(f => {
          const j = g.fotos.indexOf(f);
          html += f.thumb
            ? `<img class="foto-thumb" loading="lazy" src="${escapeHtml(f.thumb)}" onclick="fotoOpenViewer(${gi},${j})">`
            : `<div class="foto-thumb foto-leeg" onclick="fotoOpenViewer(${gi},${j})">🖼️</div>`;
        });
        html += '</div>';
      });
    }
    return html + '</div>';
  }).join('');
}

function fotosToggle(gi) {
  const g = _fotoGroepen[gi];
  if (!g) return;
  let nu = _fotoOpen[g.sleutel];
  if (nu === undefined) nu = _fotoGroepen.length === 1 || !!(document.getElementById('fotos-zoek').value || '').trim();
  _fotoOpen[g.sleutel] = !nu;
  fotosRender();
}

function fotosWisFilters() {
  ['fotos-zoek', 'fotos-van', 'fotos-tot'].forEach(id => { const e = document.getElementById(id); if (e) e.value = ''; });
  fotosRender();
}

// ── Vergroten (volledige foto pas ophalen na tikken) ──────────
let _fotoViewer = { lijst: [], idx: 0, teller: 0 };

async function fotoDownloadUrl(f) {
  if (f.dl && Date.now() - f.dlTijd < 20 * 60000) return f.dl;
  const item = await graphGet(`/drives/${f.driveId}/items/${f.id}`);
  const url = item['@microsoft.graph.downloadUrl'];
  if (!url) throw new Error('Geen downloadlink');
  f.dl = url; f.dlTijd = Date.now();
  return url;
}

function fotoOpenViewer(gi, j) {
  const g = _fotoGroepen[gi];
  if (!g) return;
  _fotoViewer.lijst = g.fotos;
  _fotoViewer.idx = j;
  document.getElementById('foto-viewer').classList.add('open');
  fotoToonViewer();
}

function fotoSluitViewer() {
  document.getElementById('foto-viewer').classList.remove('open');
  _fotoViewer.teller++;
}

function fotoViewerVolgende(stap) {
  const n = _fotoViewer.lijst.length;
  if (n < 2) return;
  _fotoViewer.idx = (_fotoViewer.idx + stap + n) % n;
  fotoToonViewer();
}

async function fotoToonViewer() {
  const f = _fotoViewer.lijst[_fotoViewer.idx];
  const img = document.getElementById('foto-viewer-img');
  const status = document.getElementById('foto-viewer-status');
  if (!f) return;
  const mijn = ++_fotoViewer.teller;
  document.getElementById('foto-viewer-titel').textContent =
    `${_fotoViewer.idx + 1} / ${_fotoViewer.lijst.length} · ${fotoLeesbaar(f.veilig)} · ${fotoDatumNL(f.datum)}`;
  img.style.display = 'none';
  status.style.display = 'block';
  status.textContent = '⏳ Foto ophalen...';
  try {
    const url = await fotoDownloadUrl(f);
    if (mijn !== _fotoViewer.teller) return;
    img.onload = () => { if (mijn === _fotoViewer.teller) { status.style.display = 'none'; img.style.display = 'block'; } };
    img.onerror = () => { if (mijn === _fotoViewer.teller) status.textContent = '⚠️ Foto kon niet worden geladen'; };
    img.src = url;
  } catch (e) {
    if (mijn === _fotoViewer.teller) status.textContent = '⚠️ ' + vertaalFout(e);
  }
}

// Veegbewegingen: links/rechts = vorige/volgende foto
(function() {
  let startX = null;
  document.addEventListener('touchstart', e => {
    const v = document.getElementById('foto-viewer');
    if (v && v.classList.contains('open') && e.touches.length === 1) startX = e.touches[0].clientX;
  }, { passive: true });
  document.addEventListener('touchend', e => {
    if (startX === null) return;
    const dx = e.changedTouches[0].clientX - startX;
    startX = null;
    if (Math.abs(dx) > 60) fotoViewerVolgende(dx < 0 ? 1 : -1);
  }, { passive: true });
})();
