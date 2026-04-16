'use strict';

// ── CONFIG ─────────────────────────────────────────
const DATA_URL = './data.json.gz';
const PAGE_SIZE = 60;

// ── ESTADO ─────────────────────────────────────────
let RAW = [];
let FILTERED_GENERAL = [];
let pagGeneral = 1;
let sortCol = -1, sortDir = 1;
let chartInstances = {};
let carrierCharts = {};

const CARRIER_ICONS = {
  COORDINADORA: '📦', LOGICUARTAS: '🚚', SERVIENTREGA: '🏃', VELOCES: '⚡', INTERRAPIDISIMO: '📦'
};
const CARRIER_COLORS = {
  COORDINADORA: '#60a5fa', LOGICUARTAS: '#a78bfa', SERVIENTREGA: '#f97316', VELOCES: '#C5D336', INTERRAPIDISIMO: '#005fa8'
};

// ── PALETA CHARTS ──────────────────────────────────
const PALETTE = ['#C5D336','#60a5fa','#f97316','#a78bfa','#4ade80','#f43f5e','#facc15','#2dd4bf'];

// ══════════════════════════════════════════════════
//  AUTH — Microsoft (MSAL)
// ══════════════════════════════════════════════════
const MSAL_CONFIG = {
  auth: {
    clientId: 'febe226c-0265-4fb2-b34e-3beebbb9fee8',
    authority: 'https://login.microsoftonline.com/af1a17b2-5d34-4f58-8b6c-6b94c6cd87ea',
    redirectUri: window.location.origin + window.location.pathname,
    navigateToLoginRequestUrl: true
  },
  cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false }
};
const MS_SCOPES = ['openid', 'profile', 'email', 'User.Read'];
const ALLOWED_DOMAIN = 'lineacom.co';

let _msalInstance = null;
let _sessionTimer = null;

async function _loadMSAL() {
  if (window.msal) return window.msal;
  const URLS = [
    'https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js',
    'https://alcdn.msftauth.net/browser/2.38.3/js/msal-browser.min.js',
    'https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/lib/msal-browser.min.js',
  ];
  for (const url of URLS) {
    try {
      await new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = url; s.onload = resolve; s.onerror = reject;
        document.head.appendChild(s);
      });
      if (window.msal) return window.msal;
    } catch (_) {}
  }
  throw new Error('No se pudo cargar MSAL.');
}

async function doMicrosoftLogin() {
  const btn = document.getElementById('msLoginBtn');
  if (btn) { btn.disabled = true; btn.textContent = 'Conectando con Microsoft...'; }
  document.getElementById('login-error').style.display = 'none';
  try {
    const msal = await _loadMSAL();
    _msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
    await _msalInstance.initialize();
    const result = await _msalInstance.loginPopup({ scopes: MS_SCOPES, prompt: 'select_account' });
    const email = result.account?.username || '';
    const domain = email.split('@')[1]?.toLowerCase();
    if (domain !== ALLOWED_DOMAIN) {
      await _msalInstance.logoutPopup({ account: result.account }).catch(() => {});
      _showLoginError(`Acceso denegado. Solo cuentas @${ALLOWED_DOMAIN}.<br><small>Tu cuenta: ${email}</small>`);
      return;
    }
    let displayName = result.account.name || email;
    let jobTitle = '';
    let photo = null;
    try {
      const gt = await _msalInstance.acquireTokenSilent({ scopes: ['User.Read'], account: result.account });
      const pr = await fetch('https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle', {
        headers: { Authorization: `Bearer ${gt.accessToken}` }
      });
      if (pr.ok) { const p = await pr.json(); displayName = p.displayName||displayName; jobTitle = p.jobTitle||''; }
      const phr = await fetch('https://graph.microsoft.com/v1.0/me/photo/$value', {
        headers: { Authorization: `Bearer ${gt.accessToken}` }
      });
      if (phr.ok) {
        const blob = await phr.blob();
        photo = await new Promise(res => { const r = new FileReader(); r.onload = ()=>res(r.result); r.readAsDataURL(blob); });
      }
    } catch (_) {}
    _enterApp({ name: displayName, role: jobTitle, email, photo });
  } catch (err) {
    _showLoginError(err.errorCode === 'user_cancelled' ? 'Inicio cancelado.' : 'Error al conectar con Microsoft.');
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 23 23" style="width:20px;height:20px;flex-shrink:0;"><rect x="1" y="1" width="10" height="10" fill="#f25022"/><rect x="12" y="1" width="10" height="10" fill="#7fba00"/><rect x="1" y="12" width="10" height="10" fill="#00a4ef"/><rect x="12" y="12" width="10" height="10" fill="#ffb900"/></svg> Iniciar sesión con Microsoft`;
    }
  }
}

function _showLoginError(msg) {
  const e = document.getElementById('login-error');
  e.innerHTML = '⚠ ' + msg;
  e.style.display = 'block';
}

function _enterApp(profile) {
  const wa = document.getElementById('user-avatar-wrap');
  const wn = document.getElementById('user-name-el');
  const wr = document.getElementById('user-role-el');
  if (profile.photo) {
    wa.innerHTML = `<img src="${profile.photo}" style="width:100%;height:100%;object-fit:cover;border-radius:50%;">`;
  } else {
    wa.textContent = (profile.name||'U').split(' ').map(w=>w[0]).join('').substring(0,2).toUpperCase();
  }
  if (wn) wn.textContent = profile.name||'';
  if (wr) wr.textContent = profile.role||'Lineacom';

  document.getElementById('login-screen').classList.add('hidden');
  document.getElementById('footer-year').textContent = `v3.0 · ${new Date().getFullYear()}`;

  _sessionTimer = setTimeout(() => {
    alert('⏰ Sesión expirada.');
    doLogout();
  }, 3600000);

  _showDataLoadingScreen();
  loadData();
}

function _showDataLoadingScreen() {
  document.getElementById('data-loading-screen').classList.remove('hidden');
}

function _hideDataLoadingScreen() {
  document.getElementById('data-loading-screen').classList.add('hidden');
  document.getElementById('app').classList.add('visible');
}

function _setLoadProgress(pct, statusText) {
  const bar = document.getElementById('dl-progress-bar');
  const status = document.getElementById('dl-status-text');
  if (bar) bar.style.width = pct + '%';
  if (status) status.textContent = statusText || '';
}

function doLogout() {
  clearTimeout(_sessionTimer);
  if (_msalInstance) {
    const accs = _msalInstance.getAllAccounts();
    if (accs.length) _msalInstance.logoutPopup({ account: accs[0] }).catch(() => {});
  }
  document.getElementById('app').classList.remove('visible');
  document.getElementById('data-loading-screen').classList.add('hidden');
  document.getElementById('login-screen').classList.remove('hidden');
  _setLoadProgress(0, 'Iniciando...');
  document.getElementById('dl-sub-text').textContent = 'Descomprimiendo y procesando datos de guías...';
  RAW = []; FILTERED_GENERAL = [];
}

// ══════════════════════════════════════════════════
//  CARGA DE DATOS (GitHub → GZ → JSON)
// ══════════════════════════════════════════════════
function showLoading(msg='Cargando datos...') {
  document.getElementById('loading-bar').classList.add('visible');
}
function hideLoading() {
  document.getElementById('loading-bar').classList.remove('visible');
}

async function loadData() {
  _setLoadProgress(5, 'Leyendo archivo local data.json.gz...');
  document.getElementById('data-summary') && (document.getElementById('data-summary').textContent = 'Procesando datos...');
  try {
    _setLoadProgress(15, 'Descargando data.json.gz...');
    const res = await fetch(DATA_URL + '?t=' + Date.now());
    if (!res.ok) throw new Error(`HTTP ${res.status} — archivo no encontrado`);

    _setLoadProgress(35, 'Archivo recibido, descomprimiendo...');
    const ab = await res.arrayBuffer();

    _setLoadProgress(50, 'Descomprimiendo gzip...');
    const ds = new DecompressionStream('gzip');
    const writer = ds.writable.getWriter();
    writer.write(new Uint8Array(ab));
    writer.close();

    const reader = ds.readable.getReader();
    const chunks = [];
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }
    const total = chunks.reduce((s, c) => s + c.length, 0);
    const merged = new Uint8Array(total);
    let offset = 0;
    for (const c of chunks) { merged.set(c, offset); offset += c.length; }

    _setLoadProgress(72, 'Parseando JSON...');
    const jsonStr = new TextDecoder('utf-8').decode(merged);
    const payload = JSON.parse(jsonStr);

    _setLoadProgress(88, 'Construyendo dashboard...');
    RAW = payload.guias || [];
    const ts = payload.generado_en || '—';
    document.getElementById('last-update').textContent = `Actualizado: ${ts}`;
    document.getElementById('data-summary').textContent =
      `${RAW.length.toLocaleString()} registros históricos · ${ts}`;

    initDashboard();

    _setLoadProgress(100, `✓ ${RAW.length.toLocaleString()} guías cargadas`);
    document.getElementById('dl-sub-text').textContent = '¡Listo! Abriendo el dashboard...';

    await new Promise(r => setTimeout(r, 600));
    _hideDataLoadingScreen();

  } catch (err) {
    console.error('[loadData]', err);
    _setLoadProgress(0, '');
    document.getElementById('dl-sub-text').textContent = '⚠ Error al cargar datos';
    document.getElementById('dl-status-text').textContent = err.message;
    document.getElementById('data-summary') && (document.getElementById('data-summary').textContent = '⚠ Error al cargar datos: ' + err.message);
    await new Promise(r => setTimeout(r, 2500));
    _hideDataLoadingScreen();
  }
}

async function reloadData() {
  _showDataLoadingScreen();
  document.getElementById('dl-sub-text').textContent = 'Actualizando datos...';
  await loadData();
}

// ══════════════════════════════════════════════════
//  DASHBOARD INIT
// ══════════════════════════════════════════════════
function initDashboard() {
  const transportadoras = ['COORDINADORA', 'LOGICUARTAS', 'SERVIENTREGA', 'VELOCES', 'INTERRAPIDISIMO'];

  document.getElementById('badge-general').textContent = RAW.length;
  transportadoras.forEach(t => {
    const cnt = RAW.filter(r => r.TRANSPORTADORA === t).length;
    const el = document.getElementById('badge-' + t);
    if (el) el.textContent = cnt;
  });

  const estados = [...new Set(RAW.map(r => r.ESTADO).filter(Boolean))].sort();
  const selEstado = document.getElementById('filter-estado');
  estados.forEach(e => {
    const opt = document.createElement('option');
    opt.value = e; opt.textContent = e;
    selEstado.appendChild(opt);
  });

  // Poblar ciudades origen
  const ciudadesOrigen = [...new Set(RAW.map(r => r.CIUDAD_ORIGEN).filter(Boolean))].sort();
  const selCiudadOrigen = document.getElementById('filter-ciudad-origen');
  if (selCiudadOrigen) {
    ciudadesOrigen.forEach(c => {
      const opt = document.createElement('option');
      opt.value = c; opt.textContent = c;
      selCiudadOrigen.appendChild(opt);
    });
  }

  renderGeneral();
  transportadoras.forEach(t => renderCarrier(t));
}

// ══════════════════════════════════════════════════
//  TAB SWITCHING
// ══════════════════════════════════════════════════
function switchTab(tab) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.toggle('active', c.id === 'tab-' + tab));
}

// ══════════════════════════════════════════════════
//  MODAL DETALLE KPI
// ══════════════════════════════════════════════════
function openKpiModal(titulo, rows) {
  let modal = document.getElementById('kpi-modal-overlay');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'kpi-modal-overlay';
    modal.className = 'kpi-modal-overlay hidden';
    modal.innerHTML = `
      <div class="kpi-modal">
        <div class="kpi-modal-header">
          <div class="kpi-modal-title">
            <span id="kpi-modal-titulo">—</span>
            <span class="kpi-modal-badge" id="kpi-modal-count">0</span>
          </div>
          <button class="kpi-modal-close" onclick="closeKpiModal()">✕</button>
        </div>
        <div class="kpi-modal-search">
          <input type="text" id="kpi-modal-search-input" placeholder="Buscar guía, destinatario, ciudad..." oninput="filterKpiModal()">
        </div>
        <div class="kpi-modal-body" id="kpi-modal-body"></div>
        <div class="kpi-modal-footer">
          <span id="kpi-modal-footer-count"></span>
          <button class="btn-primary" style="font-size:12px;padding:7px 14px;" onclick="exportKpiModal()">↓ Exportar Excel</button>
        </div>
      </div>`;
    modal.addEventListener('click', function(e){ if(e.target === modal) closeKpiModal(); });
    document.body.appendChild(modal);
  }

  window._kpiModalRows = rows;
  window._kpiModalFiltered = rows;
  document.getElementById('kpi-modal-titulo').textContent = titulo;
  document.getElementById('kpi-modal-count').textContent = rows.length.toLocaleString();
  document.getElementById('kpi-modal-search-input').value = '';
  renderKpiModalTable(rows);
  modal.classList.remove('hidden');
  document.body.style.overflow = 'hidden';
}

function closeKpiModal() {
  const modal = document.getElementById('kpi-modal-overlay');
  if (modal) modal.classList.add('hidden');
  document.body.style.overflow = '';
}

function filterKpiModal() {
  const q = (document.getElementById('kpi-modal-search-input')?.value || '').toLowerCase().trim();
  const rows = window._kpiModalRows || [];
  window._kpiModalFiltered = !q ? rows : rows.filter(r => {
    const hay = [r.GUIA, r.NOMBRE_DESTINATARIO, r.CIUDAD_DESTINO, r.TRANSPORTADORA, r.ESTADO, r.DOCUMENTO, r.NOVEDAD]
      .map(v => (v||'').toLowerCase()).join(' ');
    return hay.includes(q);
  });
  renderKpiModalTable(window._kpiModalFiltered);
}

function renderKpiModalTable(rows) {
  const body = document.getElementById('kpi-modal-body');
  const footer = document.getElementById('kpi-modal-footer-count');
  if (!body) return;

  if (footer) footer.textContent = `${rows.length.toLocaleString()} registros`;

  if (!rows.length) {
    body.innerHTML = '<div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-text">Sin resultados.</div></div>';
    return;
  }

  const cols = [
    { label: 'Guía',           fn: r => `<span class="cell-guia">${esc(r.GUIA||'—')}</span>` },
    { label: 'Transportadora', fn: r => `<span style="color:${CARRIER_COLORS[r.TRANSPORTADORA]||'#aaa'};font-weight:600;font-size:11px;">${esc(r.TRANSPORTADORA||'—')}</span>` },
    { label: 'Estado',         fn: r => statusPill(r.ESTADO||'') },
    { label: 'Destinatario',   fn: r => esc(r.NOMBRE_DESTINATARIO||'—') },
    { label: 'Ciudad Destino', fn: r => esc(r.CIUDAD_DESTINO||'—') },
    { label: 'Último Mov.',    fn: r => esc(r.ULTIMO_MOVIMIENTO||'—') },
    { label: 'Novedad',        fn: r => { const n=r.NOVEDAD||''; return n && n!=='NO APLICA' ? `<span style="color:var(--warn);font-size:11px;">${esc(n.substring(0,80))}</span>` : '<span style="color:var(--text-muted)">—</span>'; } },
    { label: 'F. Entrega',     fn: r => esc(r.FECHA_ENTREGA||'—') },
  ];

  body.innerHTML = `
    <table>
      <thead><tr>${cols.map(c=>`<th>${c.label}</th>`).join('')}</tr></thead>
      <tbody>${rows.map(r=>`<tr>${cols.map(c=>`<td>${c.fn(r)}</td>`).join('')}</tr>`).join('')}</tbody>
    </table>`;
}

function exportKpiModal() {
  const rows = window._kpiModalFiltered || [];
  const titulo = document.getElementById('kpi-modal-titulo')?.textContent || 'detalle';
  _exportXlsx(_toExcelData(rows), titulo, `Lineacom_${titulo.replace(/\s+/g,'_')}`);
}

// Keyboard close
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeKpiModal(); });

// ══════════════════════════════════════════════════
//  RENDER GENERAL
// ══════════════════════════════════════════════════
function renderGeneral() {
  const data = RAW;
  const total = data.length;
  const { entregadas, transito, novedad, pendientes, devueltas, canceladas } = contarEstados(data);

  const kpis = [
    { label: 'Total Guías',  value: total.toLocaleString(),       cls: 'lima',    icon: '📋', estado: null },
    { label: 'Entregadas',   value: entregadas.toLocaleString(),  cls: 'ok',      icon: '✅',  sub: pct(entregadas, total),  estado: 'entregada' },
    { label: 'En Tránsito',  value: transito.toLocaleString(),    cls: 'transit', icon: '🚚', sub: pct(transito, total),    estado: 'transito' },
    { label: 'Con Novedad',  value: novedad.toLocaleString(),     cls: 'warn',    icon: '⚠️', sub: pct(novedad, total),     estado: 'novedad' },
    { label: 'Pendientes',   value: pendientes.toLocaleString(),  cls: 'info',    icon: '⏳', sub: pct(pendientes, total),  estado: 'pendiente' },
    { label: 'Devueltas',    value: devueltas.toLocaleString(),   cls: 'danger',  icon: '↩️', sub: pct(devueltas, total),   estado: 'devuelta' },
    { label: 'Canceladas',   value: canceladas.toLocaleString(),  cls: 'muted',   icon: '🚫', sub: pct(canceladas, total),  estado: 'cancelada' },
  ];

  document.getElementById('kpis-general').innerHTML = kpis.map(k => `
    <div class="kpi-card" ${k.estado ? `data-kpi="${k.estado}" onclick="kpiClick('${k.estado}', '${k.label}', null)"` : ''} title="${k.estado ? 'Clic para ver detalle' : ''}">
      <div class="kpi-icon">${k.icon}</div>
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value ${k.cls}">${k.value}</div>
      ${k.sub ? `<div class="kpi-sub">${k.sub} del total</div>` : ''}
    </div>`).join('');

  renderChartEstados('chart-estados-general', data);
  renderChartTransportadoras('chart-transportadoras', data);
  renderChartTimeline('chart-timeline', data);
  renderChartTasaEntrega('chart-tasa-entrega', data);
  renderChartNovedadesTipo('chart-novedades-tipo', data);
  renderChartCiudades('chart-ciudades-global', data);
  renderChartDevNovCarrier('chart-dev-nov-carrier', data);
  renderChartPendTransito('chart-pend-transito', data);
  renderChartTiempoEntrega('chart-tiempo-entrega', data);

  FILTERED_GENERAL = [...data];
  pagGeneral = 1;
  renderTableGeneral();
}

// Función para abrir modal desde KPI — pestaña general
function kpiClick(estado, label, carrier) {
  let rows;
  if (carrier) {
    const base = RAW.filter(r => r.TRANSPORTADORA === carrier);
    rows = filtrarPorEstadoNorm(base, estado);
  } else {
    rows = filtrarPorEstadoNorm(RAW, estado);
  }
  openKpiModal(`${label}${carrier ? ' — ' + carrier : ''}`, rows);
}

function filtrarPorEstadoNorm(data, estado) {
  return data.filter(r => normalizeEstado(r.ESTADO) === estado);
}

function filterGeneral() {
  const search = (document.getElementById('search-general').value || '').toLowerCase().trim();
  const trans = document.getElementById('filter-transportadora').value;
  const estado = document.getElementById('filter-estado').value;
  const estadoNorm = document.getElementById('filter-estado-norm')?.value || '';
  const ciudad = (document.getElementById('filter-ciudad')?.value || '').toLowerCase().trim();
  const fechaDesde = document.getElementById('filter-fecha-desde')?.value || '';
  const fechaHasta = document.getElementById('filter-fecha-hasta')?.value || '';
  const guia = (document.getElementById('filter-guia')?.value || '').toLowerCase().trim();
  const ciudadOrigen = document.getElementById('filter-ciudad-origen')?.value || '';

  FILTERED_GENERAL = RAW.filter(r => {
    if (trans && r.TRANSPORTADORA !== trans) return false;
    if (estado && r.ESTADO !== estado) return false;
    if (estadoNorm && normalizeEstado(r.ESTADO) !== estadoNorm) return false;
    if (ciudad && !(r.CIUDAD_DESTINO||'').toLowerCase().includes(ciudad)) return false;
    if (ciudadOrigen && r.CIUDAD_ORIGEN !== ciudadOrigen) return false;
    if (fechaDesde && (r.FECHA_PROCESAMIENTO||'') < fechaDesde) return false;
    if (fechaHasta && (r.FECHA_PROCESAMIENTO||'') > fechaHasta + 'T23:59:59') return false;
    if (guia && !(r.GUIA||'').toLowerCase().includes(guia)) return false;
    if (search) {
      const hay = [r.GUIA, r.NOMBRE_DESTINATARIO, r.CIUDAD_DESTINO, r.TRANSPORTADORA, r.ESTADO, r.DOCUMENTO, r.NOVEDAD]
        .map(v => (v||'').toLowerCase()).join(' ');
      if (!hay.includes(search)) return false;
    }
    return true;
  });
  pagGeneral = 1;
  renderTableGeneral();
}

function clearFilters() {
  document.getElementById('search-general').value = '';
  document.getElementById('filter-transportadora').value = '';
  document.getElementById('filter-estado').value = '';
  const fn = document.getElementById('filter-estado-norm'); if (fn) fn.value = '';
  const fc = document.getElementById('filter-ciudad'); if (fc) fc.value = '';
  const fd = document.getElementById('filter-fecha-desde'); if (fd) fd.value = '';
  const fh = document.getElementById('filter-fecha-hasta'); if (fh) fh.value = '';
  const fg = document.getElementById('filter-guia'); if (fg) fg.value = '';
  const fo = document.getElementById('filter-ciudad-origen'); if (fo) fo.value = '';
  filterGeneral();
}

function renderTableGeneral() {
  const total = FILTERED_GENERAL.length;
  document.getElementById('count-general').textContent = `${total.toLocaleString()} registros`;
  const pages = Math.max(1, Math.ceil(total / PAGE_SIZE));
  const slice = FILTERED_GENERAL.slice((pagGeneral - 1) * PAGE_SIZE, pagGeneral * PAGE_SIZE);

  const cols = [
    { label: 'Guía',         fn: r => `<span class="cell-guia">${esc(r.GUIA||'—')}</span>` },
    { label: 'Transportadora', fn: r => `<span style="color:${CARRIER_COLORS[r.TRANSPORTADORA]||'#aaa'};font-weight:600;font-size:11px;">${esc(r.TRANSPORTADORA||'—')}</span>` },
    { label: 'Estado',       fn: r => statusPill(r.ESTADO||'') },
    { label: 'Destinatario', fn: r => esc(r.NOMBRE_DESTINATARIO||'—') },
    { label: 'Ciudad Destino', fn: r => esc(r.CIUDAD_DESTINO||'—') },
    { label: 'Último Mov.',  fn: r => esc(r.ULTIMO_MOVIMIENTO||'—') },
    { label: 'Novedad',      fn: r => { const n=r.NOVEDAD||''; return n && n!=='NO APLICA' ? `<span style="color:var(--warn);font-size:11px;">${esc(n.substring(0,80))}</span>` : '<span style="color:var(--text-muted)">—</span>'; } },
    { label: 'F. Entrega',   fn: r => esc(r.FECHA_ENTREGA||'—') },
  ];

  if (!slice.length) {
    document.getElementById('table-general-wrap').innerHTML = '<div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-text">No hay registros que coincidan con los filtros.</div></div>';
    document.getElementById('pag-general').innerHTML = '';
    return;
  }

  document.getElementById('table-general-wrap').innerHTML = `
    <table>
      <thead><tr>${cols.map(c => `<th>${c.label}</th>`).join('')}</tr></thead>
      <tbody>${slice.map(r => `<tr>${cols.map(c => `<td>${c.fn(r)}</td>`).join('')}</tr>`).join('')}</tbody>
    </table>`;

  mkPagination('pag-general', pagGeneral, pages, p => { pagGeneral = p; renderTableGeneral(); });
}

// ══════════════════════════════════════════════════
//  RENDER POR TRANSPORTADORA
// ══════════════════════════════════════════════════
function renderCarrier(carrier) {
  const data = RAW.filter(r => r.TRANSPORTADORA === carrier);
  const container = document.getElementById('content-' + carrier);
  if (!container) return;

  const total = data.length;
  const { entregadas, transito, novedad, pendientes, devueltas, canceladas } = contarEstados(data);
  const pctEntr = total ? Math.round(entregadas / total * 100) : 0;

  container.innerHTML = `
    <div class="carrier-banner">
      <div class="carrier-icon">${CARRIER_ICONS[carrier]||'📋'}</div>
      <div>
        <div class="carrier-name">${carrier}</div>
        <div class="carrier-total">${total.toLocaleString()} guías registradas · ${pctEntr}% entregadas</div>
      </div>
      <div style="margin-left:auto;">
        <button class="btn-primary" onclick="exportarCarrier('${carrier}')">↓ Exportar Excel</button>
      </div>
    </div>

    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-icon">📋</div><div class="kpi-label">Total</div><div class="kpi-value lima">${total.toLocaleString()}</div></div>
      <div class="kpi-card" data-kpi="entregada" onclick="kpiClick('entregada','Entregadas','${carrier}')"><div class="kpi-icon">✅</div><div class="kpi-label">Entregadas</div><div class="kpi-value ok">${entregadas.toLocaleString()}</div><div class="kpi-sub">${pct(entregadas,total)}</div></div>
      <div class="kpi-card" data-kpi="transito" onclick="kpiClick('transito','En Tránsito','${carrier}')"><div class="kpi-icon">🚚</div><div class="kpi-label">En Tránsito</div><div class="kpi-value transit">${transito.toLocaleString()}</div><div class="kpi-sub">${pct(transito,total)}</div></div>
      <div class="kpi-card" data-kpi="novedad" onclick="kpiClick('novedad','Novedad','${carrier}')"><div class="kpi-icon">⚠️</div><div class="kpi-label">Novedad</div><div class="kpi-value warn">${novedad.toLocaleString()}</div><div class="kpi-sub">${pct(novedad,total)}</div></div>
      <div class="kpi-card" data-kpi="pendiente" onclick="kpiClick('pendiente','Pendientes','${carrier}')"><div class="kpi-icon">⏳</div><div class="kpi-label">Pendientes</div><div class="kpi-value info">${pendientes.toLocaleString()}</div><div class="kpi-sub">${pct(pendientes,total)}</div></div>
      <div class="kpi-card" data-kpi="devuelta" onclick="kpiClick('devuelta','Devueltas','${carrier}')"><div class="kpi-icon">↩️</div><div class="kpi-label">Devueltas</div><div class="kpi-value danger">${devueltas.toLocaleString()}</div><div class="kpi-sub">${pct(devueltas,total)}</div></div>
      <div class="kpi-card" data-kpi="cancelada" onclick="kpiClick('cancelada','Canceladas','${carrier}')"><div class="kpi-icon">🚫</div><div class="kpi-label">Canceladas</div><div class="kpi-value muted">${canceladas.toLocaleString()}</div><div class="kpi-sub">${pct(canceladas,total)}</div></div>
    </div>

    <div class="chart-grid">
      <div class="chart-card">
        <div class="chart-title">Estados — ${carrier}</div>
        <div class="chart-wrap" style="height:250px;"><canvas id="chart-estados-${carrier}"></canvas></div>
      </div>
      <div class="chart-card">
        <div class="chart-title">Top 10 Ciudades Destino</div>
        <div class="chart-wrap" style="height:250px;"><canvas id="chart-ciudades-${carrier}"></canvas></div>
      </div>
    </div>

    <div class="filters-bar">
      <span class="filter-label">Filtrar</span>
      <input class="filter-input" type="text" id="search-${carrier}" placeholder="Buscar guía, destinatario..." oninput="filterCarrier('${carrier}')">
      <select class="filter-select" id="filter-estadonorm-${carrier}" onchange="filterCarrier('${carrier}')">
        <option value="">Categoría estado</option>
        <option value="entregada">✅ Entregadas</option>
        <option value="transito">🚚 En Tránsito</option>
        <option value="novedad">⚠️ Novedad</option>
        <option value="pendiente">⏳ Pendientes</option>
        <option value="devuelta">↩️ Devueltas</option>
        <option value="cancelada">🚫 Canceladas</option>
        <option value="otro">❓ Otro</option>
      </select>
      <select class="filter-select" id="filter-estado-${carrier}" onchange="filterCarrier('${carrier}')">
        <option value="">Estado exacto</option>
        ${[...new Set(data.map(r=>r.ESTADO).filter(Boolean))].sort().map(e=>`<option value="${esc(e)}">${esc(e)}</option>`).join('')}
      </select>
      <input class="filter-input" type="text" id="filter-ciudad-${carrier}" placeholder="Ciudad destino..." oninput="filterCarrier('${carrier}')" style="max-width:160px;">
      <input class="filter-input" type="date" id="filter-fechadesde-${carrier}" onchange="filterCarrier('${carrier}')" title="Desde (fecha procesamiento)" style="max-width:140px;">
      <input class="filter-input" type="date" id="filter-fechahasta-${carrier}" onchange="filterCarrier('${carrier}')" title="Hasta (fecha procesamiento)" style="max-width:140px;">
      <button class="btn-secondary" onclick="clearCarrierFilters('${carrier}')">✕ Limpiar</button>
    </div>
    <div class="table-container">
      <div class="table-header">
        <span class="table-count" id="count-${carrier}">— registros</span>
      </div>
      <div class="table-scroll">
        <div id="table-${carrier}-wrap"></div>
      </div>
      <div class="pagination" id="pag-${carrier}"></div>
    </div>
  `;

  setTimeout(() => {
    renderChartEstados('chart-estados-' + carrier, data, CARRIER_COLORS[carrier]);
    renderChartCiudades('chart-ciudades-' + carrier, data);
    renderTableCarrier(carrier, data, data);
  }, 50);
}

const _carrierFiltered = {};
const _carrierPage = {};

function filterCarrier(carrier) {
  const search = (document.getElementById('search-' + carrier)?.value || '').toLowerCase().trim();
  const estado = document.getElementById('filter-estado-' + carrier)?.value || '';
  const estadoNorm = document.getElementById('filter-estadonorm-' + carrier)?.value || '';
  const ciudad = (document.getElementById('filter-ciudad-' + carrier)?.value || '').toLowerCase().trim();
  const fechaDesde = document.getElementById('filter-fechadesde-' + carrier)?.value || '';
  const fechaHasta = document.getElementById('filter-fechahasta-' + carrier)?.value || '';
  const base = RAW.filter(r => r.TRANSPORTADORA === carrier);
  _carrierFiltered[carrier] = base.filter(r => {
    if (estado && r.ESTADO !== estado) return false;
    if (estadoNorm && normalizeEstado(r.ESTADO) !== estadoNorm) return false;
    if (ciudad && !(r.CIUDAD_DESTINO||'').toLowerCase().includes(ciudad)) return false;
    if (fechaDesde && (r.FECHA_PROCESAMIENTO||'') < fechaDesde) return false;
    if (fechaHasta && (r.FECHA_PROCESAMIENTO||'') > fechaHasta + 'T23:59:59') return false;
    if (search) {
      const hay = [r.GUIA, r.NOMBRE_DESTINATARIO, r.CIUDAD_DESTINO, r.ESTADO, r.DOCUMENTO, r.NOVEDAD]
        .map(v=>(v||'').toLowerCase()).join(' ');
      if (!hay.includes(search)) return false;
    }
    return true;
  });
  _carrierPage[carrier] = 1;
  renderTableCarrier(carrier, _carrierFiltered[carrier], RAW.filter(r=>r.TRANSPORTADORA===carrier));
}

function clearCarrierFilters(carrier) {
  const s = document.getElementById('search-' + carrier); if (s) s.value = '';
  const e = document.getElementById('filter-estado-' + carrier); if (e) e.value = '';
  const en = document.getElementById('filter-estadonorm-' + carrier); if (en) en.value = '';
  const c = document.getElementById('filter-ciudad-' + carrier); if (c) c.value = '';
  filterCarrier(carrier);
}

function renderTableCarrier(carrier, base, _unused) {
  if (!_carrierFiltered[carrier]) _carrierFiltered[carrier] = base;
  if (!_carrierPage[carrier]) _carrierPage[carrier] = 1;
  const fd = _carrierFiltered[carrier];
  const page = _carrierPage[carrier];
  const pages = Math.max(1, Math.ceil(fd.length / PAGE_SIZE));
  const slice = fd.slice((page-1)*PAGE_SIZE, page*PAGE_SIZE);

  document.getElementById('count-' + carrier).textContent = `${fd.length.toLocaleString()} registros`;

  const cols = [
    { label: 'Guía',        fn: r => `<span class="cell-guia">${esc(r.GUIA||'—')}</span>` },
    { label: 'Estado',      fn: r => statusPill(r.ESTADO||'') },
    { label: 'Destinatario', fn: r => esc(r.NOMBRE_DESTINATARIO||'—') },
    { label: 'Dirección',   fn: r => `<span style="font-size:11px;color:var(--text-dim);">${esc((r.DIRECCION_DESTINATARIO||'—').substring(0,50))}</span>` },
    { label: 'Ciudad Dest.',fn: r => esc(r.CIUDAD_DESTINO||'—') },
    { label: 'Origen',      fn: r => esc(r.CIUDAD_ORIGEN||'—') },
    { label: 'Último Mov.', fn: r => esc(r.ULTIMO_MOVIMIENTO||'—') },
    { label: 'F. Entrega',  fn: r => esc(r.FECHA_ENTREGA||'—') },
    { label: 'Novedad',     fn: r => { const n=r.NOVEDAD||''; return n && n!=='NO APLICA' ? `<span style="color:var(--warn);font-size:11px;">${esc(n.substring(0,70))}</span>` : '<span style="color:var(--text-muted)">—</span>'; } },
  ];

  if (!slice.length) {
    document.getElementById('table-' + carrier + '-wrap').innerHTML =
      '<div class="empty-state"><div class="empty-icon">📭</div><div class="empty-text">Sin resultados.</div></div>';
    document.getElementById('pag-' + carrier).innerHTML = '';
    return;
  }

  document.getElementById('table-' + carrier + '-wrap').innerHTML = `
    <table>
      <thead><tr>${cols.map(c=>`<th>${c.label}</th>`).join('')}</tr></thead>
      <tbody>${slice.map(r=>`<tr>${cols.map(c=>`<td>${c.fn(r)}</td>`).join('')}</tr>`).join('')}</tbody>
    </table>`;

  mkPagination('pag-' + carrier, page, pages, p => changeCarrierPage(carrier, p));
}

// Función global para evitar ReferenceError en pagination closures
window.changeCarrierPage = function(carrier, p) {
  _carrierPage[carrier] = p;
  renderTableCarrier(carrier, _carrierFiltered[carrier], RAW.filter(r => r.TRANSPORTADORA === carrier));
};

// ══════════════════════════════════════════════════
//  CHARTS
// ══════════════════════════════════════════════════
const CHART_DEFAULTS = {
  responsive: true, maintainAspectRatio: false,
  plugins: {
    legend: { labels: { color: '#7a8099', font: { family: 'DM Sans', size: 11 } } },
    tooltip: {
      backgroundColor: 'rgba(13,15,26,.95)',
      titleColor: '#C5D336',
      bodyColor: '#e8eaf0',
      borderColor: 'rgba(197,211,54,.2)',
      borderWidth: 1, padding: 12,
    }
  }
};

function destroyChart(id) {
  if (chartInstances[id]) { chartInstances[id].destroy(); delete chartInstances[id]; }
}

function renderChartEstados(canvasId, data, accentColor) {
  destroyChart(canvasId);
  const counts = {};
  data.forEach(r => { const e = r.ESTADO||'SIN ESTADO'; counts[e] = (counts[e]||0)+1; });
  const sorted = Object.entries(counts).sort((a,b) => b[1]-a[1]).slice(0, 10);
  const ctx = document.getElementById(canvasId)?.getContext('2d');
  if (!ctx) return;
  chartInstances[canvasId] = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: sorted.map(([k])=>k),
      datasets: [{ data: sorted.map(([,v])=>v), backgroundColor: PALETTE, borderWidth: 0, hoverOffset: 8 }]
    },
    options: { ...CHART_DEFAULTS, cutout: '65%' }
  });
}

function renderChartTransportadoras(canvasId, data) {
  destroyChart(canvasId);
  const counts = {};
  data.forEach(r => { const t = r.TRANSPORTADORA||'OTRO'; counts[t] = (counts[t]||0)+1; });
  const entries = Object.entries(counts).sort((a,b)=>b[1]-a[1]);
  const ctx = document.getElementById(canvasId)?.getContext('2d');
  if (!ctx) return;
  chartInstances[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: entries.map(([k])=>k),
      datasets: [{
        data: entries.map(([,v])=>v),
        backgroundColor: entries.map(([k]) => CARRIER_COLORS[k]||'#C5D336'),
        borderRadius: 8, borderSkipped: false,
      }]
    },
    options: {
      ...CHART_DEFAULTS,
      scales: {
        x: { ticks: { color: '#7a8099', font: { family: 'DM Sans', size: 11 } }, grid: { color: 'rgba(255,255,255,.04)' } },
        y: { ticks: { color: '#7a8099', font: { family: 'DM Sans', size: 11 } }, grid: { color: 'rgba(255,255,255,.04)' } }
      },
      plugins: { ...CHART_DEFAULTS.plugins, legend: { display: false } }
    }
  });
}

function renderChartTimeline(canvasId, data) {
  destroyChart(canvasId);
  const byDate = {};
  data.forEach(r => {
    const d = (r.FECHA_PROCESAMIENTO||'').substring(0, 10);
    if (d) byDate[d] = (byDate[d]||0)+1;
  });
  const sorted = Object.entries(byDate).sort((a,b)=>a[0].localeCompare(b[0]));
  const ctx = document.getElementById(canvasId)?.getContext('2d');
  if (!ctx) return;
  chartInstances[canvasId] = new Chart(ctx, {
    type: 'line',
    data: {
      labels: sorted.map(([k])=>k),
      datasets: [{
        label: 'Registros procesados',
        data: sorted.map(([,v])=>v),
        borderColor: '#C5D336',
        backgroundColor: 'rgba(197,211,54,.08)',
        fill: true, tension: 0.4,
        pointRadius: 3, pointBackgroundColor: '#C5D336', borderWidth: 2,
      }]
    },
    options: {
      ...CHART_DEFAULTS,
      scales: {
        x: { ticks: { color: '#7a8099', maxTicksLimit: 12, font:{family:'DM Sans',size:10} }, grid: { color: 'rgba(255,255,255,.04)' } },
        y: { ticks: { color: '#7a8099', font:{family:'DM Sans',size:10} }, grid: { color: 'rgba(255,255,255,.04)' } }
      }
    }
  });
}

function renderChartCiudades(canvasId, data) {
  destroyChart(canvasId);
  const counts = {};
  data.forEach(r => { const c = r.CIUDAD_DESTINO; if (c && c !== 'NO APLICA') counts[c] = (counts[c]||0)+1; });
  const top = Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const ctx = document.getElementById(canvasId)?.getContext('2d');
  if (!ctx) return;
  chartInstances[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: top.map(([k])=>k),
      datasets: [{
        data: top.map(([,v])=>v),
        backgroundColor: 'rgba(197,211,54,.7)',
        borderRadius: 6, borderSkipped: false,
      }]
    },
    options: {
      ...CHART_DEFAULTS,
      indexAxis: 'y',
      scales: {
        x: { ticks:{color:'#7a8099',font:{family:'DM Sans',size:10}}, grid:{color:'rgba(255,255,255,.04)'} },
        y: { ticks:{color:'#e8eaf0',font:{family:'DM Sans',size:11}}, grid:{color:'rgba(255,255,255,.04)'} }
      },
      plugins: { ...CHART_DEFAULTS.plugins, legend:{display:false} }
    }
  });
}

function renderChartTasaEntrega(canvasId, data) {
  destroyChart(canvasId);
  const carriers = ['COORDINADORA','LOGICUARTAS','SERVIENTREGA','VELOCES','INTERRAPIDISIMO'];
  const labels = [], entregPct = [], novedPct = [], transitPct = [];
  carriers.forEach(c => {
    const rows = data.filter(r => r.TRANSPORTADORA === c);
    if (!rows.length) return;
    const { entregadas, transito, novedad } = contarEstados(rows);
    labels.push(c);
    entregPct.push(+(entregadas / rows.length * 100).toFixed(1));
    transitPct.push(+(transito / rows.length * 100).toFixed(1));
    novedPct.push(+(novedad / rows.length * 100).toFixed(1));
  });
  const ctx = document.getElementById(canvasId)?.getContext('2d');
  if (!ctx) return;
  chartInstances[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Entregadas %', data: entregPct, backgroundColor: 'rgba(74,222,128,.75)', borderRadius: 6 },
        { label: 'En Tránsito %', data: transitPct, backgroundColor: 'rgba(167,139,250,.75)', borderRadius: 6 },
        { label: 'Novedad %', data: novedPct, backgroundColor: 'rgba(250,204,21,.75)', borderRadius: 6 },
      ]
    },
    options: {
      ...CHART_DEFAULTS,
      scales: {
        x: { ticks:{color:'#7a8099',font:{family:'DM Sans',size:11}}, grid:{color:'rgba(255,255,255,.04)'} },
        y: { ticks:{color:'#7a8099',font:{family:'DM Sans',size:10},callback:v=>v+'%'}, grid:{color:'rgba(255,255,255,.04)'}, max:100 }
      }
    }
  });
}

function renderChartNovedadesTipo(canvasId, data) {
  destroyChart(canvasId);
  // Agrupar novedades por el campo NOVEDAD (solo registros con estado novedad)
  const novedadRows = data.filter(r => normalizeEstado(r.ESTADO) === 'novedad');
  const counts = {};
  novedadRows.forEach(r => {
    const v = (r.NOVEDAD || r.ESTADO || 'SIN DESCRIPCIÓN').trim();
    const key = v.substring(0, 40);
    counts[key] = (counts[key]||0)+1;
  });
  const top = Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0, 8);
  const ctx = document.getElementById(canvasId)?.getContext('2d');
  if (!ctx) return;
  chartInstances[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: top.map(([k])=>k),
      datasets: [{
        data: top.map(([,v])=>v),
        backgroundColor: PALETTE.map(c=>c+'cc'),
        borderRadius: 6, borderSkipped: false,
      }]
    },
    options: {
      ...CHART_DEFAULTS,
      indexAxis: 'y',
      scales: {
        x: { ticks:{color:'#7a8099',font:{family:'DM Sans',size:10}}, grid:{color:'rgba(255,255,255,.04)'} },
        y: { ticks:{color:'#e8eaf0',font:{family:'DM Sans',size:10},maxTicksLimit:10}, grid:{color:'rgba(255,255,255,.04)'} }
      },
      plugins: { ...CHART_DEFAULTS.plugins, legend:{display:false} }
    }
  });
}

function renderChartEstadosCat(canvasId, data) {
  // Gráfico de barras agrupadas por categoría normalizada
  destroyChart(canvasId);
  const cats = ['entregada','transito','novedad','pendiente','devuelta','cancelada','otro'];
  const labels = ['Entregadas','En Tránsito','Novedad','Pendientes','Devueltas','Canceladas','Otro'];
  const colors = ['#4ade80','#a78bfa','#facc15','#60a5fa','#f87171','#4a5068','#7a8099'];
  const counts = cats.map(() => 0);
  data.forEach(r => {
    const idx = cats.indexOf(normalizeEstado(r.ESTADO));
    if (idx >= 0) counts[idx]++;
  });
  const ctx = document.getElementById(canvasId)?.getContext('2d');
  if (!ctx) return;
  chartInstances[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        data: counts,
        backgroundColor: colors.map(c=>c+'cc'),
        borderRadius: 8, borderSkipped: false,
      }]
    },
    options: {
      ...CHART_DEFAULTS,
      scales: {
        x: { ticks:{color:'#7a8099',font:{family:'DM Sans',size:11}}, grid:{color:'rgba(255,255,255,.04)'} },
        y: { ticks:{color:'#7a8099',font:{family:'DM Sans',size:10}}, grid:{color:'rgba(255,255,255,.04)'} }
      },
      plugins: { ...CHART_DEFAULTS.plugins, legend:{display:false} }
    }
  });
}

// ══════════════════════════════════════════════════
//  NORMALIZACIÓN DE ESTADOS
// ══════════════════════════════════════════════════
function normalizeEstado(raw) {
  // Guard: valores vacíos, booleanos falsos, o literales "false"/"null"/"undefined"
  if (raw === null || raw === undefined || raw === false ||
      raw === 'false' || raw === 'False' || raw === 'FALSE' ||
      raw === 'null' || raw === 'undefined' || raw === '') return 'otro';

  const e = String(raw).toUpperCase().trim();
  if (!e) return 'otro';

  // ENTREGADAS A REMITENTE → devuelta (debe ir ANTES que la regla genérica de ENTREGAD)
  if (/ENTREGADO A REMITENTE|ENTREGADA A REMITENTE/.test(e)) return 'devuelta';

  // ENTREGADAS — primero, antes de cualquier otra regla
  if (/^ENTREGAD/.test(e)) return 'entregada';
  // Aliases explícitos por si acaso
  if (e === 'ENTREGADO' || e === 'ENTREGADA') return 'entregada';

  // DEVOLUCIONES
  if (/DEVOLU|DEVUELT|PROCESO DE DEVOL|EN PROCESO DE DEV/.test(e)) return 'devuelta';

  // NOVEDADES — SOLO estados que son puramente novedad, nunca entregados
  if (/^NOVEDAD$|^NOVEDAD \(|CERRADO POR INCIDENCIA|^SINIESTRO$|^SINIESTRAD/.test(e)) return 'novedad';

  // TRÁNSITO
  if (
    /^DESPACH|EN TRANSPORTE|EN TERMINAL|EN REPARTO|^EN CAMINO$|^EN RUTA$|^EN TRANSITO$|EN TR[AÁ]N|EN PUNTO|EN PROCESAMIENTO|^BODEGA/.test(e) ||
    /^En Camino$/i.test(String(raw))
  ) return 'transito';

  // CANCELADAS
  if (/^CANCELADO$|^NO APLICA$/.test(e)) return 'cancelada';

  // PENDIENTES — incluye EN ALISTAMIENTO DEL CLIENTE y RECIBIDO DEL CLIENTE
  if (/^INGRESADO$|^GU[IÍ]A GENERADA$|^RECIBIDO DEL CLIENTE$|^EN ALISTAMIENTO DEL CLIENTE$|^PEND/.test(e)) return 'pendiente';

  // INTERRAPIDISIMO ADICIONALES
  if (/^EN PROCESAMIENTO$|^EN CAMINO/.test(e)) return 'transito';
  if (/T[ÚU] ENV[ÍI]O FUE ENTREGADO|YA PUEDES RECOGER TU ENV[ÍI]O/.test(e)) return 'entregada';

  return 'otro';
}

function contarEstados(data) {
  let entregadas = 0, transito = 0, novedad = 0, pendientes = 0, devueltas = 0, canceladas = 0, otro = 0;
  data.forEach(r => {
    const n = normalizeEstado(r.ESTADO);
    if      (n === 'entregada')  entregadas++;
    else if (n === 'transito')   transito++;
    else if (n === 'novedad')    novedad++;
    else if (n === 'devuelta')   devueltas++;
    else if (n === 'cancelada')  canceladas++;
    else if (n === 'pendiente')  pendientes++;
    else                         otro++;
  });
  return { entregadas, transito, novedad, pendientes, devueltas, canceladas, otro };
}

function pct(n, total) {
  if (!total) return '0%';
  return Math.round(n / total * 100) + '%';
}

function esc(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function statusPill(estado) {
  const n = normalizeEstado(estado);
  let cls = 'muted';
  if      (n === 'entregada')  cls = 'ok';
  else if (n === 'novedad')    cls = 'warn';
  else if (n === 'devuelta')   cls = 'danger';
  else if (n === 'transito')   cls = 'transit';
  else if (n === 'pendiente')  cls = 'info';
  else if (n === 'cancelada')  cls = 'muted';
  return `<span class="pill pill-${cls}">${esc((estado||'—').substring(0,35))}</span>`;
}

function mkPagination(containerId, current, pages, onPage) {
  const el = document.getElementById(containerId);
  if (!el) return;
  if (pages <= 1) { el.innerHTML = ''; return; }

  const range = [];
  for (let i = Math.max(1, current-2); i <= Math.min(pages, current+2); i++) range.push(i);

  el.innerHTML = `
    <button class="pag-btn" onclick="${onPage.toString().replace('p => ', '').replace('onPage', '')}(${current - 1})" ${current <= 1 ? 'disabled' : ''}>←</button>
    ${current > 3 ? `<button class="pag-btn" onclick="${onPage.toString().replace('p => ', '').replace('onPage', '')}(1)">1</button><span style="color:var(--text-muted);padding:0 4px;">…</span>` : ''}
    ${range.map(p => `<button class="pag-btn ${p === current ? 'active' : ''}" onclick="${onPage.toString().replace('p => ', '').replace('onPage', '')}(${p})">${p}</button>`).join('')}
    ${current < pages - 2 ? `<span style="color:var(--text-muted);padding:0 4px;">…</span><button class="pag-btn" onclick="${onPage.toString().replace('p => ', '').replace('onPage', '')}(${pages})">${pages}</button>` : ''}
    <button class="pag-btn" onclick="${onPage.toString().replace('p => ', '').replace('onPage', '')}(${current + 1})" ${current >= pages ? 'disabled' : ''}>→</button>
    <span style="font-size:11px;color:var(--text-muted);margin-left:8px;">${current} / ${pages}</span>
  `;
}

// ══════════════════════════════════════════════════
//  NUEVAS GRÁFICAS
// ══════════════════════════════════════════════════

// Devoluciones y Novedades absolutas por transportadora
function renderChartDevNovCarrier(canvasId, data) {
  destroyChart(canvasId);
  const carriers = ['COORDINADORA','LOGICUARTAS','SERVIENTREGA','VELOCES','INTERRAPIDISIMO'];
  const devData = [], novData = [], cancData = [], labels = [];
  carriers.forEach(c => {
    const rows = data.filter(r => r.TRANSPORTADORA === c);
    if (!rows.length) return;
    const { devueltas, novedad, canceladas } = contarEstados(rows);
    labels.push(c);
    devData.push(devueltas);
    novData.push(novedad);
    cancData.push(canceladas);
  });
  const ctx = document.getElementById(canvasId)?.getContext('2d');
  if (!ctx) return;
  chartInstances[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Devueltas', data: devData, backgroundColor: 'rgba(248,113,113,.8)', borderRadius: 6 },
        { label: 'Novedad',   data: novData, backgroundColor: 'rgba(250,204,21,.8)',  borderRadius: 6 },
        { label: 'Canceladas',data: cancData,backgroundColor: 'rgba(74,80,104,.8)',   borderRadius: 6 },
      ]
    },
    options: {
      ...CHART_DEFAULTS,
      scales: {
        x: { ticks:{color:'#7a8099',font:{family:'DM Sans',size:11}}, grid:{color:'rgba(255,255,255,.04)'} },
        y: { ticks:{color:'#7a8099',font:{family:'DM Sans',size:10}}, grid:{color:'rgba(255,255,255,.04)'} }
      }
    }
  });
}

// Pendientes vs En Tránsito (donut por categoría)
function renderChartPendTransito(canvasId, data) {
  destroyChart(canvasId);
  const { transito, pendientes, entregadas, novedad, devueltas, canceladas, otro } = contarEstados(data);
  const ctx = document.getElementById(canvasId)?.getContext('2d');
  if (!ctx) return;
  chartInstances[canvasId] = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['Entregadas','En Tránsito','Pendientes','Novedad','Devueltas','Canceladas','Otro'],
      datasets: [{
        data: [entregadas, transito, pendientes, novedad, devueltas, canceladas, otro],
        backgroundColor: [
          'rgba(74,222,128,.85)',
          'rgba(167,139,250,.85)',
          'rgba(96,165,250,.85)',
          'rgba(250,204,21,.85)',
          'rgba(248,113,113,.85)',
          'rgba(74,80,104,.85)',
          'rgba(122,128,153,.85)',
        ],
        borderWidth: 0, hoverOffset: 8
      }]
    },
    options: { ...CHART_DEFAULTS, cutout: '60%' }
  });
}

// Tiempo promedio de entrega por transportadora
function renderChartTiempoEntrega(canvasId, data) {
  destroyChart(canvasId);
  const carriers = ['COORDINADORA','LOGICUARTAS','SERVIENTREGA','VELOCES','INTERRAPIDISIMO'];
  const labels = [], avgDays = [];
  carriers.forEach(c => {
    const rows = data.filter(r => r.TRANSPORTADORA === c && r.FECHA_DESPACHO && r.FECHA_ENTREGA && normalizeEstado(r.ESTADO) === 'entregada');
    if (rows.length < 3) return;
    const diffs = rows.map(r => {
      const d1 = new Date(r.FECHA_DESPACHO);
      const d2 = new Date(r.FECHA_ENTREGA);
      if (isNaN(d1) || isNaN(d2) || d2 < d1) return null;
      return (d2 - d1) / 86400000;
    }).filter(d => d !== null && d >= 0 && d <= 60);
    if (!diffs.length) return;
    const avg = diffs.reduce((a,b) => a+b, 0) / diffs.length;
    labels.push(c);
    avgDays.push(+avg.toFixed(1));
  });
  const ctx = document.getElementById(canvasId)?.getContext('2d');
  if (!ctx) return;
  chartInstances[canvasId] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Días promedio',
        data: avgDays,
        backgroundColor: labels.map(l => (CARRIER_COLORS[l]||'#C5D336') + 'cc'),
        borderRadius: 8, borderSkipped: false,
      }]
    },
    options: {
      ...CHART_DEFAULTS,
      scales: {
        x: { ticks:{color:'#7a8099',font:{family:'DM Sans',size:11}}, grid:{color:'rgba(255,255,255,.04)'} },
        y: { ticks:{color:'#7a8099',font:{family:'DM Sans',size:10},callback:v=>v+' d'}, grid:{color:'rgba(255,255,255,.04)'}, beginAtZero:true }
      },
      plugins: { ...CHART_DEFAULTS.plugins, legend:{display:false} }
    }
  });
}

// ══════════════════════════════════════════════════
//  EXPORTAR EXCEL
// ══════════════════════════════════════════════════
function _toExcelData(rows) {
  return rows.map(r => ({
    'Guía':               r.GUIA||'',
    'Documento':          r.DOCUMENTO||'',
    'Transportadora':     r.TRANSPORTADORA||'',
    'Estado':             r.ESTADO||'',
    'Último Movimiento':  r.ULTIMO_MOVIMIENTO||'',
    'Novedad':            r.NOVEDAD||'',
    'Ciudad Origen':      r.CIUDAD_ORIGEN||'',
    'Ciudad Destino':     r.CIUDAD_DESTINO||'',
    'Nombre Destinatario':r.NOMBRE_DESTINATARIO||'',
    'Dirección':          r.DIRECCION_DESTINATARIO||'',
    'Teléfono':           r.TELEFONO_DESTINATARIO||'',
    'Fecha Despacho':     r.FECHA_DESPACHO||'',
    'Fecha Entrega':      r.FECHA_ENTREGA||'',
    'Fecha Procesamiento':r.FECHA_PROCESAMIENTO||'',
  }));
}

function _exportXlsx(data, sheetName, filename) {
  if (!data.length) { alert('Sin datos para exportar.'); return; }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = Object.keys(data[0]).map(k => ({
    wch: Math.max(k.length + 2, ...data.slice(0,50).map(r => String(r[k]||'').length))
  }));
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, `${filename}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

function exportarExcelGeneral() {
  _exportXlsx(_toExcelData(RAW), 'Todas las Guías', 'Tracking_Lineacom_Total');
}

function exportarExcelFiltrado() {
  _exportXlsx(_toExcelData(FILTERED_GENERAL), 'Filtrado', 'Tracking_Lineacom_Filtrado');
}

function exportarCarrier(carrier) {
  const data = _carrierFiltered[carrier] || RAW.filter(r => r.TRANSPORTADORA === carrier);
  _exportXlsx(_toExcelData(data), carrier, `Tracking_${carrier}`);
}