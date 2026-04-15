// ══════════════════════════════════════════════════
//  MÓDULO: CONSULTAR GUÍAS — Individual y Masiva
//  Se agrega al final de App.js existente
// ══════════════════════════════════════════════════

// ── CONSULTA INDIVIDUAL ────────────────────────────
function consultarGuiaIndividual() {
  const input = document.getElementById('consulta-guia-input');
  if (!input) return;
  const guia = input.value.trim().toUpperCase();
  const resultDiv = document.getElementById('consulta-guia-result');
  if (!resultDiv) return;

  if (!guia) {
    resultDiv.innerHTML = '';
    return;
  }

  if (!RAW.length) {
    resultDiv.innerHTML = `<div class="guia-not-found"><div class="guia-not-found-icon">⚠️</div><div class="guia-not-found-title">Datos no disponibles</div><div class="guia-not-found-sub">Los datos aún no han cargado. Por favor espera.</div></div>`;
    return;
  }

  // Buscar por número de guía (coincidencia exacta o parcial)
  const matches = RAW.filter(r =>
    (r.GUIA||'').toUpperCase().includes(guia) ||
    (r.DOCUMENTO||'').toUpperCase().includes(guia)
  );

  if (!matches.length) {
    resultDiv.innerHTML = `
      <div class="guia-not-found">
        <div class="guia-not-found-icon">🔍</div>
        <div class="guia-not-found-title">Guía no encontrada</div>
        <div class="guia-not-found-sub">No se encontró la guía <span style="font-family:'Space Mono',monospace;color:var(--lima);">${esc(guia)}</span> en el sistema.<br>Verifica el número e intenta de nuevo.</div>
      </div>`;
    return;
  }

  const r = matches[0]; // Si hay múltiples, mostramos la primera y tabla si hay más
  const carrierColor = CARRIER_COLORS[r.TRANSPORTADORA] || '#aaa';
  const novedad = (r.NOVEDAD || '') !== '' && r.NOVEDAD !== 'NO APLICA' && r.NOVEDAD !== 'no aplica';

  let html = `
    <div class="guia-result-card">
      <div class="guia-result-header">
        <div>
          <div class="guia-result-num">${esc(r.GUIA || guia)}</div>
          <div class="guia-result-carrier" style="color:${carrierColor};">
            ${CARRIER_ICONS[r.TRANSPORTADORA] || '📦'} ${esc(r.TRANSPORTADORA || '—')}
            ${r.DOCUMENTO ? `<span style="color:var(--text-muted);margin-left:8px;">Doc: ${esc(r.DOCUMENTO)}</span>` : ''}
          </div>
        </div>
        <div>${statusPill(r.ESTADO || '')}</div>
      </div>
      ${novedad ? `
      <div class="guia-novedad-banner">
        <div class="guia-novedad-icon">⚠️</div>
        <div class="guia-novedad-text"><strong>Novedad:</strong> ${esc(r.NOVEDAD)}</div>
      </div>` : ''}
      <div class="guia-fields-grid">
        ${_guiaField('Último Movimiento', r.ULTIMO_MOVIMIENTO)}
        ${_guiaField('Destinatario', r.NOMBRE_DESTINATARIO)}
        ${_guiaField('Ciudad Destino', r.CIUDAD_DESTINO)}
        ${_guiaField('Ciudad Origen', r.CIUDAD_ORIGEN)}
        ${_guiaField('Dirección', r.DIRECCION_DESTINATARIO)}
        ${_guiaField('Teléfono', r.TELEFONO_DESTINATARIO)}
        ${_guiaField('Fecha Despacho', r.FECHA_DESPACHO)}
        ${_guiaField('Fecha Entrega', r.FECHA_ENTREGA)}
        ${_guiaField('Fecha Procesamiento', r.FECHA_PROCESAMIENTO)}
      </div>
    </div>`;

  // Si hay más de un resultado, mostrar tabla de todos
  if (matches.length > 1) {
    html += `
      <div style="margin-top:20px;">
        <div class="table-header" style="background:var(--surface);border-radius:var(--radius-sm) var(--radius-sm) 0 0;border:1px solid var(--border);border-bottom:none;padding:14px 20px;">
          <span class="table-count">${matches.length} registros encontrados con "${esc(guia)}"</span>
          <button class="btn-secondary" style="font-size:12px;padding:6px 14px;" onclick="_exportXlsx(_toExcelData(window._consultaMatches||[]),'Consulta','Consulta_Guias')">↓ Exportar todos</button>
        </div>
        <div class="table-container" style="border-radius:0 0 var(--radius-sm) var(--radius-sm);">
          <div class="table-scroll">
            <table>
              <thead><tr>
                <th>Guía</th><th>Transportadora</th><th>Estado</th>
                <th>Destinatario</th><th>Ciudad Destino</th><th>Último Mov.</th><th>Novedad</th><th>F. Entrega</th>
              </tr></thead>
              <tbody>
                ${matches.map(m => `<tr>
                  <td><span class="cell-guia">${esc(m.GUIA||'—')}</span></td>
                  <td><span style="color:${CARRIER_COLORS[m.TRANSPORTADORA]||'#aaa'};font-weight:600;font-size:11px;">${esc(m.TRANSPORTADORA||'—')}</span></td>
                  <td>${statusPill(m.ESTADO||'')}</td>
                  <td>${esc(m.NOMBRE_DESTINATARIO||'—')}</td>
                  <td>${esc(m.CIUDAD_DESTINO||'—')}</td>
                  <td>${esc(m.ULTIMO_MOVIMIENTO||'—')}</td>
                  <td>${m.NOVEDAD && m.NOVEDAD!=='NO APLICA' ? `<span style="color:var(--warn);font-size:11px;">${esc((m.NOVEDAD||'').substring(0,60))}</span>` : '<span style="color:var(--text-muted)">—</span>'}</td>
                  <td>${esc(m.FECHA_ENTREGA||'—')}</td>
                </tr>`).join('')}
              </tbody>
            </table>
          </div>
        </div>
      </div>`;
    window._consultaMatches = matches;
  }

  resultDiv.innerHTML = html;
}

function _guiaField(label, value) {
  const val = value && value !== 'NO APLICA' ? value : '—';
  return `<div class="guia-field">
    <div class="guia-field-label">${label}</div>
    <div class="guia-field-value">${esc(val)}</div>
  </div>`;
}

// Enter key para consulta individual
function _consultaKeydown(e) {
  if (e.key === 'Enter') consultarGuiaIndividual();
}

// ── CONSULTA MASIVA ────────────────────────────────
let _masivaGuias   = [];   // guías leídas del Excel
let _masivaResultados = []; // registros encontrados
let _masivaNoEncontradas = []; // guías no encontradas
let _masivaViewTab = 'encontradas'; // 'encontradas' | 'no-encontradas'
let _masivaFiltered = [];
let _masivaPag = 1;

// Descarga la plantilla Excel
function descargarPlantillaMasiva() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['NUMERO_GUIA'],
    ['Ej: 1234567890'],
    ['Ej: ABC9876543'],
  ]);
  ws['!cols'] = [{ wch: 22 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Guías');
  XLSX.writeFile(wb, 'Plantilla_Consulta_Masiva.xlsx');
}

// Leer Excel de guías subido
function onMasivaFileChange(e) {
  const file = e.target.files[0];
  if (!file) return;
  _processMasivaFile(file);
}

function onMasivaDrop(e) {
  e.preventDefault();
  const dz = document.getElementById('masiva-drop-zone');
  if (dz) dz.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (!file) return;
  _processMasivaFile(file);
}

function _processMasivaFile(file) {
  const nameEl = document.getElementById('masiva-file-name');
  if (nameEl) nameEl.textContent = file.name;

  const reader = new FileReader();
  reader.onload = function(ev) {
    try {
      const wb = XLSX.read(ev.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

      // Buscar columna de guías (cualquier columna que diga GUIA, NUMERO, TRACKING…)
      const header = (rows[0] || []).map(c => String(c||'').toUpperCase().trim());
      let colIdx = header.findIndex(h =>
        h.includes('GUIA') || h.includes('GUÍA') || h.includes('NUMERO') ||
        h.includes('NÚMERO') || h.includes('TRACKING') || h.includes('CÓDIGO') ||
        h.includes('CODIGO')
      );
      if (colIdx === -1) colIdx = 0; // fallback: primera columna

      const guias = rows.slice(1)
        .map(row => String(row[colIdx] || '').trim().toUpperCase())
        .filter(g => g && g.length > 2 && !g.includes('EJ:'));

      _masivaGuias = [...new Set(guias)]; // dedup
      const countEl = document.getElementById('masiva-guias-count');
      if (countEl) countEl.textContent = `${_masivaGuias.length} guías leídas`;

      // Habilitar botón consultar
      const btn = document.getElementById('btn-consultar-masiva');
      if (btn) btn.disabled = false;

      // Resetear resultados anteriores
      _resetMasivaResults();

    } catch(err) {
      alert('Error leyendo el archivo: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function _resetMasivaResults() {
  _masivaResultados = [];
  _masivaNoEncontradas = [];
  _masivaFiltered = [];
  const resDiv = document.getElementById('masiva-results');
  if (resDiv) resDiv.innerHTML = '';
  const progCard = document.getElementById('masiva-progress-card');
  if (progCard) progCard.classList.remove('visible');
}

// Ejecutar consulta masiva
async function ejecutarConsultaMasiva() {
  if (!_masivaGuias.length) return;
  if (!RAW.length) {
    alert('Los datos aún no han cargado. Espera a que el dashboard termine de cargar.');
    return;
  }

  const btn = document.getElementById('btn-consultar-masiva');
  if (btn) btn.disabled = true;

  // Mostrar progreso
  const progCard = document.getElementById('masiva-progress-card');
  const progBar  = document.getElementById('masiva-prog-bar-fill');
  const progStat = document.getElementById('masiva-prog-stats');
  if (progCard) progCard.classList.add('visible');

  _masivaResultados = [];
  _masivaNoEncontradas = [];

  const total = _masivaGuias.length;

  for (let i = 0; i < total; i++) {
    const g = _masivaGuias[i];
    const found = RAW.filter(r =>
      (r.GUIA||'').toUpperCase() === g ||
      (r.DOCUMENTO||'').toUpperCase() === g ||
      (r.GUIA||'').toUpperCase().includes(g) ||
      (r.DOCUMENTO||'').toUpperCase().includes(g)
    );

    if (found.length) {
      _masivaResultados.push(...found);
    } else {
      _masivaNoEncontradas.push(g);
    }

    // Actualizar barra cada 10 o al final
    if (i % 10 === 0 || i === total - 1) {
      const pct = Math.round((i + 1) / total * 100);
      if (progBar)  progBar.style.width  = pct + '%';
      if (progStat) progStat.textContent = `${i + 1} / ${total} guías procesadas · ${_masivaResultados.length} encontradas · ${_masivaNoEncontradas.length} no encontradas`;
      // yield para no bloquear UI
      await new Promise(r => setTimeout(r, 0));
    }
  }

  if (btn) btn.disabled = false;

  // Deduplicar resultados (una guía puede tener múltiples registros históricos)
  _masivaFiltered = [..._masivaResultados];
  _masivaPag = 1;
  _renderMasivaResults();
}

function _renderMasivaResults() {
  const resDiv = document.getElementById('masiva-results');
  if (!resDiv) return;

  const totalGuias    = _masivaGuias.length;
  const encontradas   = [...new Set(_masivaResultados.map(r => r.GUIA||r.DOCUMENTO||''))].length;
  const noEncontradas = _masivaNoEncontradas.length;
  const { entregadas, transito, novedad, pendientes, devueltas } = contarEstados(_masivaResultados);

  resDiv.innerHTML = `
    <!-- KPIs -->
    <div class="masiva-kpi-row">
      <div class="masiva-kpi">
        <div class="masiva-kpi-icon">📋</div>
        <div class="masiva-kpi-val">${totalGuias.toLocaleString()}</div>
        <div class="masiva-kpi-lbl">Consultadas</div>
      </div>
      <div class="masiva-kpi">
        <div class="masiva-kpi-icon">✅</div>
        <div class="masiva-kpi-val" style="color:var(--ok);">${encontradas.toLocaleString()}</div>
        <div class="masiva-kpi-lbl">Encontradas</div>
      </div>
      <div class="masiva-kpi">
        <div class="masiva-kpi-icon">❌</div>
        <div class="masiva-kpi-val" style="color:var(--danger);">${noEncontradas.toLocaleString()}</div>
        <div class="masiva-kpi-lbl">No encontradas</div>
      </div>
      <div class="masiva-kpi">
        <div class="masiva-kpi-icon">🚚</div>
        <div class="masiva-kpi-val" style="color:var(--transit);">${entregadas.toLocaleString()}</div>
        <div class="masiva-kpi-lbl">Entregadas</div>
      </div>
      <div class="masiva-kpi">
        <div class="masiva-kpi-icon">⚠️</div>
        <div class="masiva-kpi-val" style="color:var(--warn);">${novedad.toLocaleString()}</div>
        <div class="masiva-kpi-lbl">Con Novedad</div>
      </div>
      <div class="masiva-kpi">
        <div class="masiva-kpi-icon">⏳</div>
        <div class="masiva-kpi-val" style="color:var(--info);">${pendientes.toLocaleString()}</div>
        <div class="masiva-kpi-lbl">Pendientes</div>
      </div>
      <div class="masiva-kpi">
        <div class="masiva-kpi-icon">↩️</div>
        <div class="masiva-kpi-val" style="color:var(--danger);">${devueltas.toLocaleString()}</div>
        <div class="masiva-kpi-lbl">Devueltas</div>
      </div>
    </div>

    <!-- Switch tabs + botones exportar -->
    <div class="masiva-results-header">
      <div class="masiva-tabs" id="masiva-inner-tabs">
        <button class="masiva-tab-btn ${_masivaViewTab==='encontradas'?'active':''}" onclick="switchMasivaTab('encontradas')">
          ✅ Encontradas (${_masivaResultados.length})
        </button>
        <button class="masiva-tab-btn ${_masivaViewTab==='no-encontradas'?'active':''}" onclick="switchMasivaTab('no-encontradas')">
          ❌ No encontradas (${noEncontradas})
        </button>
      </div>
      <div class="masiva-export-row">
        <button class="btn-secondary" onclick="exportarMasivaEncontradas()">↓ Excel encontradas</button>
        ${noEncontradas ? `<button class="btn-secondary" style="border-color:rgba(248,113,113,.3);color:var(--danger);" onclick="exportarMasivaNoEncontradas()">↓ Excel no encontradas</button>` : ''}
      </div>
    </div>

    <!-- Contenido según tab -->
    <div id="masiva-tab-content">
      ${_buildMasivaTabContent()}
    </div>
  `;
}

function switchMasivaTab(tab) {
  _masivaViewTab = tab;
  // Actualizar botones
  document.querySelectorAll('.masiva-tab-btn').forEach(b => {
    b.classList.toggle('active', b.textContent.includes(tab === 'encontradas' ? 'Encontradas' : 'No encontradas'));
  });
  const content = document.getElementById('masiva-tab-content');
  if (content) content.innerHTML = _buildMasivaTabContent();
}

function _buildMasivaTabContent() {
  if (_masivaViewTab === 'no-encontradas') {
    return _buildNoEncontradasPanel();
  }
  return _buildEncontradasPanel();
}

function _buildEncontradasPanel() {
  if (!_masivaResultados.length) {
    return `<div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-text">No se encontraron registros.</div></div>`;
  }

  // Filtro rápido
  const PAGE = 50;
  const totalPages = Math.max(1, Math.ceil(_masivaFiltered.length / PAGE));
  const slice = _masivaFiltered.slice((_masivaPag - 1) * PAGE, _masivaPag * PAGE);

  const cols = [
    { label: 'Guía',           fn: r => `<span class="cell-guia">${esc(r.GUIA||'—')}</span>` },
    { label: 'Transportadora', fn: r => `<span style="color:${CARRIER_COLORS[r.TRANSPORTADORA]||'#aaa'};font-weight:600;font-size:11px;">${esc(r.TRANSPORTADORA||'—')}</span>` },
    { label: 'Estado',         fn: r => statusPill(r.ESTADO||'') },
    { label: 'Destinatario',   fn: r => esc(r.NOMBRE_DESTINATARIO||'—') },
    { label: 'Ciudad Destino', fn: r => esc(r.CIUDAD_DESTINO||'—') },
    { label: 'Último Mov.',    fn: r => esc(r.ULTIMO_MOVIMIENTO||'—') },
    { label: 'Novedad',        fn: r => { const n=r.NOVEDAD||''; return n && n!=='NO APLICA' ? `<span style="color:var(--warn);font-size:11px;">${esc(n.substring(0,60))}</span>` : '<span style="color:var(--text-muted)">—</span>'; }},
    { label: 'F. Entrega',     fn: r => esc(r.FECHA_ENTREGA||'—') },
  ];

  let pagHtml = '';
  if (totalPages > 1) {
    const range = [];
    for (let i = Math.max(1, _masivaPag-2); i <= Math.min(totalPages, _masivaPag+2); i++) range.push(i);
    pagHtml = `<div class="pagination">
      <button class="pag-btn" onclick="_masivaPag=Math.max(1,_masivaPag-1);_refreshMasivaTable()" ${_masivaPag<=1?'disabled':''}>←</button>
      ${range.map(p=>`<button class="pag-btn ${p===_masivaPag?'active':''}" onclick="_masivaPag=${p};_refreshMasivaTable()">${p}</button>`).join('')}
      <button class="pag-btn" onclick="_masivaPag=Math.min(${totalPages},_masivaPag+1);_refreshMasivaTable()" ${_masivaPag>=totalPages?'disabled':''}>→</button>
      <span style="font-size:11px;color:var(--text-muted);margin-left:8px;">${_masivaPag} / ${totalPages}</span>
    </div>`;
  }

  return `
    <!-- Filtro rápido encontradas -->
    <div class="filters-bar" style="margin-bottom:12px;">
      <span class="filter-label">Filtrar</span>
      <input class="filter-input" type="text" placeholder="Guía, destinatario, ciudad..." oninput="_filterMasiva(this.value)" style="min-width:280px;">
      <select class="filter-select" onchange="_filterMasivaEstado(this.value)">
        <option value="">Todos los estados</option>
        <option value="entregada">✅ Entregadas</option>
        <option value="transito">🚚 En Tránsito</option>
        <option value="novedad">⚠️ Novedad</option>
        <option value="pendiente">⏳ Pendientes</option>
        <option value="devuelta">↩️ Devueltas</option>
      </select>
    </div>
    <div class="table-container">
      <div class="table-header">
        <span class="table-count">${_masivaFiltered.length.toLocaleString()} registros</span>
      </div>
      <div class="table-scroll" id="masiva-table-scroll">
        <table>
          <thead><tr>${cols.map(c=>`<th>${c.label}</th>`).join('')}</tr></thead>
          <tbody>${slice.map(r=>`<tr>${cols.map(c=>`<td>${c.fn(r)}</td>`).join('')}</tr>`).join('')}</tbody>
        </table>
      </div>
      ${pagHtml}
    </div>`;
}

function _buildNoEncontradasPanel() {
  if (!_masivaNoEncontradas.length) {
    return `<div class="empty-state"><div class="empty-icon">🎉</div><div class="empty-text">¡Todas las guías fueron encontradas!</div></div>`;
  }
  return `
    <div class="no-enc-list">
      <div class="no-enc-header">
        <span style="font-size:18px;">❌</span>
        <div>
          <div class="no-enc-title">Guías no encontradas en el sistema</div>
          <div class="no-enc-count">${_masivaNoEncontradas.length} guías sin registro</div>
        </div>
      </div>
      <div class="no-enc-grid">
        ${_masivaNoEncontradas.map(g => `<div class="no-enc-chip">${esc(g)}</div>`).join('')}
      </div>
    </div>`;
}

// Filtro live en tabla de resultados masivos
let _masivaTextFilter = '';
let _masivaEstadoFilter = '';

function _filterMasiva(val) {
  _masivaTextFilter = val.toLowerCase().trim();
  _applyMasivaFilters();
}
function _filterMasivaEstado(val) {
  _masivaEstadoFilter = val;
  _applyMasivaFilters();
}
function _applyMasivaFilters() {
  _masivaFiltered = _masivaResultados.filter(r => {
    if (_masivaEstadoFilter && normalizeEstado(r.ESTADO) !== _masivaEstadoFilter) return false;
    if (_masivaTextFilter) {
      const hay = [r.GUIA, r.NOMBRE_DESTINATARIO, r.CIUDAD_DESTINO, r.TRANSPORTADORA, r.ESTADO, r.DOCUMENTO]
        .map(v => (v||'').toLowerCase()).join(' ');
      if (!hay.includes(_masivaTextFilter)) return false;
    }
    return true;
  });
  _masivaPag = 1;
  _refreshMasivaTable();
}

function _refreshMasivaTable() {
  // Re-render solo el panel de encontradas sin destruir el resto
  if (_masivaViewTab === 'encontradas') {
    const content = document.getElementById('masiva-tab-content');
    if (content) content.innerHTML = _buildEncontradasPanel();
  }
}

// Exportar encontradas a Excel
function exportarMasivaEncontradas() {
  if (!_masivaResultados.length) { alert('Sin resultados para exportar.'); return; }
  _exportXlsx(_toExcelData(_masivaResultados), 'Encontradas', 'Consulta_Masiva_Encontradas');
}

// Exportar no-encontradas a Excel
function exportarMasivaNoEncontradas() {
  if (!_masivaNoEncontradas.length) { alert('Todas las guías fueron encontradas.'); return; }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(_masivaNoEncontradas.map(g => ({ 'NUMERO_GUIA': g, 'ESTADO': 'NO ENCONTRADA' })));
  ws['!cols'] = [{ wch: 28 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, ws, 'No Encontradas');
  XLSX.writeFile(wb, `Consulta_Masiva_NoEncontradas_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// Drag-and-drop helpers
function _masivaDragOver(e) {
  e.preventDefault();
  const dz = document.getElementById('masiva-drop-zone');
  if (dz) dz.classList.add('drag-over');
}
function _masivaDragLeave() {
  const dz = document.getElementById('masiva-drop-zone');
  if (dz) dz.classList.remove('drag-over');
}