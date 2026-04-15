/* ============================================================
   AMBRAC Agenda de Visitas — app.js
   Stack: Supabase JS v2 (CDN) + vanilla JS, sem build
   ============================================================ */

const SUPABASE_URL      = 'https://xwvzqwjuicmvcwxjgwcp.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inh3dnpxd2p1aWNtdmN3eGpnd2NwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzUyNDEyODYsImV4cCI6MjA5MDgxNzI4Nn0.TaoeuDwu9rlpXLpkTk5VFFz8tqEJlPEXc8e3_cp_SKo';

const sb = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// ---- State ----
let currentDate = todayISO();
let allRecords  = [];
let editingId   = null;

// ---- Bootstrap ----
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('datePicker').value = currentDate;
  updateDayLabel();
  bindEvents();
  loadDay(currentDate);
});

// ---- Helpers ----
function todayISO() {
  return new Date().toISOString().split('T')[0];
}

const WEEKDAYS = [
  'domingo', 'segunda-feira', 'terça-feira',
  'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'
];

function updateDayLabel() {
  const d = new Date(currentDate + 'T12:00:00');
  document.getElementById('dayLabel').textContent = WEEKDAYS[d.getDay()];
}

function esc(str) {
  if (!str && str !== 0) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ---- Events ----
function bindEvents() {
  document.getElementById('btnPrev').addEventListener('click', () => shiftDay(-1));
  document.getElementById('btnNext').addEventListener('click', () => shiftDay(1));
  document.getElementById('datePicker').addEventListener('change', e => {
    currentDate = e.target.value;
    updateDayLabel();
    loadDay(currentDate);
  });

  document.getElementById('btnAdd').addEventListener('click', () => openModal(null));
  document.getElementById('modalClose').addEventListener('click', closeModal);
  document.getElementById('modalOverlay').addEventListener('click', closeModal);
  document.getElementById('btnCancel').addEventListener('click', closeModal);
  document.getElementById('btnDelete').addEventListener('click', deleteRecord);
  document.getElementById('form').addEventListener('submit', saveRecord);

  // Import
  document.getElementById('btnImport').addEventListener('click', openImport);
  document.getElementById('importClose').addEventListener('click', closeImport);
  document.getElementById('importOverlay').addEventListener('click', closeImport);
  document.getElementById('csvFile').addEventListener('change', onFileChosen);
  document.getElementById('btnImportBack').addEventListener('click', importGoStep1);
  document.getElementById('btnImportRun').addEventListener('click', runImport);
  document.getElementById('btnImportClose').addEventListener('click', closeImport);
  bindDropZone();

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeModal();
  });

  ['filtStatus', 'filtCategoria', 'filtAutor', 'filtSearch'].forEach(id => {
    document.getElementById(id).addEventListener('input', renderCards);
  });
  document.getElementById('btnClearFilters').addEventListener('click', clearFilters);
}

function shiftDay(n) {
  const d = new Date(currentDate + 'T12:00:00');
  d.setDate(d.getDate() + n);
  currentDate = d.toISOString().split('T')[0];
  document.getElementById('datePicker').value = currentDate;
  updateDayLabel();
  loadDay(currentDate);
}

// ---- Data ----
async function loadDay(date) {
  showLoading(true);

  const { data, error } = await sb
    .from('agendamentos')
    .select('*')
    .eq('data_visita', date)
    .order('ordem', { nullsFirst: false });

  showLoading(false);

  if (error) {
    console.error('Supabase error:', error);
    document.getElementById('empty').textContent = 'Erro ao carregar dados. Verifique o console.';
    document.getElementById('empty').style.display = 'block';
    return;
  }

  allRecords = data || [];
  populateAutorFilter();
  renderCards();
}

function showLoading(on) {
  document.getElementById('loading').style.display = on ? 'block' : 'none';
  document.getElementById('cards').style.display   = on ? 'none'  : 'grid';
  if (on) document.getElementById('empty').style.display = 'none';
}

function populateAutorFilter() {
  const autores = [...new Set(allRecords.map(r => r.autor).filter(Boolean))].sort();
  const sel = document.getElementById('filtAutor');
  const cur = sel.value;
  sel.innerHTML = '<option value="">Todos os autores</option>';
  autores.forEach(a => {
    const o = document.createElement('option');
    o.value = a;
    o.textContent = a;
    if (a === cur) o.selected = true;
    sel.appendChild(o);
  });
}

function clearFilters() {
  ['filtStatus', 'filtCategoria', 'filtAutor'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('filtSearch').value = '';
  renderCards();
}

function getFiltered() {
  const status    = document.getElementById('filtStatus').value;
  const categoria = document.getElementById('filtCategoria').value;
  const autor     = document.getElementById('filtAutor').value;
  const search    = document.getElementById('filtSearch').value.toLowerCase().trim();

  return allRecords.filter(r => {
    if (status    && r.status_documentacao !== status)    return false;
    if (categoria && r.categoria           !== categoria) return false;
    if (autor     && r.autor               !== autor)     return false;
    if (search) {
      const hay = [
        r.razao_social, r.cnpj, r.cidade, r.grupo,
        r.unidade, r.visita, r.autor, r.endereco
      ].join(' ').toLowerCase();
      if (!hay.includes(search)) return false;
    }
    return true;
  });
}

// ---- Render ----
function renderCards() {
  const filtered = getFiltered();
  const grid  = document.getElementById('cards');
  const empty = document.getElementById('empty');

  // Count
  document.getElementById('countBadge').textContent =
    `${filtered.length} empresa${filtered.length !== 1 ? 's' : ''}`;

  // Uber total (parse "R$ 8,94" etc.)
  const uberTotal = filtered.reduce((acc, r) => {
    if (!r.uber) return acc;
    const m = r.uber.replace(/\s/g, '').match(/[\d]+[,.][\d]{2}/);
    if (m) {
      const val = parseFloat(m[0].replace(',', '.'));
      if (!isNaN(val)) acc += val;
    }
    return acc;
  }, 0);
  document.getElementById('uberTotal').textContent =
    uberTotal > 0 ? `Uber total: R$ ${uberTotal.toFixed(2).replace('.', ',')}` : '';

  if (filtered.length === 0) {
    grid.innerHTML = '';
    empty.style.display = 'block';
    empty.textContent   = allRecords.length === 0
      ? 'Nenhum agendamento para este dia.'
      : 'Nenhum resultado para os filtros selecionados.';
    return;
  }
  empty.style.display = 'none';

  grid.innerHTML = filtered.map(cardHTML).join('');
  grid.querySelectorAll('.card').forEach(el => {
    el.addEventListener('click', () => openModal(el.dataset.id));
  });
}

// ---- Status → CSS class ----
function statusClass(s) {
  if (!s) return 's-DEFAULT';
  const u = s.toUpperCase();
  if (u.includes('NOVA'))                          return 's-NOVA';
  if (u.includes('VENCIDA'))                       return 's-VENCIDA';
  if (u.includes('ENVIADO'))                       return 's-ENVIADO';
  if (u.includes('VÁLIDA') || u.includes('VALIDA')) return 's-VALIDA';
  if (u.includes('ANDAMENTO'))                     return 's-ANDAMENTO';
  if (u.includes('MIGRA'))                         return 's-MIGRACAO';
  if (u.includes('NÃO ENVIAR') || u.includes('NAO ENVIAR')) return 's-NAO';
  if (u.includes('DESMARCADO'))                    return 's-DESMARC';
  return 's-DEFAULT';
}

function visitaClass(v) {
  if (!v) return '';
  const u = v.toUpperCase();
  if (u === 'OK') return 'visita-ok';
  if (u.includes('NÃO') || u.includes('NAO') || u.includes('DESMARC')) return 'visita-nao';
  return '';
}

// ---- Card HTML ----
function cardHTML(r) {
  const sc = statusClass(r.status_documentacao);
  const vc = visitaClass(r.visitas_feitas);

  const topBadges = [
    r.ordem    ? `<span class="badge badge-ordem">#${esc(r.ordem)}</span>` : '',
    r.categoria ? `<span class="badge badge-cat">${esc(r.categoria)}</span>` : '',
  ].filter(Boolean).join('');

  const statusBadge = r.status_documentacao
    ? `<span class="badge status-badge ${sc}">${esc(r.status_documentacao)}</span>`
    : '';

  const infoRows = [
    r.cidade     ? `<span class="card-info-row"><span class="card-icon">&#128205;</span>${esc(r.cidade)}</span>` : '',
    r.visita     ? `<span class="card-info-row"><span class="card-icon">&#128222;</span>${esc(r.visita)}</span>` : '',
    r.visitas_feitas
      ? `<span class="card-info-row ${vc}"><span class="card-icon">&#10003;</span>${esc(r.visitas_feitas)}</span>`
      : '',
  ].filter(Boolean).join('');

  const graziTag = r.grazi ? '<span class="card-grazi">Grazi</span>' : '';
  const uberTag  = r.uber && r.uber.trim()
    ? `<span class="card-uber">${esc(r.uber.trim())}</span>`
    : '';

  return `
    <div class="card" data-id="${esc(r.id)}" tabindex="0" role="button" aria-label="${esc(r.razao_social)}">
      <div class="card-top">
        <div class="card-left-badges">${topBadges}</div>
        ${statusBadge}
      </div>
      <div class="card-title">${esc(r.razao_social) || '—'}</div>
      <div class="card-cnpj">${esc(r.cnpj) || ''}${r.grupo ? ` · ${esc(r.grupo)}` : ''}</div>
      <div class="card-info">${infoRows}</div>
      <div class="card-footer">
        <span class="card-autor">${esc(r.autor) || ''}${graziTag}</span>
        ${uberTag}
      </div>
    </div>
  `;
}

// ---- Modal ----
function openModal(id) {
  editingId = id;
  const form      = document.getElementById('form');
  const deleteBtn = document.getElementById('btnDelete');

  if (id) {
    const r = allRecords.find(x => x.id === id);
    if (!r) return;
    document.getElementById('modalTitle').textContent = r.razao_social || 'Editar Visita';
    deleteBtn.style.display = 'inline-block';
    populateForm(r);
  } else {
    document.getElementById('modalTitle').textContent = 'Nova Visita';
    deleteBtn.style.display = 'none';
    form.reset();
    form.querySelector('[name="data_visita"]').value = currentDate;
  }

  document.getElementById('modal').style.display = 'flex';
  document.body.style.overflow = 'hidden';
  // focus first field
  setTimeout(() => {
    const first = form.querySelector('input:not([type="checkbox"]), select');
    if (first) first.focus();
  }, 50);
}

function closeModal() {
  document.getElementById('modal').style.display = 'none';
  document.body.style.overflow = '';
  editingId = null;
}

function populateForm(r) {
  const form = document.getElementById('form');
  form.reset();

  Object.entries(r).forEach(([key, val]) => {
    const el = form.elements[key];
    if (!el || el.type === 'submit' || el.type === 'button') return;
    if (el.type === 'checkbox') {
      el.checked = val === true || val === 'Sim' || val === 'sim';
    } else {
      el.value = val != null ? val : '';
    }
  });
}

function formToData() {
  const form = document.getElementById('form');
  const data = {};

  // Text / select / textarea fields via FormData
  new FormData(form).forEach((v, k) => {
    data[k] = v.trim() === '' ? null : v.trim();
  });

  // Checkbox (not in FormData when unchecked)
  data.grazi = !!form.elements['grazi'].checked;

  // Parse ints
  if (data.ordem)         data.ordem         = parseInt(data.ordem)         || null;
  if (data.classificacao) data.classificacao = parseInt(data.classificacao) || null;

  data.atualizado_em = new Date().toISOString();
  return data;
}

// ---- Save ----
async function saveRecord(e) {
  e.preventDefault();
  const data   = formToData();
  const btn    = document.getElementById('btnSave');
  btn.textContent = 'Salvando…';
  btn.disabled    = true;

  let error;
  if (editingId) {
    ({ error } = await sb.from('agendamentos').update(data).eq('id', editingId));
  } else {
    ({ error } = await sb.from('agendamentos').insert(data));
  }

  btn.textContent = 'Salvar';
  btn.disabled    = false;

  if (error) {
    alert('Erro ao salvar:\n' + error.message);
    return;
  }

  closeModal();
  await loadDay(currentDate);
}

// ---- Delete ----
async function deleteRecord() {
  if (!editingId) return;
  const r = allRecords.find(x => x.id === editingId);
  const nome = r ? r.razao_social : 'este agendamento';
  if (!confirm(`Excluir "${nome}"? Esta ação não pode ser desfeita.`)) return;

  const { error } = await sb.from('agendamentos').delete().eq('id', editingId);
  if (error) {
    alert('Erro ao excluir:\n' + error.message);
    return;
  }

  closeModal();
  await loadDay(currentDate);
}

// ============================================================
// IMPORT CSV
// ============================================================

let parsedRows = [];

// ---- Open / Close ----
function openImport() {
  importGoStep1();
  document.getElementById('importModal').style.display = 'flex';
  document.body.style.overflow = 'hidden';
}

function closeImport() {
  document.getElementById('importModal').style.display = 'none';
  document.body.style.overflow = '';
  parsedRows = [];
}

function importGoStep1() {
  document.getElementById('importStep1').style.display = 'block';
  document.getElementById('importStep2').style.display = 'none';
  document.getElementById('importStep3').style.display = 'none';
  document.getElementById('csvFile').value = '';
  parsedRows = [];
}

// ---- Drop zone ----
function bindDropZone() {
  const zone = document.getElementById('dropZone');
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  });
}

function onFileChosen(e) {
  const file = e.target.files[0];
  if (file) processFile(file);
}

// ---- File processing ----
function processFile(file) {
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      parsedRows = parseCSV(ev.target.result);
      showPreview(parsedRows);
    } catch (err) {
      alert('Erro ao ler o arquivo:\n' + err.message);
    }
  };
  reader.readAsText(file, 'UTF-8');
}

// ---- CSV Parser ----
function parseCSVLine(line) {
  const cols = [];
  let cur = '';
  let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') { inQ = !inQ; }
    else if (ch === ';' && !inQ) { cols.push(cur.trim()); cur = ''; }
    else { cur += ch; }
  }
  cols.push(cur.trim());
  return cols;
}

function nullify(v) {
  if (!v) return null;
  const t = v.trim();
  return t === '' || t === '-' ? null : t;
}

function parseCSV(text) {
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const rows = [];
  let dateCtx = null;

  for (const raw of lines) {
    const line = raw.trim();
    if (!line) continue;

    // Date header: "AGENDAMENTO DIA 30/03/2026 - SEGUNDA-FEIRA"
    if (/^AGENDAMENTO DIA/i.test(line)) {
      const m = line.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (m) dateCtx = `${m[3]}-${m[2]}-${m[1]}`;
      continue;
    }

    const cols = parseCSVLine(line);

    // Skip column header row
    if (cols[0] && cols[0].toUpperCase() === 'ORDEM') continue;

    // Skip rows without razao_social (col 5) or without date context
    const razao = nullify(cols[5]);
    if (!razao || !dateCtx) continue;

    const grazi = (nullify(cols[13]) || '').toLowerCase() === 'sim';

    rows.push({
      data_visita:              dateCtx,
      ordem:                    parseInt(nullify(cols[0])) || null,
      categoria:                nullify(cols[1]),
      classificacao:            parseInt(nullify(cols[2])) || null,
      cnpj:                     nullify(cols[3]),
      grupo:                    nullify(cols[4]),
      razao_social:             razao,
      unidade:                  nullify(cols[6]),
      servicos_contrato:        nullify(cols[7]),
      data_documentacao:        nullify(cols[8]),
      visitas_feitas:           nullify(cols[9]),
      psicossocial:             nullify(cols[10]),
      status_documentacao:      nullify(cols[11]),
      autor:                    nullify(cols[12]),
      grazi,
      responsavel_documentacao: nullify(cols[14]),
      tipo_contrato:            nullify(cols[15]),
      data_contrato:            nullify(cols[16]),
      planilha_empresa:         nullify(cols[17]),
      comercial_responsavel:    nullify(cols[18]),
      ordem_servico_unisyst:    nullify(cols[19]),
      visita:                   nullify(cols[20]),
      cidade:                   nullify(cols[21]),
      endereco:                 nullify(cols[22]),
      email:                    nullify(cols[23]),
      telefone:                 nullify(cols[24]),
      uber:                     nullify(cols[25]),
      data_pagamento:           nullify(cols[26]),
      atualizado_em:            new Date().toISOString(),
      atualizado_por:           'importação CSV',
    });
  }

  return rows;
}

// ---- Preview ----
function showPreview(rows) {
  if (rows.length === 0) {
    alert('Nenhum registro válido encontrado no arquivo. Verifique se o separador é ponto-e-vírgula (;).');
    return;
  }

  const dias = new Set(rows.map(r => r.data_visita)).size;
  document.getElementById('importSummary').innerHTML =
    `<span>&#128202; <strong>${rows.length}</strong> registros encontrados</span>` +
    `<span>&#128197; <strong>${dias}</strong> dias de agendamento</span>`;

  const tbody = document.querySelector('#importPreview tbody');
  tbody.innerHTML = rows.map(r => `
    <tr>
      <td>${r.data_visita ? r.data_visita.split('-').reverse().join('/') : ''}</td>
      <td>${esc(r.ordem ?? '')}</td>
      <td>${esc(r.categoria ?? '')}</td>
      <td>${esc(r.razao_social)}</td>
      <td>${esc(r.status_documentacao ?? '')}</td>
      <td>${esc(r.cidade ?? '')}</td>
      <td>${esc(r.autor ?? '')}</td>
    </tr>
  `).join('');

  document.getElementById('importStep1').style.display = 'none';
  document.getElementById('importStep2').style.display = 'block';
  document.getElementById('importProgress').style.display = 'none';
  document.getElementById('btnImportRun').disabled = false;
  document.getElementById('btnImportRun').textContent = 'Importar tudo';
}

// ---- Run import ----
async function runImport() {
  if (!parsedRows.length) return;

  const btn = document.getElementById('btnImportRun');
  btn.disabled = true;
  btn.textContent = 'Importando…';

  const progress = document.getElementById('importProgress');
  const fill     = document.getElementById('progressFill');
  const label    = document.getElementById('progressLabel');
  progress.style.display = 'block';

  const BATCH = 50;
  let inserted = 0;
  let errors   = 0;

  for (let i = 0; i < parsedRows.length; i += BATCH) {
    const chunk = parsedRows.slice(i, i + BATCH);
    const { error } = await sb.from('agendamentos').insert(chunk);
    if (error) {
      console.error('Import batch error:', error);
      errors += chunk.length;
    } else {
      inserted += chunk.length;
    }
    const pct = Math.round(((i + chunk.length) / parsedRows.length) * 100);
    fill.style.width  = pct + '%';
    label.textContent = `Importando… ${inserted} de ${parsedRows.length}`;
  }

  document.getElementById('importStep2').style.display = 'none';
  document.getElementById('importStep3').style.display = 'block';
  document.getElementById('importDoneMsg').textContent =
    errors === 0
      ? `${inserted} registros importados com sucesso.`
      : `${inserted} importados · ${errors} com erro (verifique o console).`;

  await loadDay(currentDate);
}