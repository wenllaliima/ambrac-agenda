const SUPABASE_URL      = 'https://xwvzqwjuicmvcwxjgwcp.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inh3dnpxd2p1aWNtdmN3eGpnd2NwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzUyNDEyODYsImV4cCI6MjA5MDgxNzI4Nn0.TaoeuDwu9rlpXLpkTk5VFFz8tqEJlPEXc8e3_cp_SKo';
const sb = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

let currentDate   = todayISO();
let allRecords    = [];
let editingId     = null;
let pendentesMode = false;
let todasMode     = false;
let parsedRows    = [];

const WEEKDAYS = ['domingo','segunda-feira','terça-feira','quarta-feira','quinta-feira','sexta-feira','sábado'];

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('datePicker').value = currentDate;
  updateDayLabel();
  bindEvents();
  loadDay(currentDate);
});

function todayISO() { return new Date().toISOString().split('T')[0]; }

function updateDayLabel() {
  const d = new Date(currentDate + 'T12:00:00');
  document.getElementById('dayLabel').textContent = WEEKDAYS[d.getDay()];
}

function esc(s) {
  if (s == null || s === false) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── EVENTS ──────────────────────────────────────────────────────────────────
function bindEvents() {
  document.getElementById('btnPrev').addEventListener('click', () => shiftDay(-1));
  document.getElementById('btnNext').addEventListener('click', () => shiftDay(1));

  document.getElementById('datePicker').addEventListener('change', e => {
    currentDate = e.target.value;
    updateDayLabel();
    setPendentes(false);
    setTodas(false);
    loadDay(currentDate);
  });

  document.getElementById('btnPendentes').addEventListener('click', () => {
    const on = !pendentesMode;
    setTodas(false);
    setPendentes(on);
    on ? loadPendentes() : loadDay(currentDate);
  });

  document.getElementById('btnTodas').addEventListener('click', () => {
    const on = !todasMode;
    setPendentes(false);
    setTodas(on);
    on ? loadAll() : loadDay(currentDate);
  });

  document.getElementById('btnAdd').addEventListener('click', () => openModal(null));
  document.getElementById('modalClose').addEventListener('click', closeModal);
  document.getElementById('modalOverlay').addEventListener('click', closeModal);
  document.getElementById('btnCancel').addEventListener('click', closeModal);
  document.getElementById('btnDelete').addEventListener('click', deleteRecord);
  document.getElementById('form').addEventListener('submit', saveRecord);
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

  document.getElementById('btnDash').addEventListener('click', () => {
    if (document.getElementById('dashView').style.display !== 'none') { closeDash(); }
    else { openDash(); }
  });
  document.getElementById('btnExport').addEventListener('click', exportXLSX);
  document.getElementById('btnImport').addEventListener('click', openImport);
  document.getElementById('importClose').addEventListener('click', closeImport);
  document.getElementById('importOverlay').addEventListener('click', closeImport);
  document.getElementById('csvFile').addEventListener('change', e => { if (e.target.files[0]) processFile(e.target.files[0]); });
  document.getElementById('btnImportBack').addEventListener('click', importStep1);
  document.getElementById('btnImportRun').addEventListener('click', runImport);
  document.getElementById('btnImportClose').addEventListener('click', closeImport);
  bindDropZone();

  document.getElementById('fillFile').addEventListener('change', e => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const bytes = new Uint8Array(ev.target.result);
        const bom = bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF;
        fillFormFromCSV(new TextDecoder(bom ? 'UTF-8' : 'windows-1252').decode(ev.target.result));
      } catch { showFillMsg('Erro ao ler arquivo', true); }
      e.target.value = '';
    };
    reader.readAsArrayBuffer(file);
  });

  ['filtStatus','filtCategoria','filtAutor','filtSearch'].forEach(id =>
    document.getElementById(id).addEventListener('input', renderCards));
  document.getElementById('btnClearFilters').addEventListener('click', clearFilters);
}

function shiftDay(n) {
  const d = new Date(currentDate + 'T12:00:00');
  d.setDate(d.getDate() + n);
  currentDate = d.toISOString().split('T')[0];
  document.getElementById('datePicker').value = currentDate;
  updateDayLabel();
  setPendentes(false);
  setTodas(false);
  loadDay(currentDate);
}

function setPendentes(on) {
  pendentesMode = on;
  document.getElementById('btnPendentes').classList.toggle('btn-active', on);
  const nav = document.getElementById('dateNav');
  if (nav) nav.style.opacity = on ? '.4' : '';
}

function setTodas(on) {
  todasMode = on;
  document.getElementById('btnTodas').classList.toggle('btn-active', on);
  const nav = document.getElementById('dateNav');
  if (nav) nav.style.opacity = on ? '.4' : '';
}

// ── DATA ────────────────────────────────────────────────────────────────────
async function loadDay(date) {
  showLoading(true);
  const { data, error } = await sb.from('agendamentos').select('*').eq('data_visita', date).order('ordem', { nullsFirst: false });
  showLoading(false);
  if (error) { showEmpty('Erro ao carregar dados.'); return; }
  allRecords = data || [];
  buildAutorFilter();
  renderCards();
}

async function loadPendentes() {
  showLoading(true);
  const { data, error } = await sb.from('agendamentos').select('*').is('data_visita', null).order('razao_social');
  showLoading(false);
  if (error) { showEmpty('Erro ao carregar dados.'); return; }
  allRecords = data || [];
  buildAutorFilter();
  renderCards();
}

async function loadAll() {
  showLoading(true);
  const { data, error } = await sb.from('agendamentos').select('*').order('data_visita', { ascending: true }).order('ordem', { nullsFirst: false });
  showLoading(false);
  if (error) { showEmpty('Erro ao carregar dados.'); return; }
  allRecords = data || [];
  buildAutorFilter();
  renderCards();
}

function showLoading(on) {
  document.getElementById('loading').style.display = on ? 'block' : 'none';
  document.getElementById('cards').style.display   = on ? 'none' : 'grid';
  if (on) document.getElementById('empty').style.display = 'none';
}

function showEmpty(msg) {
  document.getElementById('empty').textContent = msg;
  document.getElementById('empty').style.display = 'block';
}

function buildAutorFilter() {
  const autores = [...new Set(allRecords.map(r => r.autor).filter(Boolean))].sort();
  const sel = document.getElementById('filtAutor');
  const cur = sel.value;
  sel.innerHTML = '<option value="">Todos os autores</option>';
  autores.forEach(a => {
    const o = document.createElement('option');
    o.value = a; o.textContent = a;
    if (a === cur) o.selected = true;
    sel.appendChild(o);
  });
}

function clearFilters() {
  ['filtStatus','filtCategoria','filtAutor'].forEach(id => document.getElementById(id).value = '');
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
      const hay = [r.razao_social, r.cnpj, r.cidade, r.grupo, r.unidade, r.visita, r.autor, r.endereco].join(' ').toLowerCase();
      if (!hay.includes(search)) return false;
    }
    return true;
  });
}

// ── RENDER ───────────────────────────────────────────────────────────────────
function renderCards() {
  const filtered = getFiltered();
  const grid  = document.getElementById('cards');
  const empty = document.getElementById('empty');

  document.getElementById('countBadge').textContent = `${filtered.length} empresa${filtered.length !== 1 ? 's' : ''}`;

  const uberTotal = filtered.reduce((acc, r) => {
    if (!r.uber) return acc;
    const m = r.uber.replace(/\s/g,'').match(/\d+[,.]\d{2}/);
    if (m) { const v = parseFloat(m[0].replace(',','.')); if (!isNaN(v)) acc += v; }
    return acc;
  }, 0);
  document.getElementById('uberTotal').textContent = uberTotal > 0 ? `Uber: R$ ${uberTotal.toFixed(2).replace('.',',')}` : '';

  if (filtered.length === 0) {
    grid.innerHTML = '';
    empty.style.display = 'block';
    empty.textContent = allRecords.length === 0 ? 'Nenhum agendamento para este dia.' : 'Nenhum resultado para os filtros.';
    return;
  }
  empty.style.display = 'none';
  grid.innerHTML = filtered.map(cardHTML).join('');
  grid.querySelectorAll('.card').forEach(el =>
    el.addEventListener('click', () => openModal(el.dataset.id)));
}

function statusClass(s) {
  if (!s) return 's-DEFAULT';
  const u = s.toUpperCase();
  if (u.includes('NOVA'))                                   return 's-NOVA';
  if (u.includes('VENCIDA'))                                return 's-VENCIDA';
  if (u.includes('ENVIADO'))                                return 's-ENVIADO';
  if (u.includes('VÁLIDA') || u.includes('VALIDA'))         return 's-VALIDA';
  if (u.includes('ANDAMENTO'))                              return 's-ANDAMENTO';
  if (u.includes('MIGRA'))                                  return 's-MIGRACAO';
  if (u.includes('NÃO ENVIAR') || u.includes('NAO ENVIAR')) return 's-NAO';
  if (u.includes('DESMARCADO'))                             return 's-DESMARC';
  return 's-DEFAULT';
}

function cardHTML(r) {
  // Topo
  const cClass = !r.tipo_contrato ? '' : r.tipo_contrato.toUpperCase().includes('MENSAL') ? 'badge-mensal' : 'badge-avulso';
  const topLeft = [
    r.ordem         ? `<span class="badge badge-ordem">#${esc(r.ordem)}</span>` : '',
    r.tipo_contrato ? `<span class="badge ${cClass}">${esc(r.tipo_contrato)}</span>` : '',
  ].filter(Boolean).join('');

  const statusBadge = r.status_documentacao
    ? `<span class="badge status-badge ${statusClass(r.status_documentacao)}">${esc(r.status_documentacao)}</span>` : '';

  // Visita
  let vcls = '';
  if (r.visitas_feitas) {
    const u = r.visitas_feitas.toUpperCase();
    if (u === 'OK') vcls = 'row-ok';
    else if (u.includes('REAG')) vcls = 'row-reag';
    else if (u.includes('NÃO') || u.includes('NAO') || u.includes('DESMARC')) vcls = 'row-nao';
  }

  // Psico
  const pcls = !r.psicossocial ? 'row-psic-pendente'
    : r.psicossocial.toUpperCase().includes('ENVIADO') ? 'row-psic-ok' : 'row-psic-outro';

  // Email
  let emailRow = '';
  if (r.email) {
    const emails = r.email.split(/[,;]/).map(e => e.trim()).filter(Boolean);
    const more = emails.length > 1 ? `<span class="card-email-more">+${emails.length - 1}</span>` : '';
    emailRow = `<div class="card-row"><span class="row-icon">&#9993;</span><span class="row-text row-email">${esc(emails[0])}</span>${more}</div>`;
  }

  const graziTag = r.grazi ? '<span class="card-grazi">Grazi</span>' : '';
  const uberTag  = r.uber && r.uber.trim()
    ? `<span class="card-uber">${esc(r.uber.trim())}${r.data_pagamento ? ' · ' + esc(r.data_pagamento) : ''}</span>` : '';

  return `<div class="card" data-id="${esc(r.id)}" tabindex="0" role="button">
    <div class="card-top">
      <div class="card-top-left">${topLeft}</div>
      ${statusBadge}
    </div>
    <div class="card-title">${esc(r.razao_social) || '—'}</div>
    <div class="card-cnpj">${esc(r.cnpj) || '—'}${r.grupo ? ' &middot; ' + esc(r.grupo) : ''}</div>
    <div class="card-rows">
      ${r.visita ? `<div class="card-row"><span class="row-icon">&#128222;</span><span class="row-text">${esc(r.visita)}</span></div>` : ''}
      ${emailRow}
      ${r.visitas_feitas ? `<div class="card-row ${vcls}"><span class="row-icon">&#10003;</span><span class="row-text">${esc(r.visitas_feitas)}</span></div>` : ''}
      <div class="card-row ${pcls}"><span class="row-icon">&#129504;</span><span class="row-text">Psico: ${esc(r.psicossocial || 'pendente')}</span></div>
    </div>
    <div class="card-footer">
      <span class="card-autor">${esc(r.autor) || ''}${graziTag}</span>
      ${uberTag}
    </div>
  </div>`;
}

// ── MODAL ────────────────────────────────────────────────────────────────────
function openModal(id) {
  editingId = id;
  const form = document.getElementById('form');
  const del  = document.getElementById('btnDelete');
  if (id) {
    const r = allRecords.find(x => x.id === id);
    if (!r) return;
    document.getElementById('modalTitle').textContent = r.razao_social || 'Editar';
    del.style.display = 'inline-block';
    populateForm(r);
  } else {
    document.getElementById('modalTitle').textContent = 'Nova Empresa';
    del.style.display = 'none';
    form.reset();
    if (!pendentesMode) form.querySelector('[name="data_visita"]').value = currentDate;
  }
  document.getElementById('modal').style.display = 'flex';
  document.body.style.overflow = 'hidden';
  setTimeout(() => { const f = form.querySelector('input:not([type="checkbox"]),select'); if (f) f.focus(); }, 50);
}

function closeModal() {
  document.getElementById('modal').style.display = 'none';
  document.body.style.overflow = '';
  editingId = null;
}

function populateForm(r) {
  const form = document.getElementById('form');
  form.reset();
  Object.entries(r).forEach(([k, v]) => {
    const el = form.elements[k];
    if (!el || el.type === 'submit' || el.type === 'button') return;
    if (el.type === 'checkbox') el.checked = v === true || v === 'Sim' || v === 'sim';
    else el.value = v != null ? v : '';
  });
}

function formToData() {
  const form = document.getElementById('form');
  const data = {};
  new FormData(form).forEach((v, k) => { data[k] = v.trim() === '' ? null : v.trim(); });
  data.grazi = !!form.elements['grazi'].checked;
  if (data.ordem)         data.ordem         = parseInt(data.ordem)         || null;
  if (data.classificacao) data.classificacao = parseInt(data.classificacao) || null;
  data.atualizado_em = new Date().toISOString();
  return data;
}

async function saveRecord(e) {
  e.preventDefault();
  const data = formToData();
  const btn  = document.getElementById('btnSave');
  btn.textContent = 'Salvando…'; btn.disabled = true;
  let error;
  if (editingId) ({ error } = await sb.from('agendamentos').update(data).eq('id', editingId));
  else           ({ error } = await sb.from('agendamentos').insert(data));
  btn.textContent = 'Salvar'; btn.disabled = false;
  if (error) { alert('Erro ao salvar:\n' + error.message); return; }
  closeModal();
  pendentesMode ? loadPendentes() : loadDay(currentDate);
}

async function deleteRecord() {
  if (!editingId) return;
  const r = allRecords.find(x => x.id === editingId);
  if (!confirm(`Arquivar "${r ? r.razao_social : 'este registro'}"?`)) return;
  const { error } = await sb.from('agendamentos').delete().eq('id', editingId);
  if (error) { alert('Erro:\n' + error.message); return; }
  closeModal();
  pendentesMode ? loadPendentes() : loadDay(currentDate);
}

// ── IMPORT ───────────────────────────────────────────────────────────────────
function openImport()  { importStep1(); document.getElementById('importModal').style.display = 'flex'; document.body.style.overflow = 'hidden'; }
function closeImport() { document.getElementById('importModal').style.display = 'none'; document.body.style.overflow = ''; parsedRows = []; }

function importStep1() {
  document.getElementById('importStep1').style.display = 'block';
  document.getElementById('importStep2').style.display = 'none';
  document.getElementById('importStep3').style.display = 'none';
  document.getElementById('csvFile').value = '';
  parsedRows = [];
}

function bindDropZone() {
  const z = document.getElementById('dropZone');
  z.addEventListener('dragover', e => { e.preventDefault(); z.classList.add('drag-over'); });
  z.addEventListener('dragleave', () => z.classList.remove('drag-over'));
  z.addEventListener('drop', e => { e.preventDefault(); z.classList.remove('drag-over'); if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]); });
}

function processFile(file) {
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const bytes = new Uint8Array(ev.target.result);
      const bom = bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF;
      parsedRows = parseCSV(new TextDecoder(bom ? 'UTF-8' : 'windows-1252').decode(ev.target.result));
      showPreview(parsedRows);
    } catch (err) { alert('Erro ao ler arquivo:\n' + err.message); }
  };
  reader.readAsArrayBuffer(file);
}

function parseCSVLine(line) {
  const cols = []; let cur = '', inQ = false;
  for (const ch of line) {
    if (ch === '"') inQ = !inQ;
    else if (ch === ';' && !inQ) { cols.push(cur.trim()); cur = ''; }
    else cur += ch;
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
  const rows = []; let dateCtx = null;
  for (const raw of text.replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n')) {
    const line = raw.trim();
    if (!line) continue;
    if (/^AGENDAMENTO DIA/i.test(line)) {
      const m = line.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (m) dateCtx = `${m[3]}-${m[2]}-${m[1]}`;
      continue;
    }
    const cols = parseCSVLine(line);
    if (cols[0] && cols[0].toUpperCase() === 'ORDEM') continue;
    const razao = nullify(cols[5]);
    if (!razao || !dateCtx) continue;
    // Colunas da planilha ROTA:
    // 0:ORDEM 1:CATEGORIA 2:CLASSIFICAÇÃO 3:CNPJ 4:GRUPO 5:RAZÃO SOCIAL
    // 6:UNIDADE 7:SERVIÇOS 8:DATA DOC 9:VISITAS FEITAS 10:PSICOSSOCIAL
    // 11:STATUS 12:VISITA(contato) 13:CIDADE 14:ENDEREÇO 15:E-MAIL 16:TELEFONE 17:UBER
    rows.push({
      data_visita:         dateCtx,
      ordem:               parseInt(nullify(cols[0]))  || null,
      categoria:           nullify(cols[1]),
      classificacao:       parseInt(nullify(cols[2]))  || null,
      cnpj:                nullify(cols[3]),
      grupo:               nullify(cols[4]),
      razao_social:        razao,
      unidade:             nullify(cols[6]),
      servicos_contrato:   nullify(cols[7]),
      data_documentacao:   nullify(cols[8]),
      visitas_feitas:      nullify(cols[9]),
      psicossocial:        nullify(cols[10]),
      status_documentacao: nullify(cols[11]),
      visita:              nullify(cols[12]),
      cidade:              nullify(cols[13]),
      endereco:            nullify(cols[14]),
      email:               nullify(cols[15]),
      telefone:            nullify(cols[16]),
      uber:                nullify(cols[17]),
      grazi:               false,
      atualizado_em:       new Date().toISOString(),
      atualizado_por:      'importação CSV',
    });
  }
  return rows;
}

function showPreview(rows) {
  if (!rows.length) { alert('Nenhum registro válido. Verifique se o separador é ponto-e-vírgula (;).'); return; }
  const dias = new Set(rows.map(r => r.data_visita)).size;
  document.getElementById('importSummary').innerHTML =
    `<span>&#128202; <strong>${rows.length}</strong> registros</span><span>&#128197; <strong>${dias}</strong> dias</span>`;
  document.querySelector('#importPreview tbody').innerHTML = rows.map(r => `
    <tr>
      <td>${r.data_visita ? r.data_visita.split('-').reverse().join('/') : ''}</td>
      <td>${esc(r.ordem ?? '')}</td><td>${esc(r.categoria ?? '')}</td>
      <td>${esc(r.razao_social)}</td><td>${esc(r.status_documentacao ?? '')}</td>
      <td>${esc(r.cidade ?? '')}</td><td>${esc(r.autor ?? '')}</td>
    </tr>`).join('');
  document.getElementById('importStep1').style.display = 'none';
  document.getElementById('importStep2').style.display = 'block';
  document.getElementById('importProgress').style.display = 'none';
  document.getElementById('btnImportRun').disabled = false;
  document.getElementById('btnImportRun').textContent = 'Importar tudo';
}

async function runImport() {
  if (!parsedRows.length) return;
  const btn   = document.getElementById('btnImportRun');
  const fill  = document.getElementById('progressFill');
  const label = document.getElementById('progressLabel');
  btn.disabled = true; btn.textContent = 'Importando…';
  document.getElementById('importProgress').style.display = 'block';

  // 1. Apagar registros existentes para cada data do arquivo
  const dates = [...new Set(parsedRows.map(r => r.data_visita).filter(Boolean))];
  label.textContent = `Limpando ${dates.length} dia(s)…`;
  for (const date of dates) {
    const { error } = await sb.from('agendamentos').delete().eq('data_visita', date);
    if (error) { alert(`Erro ao limpar ${date}:\n` + error.message); btn.disabled = false; btn.textContent = 'Importar tudo'; return; }
  }

  // 2. Inserir novos registros em lotes
  const BATCH = 50; let inserted = 0, errors = 0;
  for (let i = 0; i < parsedRows.length; i += BATCH) {
    const chunk = parsedRows.slice(i, i + BATCH);
    const { error } = await sb.from('agendamentos').insert(chunk);
    if (error) { alert('Erro ao inserir:\n' + error.message); errors += chunk.length; }
    else inserted += chunk.length;
    fill.style.width  = Math.round(((i + chunk.length) / parsedRows.length) * 100) + '%';
    label.textContent = `Importando… ${inserted} de ${parsedRows.length}`;
  }

  document.getElementById('importStep2').style.display = 'none';
  document.getElementById('importStep3').style.display = 'block';
  document.getElementById('importDoneMsg').textContent = errors === 0
    ? `${inserted} registros importados com sucesso.`
    : `${inserted} importados · ${errors} com erro.`;
  loadDay(currentDate);
}

// ── DASHBOARD ────────────────────────────────────────────────────────────────
let _dashCharts = [];

async function openDash() {
  const view = document.getElementById('dashView');
  view.style.display = 'block';
  document.getElementById('main').style.display = 'none';
  document.getElementById('btnDash').classList.add('btn-active');
  view.innerHTML = '<div class="state-msg">Carregando dados…</div>';
  const { data, error } = await sb.from('agendamentos').select('*');
  if (error) { view.innerHTML = '<div class="state-msg">Erro ao carregar dados.</div>'; return; }
  renderDash(data || []);
}

function closeDash() {
  document.getElementById('dashView').style.display = 'none';
  document.getElementById('main').style.display = 'block';
  document.getElementById('btnDash').classList.remove('btn-active');
  _dashCharts.forEach(c => c.destroy());
  _dashCharts = [];
}

function countBy(arr, key) {
  return arr.reduce((acc, r) => {
    const v = r[key] || '(sem dado)';
    acc[v] = (acc[v] || 0) + 1;
    return acc;
  }, {});
}

function renderDash(d) {
  const totalEmpresas = d.length;

  // Uber
  const uberTotal = d.reduce((acc, r) => {
    if (!r.uber) return acc;
    const m = r.uber.replace(/\s/g,'').match(/[\d]+[,.][\d]{2}/);
    if (m) { const v = parseFloat(m[0].replace(',','.')); if (!isNaN(v)) acc += v; }
    return acc;
  }, 0);
  const comUber = d.filter(r => r.uber && r.uber.trim()).length;
  const semUber = totalEmpresas - comUber;

  // Visitas
  const visitasOK   = d.filter(r => r.visitas_feitas && r.visitas_feitas.toUpperCase() === 'OK').length;
  const visitasReag = d.filter(r => r.visitas_feitas && r.visitas_feitas.toUpperCase().includes('REAG')).length;
  const visitasNao  = d.filter(r => r.visitas_feitas && (r.visitas_feitas.toUpperCase().includes('NÃO') || r.visitas_feitas.toUpperCase().includes('NAO'))).length;
  const visitasSem  = totalEmpresas - visitasOK - visitasReag - visitasNao;

  // Documentação
  const docsOK     = d.filter(r => r.status_documentacao && ['ENVIADO','VÁLIDA'].includes(r.status_documentacao)).length;
  const docsVencid = d.filter(r => r.status_documentacao === 'VENCIDA').length;
  const docsNova   = d.filter(r => r.status_documentacao === 'NOVA EMPRESA').length;

  // Psicossocial
  const psicEnviado  = d.filter(r => r.psicossocial && r.psicossocial.toUpperCase().includes('ENVIADO')).length;
  const psicPendente = d.filter(r => !r.psicossocial || r.psicossocial.trim() === '').length;
  const psicOutro    = totalEmpresas - psicEnviado - psicPendente;

  const byStatus   = countBy(d, 'status_documentacao');
  const byCateg    = countBy(d, 'categoria');
  const byContrato = countBy(d, 'tipo_contrato');
  const byCidade   = countBy(d, 'cidade');
  const topCidades = Object.entries(byCidade).filter(e => e[0] !== '(sem dado)').sort((a,b)=>b[1]-a[1]).slice(0,10);

  const uberByDate = {};
  d.forEach(r => {
    if (!r.data_visita || !r.uber) return;
    const m = r.uber.replace(/\s/g,'').match(/[\d]+[,.][\d]{2}/);
    if (!m) return;
    const v = parseFloat(m[0].replace(',','.'));
    if (!isNaN(v)) uberByDate[r.data_visita] = (uberByDate[r.data_visita] || 0) + v;
  });
  const uberDates = Object.keys(uberByDate).sort();

  function pct(n) { return totalEmpresas ? Math.round(n / totalEmpresas * 100) : 0; }

  const view = document.getElementById('dashView');
  view.innerHTML = `
    <div class="dash-header">
      <h2 class="dash-title">Dashboard</h2>
      <button id="btnCloseDash" class="btn-ghost">&#8592; Voltar</button>
    </div>

    <div class="dash-kpis">
      <div class="dash-kpi">
        <div class="kpi-label">Total de Empresas</div>
        <div class="kpi-value">${totalEmpresas}</div>
      </div>
      <div class="dash-kpi">
        <div class="kpi-label">Total Uber</div>
        <div class="kpi-value kpi-green">R$&nbsp;${uberTotal.toFixed(2).replace('.',',')}</div>
        <div class="kpi-sub">${comUber} empresas · ${semUber} sem valor</div>
      </div>
      <div class="dash-kpi">
        <div class="kpi-label">Visitas Realizadas (OK)</div>
        <div class="kpi-value kpi-blue">${visitasOK}</div>
        <div class="kpi-sub">${pct(visitasOK)}% · Reagendado: ${visitasReag} · Não feita: ${visitasNao}</div>
      </div>
      <div class="dash-kpi">
        <div class="kpi-label">Docs Enviadas / Válidas</div>
        <div class="kpi-value kpi-yellow">${docsOK}</div>
        <div class="kpi-sub">${pct(docsOK)}% · Vencida: ${docsVencid} · Nova: ${docsNova}</div>
      </div>
      <div class="dash-kpi">
        <div class="kpi-label">Psicossocial Enviado</div>
        <div class="kpi-value kpi-green">${psicEnviado}</div>
        <div class="kpi-sub">${pct(psicEnviado)}% · Pendente: ${psicPendente} · Outro: ${psicOutro}</div>
      </div>
    </div>

    <div class="dash-charts">
      <div class="dash-chart-box">
        <div class="chart-title">Status da Documentação</div>
        <canvas id="cStatus"></canvas>
      </div>
      <div class="dash-chart-box">
        <div class="chart-title">Visitas Realizadas</div>
        <canvas id="cVisitas"></canvas>
      </div>
      <div class="dash-chart-box">
        <div class="chart-title">Psicossocial</div>
        <canvas id="cPsico"></canvas>
      </div>
      <div class="dash-chart-box">
        <div class="chart-title">Categoria</div>
        <canvas id="cCateg"></canvas>
      </div>
      <div class="dash-chart-box">
        <div class="chart-title">Tipo de Contrato</div>
        <canvas id="cContrato"></canvas>
      </div>
      <div class="dash-chart-box">
        <div class="chart-title">Top Cidades</div>
        <canvas id="cCidades"></canvas>
      </div>
      <div class="dash-chart-box dash-chart-wide">
        <div class="chart-title">Uber por Data de Visita (R$)</div>
        <canvas id="cUber"></canvas>
      </div>
    </div>
  `;

  document.getElementById('btnCloseDash').addEventListener('click', closeDash);

  const PALETTE = ['#6366f1','#86efac','#fcd34d','#fca5a5','#93c5fd','#d8b4fe','#fb923c','#34d399','#f472b6','#a3e635'];
  const tickColor = '#8892b0';
  const gridColor = '#2d3150';

  const datalabels = {
    color: '#fff',
    font: { size: 11, weight: 'bold' },
    formatter: (val, ctx) => {
      const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
      if (!total || val === 0) return '';
      return val + '\n' + Math.round(val / total * 100) + '%';
    }
  };

  function donut(id, entries) {
    return new Chart(document.getElementById(id), {
      type: 'doughnut',
      data: {
        labels: entries.map(e => e[0]),
        datasets: [{ data: entries.map(e => e[1]), backgroundColor: PALETTE, borderWidth: 0 }]
      },
      options: {
        responsive: true,
        cutout: '55%',
        plugins: {
          legend: { labels: { color: tickColor, font: { size: 11 }, boxWidth: 12 } },
          tooltip: {
            callbacks: {
              label: ctx => {
                const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
                return ` ${ctx.label}: ${ctx.raw} (${Math.round(ctx.raw/total*100)}%)`;
              }
            }
          }
        }
      }
    });
  }

  _dashCharts.push(donut('cStatus', Object.entries(byStatus).sort((a,b)=>b[1]-a[1])));

  _dashCharts.push(new Chart(document.getElementById('cVisitas'), {
    type: 'doughnut',
    data: {
      labels: ['OK', 'Reagendado', 'Não feita', 'Sem info'],
      datasets: [{ data: [visitasOK, visitasReag, visitasNao, visitasSem],
        backgroundColor: ['#86efac','#fcd34d','#fca5a5','#374151'], borderWidth: 0 }]
    },
    options: {
      responsive: true, cutout: '55%',
      plugins: {
        legend: { labels: { color: tickColor, font: { size: 11 }, boxWidth: 12 } },
        tooltip: { callbacks: { label: ctx => {
          const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
          return ` ${ctx.label}: ${ctx.raw} (${Math.round(ctx.raw/total*100)}%)`;
        }}}
      }
    }
  }));

  _dashCharts.push(new Chart(document.getElementById('cPsico'), {
    type: 'doughnut',
    data: {
      labels: ['Enviado', 'Pendente', 'Outro'],
      datasets: [{ data: [psicEnviado, psicPendente, psicOutro],
        backgroundColor: ['#86efac','#fcd34d','#93c5fd'], borderWidth: 0 }]
    },
    options: {
      responsive: true, cutout: '55%',
      plugins: {
        legend: { labels: { color: tickColor, font: { size: 11 }, boxWidth: 12 } },
        tooltip: { callbacks: { label: ctx => {
          const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
          return ` ${ctx.label}: ${ctx.raw} (${Math.round(ctx.raw/total*100)}%)`;
        }}}
      }
    }
  }));

  _dashCharts.push(donut('cCateg',    Object.entries(byCateg).sort((a,b)=>b[1]-a[1])));
  _dashCharts.push(donut('cContrato', Object.entries(byContrato).sort((a,b)=>b[1]-a[1])));

  _dashCharts.push(new Chart(document.getElementById('cCidades'), {
    type: 'bar',
    data: {
      labels: topCidades.map(e => e[0]),
      datasets: [{ data: topCidades.map(e => e[1]), backgroundColor: '#6366f1', borderRadius: 5 }]
    },
    options: {
      responsive: true, indexAxis: 'y',
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ` ${ctx.raw} empresa${ctx.raw !== 1 ? 's' : ''}` } }
      },
      scales: {
        x: { ticks: { color: tickColor, precision: 0 }, grid: { color: gridColor } },
        y: { ticks: { color: '#e2e8f0', font: { size: 11 } }, grid: { display: false } }
      }
    }
  }));

  _dashCharts.push(new Chart(document.getElementById('cUber'), {
    type: 'bar',
    data: {
      labels: uberDates.map(dt => dt.split('-').reverse().join('/')),
      datasets: [{ label: 'Uber (R$)', data: uberDates.map(dt => uberByDate[dt]), backgroundColor: '#86efac', borderRadius: 4 }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ` R$ ${ctx.raw.toFixed(2).replace('.',',')}` } }
      },
      scales: {
        x: { ticks: { color: tickColor, font: { size: 10 } }, grid: { display: false } },
        y: { ticks: { color: tickColor, callback: v => 'R$ ' + v.toFixed(0) }, grid: { color: gridColor } }
      }
    }
  }));
}

// ── EXPORTAR XLSX ─────────────────────────────────────────────────────────────
async function exportXLSX() {
  const btn = document.getElementById('btnExport');
  btn.textContent = 'Exportando…'; btn.disabled = true;
  const { data, error } = await sb.from('agendamentos').select('*')
    .order('data_visita', { ascending: true })
    .order('ordem', { nullsFirst: false });
  btn.textContent = '↓ Exportar'; btn.disabled = false;
  if (error || !data) { alert('Erro ao exportar:\n' + (error ? error.message : 'sem dados')); return; }

  const COLS = [
    ['data_visita','Data Visita'], ['ordem','Ordem'], ['categoria','Categoria'],
    ['classificacao','Classificação'], ['cnpj','CNPJ'], ['grupo','Grupo'],
    ['razao_social','Razão Social'], ['unidade','Unidade'],
    ['tipo_contrato','Tipo Contrato'], ['data_contrato','Data Contrato'],
    ['servicos_contrato','Serviços Contrato'], ['planilha_empresa','Planilha Empresa'],
    ['comercial_responsavel','Comercial Responsável'], ['data_documentacao','Data Documentação'],
    ['visitas_feitas','Visitas Feitas'], ['psicossocial','Psicossocial'],
    ['status_documentacao','Status Documentação'], ['responsavel_documentacao','Resp. Documentação'],
    ['visita','Informações Visita'], ['cidade','Cidade'], ['endereco','Endereço'],
    ['email','E-mail'], ['telefone','Telefone'], ['uber','Uber'],
    ['data_pagamento','Data Pagamento Uber'], ['grazi','Grazi'],
    ['autor','Autor'], ['atualizado_por','Atualizado Por'], ['atualizado_em','Atualizado Em'],
  ];

  const headers = COLS.map(c => c[1]);
  const rows = data.map(r => COLS.map(([k]) => {
    if (k === 'grazi') return r[k] ? 'Sim' : 'Não';
    if (k === 'atualizado_em' && r[k]) return r[k].replace('T',' ').slice(0,19);
    return r[k] ?? '';
  }));

  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols'] = COLS.map(() => ({ wch: 22 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Agendamentos');
  XLSX.writeFile(wb, `ambrac_agenda_${new Date().toISOString().split('T')[0]}.xlsx`);
}

// ── FILL FORM FROM CSV ───────────────────────────────────────────────────────
const FILL_MAP = {
  'ORDEM':'ordem','CATEGORIA':'categoria','CLASSIFICAÇÃO':'classificacao','CNPJ':'cnpj',
  'GRUPO':'grupo','RAZÃO SOCIAL':'razao_social','UNIDADE':'unidade',
  'SERVIÇOS EM CONTRATO':'servicos_contrato','DATA DA DOCUMENTAÇÃO':'data_documentacao',
  'PSICOSSOCIAL':'psicossocial','STATUS DA DOCUMENTAÇÃO':'status_documentacao',
  'AUTOR':'autor','GRAZI':'grazi','RESPONSÁVEL PELA DOCUMENTAÇÃO':'responsavel_documentacao',
  'TIPO DE CONTRATO':'tipo_contrato','DATA DO CONTRATO':'data_contrato',
  'PLANILHA DA EMPRESA':'planilha_empresa','VISITAS':'visitas_feitas',
  'COMERCIAL RESPONSÁVEL':'comercial_responsavel','O.S ABERTA UNISYST':'ordem_servico_unisyst',
  'CIDADE':'cidade','ENDEREÇO':'endereco','TELEFONE':'telefone','E-MAIL':'email',
};

function fillFormFromCSV(text) {
  const lines = text.replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n').filter(l => l.trim());
  if (lines.length < 2) { showFillMsg('Arquivo precisa ter cabeçalho e dados', true); return; }
  const headers = parseCSVLine(lines[0]).map(h => h.trim().toUpperCase());
  const values  = parseCSVLine(lines[1]);
  const form = document.getElementById('form');
  let filled = 0;
  headers.forEach((h, i) => {
    const field = FILL_MAP[h]; if (!field) return;
    const val = (values[i] || '').trim(); if (!val || val === '-') return;
    const el = form.elements[field]; if (!el) return;
    if (el.type === 'checkbox') el.checked = val.toLowerCase() === 'sim';
    else el.value = val;
    filled++;
  });
  filled === 0 ? showFillMsg('Nenhum campo reconhecido.', true) : showFillMsg(`${filled} campos preenchidos`, false);
}

function showFillMsg(msg, isErr) {
  const el = document.getElementById('fillFeedback');
  el.textContent = msg; el.className = 'fill-feedback' + (isErr ? ' err' : '');
  setTimeout(() => { el.textContent = ''; el.className = 'fill-feedback'; }, 4000);
}
