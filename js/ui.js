// ui.js — render tampilan, tabel, form, preview, navbar auto-hide
import { getAll, upsert, byId, removeById, normalizeFormData, getConfig, setConfig, getSuggestions } from './storage.js';
import { apiInit, apiHeaders, apiPullOverwrite, apiPushUpsert } from './api.js';
import { exportDOCX, exportXLSX, exportPDF } from './export.js';

// ======= Helper =======
function $(sel){ return document.querySelector(sel); }
function h(html){ const d=document.createElement('div'); d.innerHTML = html; return d.firstElementChild; }
function esc(s){ return (s==null?'':String(s)).replace(/[&<>"']/g, m=>({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[m])); }

export function initUI(){
  // Navbar link switching
  document.querySelectorAll('.nav-auto-hide').forEach(a=>{
    a.addEventListener('click', (e)=>{
      e.preventDefault();
      const target = a.getAttribute('data-target');
      showView(target);
      // auto-collapse on mobile
      const nav = $('#navMain');
      if(nav.classList.contains('show')){
        new bootstrap.Collapse(nav, {toggle:true});
      }
      document.querySelectorAll('.nav-auto-hide').forEach(x=>x.classList.remove('active'));
      a.classList.add('active');
    });
  });

  // Buttons (global)
  $('#btnExportDocx').addEventListener('click', async ()=>tryRun(()=>exportDOCX(currentId())));
  $('#btnExportXlsx').addEventListener('click', async ()=>tryRun(exportXLSX));
  $('#btnExportPdf').addEventListener('click', async ()=>tryRun(exportPDF));

  // Dashboard sync buttons
  $('#btnPush').addEventListener('click', ()=>doPush('#syncLog'));
  $('#btnPull').addEventListener('click', ()=>doPull('#syncLog'));
  $('#btnPush2').addEventListener('click', ()=>doPush('#syncLog2'));
  $('#btnPull2').addEventListener('click', ()=>doPull('#syncLog2'));
  $('#btnHeaders').addEventListener('click', ()=>doHeaders('#syncLog2'));

  // Settings
  const cfg = getConfig();
  $('#spreadsheetId').value = cfg.spreadsheetId||'';
  $('#sheetName').value = cfg.sheetName;
  $('#templateUrl').value = cfg.templateUrl;

  $('#spreadsheetId').addEventListener('change', e=>setConfig({ spreadsheetId: e.target.value.trim() }));
  $('#sheetName').addEventListener('change', e=>setConfig({ sheetName: e.target.value.trim()||'SILABUS' }));
  $('#templateUrl').addEventListener('change', e=>setConfig({ templateUrl: e.target.value.trim() }));

  $('#btnInit').addEventListener('click', async ()=>tryRun(async ()=>{
    const r = await apiInit();
    log('#syncLog2', 'INIT', r);
    $('#spreadsheetId').value = (r?.spreadsheetId)||'';
  }));

  $('#btnCopyHeaders').addEventListener('click', ()=>doHeaders('#syncLog2'));
  $('#btnResetLocal').addEventListener('click', ()=>{
    if(confirm('Hapus semua data lokal?')){ localStorage.clear(); location.reload(); }
  });

  // List & search
  $('#q').addEventListener('input', renderList);

  // Editor form buttons
  $('#btnSave').addEventListener('click', ()=>saveForm());
  $('#btnNew').addEventListener('click', ()=>loadForm(null));
  $('#btnDelete').addEventListener('click', ()=>delCurrent());

  // First render
showView('#viewDashboard');
bootstrapAndPullOnLoad();
}

// ======= Views =======
function showView(sel){
  document.querySelectorAll('.view').forEach(v=>v.classList.add('d-none'));
  document.querySelector(sel).classList.remove('d-none');
}

// ======= Logging =======
function log(where, title, obj){
  const el = $(where);
  const now = dayjs().format('YYYY-MM-DD HH:mm:ss');
  el.textContent = `[${now}] ${title} :: ${JSON.stringify(obj, null, 2)}\n` + el.textContent;
}
async function tryRun(fn){
  try{ await fn(); }catch(e){ alert(e.message||e); console.error(e); }
}

// ==== Refresh datalist & textarea suggest ====
function refreshSuggestionsUI(){
  const bag = getSuggestions();
  // isi datalist (Rumpun, Peserta)
  const dlR = document.getElementById('dlRumpun');
  const dlP = document.getElementById('dlPeserta');
  if(dlR) dlR.innerHTML = bag.rumpun.map(v=>`<option value="${esc(v)}">`).join('');
  if(dlP) dlP.innerHTML = bag.peserta.map(v=>`<option value="${esc(v)}">`).join('');

  // pasang suggestor untuk textarea (dengan sumber data dinamis)
  attachTextareaSuggest(document.getElementById('tujuan'),   ()=>getSuggestions().tujuan);
  attachTextareaSuggest(document.getElementById('materi'),   ()=>getSuggestions().materi);
  attachTextareaSuggest(document.getElementById('trainer'),  ()=>getSuggestions().trainer);
  attachTextareaSuggest(document.getElementById('evaluasi'), ()=>getSuggestions().evaluasi);
}

// ===== Dropdown suggest untuk TEXTAREA per-baris =====
let _suggMenu, _suggActive = -1, _suggList = [], _suggTarget;

function ensureSuggMenu(){
  if(_suggMenu) return _suggMenu;
  _suggMenu = document.createElement('div');
  _suggMenu.className = 'sugg-menu d-none';
  document.body.appendChild(_suggMenu);
  return _suggMenu;
}
function hideMenu(){ if(_suggMenu) _suggMenu.classList.add('d-none'); _suggActive = -1; _suggList = []; _suggTarget=null; }
function moveActive(d){
  if(!_suggList.length) return;
  _suggActive = (_suggActive + d + _suggList.length) % _suggList.length;
  [..._suggMenu.querySelectorAll('.sugg-item')].forEach((el,i)=> el.classList.toggle('active', i===_suggActive));
}
function replaceCurrentLine(textarea, replacement){
  const val = textarea.value;
  const caret = textarea.selectionStart;
  const lnStart = val.lastIndexOf('\n', caret-1) + 1;
  const lnEnd = val.indexOf('\n', caret);
  const s = Math.max(0, lnStart);
  const e = (lnEnd === -1 ? val.length : lnEnd);
  const before = val.slice(0, s);
  const after  = val.slice(e);
  const ins = replacement;
  textarea.value = before + ins + after;
  const pos = (before + ins).length;
  textarea.setSelectionRange(pos, pos);
  textarea.dispatchEvent(new Event('input', {bubbles:true}));
}
function showMenuFor(textarea, items){
  const m = ensureSuggMenu();
  m.innerHTML = '';
  _suggList = items.slice(0, 8);
  _suggActive = -1;
  _suggTarget = textarea;

  _suggList.forEach((txt, i)=>{
    const item = document.createElement('div');
    item.className = 'sugg-item';
    item.textContent = txt;
    item.dataset.val = txt;
    item.addEventListener('mousedown', (ev)=>{ // mousedown supaya tidak hilang karena blur
      ev.preventDefault();
      replaceCurrentLine(textarea, txt);
      hideMenu();
    });
    m.appendChild(item);
  });

  const rect = textarea.getBoundingClientRect();
  m.style.left = (rect.left + window.scrollX) + 'px';
  m.style.top  = (rect.bottom + window.scrollY) + 'px';
  m.style.width = rect.width + 'px';
  m.classList.toggle('d-none', _suggList.length === 0);
}
function attachTextareaSuggest(textarea, sourceFn){
  if(!textarea || textarea.dataset.suggBound === '1') return;
  textarea.dataset.suggBound = '1';

  textarea.addEventListener('input', ()=>{
    const src = (typeof sourceFn==='function' ? sourceFn() : []);
    const val = textarea.value;
    const caret = textarea.selectionStart;
    const lnStart = val.lastIndexOf('\n', caret-1) + 1;
    const typed = val.slice(lnStart, caret).trim().toLowerCase();
    if(!typed){ hideMenu(); return; }
    const matches = src.filter(s=> s.toLowerCase().includes(typed));
    if(matches.length) showMenuFor(textarea, matches);
    else hideMenu();
  });

  textarea.addEventListener('keydown', (e)=>{
    if(!_suggMenu || _suggMenu.classList.contains('d-none')) return;
    if(e.key === 'ArrowDown'){ e.preventDefault(); moveActive(1); }
    else if(e.key === 'ArrowUp'){ e.preventDefault(); moveActive(-1); }
    else if(e.key === 'Enter'){ 
      if(_suggActive >= 0){ e.preventDefault(); const txt = _suggList[_suggActive]; replaceCurrentLine(textarea, txt); hideMenu(); }
    }
    else if(e.key === 'Escape'){ hideMenu(); }
  });

  textarea.addEventListener('blur', ()=> setTimeout(hideMenu, 150));
}

async function bootstrapAndPullOnLoad(){
  try {
    // Jika belum ada spreadsheetId di config → init (akan mengembalikan ID master yang tersimpan di GAS)
    let cfg = getConfig();
    if(!cfg.spreadsheetId){
      const r = await apiInit(); // { spreadsheetId, spreadsheetUrl, ... }
      if(r?.spreadsheetId){
        setConfig({ spreadsheetId: r.spreadsheetId });
        log('#syncLog', 'BOOTSTRAP (init)', { spreadsheetId: r.spreadsheetId });
      }
    }
    // Selalu tarik data terbaru dari Sheet → overwrite localStorage
    const res = await apiPullOverwrite();
    log('#syncLog', 'AUTO PULL on load', { count: res.count });
  } catch (e){
    console.error(e);
    // tidak blok UI; user masih bisa pakai mode lokal
  } finally {
    // Render ulang setelah bootstrap/pull (agar daftar & editor langsung terisi)
    renderList(); renderStats();
    refreshSuggestionsUI();
    loadForm(getAll()[0]?.id || null);
  }
}


// ======= Stats, List, Editor =======
function renderStats(){
  const all = getAll();
  $('#statCount').textContent = all.length;
  const upd = all.map(r=>r.updated_at).sort().pop();
  $('#statUpdated').textContent = upd ? dayjs(upd).format('YYYY-MM-DD HH:mm') : '-';
}

function renderList(){
  const q = $('#q').value.toLowerCase().trim();
  const tbody = $('#tblList tbody');
  tbody.innerHTML = '';

  const data = getAll().filter(r=>{
    const hay = `${r.kode||''} ${r.judul||''} ${r.rumpun_pekerjaan||''} ${r.tipe_training||''}`.toLowerCase();
    return !q || hay.includes(q);
  });

  for (const r of data){
    const tr = document.createElement('tr');

    // ===== Aksi
    const tdAct = document.createElement('td');
    tdAct.className = 'text-nowrap';

    const btnEdit = document.createElement('button');
    btnEdit.className = 'btn btn-sm btn-outline-primary me-1';
    btnEdit.title = 'Edit';
    btnEdit.innerHTML = '<i class="bi bi-pencil"></i>';
    btnEdit.addEventListener('click', ()=>{
      loadForm(r.id);               // set form + preview
      showView('#viewEditor');      // pastikan pindah ke editor
    });

    const btnDup = document.createElement('button');
    btnDup.className = 'btn btn-sm btn-outline-secondary me-1';
    btnDup.title = 'Duplikasi';
    btnDup.innerHTML = '<i class="bi bi-files"></i>';
    btnDup.addEventListener('click', ()=>{
      const copy = (window.structuredClone ? structuredClone(r) : JSON.parse(JSON.stringify(r)));
      copy.id = `${r.id}-copy-${dayjs().format('HHmmss')}`;
      copy.version = 1;
      copy.created_at = dayjs().toISOString();
      copy.updated_at = copy.created_at;
      upsert(copy);
      renderList(); renderStats();
    });

    const btnDel = document.createElement('button');
    btnDel.className = 'btn btn-sm btn-outline-danger';
    btnDel.title = 'Hapus';
    btnDel.innerHTML = '<i class="bi bi-trash"></i>';
    btnDel.addEventListener('click', ()=>{
      if(confirm('Hapus item ini?')){
        removeById(r.id);
        renderList(); renderStats();
      }
    });

    tdAct.append(btnEdit, btnDup, btnDel);
    tr.appendChild(tdAct);

    // ===== Sel data (pakai textContent, aman dari HTML injection & tidak perlu esc)
    const cells = [
      r.kode || '',
      r.judul || '',
      r.rumpun_pekerjaan || '',
      r.tipe_training || '',
      `${r.durasi_teori_jam??0}/${r.durasi_praktik_jam??0}`,
      (r.updated_at ? dayjs(r.updated_at).format('YYYY-MM-DD') : '-'),
      r.id || ''
    ];

    // Kode, Judul, Rumpun, Tipe, Durasi (T/P), Updated, ID
    for (let i=0; i<cells.length; i++){
      const td = document.createElement('td');
      td.textContent = cells[i];
      if (i === cells.length - 1){ td.className = 'text-muted small'; } // kolom ID
      tr.appendChild(td);
    }

    tbody.appendChild(tr);
    refreshSuggestionsUI(); 
  }
}


function formData(){
  return {
    id: $('#lblId').dataset.id || undefined,
    version: Number($('#lblId').dataset.version || 1),
    program: $('#program').value,
    kode: $('#kode').value,
    tipe_training: $('#tipe_training').value,
    rumpun_pekerjaan: $('#rumpun_pekerjaan').value,
    judul: $('#judul').value,
    deskripsi: $('#deskripsi').value,
    kompetensi_acuan: $('#kompetensi_acuan').value,
    tujuan: $('#tujuan').value,
    materi: $('#materi').value,
    trainer: $('#trainer').value,
    evaluasi: $('#evaluasi').value,
    durasi_teori_jam: $('#durasi_teori_jam').value,
    durasi_praktik_jam: $('#durasi_praktik_jam').value,
    peserta: $('#peserta').value,
    catatan: $('#catatan').value,
    created_at: $('#lblId').dataset.created_at
  };
}
function setForm(r){
  $('#lblId').textContent = r ? `ID: ${r.id}` : 'ID: (baru)';
  $('#lblId').dataset.id = r?.id || '';
  $('#lblId').dataset.version = r?.version || 1;
  $('#lblId').dataset.created_at = r?.created_at || '';

  $('#program').value = r?.program || 'Akademi Tanam';
  $('#kode').value = r?.kode || '';
  $('#tipe_training').value = r?.tipe_training || 'Compulsory';
  $('#rumpun_pekerjaan').value = r?.rumpun_pekerjaan || '';
  $('#judul').value = r?.judul || '';
  $('#deskripsi').value = r?.deskripsi || '';
  $('#kompetensi_acuan').value = r?.kompetensi_acuan || '';
  $('#tujuan').value = (r?.tujuan||[]).join('\n');
  $('#materi').value = (r?.materi||[]).join('\n');
  $('#trainer').value = (r?.trainer||[]).join('\n');
  $('#evaluasi').value = (r?.evaluasi||[]).join('\n');
  $('#durasi_teori_jam').value = r?.durasi_teori_jam ?? 0;
  $('#durasi_praktik_jam').value = r?.durasi_praktik_jam ?? 0;
  $('#peserta').value = r?.peserta || '';
  $('#catatan').value = r?.catatan || '';

  renderPreview(r);
}

function loadForm(id){
  const r = id ? byId(id) : null;
  setForm(r);
  showView('#viewEditor');
}
function currentId(){ return $('#lblId').dataset.id || null; }

function saveForm(){
  // validasi minimal
  const fd = formData();
  if(!fd.kode.trim()) return alert('Kode wajib diisi.');
  if(!fd.judul.trim()) return alert('Judul wajib diisi.');

  let rec = normalizeFormData(fd);
  // version bump jika update
  if (currentId()){
    rec.id = currentId();
    rec.version = Number($('#lblId').dataset.version||1) + 1;
    rec.created_at = $('#lblId').dataset.created_at || rec.created_at;
  }
  upsert(rec);
  setForm(rec);
  renderList(); renderStats();
  refreshSuggestionsUI();
  alert('Tersimpan di lokal.');
}

// Delete current
function delCurrent(){
  const id = currentId();
  if(!id) return;
  if(confirm('Hapus silabus ini?')){
    removeById(id);
    renderList(); renderStats();
    loadForm(null);
  }
}

// Preview (ringkas, mirip template untuk PDF)
function renderPreview(r){
  const el = $('#previewPane');
  if(!r){ el.innerHTML = '<div class="text-muted">Belum ada data.</div>'; return; }
  el.innerHTML = `
    <div class="p-3">
      <h5 class="mb-1">${esc(r.judul)}</h5>
      <div class="small text-muted mb-2">${esc(r.kode)} • ${esc(r.tipe_training)} • ${esc(r.rumpun_pekerjaan)}</div>
      <p>${esc(r.deskripsi||'')}</p>

      <div class="row">
        <div class="col-md-6">
          <h6>Kompetensi Acuan</h6>
          <p>${esc(r.kompetensi_acuan||'-')}</p>
        </div>
        <div class="col-md-6">
          <h6>Peserta</h6>
          <p>${esc(r.peserta||'-')}</p>
        </div>
      </div>

      <div class="row">
        <div class="col-md-6"><h6>Tujuan</h6><ol>${(r.tujuan||[]).map(x=>`<li>${esc(x)}</li>`).join('')}</ol></div>
        <div class="col-md-6"><h6>Materi</h6><ol>${(r.materi||[]).map(x=>`<li>${esc(x)}</li>`).join('')}</ol></div>
      </div>

      <div class="row">
        <div class="col-md-4"><h6>Durasi</h6><p>Teori: ${r.durasi_teori_jam} jam<br>Praktik: ${r.durasi_praktik_jam} jam</p></div>
        <div class="col-md-4"><h6>Trainer</h6><ul>${(r.trainer||[]).map(x=>`<li>${esc(x)}</li>`).join('')}</ul></div>
        <div class="col-md-4"><h6>Evaluasi</h6><ul>${(r.evaluasi||[]).map(x=>`<li>${esc(x)}</li>`).join('')}</ul></div>
      </div>

      <div><h6>Catatan</h6><p>${esc(r.catatan||'-')}</p></div>
    </div>`;
}

// ======= Sync actions =======
async function doPush(where){
  await tryRun(async ()=>{
    const r = await apiPushUpsert();
    log(where, 'PUSH (upsert by id)', r);
  });
}
async function doPull(where){
  if(!confirm('Pull akan overwrite localStorage. Lanjutkan?')) return;
  await tryRun(async ()=>{
    const r = await apiPullOverwrite();
    log(where, 'PULL (overwrite local)', { count: r.count });
    renderList(); renderStats();
    loadForm(getAll()[0]?.id || null);
  });
}
async function doHeaders(where){
  await tryRun(async ()=>{
    const r = await apiHeaders();
    log(where, 'HEADERS SYNC', r);
  });
}

export { renderStats, renderList, loadForm, currentId };
