// storage.js — localStorage model & helpers

export const NS = 'AT_SILABUS_V1';
export const CFG = 'AT_CONFIG_V1';

const defaultConfig = {
  spreadsheetId: '',
  sheetName: 'SILABUS',
  templateUrl: './assets/Template-Silabus.docx'
};

export function getConfig(){
  const raw = localStorage.getItem(CFG);
  return raw ? { ...defaultConfig, ...JSON.parse(raw) } : { ...defaultConfig };
}
export function setConfig(patch){
  const cfg = { ...getConfig(), ...(patch||{}) };
  localStorage.setItem(CFG, JSON.stringify(cfg));
  return cfg;
}

export function getAll(){
  const raw = localStorage.getItem(NS);
  return raw ? JSON.parse(raw) : [];
}
export function setAll(arr){
  localStorage.setItem(NS, JSON.stringify(arr||[]));
}
export function upsert(record){
  const all = getAll();
  const idx = all.findIndex(r => r.id === record.id);
  if (idx >= 0) all[idx] = record; else all.unshift(record);
  setAll(all);
}
export function removeById(id){
  setAll(getAll().filter(r=>r.id!==id));
}
export function byId(id){ return getAll().find(r=>r.id===id) || null; }

export function generateId(kode, judul){
  const stamp = dayjs().format('YYYYMMDD-HHmmss');
  const slug = (String(judul||kode||'SILABUS').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/(^-|-$)/g,'')).slice(0,16);
  return `AT-${stamp}-${slug||'item'}`;
}

export function normalizeFormData(form){
  // textareas -> arrays per baris
  const splitLines = (v)=> String(v||'').split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
  const asNumber = (v)=> (v===''||v==null) ? 0 : Number(v);

  return {
    id: form.id || generateId(form.kode, form.judul),
    version: Number(form.version || 1),
    created_at: form.created_at || dayjs().toISOString(),
    updated_at: dayjs().toISOString(),

    program: String(form.program||'Akademi Tanam'),
    kode: String(form.kode||'').trim(),
    tipe_training: String(form.tipe_training||'Compulsory'),
    rumpun_pekerjaan: String(form.rumpun_pekerjaan||'').trim(),
    judul: String(form.judul||'').trim(),
    deskripsi: String(form.deskripsi||'').trim(),
    kompetensi_acuan: String(form.kompetensi_acuan||'').trim(),

    tujuan: Array.isArray(form.tujuan) ? form.tujuan : splitLines(form.tujuan),
    materi: Array.isArray(form.materi) ? form.materi : splitLines(form.materi),
    trainer: Array.isArray(form.trainer) ? form.trainer : splitLines(form.trainer),
    evaluasi: Array.isArray(form.evaluasi) ? form.evaluasi : splitLines(form.evaluasi),

    durasi_teori_jam: asNumber(form.durasi_teori_jam),
    durasi_praktik_jam: asNumber(form.durasi_praktik_jam),

    peserta: String(form.peserta||'').trim(),
    catatan: String(form.catatan||'').trim(),
  };
}

// Konversi record -> row untuk Sheet (stringify arrays ke *_json)
export function toRow(r){
  return {
    id: r.id, version: r.version, updated_at: r.updated_at, created_at: r.created_at,
    program: r.program, kode: r.kode, tipe_training: r.tipe_training, rumpun_pekerjaan: r.rumpun_pekerjaan,
    judul: r.judul, deskripsi: r.deskripsi, kompetensi_acuan: r.kompetensi_acuan,
    tujuan_json: JSON.stringify(r.tujuan||[]),
    materi_json: JSON.stringify(r.materi||[]),
    durasi_teori_jam: r.durasi_teori_jam,
    durasi_praktik_jam: r.durasi_praktik_jam,
    trainer_json: JSON.stringify(r.trainer||[]),
    peserta: r.peserta,
    evaluasi_json: JSON.stringify(r.evaluasi||[]),
    catatan: r.catatan
  };
}

// rows (dari GAS) -> record
export function fromRow(row){
  const parseArr = (v)=> {
    try{ return JSON.parse(v||'[]'); }catch(e){ return []; }
  };
  return {
    id: row.id, version: Number(row.version||1), updated_at: row.updated_at, created_at: row.created_at,
    program: row.program, kode: row.kode, tipe_training: row.tipe_training, rumpun_pekerjaan: row.rumpun_pekerjaan,
    judul: row.judul, deskripsi: row.deskripsi, kompetensi_acuan: row.kompetensi_acuan,
    tujuan: parseArr(row.tujuan_json), materi: parseArr(row.materi_json),
    durasi_teori_jam: Number(row.durasi_teori_jam||0), durasi_praktik_jam: Number(row.durasi_praktik_jam||0),
    trainer: parseArr(row.trainer_json), peserta: row.peserta,
    evaluasi: parseArr(row.evaluasi_json), catatan: row.catatan
  };
}

// Header baku
export const SHEET_HEADERS = [
  'id','version','updated_at','created_at','program','kode','tipe_training','rumpun_pekerjaan',
  'judul','deskripsi','kompetensi_acuan','tujuan_json','materi_json','durasi_teori_jam','durasi_praktik_jam',
  'trainer_json','peserta','evaluasi_json','catatan'
];
