/****************************************************
 *  Akademi Tanam – Silabus (GAS Web App)
 *  Routes: init, headers, pull, push
 *  - Storage: Google Spreadsheet
 *  - Upsert by 'id', update bila version/updated_at lokal lebih baru
 *  - Content-Type: text/plain; charset=utf-8 (hindari preflight)
 ****************************************************/

// ====== Konstanta Header Baku (harus sama dengan front-end) ======
const HEADERS = [
  'id','version','updated_at','created_at','program','kode','tipe_training','rumpun_pekerjaan',
  'judul','deskripsi','kompetensi_acuan','tujuan_json','materi_json','durasi_teori_jam','durasi_praktik_jam',
  'trainer_json','peserta','evaluasi_json','catatan'
];

// ====== Utils ======
function _json(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function getSpreadsheet_(spreadsheetId){
  try{
    if(spreadsheetId) return SpreadsheetApp.openById(spreadsheetId);
  }catch(e){}
  return null;
}
function ensureSpreadsheet_(spreadsheetId){
  // Simpan/ambil ID master di Script Properties agar lintas-device selalu sama
  const props = PropertiesService.getScriptProperties();
  let masterId = String(spreadsheetId || props.getProperty('MASTER_SPREADSHEET_ID') || '').trim();

  let ss = null;
  if (masterId) {
    try { ss = SpreadsheetApp.openById(masterId); } catch(e) { ss = null; }
  }

  // Jika belum ada master, buat SEKALI lalu simpan ke properties
  if (!ss) {
    ss = SpreadsheetApp.create('AkademiTanam_Silabus');
    masterId = ss.getId();
    props.setProperty('MASTER_SPREADSHEET_ID', masterId);
  }

  return ss;
}

function ensureSheet_(ss, sheetName){
  let sh = ss.getSheetByName(sheetName);
  if (sh) return sh;

  // Jika sheetName belum ada, cek apakah ada "Sheet1" yang benar-benar kosong
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && s1.getLastColumn() === 0) {
    s1.setName(sheetName);
    return s1;
  }
  // Jika tidak, buat sheet baru dengan nama yang diminta
  return ss.insertSheet(sheetName);
}

function ensureHeaders_(sh, headers){
  if (!Array.isArray(headers) || headers.length === 0) return;

  // Pastikan jumlah kolom minimal = jumlah header (tanpa menyentuh data)
  const maxCols = sh.getMaxColumns();
  if (maxCols < headers.length) {
    sh.insertColumnsAfter(maxCols, headers.length - maxCols);
  }

  // Tulis/Perbarui baris header (baris 1) tanpa clear isi di bawahnya
  const rangeHdr = sh.getRange(1, 1, 1, headers.length);
  const currentHdr = (sh.getLastRow() >= 1) ? rangeHdr.getValues()[0] : [];

  // Susun baris header yang diinginkan
  const outHdr = headers.slice(); // menyalin array

  // Jika header sebelumnya ada dan beda, kita timpa saja sel di baris 1
  let needWrite = false;
  if (currentHdr.length !== outHdr.length) {
    needWrite = true;
  } else {
    for (let i = 0; i < outHdr.length; i++) {
      if (currentHdr[i] !== outHdr[i]) { needWrite = true; break; }
    }
  }
  if (needWrite) {
    rangeHdr.setValues([outHdr]);
  }

  // Format: nowrap (clip), non-bold
  const dataRows = Math.max(1, sh.getMaxRows());
  const allRange = sh.getRange(1, 1, dataRows, headers.length);
  const style = SpreadsheetApp.newTextStyle().setBold(false).build();
  allRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setTextStyle(style);
}


function sheetToObjects_(sh){
  const lr = sh.getLastRow();
  if(lr < 2) return [];
  const lc = sh.getLastColumn();
  const rows = sh.getRange(1,1,lr,lc).getValues();
  const hdr = rows.shift();
  const idx = {};
  hdr.forEach((h,i)=> idx[h]=i);
  const out=[];
  rows.forEach(r=>{
    const o={};
    HEADERS.forEach(h=> o[h] = r[idx[h]] );
    out.push(o);
  });
  return out;
}
function upsertRows_(sh, incoming){
  // Build index by id
  const all = sheetToObjects_(sh);
  const idxById = {};
  all.forEach((r,i)=>{ idxById[r.id] = { rowIndex: i+2, rec: r }; }); // +2 (header + 1-based)

  let inserted = 0, updated = 0, skipped = 0;

  incoming.forEach(row=>{
    const id = String(row.id||'').trim();
    if(!id){ skipped++; return; }

    const exists = idxById[id];
    if(!exists){
      // append
      const arr = HEADERS.map(h=> row[h]);
      sh.appendRow(arr);
      inserted++;
    } else {
      // compare version/updated_at
      const old = exists.rec;
      const oldVer = Number(old.version||1), newVer = Number(row.version||1);
      const oldUpd = new Date(old.updated_at||0).getTime();
      const newUpd = new Date(row.updated_at||0).getTime();

      const isNewer = (newVer > oldVer) || (newVer===oldVer && newUpd > oldUpd);
      if(isNewer){
        const r = exists.rowIndex;
        HEADERS.forEach((h, cIdx)=> sh.getRange(r, cIdx+1).setValue(row[h]));
        updated++;
      } else {
        skipped++;
      }
    }
  });
  return { inserted, updated, skipped };
}

// ====== Router ======
function doPost(e){
  try{
    const route = String(e?.parameter?.route||'').toLowerCase();
    const body  = JSON.parse(e?.postData?.contents || '{}');
    const sheetName = String(body.sheetName||'SILABUS');
    const headers = Array.isArray(body.headers) ? body.headers : HEADERS;

if(route === 'init'){
  const ss = ensureSpreadsheet_(body.spreadsheetId);
  const sh = ensureSheet_(ss, sheetName);
  ensureHeaders_(sh, headers);
  return _json({ ok:true, spreadsheetId: ss.getId(), spreadsheetUrl: ss.getUrl(), sheetName });
}

    if(route === 'headers'){
      const ss = ensureSpreadsheet_(body.spreadsheetId);
      const sh = ensureSheet_(ss, sheetName);
      ensureHeaders_(sh, headers);
      return _json({ ok:true, sheetName, headers });
    }

if(route === 'pull'){
  const ss = ensureSpreadsheet_(body.spreadsheetId);
  const sh = ensureSheet_(ss, sheetName);
  ensureHeaders_(sh, headers);
  const rows = sheetToObjects_(sh);
  return _json({ ok:true, spreadsheetId: ss.getId(), spreadsheetUrl: ss.getUrl(), sheetName, headers, rows });
}


    if(route === 'push'){
      const ss = ensureSpreadsheet_(body.spreadsheetId);
      const sh = ensureSheet_(ss, sheetName);
      ensureHeaders_(sh, headers);
      const rows = Array.isArray(body.rows) ? body.rows : [];
      const res = upsertRows_(sh, rows);
      return _json({ ok:true, sheetName, ...res });
    }

    return _json({ ok:false, error: 'Unknown route' });
  }catch(err){
    return _json({ ok:false, error: String(err && err.message || err) });
  }
}
