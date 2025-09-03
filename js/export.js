// export.js — DOCX, XLSX, PDF exports
import { getAll, getConfig } from './storage.js';

// ==== DOCX via Docxtemplater + PizZip ====
// Template harus tersedia di /assets/Template-Silabus.docx atau sesuai config.templateUrl
export async function exportDOCX(currentId){
  const cfg = getConfig();
  const list = getAll();
  const rec = currentId ? list.find(r=>r.id===currentId) : list[0];
  if(!rec) throw new Error('Tidak ada data silabus untuk diekspor.');

  const templateRes = await fetch(cfg.templateUrl);
  if(!templateRes.ok) throw new Error('Gagal memuat template DOCX.');
  const arrayBuffer = await templateRes.arrayBuffer();

  const zip = new PizZip(arrayBuffer);
  const doc = new window.docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

  // data mapping (sesuai template placeholder)
  const data = {
    program: rec.program,
    kode: rec.kode,
    tipe_training: rec.tipe_training,
    rumpun_pekerjaan: rec.rumpun_pekerjaan,
    judul: rec.judul,
    deskripsi: rec.deskripsi,
    kompetensi_acuan: rec.kompetensi_acuan,
    tujuan: rec.tujuan,
    materi: rec.materi,
    durasi_teori_jam: String(rec.durasi_teori_jam||0),
    durasi_praktik_jam: String(rec.durasi_praktik_jam||0),
    trainer: rec.trainer,
    peserta: rec.peserta,
    evaluasi: rec.evaluasi,
    catatan: rec.catatan
  };

  doc.setData(data);
  doc.render();

  const out = doc.getZip().generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });
  saveAs(out, `Silabus_${rec.kode||'NoCode'}.docx`);
}

// ==== XLSX via SheetJS ====
export function exportXLSX(){
  if (!window.XLSX) {
    throw new Error('Library XLSX belum siap. Coba tunggu beberapa detik atau refresh halaman.');
  }
  const rows = getAll();
  if(rows.length===0) throw new Error('Tidak ada data untuk XLSX.');
  const header = ["id","version","updated_at","created_at","program","kode","tipe_training","rumpun_pekerjaan","judul","deskripsi","kompetensi_acuan","tujuan","materi","durasi_teori_jam","durasi_praktik_jam","trainer","peserta","evaluasi","catatan"];
  const sheetData = [header];
  for(const r of rows){
    sheetData.push([
      r.id, r.version, r.updated_at, r.created_at, r.program, r.kode, r.tipe_training, r.rumpun_pekerjaan,
      r.judul, r.deskripsi, r.kompetensi_acuan,
      (r.tujuan||[]).join(' | '),
      (r.materi||[]).join(' | '),
      r.durasi_teori_jam, r.durasi_praktik_jam,
      (r.trainer||[]).join(' | '),
      r.peserta,
      (r.evaluasi||[]).join(' | '),
      r.catatan
    ]);
  }
  const wb = window.XLSX.utils.book_new();
  const ws = window.XLSX.utils.aoa_to_sheet(sheetData);
  window.XLSX.utils.book_append_sheet(wb, ws, "Silabus");
  window.XLSX.writeFile(wb, "Silabus_AkademiTanam.xlsx");
}

// ==== PDF via html2pdf (render dari preview) ====
export function exportPDF(){
  const el = document.getElementById('previewPane');
  if(!el || !el.innerText.trim()) throw new Error('Preview kosong. Simpan/ubah data terlebih dahulu.');
  const opt = {
    margin: 10,
    filename: 'Silabus_AkademiTanam.pdf',
    image: { type: 'jpeg', quality: 0.95 },
    html2canvas: { scale: 2, useCORS: true },
    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
  };
  html2pdf().set(opt).from(el).save();
}
