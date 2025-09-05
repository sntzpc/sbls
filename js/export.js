// export.js — XLSX (tabular), DOCX (docx.js), PDF (HTML → html2pdf)
import { getAll } from './storage.js';

function _stamp(){ try{ return dayjs().format('YYYYMMDD-HHmmss'); }catch(_){ return Date.now(); } }
function _ensureDocx(){
  if(!window.docx) throw new Error('Library docx.js belum dimuat. Tambahkan: <script src="https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js"></script>');
  return window.docx;
}
const _esc = (s)=> (s==null ? '' : String(s));
const _arr = (v)=> Array.isArray(v) ? v.filter(Boolean) : (_esc(v) ? [_esc(v)] : []);

/* =========================
   1) XLSX — sheet tabular (tetap)
   ========================= */
export function exportXLSX(){
  if (!window.XLSX) throw new Error('Library XLSX belum siap.');
  const rows = getAll();
  if(rows.length===0) throw new Error('Tidak ada data untuk XLSX.');

  const header = [
    "id","version","updated_at","created_at","program","kode","tipe_training","rumpun_pekerjaan",
    "judul","deskripsi","kompetensi_acuan","tujuan","materi","durasi_teori_jam","durasi_praktik_jam",
    "trainer","peserta","evaluasi","catatan"
  ];
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
  // sedikit lebar kolom agar enak dibaca
  ws['!cols'] = header.map(()=>({ wch: 22 }));
  window.XLSX.utils.book_append_sheet(wb, ws, "SILABUS");
  window.XLSX.writeFile(wb, `Silabus_AkademiTanam_${_stamp()}.xlsx`);
}

/* =========================================
   2) DOCX — build tabel dengan docx.js (ONE)
   ========================================= */
export async function exportDOCX(currentId){
  const list = getAll();
  const rec  = currentId ? list.find(r=>r.id===currentId) : list[0];
  if(!rec) throw new Error('Tidak ada data silabus untuk diekspor.');
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
          WidthType, AlignmentType, BorderStyle, VerticalAlign } = _ensureDocx();

  // helpers
  const cellTxt = (txt, opts={})=> new TableCell({
    width: { size: opts.w || 50, type: WidthType.PERCENTAGE },
    verticalAlign: VerticalAlign.CENTER,
    borders: { top:{style:BorderStyle.SINGLE,size:6,color:"000000"},
               bottom:{style:BorderStyle.SINGLE,size:6,color:"000000"},
               left:{style:BorderStyle.SINGLE,size:6,color:"000000"},
               right:{style:BorderStyle.SINGLE,size:6,color:"000000"} },
    shading: opts.shade ? { fill: "DDDDDD" } : undefined,
    children: [ new Paragraph({ children: [ new TextRun(String(txt||"")) ] }) ],
  });
  const bullets = (arr)=> _arr(arr).map(t=> new Paragraph({
    text: t, bullet: { level: 0 }
  }));
  const para = (t)=> new Paragraph({ children: [ new TextRun(String(t||"")) ] });

  // nested table Durasi
  const durasiTbl = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          cellTxt("Praktek", { shade:true, w:50 }), cellTxt("Teori", { shade:true, w:50 })
        ]
      }),
      new TableRow({
        children: [
          cellTxt(_esc(rec.durasi_praktik_jam??''), { w:50 }),
          cellTxt(_esc(rec.durasi_teori_jam??''), { w:50 })
        ]
      })
    ]
  });

  // tabel utama (2 kolom label/isi) + beberapa baris 4 kolom (tipe/rumpun)
  const rows = [];
  rows.push(new TableRow({ children: [ cellTxt("No CODE", {shade:true, w:32}), cellTxt(rec.kode, {w:68}) ]}));
  rows.push(new TableRow({ children: [
    cellTxt("Tipe Training", {shade:true, w:32}),
    cellTxt(rec.tipe_training, {w:34}),
    cellTxt("Rumpun Pekerjaan", {shade:true, w:16}), // pecah 4 kolom agar mirip
    cellTxt(rec.rumpun_pekerjaan, {w:18}),
  ]}));
  rows.push(new TableRow({ children: [ cellTxt("Judul", {shade:true, w:32}), cellTxt(rec.judul, {w:68}) ]}));
  rows.push(new TableRow({ children: [ cellTxt("Deskripsi", {shade:true, w:32}), new TableCell({
    width:{ size:68, type:WidthType.PERCENTAGE }, borders:{ top:{style:BorderStyle.SINGLE,size:6,color:"000000"},
    bottom:{style:BorderStyle.SINGLE,size:6,color:"000000"}, left:{style:BorderStyle.SINGLE,size:6,color:"000000"},
    right:{style:BorderStyle.SINGLE,size:6,color:"000000"} }, children: [ para(rec.deskripsi) ] }) ]}));
  rows.push(new TableRow({ children: [ cellTxt("Kompetensi Acuan", {shade:true, w:32}), new TableCell({
    width:{ size:68, type:WidthType.PERCENTAGE }, borders:{ top:{style:BorderStyle.SINGLE,size:6,color:"000000"},
    bottom:{style:BorderStyle.SINGLE,size:6,color:"000000"}, left:{style:BorderStyle.SINGLE,size:6,color:"000000"},
    right:{style:BorderStyle.SINGLE,size:6,color:"000000"} }, children: [ para(rec.kompetensi_acuan) ] }) ]}));
  rows.push(new TableRow({ children: [ cellTxt("Tujuan", {shade:true, w:32}), new TableCell({
    width:{ size:68, type:WidthType.PERCENTAGE }, borders:{ top:{style:BorderStyle.SINGLE,size:6,color:"000000"},
    bottom:{style:BorderStyle.SINGLE,size:6,color:"000000"}, left:{style:BorderStyle.SINGLE,size:6,color:"000000"},
    right:{style:BorderStyle.SINGLE,size:6,color:"000000"} }, children: bullets(rec.tujuan) }) ]}));
  rows.push(new TableRow({ children: [ cellTxt("Materi", {shade:true, w:32}), new TableCell({
    width:{ size:68, type:WidthType.PERCENTAGE }, borders:{ top:{style:BorderStyle.SINGLE,size:6,color:"000000"},
    bottom:{style:BorderStyle.SINGLE,size:6,color:"000000"}, left:{style:BorderStyle.SINGLE,size:6,color:"000000"},
    right:{style:BorderStyle.SINGLE,size:6,color:"000000"} }, children: bullets(rec.materi) }) ]}));
  rows.push(new TableRow({ children: [ cellTxt("Durasi (jam)", {shade:true, w:32}), new TableCell({
    width:{ size:68, type:WidthType.PERCENTAGE }, borders:{ top:{style:BorderStyle.SINGLE,size:6,color:"000000"},
    bottom:{style:BorderStyle.SINGLE,size:6,color:"000000"}, left:{style:BorderStyle.SINGLE,size:6,color:"000000"},
    right:{style:BorderStyle.SINGLE,size:6,color:"000000"} }, children: [ durasiTbl ] }) ]}));
  rows.push(new TableRow({ children: [ cellTxt("Trainer", {shade:true, w:32}), new TableCell({
    width:{ size:68, type:WidthType.PERCENTAGE }, borders:{ top:{style:BorderStyle.SINGLE,size:6,color:"000000"},
    bottom:{style:BorderStyle.SINGLE,size:6,color:"000000"}, left:{style:BorderStyle.SINGLE,size:6,color:"000000"},
    right:{style:BorderStyle.SINGLE,size:6,color:"000000"} }, children: bullets(rec.trainer) }) ]}));
  rows.push(new TableRow({ children: [ cellTxt("Peserta", {shade:true, w:32}), cellTxt(rec.peserta, {w:68}) ]}));
  rows.push(new TableRow({ children: [ cellTxt("Evaluasi/Pengukuran", {shade:true, w:32}), new TableCell({
    width:{ size:68, type:WidthType.PERCENTAGE }, borders:{ top:{style:BorderStyle.SINGLE,size:6,color:"000000"},
    bottom:{style:BorderStyle.SINGLE,size:6,color:"000000"}, left:{style:BorderStyle.SINGLE,size:6,color:"000000"},
    right:{style:BorderStyle.SINGLE,size:6,color:"000000"} }, children: bullets(rec.evaluasi) }) ]}));
  rows.push(new TableRow({ children: [ cellTxt("Catatan", {shade:true, w:32}), cellTxt(rec.catatan, {w:68}) ]}));

  const table = new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows });
  const doc = new Document({
    sections: [{ properties: { page: { margin: { top: 720, right: 720, bottom: 720, left: 720 } } }, children: [ table ] }]
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `Silabus_${_esc(rec.kode)||'NoCode'}.docx`);
}

/* =============================================
   3) DOCX (ALL) → ZIP (pakai docx.js + JSZip)
   ============================================= */
export async function exportDOCXAll(onProgress){
  if(!window.JSZip) throw new Error('JSZip belum dimuat.');
  const list = getAll();
  if(!list.length) throw new Error('Tidak ada data silabus.');

  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
          WidthType, AlignmentType, BorderStyle, VerticalAlign } = _ensureDocx();

  const zip = new JSZip();
  const total = list.length;

  let i = 0;
  for(const rec of list){
    i++;
    const fname = `Silabus_${_esc(rec.kode)||('NoCode_'+i)}.docx`;
    if (typeof onProgress === 'function') onProgress(i, total, fname);

    // build doc sama seperti single
    const cellTxt = (txt, opts={})=> new TableCell({
      width: { size: opts.w || 50, type: WidthType.PERCENTAGE },
      verticalAlign: VerticalAlign.CENTER,
      borders: { top:{style:BorderStyle.SINGLE,size:6,color:"000000"},
                 bottom:{style:BorderStyle.SINGLE,size:6,color:"000000"},
                 left:{style:BorderStyle.SINGLE,size:6,color:"000000"},
                 right:{style:BorderStyle.SINGLE,size:6,color:"000000"} },
      shading: opts.shade ? { fill: "DDDDDD" } : undefined,
      children: [ new Paragraph({ children: [ new TextRun(String(txt||"")) ] }) ],
    });
    const bullets = (arr)=> _arr(arr).map(t=> new Paragraph({ text: t, bullet: { level: 0 } }));

    const durasiTbl = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [ cellTxt("Praktek", { shade:true, w:50 }), cellTxt("Teori", { shade:true, w:50 }) ] }),
        new TableRow({ children: [ cellTxt(_esc(rec.durasi_praktik_jam??''), { w:50 }), cellTxt(_esc(rec.durasi_teori_jam??''), { w:50 }) ] })
      ]
    });

    const rows = [];
    rows.push(new TableRow({ children: [ cellTxt("No CODE", {shade:true, w:32}), cellTxt(rec.kode, {w:68}) ]}));
    rows.push(new TableRow({ children: [
      cellTxt("Tipe Training", {shade:true, w:32}),
      cellTxt(rec.tipe_training, {w:34}),
      cellTxt("Rumpun Pekerjaan", {shade:true, w:16}),
      cellTxt(rec.rumpun_pekerjaan, {w:18}),
    ]}));
    rows.push(new TableRow({ children: [ cellTxt("Judul", {shade:true, w:32}), cellTxt(rec.judul, {w:68}) ]}));
    rows.push(new TableRow({ children: [ cellTxt("Deskripsi", {shade:true, w:32}), cellTxt(rec.deskripsi, {w:68}) ]}));
    rows.push(new TableRow({ children: [ cellTxt("Kompetensi Acuan", {shade:true, w:32}), cellTxt(rec.kompetensi_acuan, {w:68}) ]}));
    rows.push(new TableRow({ children: [ cellTxt("Tujuan", {shade:true, w:32}), new TableCell({
      width:{ size:68, type:WidthType.PERCENTAGE }, borders:{ top:{style:BorderStyle.SINGLE,size:6,color:"000000"},
      bottom:{style:BorderStyle.SINGLE,size:6,color:"000000"}, left:{style:BorderStyle.SINGLE,size:6,color:"000000"},
      right:{style:BorderStyle.SINGLE,size:6,color:"000000"} }, children: bullets(rec.tujuan) }) ]}));
    rows.push(new TableRow({ children: [ cellTxt("Materi", {shade:true, w:32}), new TableCell({
      width:{ size:68, type:WidthType.PERCENTAGE }, borders:{ top:{style:BorderStyle.SINGLE,size:6,color:"000000"},
      bottom:{style:BorderStyle.SINGLE,size:6,color:"000000"}, left:{style:BorderStyle.SINGLE,size:6,color:"000000"},
      right:{style:BorderStyle.SINGLE,size:6,color:"000000"} }, children: bullets(rec.materi) }) ]}));
    rows.push(new TableRow({ children: [ cellTxt("Durasi (jam)", {shade:true, w:32}), new TableCell({
      width:{ size:68, type:WidthType.PERCENTAGE }, borders:{ top:{style:BorderStyle.SINGLE,size:6,color:"000000"},
      bottom:{style:BorderStyle.SINGLE,size:6,color:"000000"}, left:{style:BorderStyle.SINGLE,size:6,color:"000000"},
      right:{style:BorderStyle.SINGLE,size:6,color:"000000"} }, children: [ durasiTbl ] }) ]}));
    rows.push(new TableRow({ children: [ cellTxt("Trainer", {shade:true, w:32}), new TableCell({
      width:{ size:68, type:WidthType.PERCENTAGE }, borders:{ top:{style:BorderStyle.SINGLE,size:6,color:"000000"},
      bottom:{style:BorderStyle.SINGLE,size:6,color:"000000"}, left:{style:BorderStyle.SINGLE,size:6,color:"000000"},
      right:{style:BorderStyle.SINGLE,size:6,color:"000000"} }, children: bullets(rec.trainer) }) ]}));
    rows.push(new TableRow({ children: [ cellTxt("Peserta", {shade:true, w:32}), cellTxt(rec.peserta, {w:68}) ]}));
    rows.push(new TableRow({ children: [ cellTxt("Evaluasi/Pengukuran", {shade:true, w:32}), new TableCell({
      width:{ size:68, type:WidthType.PERCENTAGE }, borders:{ top:{style:BorderStyle.SINGLE,size:6,color:"000000"},
      bottom:{style:BorderStyle.SINGLE,size:6,color:"000000"}, left:{style:BorderStyle.SINGLE,size:6,color:"000000"},
      right:{style:BorderStyle.SINGLE,size:6,color:"000000"} }, children: bullets(rec.evaluasi) }) ]}));
    rows.push(new TableRow({ children: [ cellTxt("Catatan", {shade:true, w:32}), cellTxt(rec.catatan, {w:68}) ]}));

    const table = new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows });
    const doc = new Document({
      sections: [{ properties: { page: { margin: { top: 720, right: 720, bottom: 720, left: 720 } } }, children: [ table ] }]
    });

    // eslint-disable-next-line no-await-in-loop
    const blob = await Packer.toBlob(doc);
    zip.file(fname, blob);
  }

  const zipBlob = await zip.generateAsync({ type: 'blob' });
  saveAs(zipBlob, `Silabus_DOCX_${_stamp()}.zip`);
}

/* =========================================
   4) PDF — HTML builder (ONE) & ZIP (ALL)
   ========================================= */
function _buildPrintHTML(rec){
  const esc = (s)=> (s==null ? '' : String(s)
      .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'));
  const lines = (arr)=> _arr(arr).map(x=>`<div>• ${esc(x)}</div>`).join('') || '&nbsp;';
  const join  = (arr, sep='<br/>')=> _arr(arr).map(esc).join(sep) || '&nbsp;';

  // no fixed height to avoid blank
  return `
  <style>
    .page { width: 210mm; padding: 25.4mm; box-sizing: border-box; background:#fff; color:#000; font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt; line-height: 1.25; }
    table.doc { width: 100%; border-collapse: collapse; table-layout: fixed; }
    table.doc td { border: 0.75pt solid #000; padding: 6pt; vertical-align: top; word-wrap: break-word; }
    .lbl { width: 32%; font-weight: 600; background: #e5e5e5; }
    .val { width: 68%; }
    .title { font-weight: 700; font-size: 12pt; }
  </style>
  <div class="page">
    <table class="doc">
      <tr><td class="lbl">No CODE</td><td class="val title">${esc(rec.kode)||'&nbsp;'}</td></tr>
      <tr>
        <td class="lbl">Tipe Training</td><td class="val">${esc(rec.tipe_training)||'&nbsp;'}</td>
      </tr>
      <tr><td class="lbl">Rumpun Pekerjaan</td><td class="val">${esc(rec.rumpun_pekerjaan)||'&nbsp;'}</td></tr>
      <tr><td class="lbl">Judul</td><td class="val title">${esc(rec.judul)||'&nbsp;'}</td></tr>
      <tr><td class="lbl">Deskripsi</td><td class="val">${esc(rec.deskripsi)||'&nbsp;'}</td></tr>
      <tr><td class="lbl">Kompetensi Acuan</td><td class="val">${esc(rec.kompetensi_acuan)||'&nbsp;'}</td></tr>
      <tr><td class="lbl">Tujuan</td><td class="val">${lines(rec.tujuan)}</td></tr>
      <tr><td class="lbl">Materi</td><td class="val">${lines(rec.materi)}</td></tr>
      <tr>
        <td class="lbl">Durasi (jam)</td>
        <td class="val">
          <table style="width:100%; border-collapse: collapse;">
            <tr>
              <td style="border:0.75pt solid #000; width:50%; padding:4pt;"><b>Praktek</b><br/>${esc(rec.durasi_praktik_jam ?? '')||'&nbsp;'}</td>
              <td style="border:0.75pt solid #000; width:50%; padding:4pt;"><b>Teori</b><br/>${esc(rec.durasi_teori_jam ?? '')||'&nbsp;'}</td>
            </tr>
          </table>
        </td>
      </tr>
      <tr><td class="lbl">Trainer</td><td class="val">${join(rec.trainer)}</td></tr>
      <tr><td class="lbl">Peserta</td><td class="val">${esc(rec.peserta)||'&nbsp;'}</td></tr>
      <tr><td class="lbl">Evaluasi/Pengukuran</td><td class="val">${lines(rec.evaluasi)}</td></tr>
      <tr><td class="lbl">Catatan</td><td class="val">${esc(rec.catatan)||'&nbsp;'}</td></tr>
    </table>
  </div>`;
}

export function exportPDF(currentId){
  const list = getAll();
  const rec  = currentId ? list.find(r=>r.id===currentId) : list[0];
  if(!rec) throw new Error('Tidak ada data silabus untuk diekspor.');

  const html = _buildPrintHTML(rec);
  const holder = document.createElement('div');
  holder.style.position = 'fixed';
  holder.style.left = '-99999px';
  holder.style.top = '0';
  holder.innerHTML = html;
  document.body.appendChild(holder);

  const opt = {
    margin: [10, 10, 10, 10],
    filename: `Silabus_${_esc(rec.kode)||'NoCode'}.pdf`,
    image: { type: 'jpeg', quality: 0.98 },
    html2canvas: { scale: 2, useCORS: true, backgroundColor: '#ffffff' },
    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
  };

  html2pdf().set(opt).from(holder).save().then(()=> holder.remove()).catch(e=>{ holder.remove(); throw e; });
}

export async function exportPDFAll(onProgress){
  if(!window.JSZip) throw new Error('JSZip belum dimuat.');
  const list = getAll();
  if(!list.length) throw new Error('Tidak ada data silabus.');

  const zip = new JSZip();
  const total = list.length;

  let i = 0;
  for(const rec of list){
    i++;
    const fname = `Silabus_${_esc(rec.kode)||('NoCode_'+i)}.pdf`;
    if (typeof onProgress === 'function') onProgress(i, total, fname);

    const html = _buildPrintHTML(rec);
    const holder = document.createElement('div');
    holder.style.position = 'fixed';
    holder.style.left = '-99999px';
    holder.style.top = '0';
    holder.innerHTML = html;
    document.body.appendChild(holder);

    const opt = {
      margin: [10, 10, 10, 10],
      filename: fname,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true, backgroundColor: '#ffffff' },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    // eslint-disable-next-line no-await-in-loop
    const blob = await html2pdf().set(opt).from(holder).toPdf().get('pdf').then(pdf => pdf.output('blob')).finally(()=> holder.remove());
    zip.file(fname, blob);
  }

  const zipBlob = await zip.generateAsync({ type: 'blob' });
  saveAs(zipBlob, `Silabus_PDF_${_stamp()}.zip`);
}
