const JSZip = globalThis.JSZip;
if (!JSZip) throw new Error('JSZip gagal dimuat. Pastikan jszip.min.js ikut di-upload.');
import { analyzeDocx, repairDocx } from './processor.js';
let currentFiles = [];
let currentFileIndex = 0;
let currentAnalysis = null;
let currentResult = null;
let batchBlob = null;
const currentFile = () => currentFiles[currentFileIndex] || null;
const defaultOptions = {
    documentType: 'auto', structureMode: 'auto', numberingProfile: 'uin-madura',
    mode: 'safe', paper: 'A4', marginTopCm: 4, marginRightCm: 3, marginBottomCm: 3, marginLeftCm: 4,
    normalizeMargins: true,
    repairPageNumbering: true,
    createToc: true,
    repairChapterPagination: true,
    repairHeadings: true,
    normalizeBodyFormatting: false,
    createTableList: false,
    createFigureList: false,
    cleanBlankParagraphs: true,
    validateDocument: true,
    removeLikelyManualToc: true,
    includeFrontMatterInToc: true,
    fontFamily: 'Times New Roman', fontSizePt: 12, lineSpacing: 1.5, firstLineIndentCm: 1.27,
    headingOverrides: {},
};
let uiOptions = { ...defaultOptions };
const icons = {
    upload: '<svg viewBox="0 0 24 24"><path d="M12 16V4m0 0L7.5 8.5M12 4l4.5 4.5M5 14v4a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-4"/></svg>',
    check: '<svg viewBox="0 0 24 24"><path d="m5 12 4 4L19 6"/></svg>',
    warn: '<svg viewBox="0 0 24 24"><path d="M12 9v4m0 4h.01M10.3 4.4 2.5 18a1 1 0 0 0 .9 1.5h17.2a1 1 0 0 0 .9-1.5L13.7 4.4a1 1 0 0 0-1.7 0Z"/></svg>',
    magic: '<svg viewBox="0 0 24 24"><path d="m15 4 5 5L9 20H4v-5L15 4Zm-7-1 1 3 3 1-3 1-1 3-1-3-3-1 3-1 1-3Z"/></svg>',
    download: '<svg viewBox="0 0 24 24"><path d="M12 4v11m0 0 4-4m-4 4-4-4M5 20h14"/></svg>',
    doc: '<svg viewBox="0 0 24 24"><path d="M7 3h7l4 4v14H7zM14 3v5h5M9.5 12h5M9.5 16h5"/></svg>',
    layers: '<svg viewBox="0 0 24 24"><path d="m12 3 9 5-9 5-9-5 9-5Zm-9 10 9 5 9-5M3 17l9 5 9-5"/></svg>',
};
const icon = (name) => icons[name] || '';
const esc = (s) => s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[c] || c));
function render() {
    document.querySelector('#app').innerHTML = `
    <div class="shell">
      <header class="topbar">
        <div class="brand"><div class="logo">Z</div><div><strong>ZAIN.NET</strong><span>Perapih Dokumen Akademik V4</span></div></div>
        <div class="badge">DOCX/XML • LOCAL PROCESSING • TANPA API</div>
      </header>
      <main>
        <section class="hero">
          <div><div class="eyebrow">ACADEMIC DOCX ENGINE V4</div><h1>Rapikan <em>Skripsi, Proposal & Makalah</em> dengan mesin yang berbeda.</h1><p>Upload DOCX, pilih jenis dokumen atau Auto Deteksi, lalu aplikasi memakai struktur, heading, TOC dan penomoran yang sesuai. V4 menambahkan pemeriksaan anti-rusak sebelum hasil boleh diunduh.</p></div>
          <div class="hero-card"><span class="mini-label">ANTI-RUSAK V4</span><div class="preset-grid"><span>Skripsi • BAB I–V</span><span>Proposal • BAB / Tanpa BAB</span><span>Makalah • Panjang / Pendek</span><span>Integritas konten</span></div></div>
        </section>

        <section class="workspace">
          <div class="panel upload-panel">
            <div class="panel-head"><div><span class="step">01</span><h2>Upload dokumen</h2></div><span class="soft">Bisa banyak file</span></div>
            <label class="dropzone ${currentFiles.length ? 'has-file' : ''}" id="dropzone">
              <input id="fileInput" type="file" multiple accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" />
              <div class="drop-icon">${icon(currentFiles.length ? 'doc' : 'upload')}</div>
              <strong>${currentFiles.length ? `${currentFiles.length} file DOCX dipilih` : 'Tarik file DOCX ke sini'}</strong>
              <span>${currentFiles.length ? esc(currentFile()?.name || '') : 'atau klik untuk memilih satu / beberapa dokumen akademik'}</span>
            </label>
            ${fileTabsHtml()}
            <button class="btn secondary" id="analyzeBtn" ${currentFile() ? '' : 'disabled'}>Analisis File Aktif</button>
          </div>

          <div class="panel settings-panel">
            <div class="panel-head"><div><span class="step">02</span><h2>Pilih fitur yang dipakai</h2></div><div class="head-actions"><button class="text-btn" id="coreBtn">Inti Aman</button><button class="text-btn" id="allBtn">Semua (Aman)</button><button class="text-btn" id="noneBtn">Kosongkan</button></div></div>
            ${settingsHtml()}
          </div>
        </section>

        ${analysisHtml()}
        ${resultHtml()}
      </main>
      <footer>ZAIN.NET • Pemrosesan DOCX berlangsung lokal di browser • Selalu simpan file asli sebagai backup.</footer>
    </div>`;
    bindEvents();
}
function fileTabsHtml() {
    if (!currentFiles.length)
        return '';
    return `<div class="file-tabs">${currentFiles.map((f, i) => `<button class="file-tab ${i === currentFileIndex ? 'active' : ''}" data-file-index="${i}" title="${esc(f.name)}"><b>${i + 1}</b><span>${esc(f.name)}</span></button>`).join('')}</div>`;
}
function settingsHtml() {
    const o = uiOptions;
    const sel = (v, current) => v === current ? 'selected' : '';
    const modules = [
        ['normalizeMargins', 'Margin & ukuran kertas', 'Seragamkan margin tanpa merusak section landscape pada Mode Aman.', o.normalizeMargins],
        ['repairPageNumbering', 'Penomoran sesuai jenis dokumen', 'Skripsi, Proposal, dan Makalah memakai mesin nomor halaman yang berbeda.', o.repairPageNumbering],
        ['createToc', 'Daftar Isi TOC Level 1–3', 'Heading disesuaikan dengan struktur BAB atau Tanpa BAB.', o.createToc],
        ['repairChapterPagination', 'Halaman awal & struktur utama', 'Rapikan awal halaman, BAB, atau awal isi proposal/makalah tanpa mengandalkan Enter.', o.repairChapterPagination],
        ['repairHeadings', 'Rapikan judul & subjudul', 'Bersihkan Enter/Shift+Enter dan Spacing Before/After yang terlalu jauh.', o.repairHeadings],
        ['normalizeBodyFormatting', 'Format isi dokumen', 'Font, ukuran, spasi, justify, dan first-line indent secara konservatif.', o.normalizeBodyFormatting],
        ['createTableList', 'Daftar Tabel otomatis', 'Deteksi caption “Tabel …” dan buat daftar dinamis Word.', o.createTableList],
        ['createFigureList', 'Daftar Gambar otomatis', 'Deteksi caption “Gambar …” dan buat daftar dinamis Word.', o.createFigureList],
        ['cleanBlankParagraphs', 'Bersihkan Enter berlebih (isi)', 'Hanya bagian isi utama. Cover dan halaman awal dilindungi.', o.cleanBlankParagraphs],
        ['validateDocument', 'Validasi + anti-rusak', 'Cek struktur, teks utama, tabel, gambar, media, footnote/endnote sebelum download.', o.validateDocument],
    ];
    return `
    <div class="doc-profile-box">
      <div class="profile-title"><strong>Jenis dokumen</strong><small>Pilih manual jika Auto Deteksi ragu.</small></div>
      <div class="preset-row v4-profile">
        <label class="field wide"><span>Dokumen</span><select id="documentType">
          <option value="auto" ${sel('auto',o.documentType)}>Auto Deteksi</option>
          <option value="skripsi" ${sel('skripsi',o.documentType)}>Skripsi</option>
          <option value="proposal" ${sel('proposal',o.documentType)}>Proposal</option>
          <option value="makalah" ${sel('makalah',o.documentType)}>Makalah</option>
        </select></label>
        <label class="field wide"><span>Struktur</span><select id="structureMode">
          <option value="auto" ${sel('auto',o.structureMode)}>Auto Deteksi BAB</option>
          <option value="with-bab" ${sel('with-bab',o.structureMode)}>Dengan BAB</option>
          <option value="without-bab" ${sel('without-bab',o.structureMode)}>Tanpa BAB</option>
        </select></label>
        <label class="field wide"><span>Profil penomoran</span><select id="numberingProfile">
          <option value="uin-madura" ${sel('uin-madura',o.numberingProfile)}>UIN Madura / Akademik</option>
          <option value="zain-legacy" ${sel('zain-legacy',o.numberingProfile)}>ZAIN.NET Skripsi lama</option>
          <option value="simple" ${sel('simple',o.numberingProfile)}>Sederhana: Cover → Isi 1</option>
        </select></label>
      </div>
      <div class="profile-hints"><span><b>Skripsi:</b> bagian awal Romawi, BAB I mulai 1.</span><span><b>Proposal:</b> bisa tanpa BAB; bagian A/B/C menjadi struktur utama.</span><span><b>Makalah:</b> cover tanpa nomor; Kata Pengantar/TOC Romawi bila ada.</span></div>
    </div>
    <div class="preset-row">
      <label class="field wide"><span>Mode</span><select id="mode"><option value="safe" ${sel('safe',o.mode)}>Aman / konservatif</option><option value="total" ${sel('total',o.mode)}>Total / lebih agresif</option></select></label>
      <label class="field wide"><span>Preset margin</span><select id="preset"><option value="uin" selected>UIN Madura 4-4-3-3</option><option value="zain">ZAIN.NET 4-3-3-3</option><option value="custom">Custom</option></select></label>
      <label class="field"><span>Kertas</span><select id="paper"><option value="A4" ${sel('A4',o.paper)}>A4</option><option value="F4" ${sel('F4',o.paper)}>F4</option></select></label>
    </div>
    <div class="settings-grid margin-grid">
      <label class="field"><span>Kiri (cm)</span><input id="mLeft" type="number" step="0.1" value="${o.marginLeftCm}" /></label>
      <label class="field"><span>Atas (cm)</span><input id="mTop" type="number" step="0.1" value="${o.marginTopCm}" /></label>
      <label class="field"><span>Kanan (cm)</span><input id="mRight" type="number" step="0.1" value="${o.marginRightCm}" /></label>
      <label class="field"><span>Bawah (cm)</span><input id="mBottom" type="number" step="0.1" value="${o.marginBottomCm}" /></label>
    </div>
    <div class="module-grid">${modules.map((m, i) => moduleToggle(i + 1, m[0], m[1], m[2], m[3])).join('')}</div>
    <details class="advanced"><summary>Pengaturan lanjutan</summary>
      <div class="settings-grid advanced-grid">
        <label class="field"><span>Font isi</span><input id="fontFamily" value="${esc(o.fontFamily)}" /></label>
        <label class="field"><span>Ukuran</span><input id="fontSize" type="number" step="0.5" value="${o.fontSizePt}" /></label>
        <label class="field"><span>Spasi</span><select id="lineSpacing"><option value="1.5" ${sel(1.5,o.lineSpacing)}>1.5</option><option value="2" ${sel(2,o.lineSpacing)}>2.0</option><option value="1" ${sel(1,o.lineSpacing)}>1.0</option></select></label>
        <label class="field"><span>First line (cm)</span><input id="firstLine" type="number" step="0.01" value="${o.firstLineIndentCm}" /></label>
      </div>
      <div class="mini-toggles">
        ${smallToggle('includeFrontMatterInToc', 'Masukkan bagian awal yang relevan ke TOC', o.includeFrontMatterInToc)}
        ${smallToggle('removeLikelyManualToc', 'Ganti daftar isi manual yang yakin terdeteksi saat membuat TOC native', o.removeLikelyManualToc)}
      </div>
    </details>`;
}
function moduleToggle(n, id, title, sub, checked) {
    return `<label class="module-card"><input id="${id}" class="module-check" type="checkbox" ${checked ? 'checked' : ''}/><span class="module-no">${String(n).padStart(2, '0')}</span><span class="module-copy"><strong>${title}</strong><small>${sub}</small></span><span class="checkmark">${icon('check')}</span></label>`;
}
function smallToggle(id, text, checked) {
    return `<label class="mini-toggle"><input id="${id}" type="checkbox" ${checked ? 'checked' : ''}><span class="switch"></span><span>${text}</span></label>`;
}
function headingMapHtml() {
    if (!currentAnalysis)
        return '';
    const list = currentAnalysis.headingCandidates;
    if (!list.length)
        return `<div class="empty-ok">Tidak ada kandidat heading yang cukup meyakinkan. Periksa file Word secara manual sebelum membuat TOC.</div>`;
    return `<details class="heading-audit" open>
    <summary><span>${icon('layers')}</span><strong>Pemeriksa Level TOC</strong><em>${list.length} kandidat</em></summary>
    <div class="audit-help">Pemeriksa TOC V4 menyesuaikan jenis dokumen. Untuk Skripsi/Dengan BAB: BAB = Level 1, Sub BAB = Level 2. Untuk Proposal/Makalah Tanpa BAB: bagian A/B/C = Level 1, subbagian angka = Level 2. Ubah jika deteksi keliru.</div>
    <div class="heading-table">
      <div class="heading-row heading-head"><span>Judul terdeteksi</span><span>Keyakinan</span><span>Level TOC</span></div>
      ${list.map((h) => `<div class="heading-row">
        <span class="heading-title"><b>${esc(h.text)}</b><small>${esc(h.reason)}${h.chapter ? ` • BAB ${esc(h.chapter)}` : ''}</small></span>
        <span><i class="confidence ${h.confidence >= .9 ? 'high' : h.confidence >= .78 ? 'medium' : 'low'}">${Math.round(h.confidence * 100)}%</i></span>
        <span><select class="heading-select" data-heading-index="${h.paragraphIndex}">
          <option value="0">Jangan masuk TOC</option>
          <option value="1" ${h.level === 1 ? 'selected' : ''}>Level 1 — BAB/Bagian utama</option>
          <option value="2" ${h.level === 2 ? 'selected' : ''}>Level 2 — Subjudul</option>
          <option value="3" ${h.level === 3 ? 'selected' : ''}>Level 3 — Sub-subjudul</option>
        </select></span>
      </div>`).join('')}
    </div>
  </details>`;
}
function analysisHtml() {
    if (!currentAnalysis) return '';
    const a = currentAnalysis;
    const typeLabel = {skripsi:'Skripsi',proposal:'Proposal',makalah:'Makalah'}[a.structure.documentType] || a.structure.documentType;
    const structLabel = a.structure.structureMode === 'with-bab' ? 'Dengan BAB' : 'Tanpa BAB';
    const issueCards = a.issues.length ? a.issues.map((i) => `<div class="issue ${i.severity}"><div>${icon(i.severity === 'info' ? 'check' : 'warn')}</div><span><strong>${esc(i.title)}</strong><small>${esc(i.detail)}</small></span>${i.count ? `<b>${i.count}</b>` : ''}</div>`).join('') : '<div class="empty-ok">✓ Tidak ada masalah utama yang terdeteksi.</div>';
    return `<section class="analysis-section">
    <div class="section-title"><div><span class="step">03</span><h2>Analisis & validasi struktur</h2></div><span class="score ${a.validationScore >= 85 ? 'good' : a.validationScore >= 65 ? 'mid' : 'bad'}">${a.validationScore}/100</span></div>
    <div class="detected-profile"><span>Jenis terdeteksi</span><strong>${typeLabel} • ${structLabel}</strong><small>${Math.round((a.structure.detectionConfidence || 0)*100)}% — ${esc(a.structure.detectionReason || '')}</small></div>
    <div class="stats seven">
      <div><span>Section</span><strong>${a.sections}</strong></div><div><span>BAB</span><strong>${a.chapters.length}</strong></div><div><span>Level 2</span><strong>${a.subchapters}</strong></div><div><span>Level 3</span><strong>${a.subSubchapters}</strong></div><div><span>Tabel</span><strong>${a.structure.tableCaptions}</strong></div><div><span>Gambar</span><strong>${a.structure.figureCaptions}</strong></div><div><span>TOC</span><strong>${a.hasNativeToc ? 'Ada' : 'Belum'}</strong></div>
    </div>
    <div class="structure-line"><span>Awal isi: <b>${esc((a.structure.mainStartText || 'Belum terdeteksi').slice(0,70))}</b></span><span>Bagian awal: <b>${a.structure.frontMatter.length}</b></span><span>Daftar pustaka: <b>${a.structure.bibliographyFound ? 'Ada' : 'Belum'}</b></span><span>Enter berlebih: <b>${a.structure.likelyBlankParagraphs}</b></span></div>
    <div class="issues">${issueCards}</div>
    ${headingMapHtml()}
    <div class="action-row">
      <button class="btn primary large" id="repairBtn">${icon('magic')} PROSES FILE AKTIF</button>
      ${currentFiles.length > 1 ? `<button class="btn secondary large" id="batchBtn">${icon('layers')} PROSES ${currentFiles.length} FILE → ZIP</button>` : ''}
    </div>
    <p class="safety-note"><b>Mode Aman direkomendasikan.</b> V4 menjalankan pemeriksaan anti-rusak pada setiap tahap. Jika teks utama, tabel, drawing atau aset DOCX terdeteksi hilang, tahap dibatalkan atau hasil diblokir.</p>
  </section>`;
}
function resultHtml() {
    if (!currentResult && !batchBlob) return '';
    if (batchBlob && !currentResult) return `<section class="result-section"><div class="result-head"><div class="success-icon">${icon('check')}</div><div><span>BATCH SELESAI</span><h2>Semua file sudah diproses</h2><p>Hasil dibungkus menjadi satu ZIP.</p></div></div><button class="btn primary large" id="downloadBatchBtn">${icon('download')} Download ZIP Hasil Batch</button></section>`;
    const r = currentResult;
    const changeRows = r.changes.map((c) => `<li>${icon('check')}<span>${esc(c)}</span></li>`).join('');
    const warningRows = r.warnings.map((c) => `<li class="warning">${icon('warn')}<span>${esc(c)}</span></li>`).join('');
    return `<section class="result-section">
    <div class="result-head"><div class="success-icon">${icon('check')}</div><div><span>SELESAI • V4</span><h2>File hasil sudah dibuat</h2><p>${esc(r.outputName)}</p></div><div class="result-score"><small>Skor akhir</small><b>${r.analysisAfter.validationScore}/100</b></div></div>
    <div class="integrity-ok">${icon('check')} <span><strong>Anti-rusak lulus</strong><small>${r.integrity?.protectedText || 0} teks unik terlindungi • ${r.integrity?.tables || 0} tabel • ${r.integrity?.drawings || 0} drawing</small></span></div>
    <div class="compare"><div><span>SEBELUM</span><strong>${currentAnalysis?.issues.length || 0} catatan</strong><small>Level 2 ${currentAnalysis?.subchapters || 0} • Level 3 ${currentAnalysis?.subSubchapters || 0}</small></div><div class="arrow">→</div><div><span>SESUDAH</span><strong>${r.analysisAfter.issues.length} catatan</strong><small>Level 2 ${r.analysisAfter.subchapters} • Level 3 ${r.analysisAfter.subSubchapters}</small></div></div>
    <ul class="change-list">${changeRows}${warningRows}</ul>
    <div class="download-row"><button class="btn primary" id="downloadBtn">${icon('download')} Download DOCX Hasil</button><button class="btn secondary" id="downloadOriginalBtn">Download File Asli</button></div>
    <div class="word-note"><strong>Setelah dibuka di Microsoft Word:</strong> tekan <b>Ctrl+A → F9</b>. Untuk Daftar Isi pilih <b>Update Entire Table</b>. Nomor halaman aktual TOC dihitung oleh mesin layout Microsoft Word.</div>
  </section>`;
}
function headingOverrides() {
    const out = {};
    document.querySelectorAll('.heading-select').forEach((s) => { out[s.dataset.headingIndex || ''] = Number(s.value); });
    return out;
}
function readOptions(withOverrides = true) {
    const b = (id) => Boolean(document.querySelector(`#${id}`)?.checked);
    const dom = document.querySelector('#documentType');
    if (!dom) return { ...uiOptions, headingOverrides: withOverrides ? (uiOptions.headingOverrides || {}) : {} };
    return {
        documentType: document.querySelector('#documentType')?.value || 'auto',
        structureMode: document.querySelector('#structureMode')?.value || 'auto',
        numberingProfile: document.querySelector('#numberingProfile')?.value || 'uin-madura',
        mode: document.querySelector('#mode')?.value || 'safe', paper: document.querySelector('#paper')?.value || 'A4',
        marginLeftCm: Number(document.querySelector('#mLeft')?.value || 4), marginTopCm: Number(document.querySelector('#mTop')?.value || 4), marginRightCm: Number(document.querySelector('#mRight')?.value || 3), marginBottomCm: Number(document.querySelector('#mBottom')?.value || 3),
        normalizeMargins: b('normalizeMargins'), repairPageNumbering: b('repairPageNumbering'), createToc: b('createToc'), repairChapterPagination: b('repairChapterPagination'), repairHeadings: b('repairHeadings'), normalizeBodyFormatting: b('normalizeBodyFormatting'), createTableList: b('createTableList'), createFigureList: b('createFigureList'), cleanBlankParagraphs: b('cleanBlankParagraphs'), validateDocument: b('validateDocument'),
        removeLikelyManualToc: b('removeLikelyManualToc'), includeFrontMatterInToc: b('includeFrontMatterInToc'),
        fontFamily: document.querySelector('#fontFamily')?.value || 'Times New Roman', fontSizePt: Number(document.querySelector('#fontSize')?.value || 12), lineSpacing: Number(document.querySelector('#lineSpacing')?.value || 1.5), firstLineIndentCm: Number(document.querySelector('#firstLine')?.value || 1.27),
        headingOverrides: withOverrides ? headingOverrides() : {},
    };
}
function captureOptions() {
    const current = readOptions(false);
    uiOptions = { ...uiOptions, ...current, headingOverrides: uiOptions.headingOverrides || {} };
    return uiOptions;
}
function setModuleSelection(kind) {
    document.querySelectorAll('.module-check').forEach((x, i) => { x.checked = kind === 'all' ? true : kind === 'none' ? false : [0,1,2,3,4,8,9].includes(i); });
    if (kind === 'all') {
        const mode = document.querySelector('#mode'); if (mode) mode.value = 'safe';
    }
    captureOptions();
}
function applyPreset(value) {
    const set = (id, v) => { const el = document.querySelector(`#${id}`); if (el) el.value = String(v); };
    if (value === 'uin') { set('mLeft', 4); set('mTop', 4); set('mRight', 3); set('mBottom', 3); }
    if (value === 'zain') { set('mLeft', 4); set('mTop', 3); set('mRight', 3); set('mBottom', 3); }
    captureOptions();
}
function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
}
function setLoading(button, loading, text) {
    if (!button)
        return;
    if (loading) {
        button.dataset.old = button.innerHTML;
        button.disabled = true;
        button.innerHTML = `<span class="spinner"></span>${text}`;
    }
    else {
        button.disabled = false;
        button.innerHTML = button.dataset.old || button.innerHTML;
    }
}
function bindEvents() {
    const input = document.querySelector('#fileInput');
    const drop = document.querySelector('#dropzone');
    input?.addEventListener('change', async () => { if (input.files?.length) await selectFiles(Array.from(input.files)); });
    drop?.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('drag'); });
    drop?.addEventListener('dragleave', () => drop.classList.remove('drag'));
    drop?.addEventListener('drop', async (e) => { e.preventDefault(); drop.classList.remove('drag'); if (e.dataTransfer?.files?.length) await selectFiles(Array.from(e.dataTransfer.files)); });
    document.querySelectorAll('.file-tab').forEach((b) => b.addEventListener('click', async () => { captureOptions(); currentFileIndex = Number(b.dataset.fileIndex || 0); currentAnalysis = null; currentResult = null; batchBlob = null; render(); await analyzeCurrent(); }));
    document.querySelector('#analyzeBtn')?.addEventListener('click', analyzeCurrent);
    document.querySelector('#repairBtn')?.addEventListener('click', repairCurrent);
    document.querySelector('#batchBtn')?.addEventListener('click', repairBatch);
    document.querySelector('#downloadBtn')?.addEventListener('click', () => currentResult && downloadBlob(currentResult.blob, currentResult.outputName));
    document.querySelector('#downloadOriginalBtn')?.addEventListener('click', () => currentFile() && downloadBlob(currentFile(), currentFile().name));
    document.querySelector('#downloadBatchBtn')?.addEventListener('click', () => batchBlob && downloadBlob(batchBlob, 'ZAINNET_HASIL_RAPI_V4_BATCH.zip'));
    document.querySelector('#coreBtn')?.addEventListener('click', () => setModuleSelection('core'));
    document.querySelector('#allBtn')?.addEventListener('click', () => setModuleSelection('all'));
    document.querySelector('#noneBtn')?.addEventListener('click', () => setModuleSelection('none'));
    document.querySelector('#preset')?.addEventListener('change', (e) => applyPreset(e.target.value));
    document.querySelectorAll('#mode,#paper,#numberingProfile,#mLeft,#mTop,#mRight,#mBottom,#fontFamily,#fontSize,#lineSpacing,#firstLine,.module-check,#includeFrontMatterInToc,#removeLikelyManualToc').forEach((el) => el.addEventListener('change', captureOptions));
    document.querySelectorAll('#documentType,#structureMode').forEach((el) => el.addEventListener('change', async () => { captureOptions(); currentAnalysis = null; currentResult = null; await analyzeCurrent(); }));
}
async function selectFiles(files) {
    const valid = files.filter((f) => /\.docx$/i.test(f.name));
    if (!valid.length)
        return alert('Pilih file Microsoft Word .DOCX.');
    currentFiles = valid;
    currentFileIndex = 0;
    currentAnalysis = null;
    currentResult = null;
    batchBlob = null;
    render();
    await analyzeCurrent();
}
async function analyzeCurrent() {
    const file = currentFile(); if (!file) return;
    const btn = document.querySelector('#analyzeBtn');
    setLoading(btn, true, 'Mendeteksi jenis dokumen & struktur...');
    try {
        captureOptions();
        currentAnalysis = await analyzeDocx(file, uiOptions);
        currentResult = null; batchBlob = null; render();
    } catch (e) {
        setLoading(btn, false, 'Analisis File Aktif');
        alert(e instanceof Error ? e.message : 'Gagal menganalisis DOCX.');
    }
}
async function repairCurrent() {
    const file = currentFile(); if (!file) return;
    const btn = document.querySelector('#repairBtn');
    setLoading(btn, true, 'Memproses V4 + pemeriksaan anti-rusak...');
    try {
        captureOptions();
        currentResult = await repairDocx(file, { ...uiOptions, headingOverrides: headingOverrides() });
        batchBlob = null; render();
        document.querySelector('.result-section')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (e) {
        setLoading(btn, false, 'PROSES FILE AKTIF');
        alert(e instanceof Error ? e.message : 'Gagal merapikan DOCX.');
    }
}
async function repairBatch() {
    if (currentFiles.length < 2) return;
    const btn = document.querySelector('#batchBtn');
    setLoading(btn, true, `Memproses 1/${currentFiles.length}...`);
    try {
        captureOptions();
        const outZip = new JSZip();
        const baseOptions = { ...uiOptions, headingOverrides: {} };
        for (let i = 0; i < currentFiles.length; i++) {
            if (btn) btn.innerHTML = `<span class="spinner"></span>Memproses ${i + 1}/${currentFiles.length}: ${esc(currentFiles[i].name)}`;
            const result = await repairDocx(currentFiles[i], baseOptions);
            outZip.file(result.outputName, result.blob);
        }
        batchBlob = await outZip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
        currentResult = null; render(); document.querySelector('.result-section')?.scrollIntoView({ behavior: 'smooth' });
    } catch (e) {
        setLoading(btn, false, 'PROSES BATCH');
        alert(e instanceof Error ? e.message : 'Gagal memproses batch.');
    }
}
render();
