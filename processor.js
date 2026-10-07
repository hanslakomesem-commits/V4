const JSZip = globalThis.JSZip;
if (!JSZip) throw new Error('JSZip gagal dimuat. Pastikan jszip.min.js ikut di-upload.');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const twipsPerCm = 1440 / 2.54;
const cmToTwips = (cm) => String(Math.round(cm * twipsPerCm));
const twipsToCm = (twips) => Number((Number(twips || '0') / twipsPerCm).toFixed(2));
const parser = new DOMParser();
const serializer = new XMLSerializer();
const FRONT_MATTER_RE = /^(DAFTAR\s+(ISI|TABEL|GAMBAR|LAMPIRAN)|ABSTRAK|ABSTRACT|KATA\s+PENGANTAR|MOTTO|PERSEMBAHAN|(?:LEMBAR|HALAMAN)\s+(JUDUL|SAMPUL|PENGESAHAN|PERSETUJUAN|PERNYATAAN)|PERNYATAAN\s+(?:KEASLIAN(?:\s+TULISAN)?|ORISINALITAS)|SURAT\s+PERNYATAAN)$/i;
const FRONT_PAGE_START_RE = /^(DAFTAR\s+(ISI|TABEL|GAMBAR|LAMPIRAN)|ABSTRAK|ABSTRACT|KATA\s+PENGANTAR|MOTTO|PERSEMBAHAN|(?:LEMBAR|HALAMAN)\s+(JUDUL|SAMPUL|PENGESAHAN|PERSETUJUAN|PERNYATAAN)|PERNYATAAN\s+(?:KEASLIAN(?:\s+TULISAN)?|ORISINALITAS)|SURAT\s+PERNYATAAN)$/i;
const END_MATTER_RE = /^(DAFTAR\s+PUSTAKA|REFERENSI|BIBLIOGRAFI|LAMPIRAN(?:\s*[-–—]?\s*LAMPIRAN)?|RIWAYAT\s+HIDUP)$/i;
const KNOWN_LEVEL2_RE = /^(KONTEKS|LATAR\s+BELAKANG|RUMUSAN\s+MASALAH|FOKUS\s+PENELITIAN|TUJUAN\s+PENELITIAN|KEGUNAAN\s+PENELITIAN|MANFAAT\s+PENELITIAN|DEFINISI\s+ISTILAH|KAJIAN\s+PENELITIAN\s+TERDAHULU|PENDEKATAN\s+DAN\s+JENIS\s+PENELITIAN|KEHADIRAN\s+PENELITI|LOKASI\s+PENELITIAN|SUMBER\s+DATA|PROSEDUR\s+PENGUMPULAN\s+DATA|TEKNIK\s+PENGUMPULAN\s+DATA|ANALISIS\s+DATA|PENGECEKAN\s+KEABSAHAN\s+DATA|PEGECEKAN\s+KEABSAHAN\s+DATA|SUMBER\s+DATA\s+(PRIMER|SEKUNDER)|OBSERVASI|WAWANCARA|DOKUMENTASI|ETIKA\s+(DEONTOLOGI|TELEOLOGI)|PERPANJANGAN\s+KEIKUTSERTAAN|KETEKUNAN\s+PENGAMATAN|TRIANGULASI|TAHAP\s+PRA\s+PENELITIAN|TAHAP\s+PROSES\s+PENELITIAN|PENYUSUNAN\s+LAPORAN|TAHAP(?:-TAHAP|\s+TAHAP)\s+PENELITIAN|PAPARAN\s+DATA(?:\s+DAN\s+TEMUAN\s+PENELITIAN)?|TEMUAN\s+PENELITIAN|PEMBAHASAN|KESIMPULAN|SARAN)$/i;
const LEVEL3_HINT_RE = /^(PENGERTIAN|DUA\s+TEORI|JENIS|MACAM|PRINSIP|FAKTOR|TUJUAN|CIRI|DASAR|TAHAPAN|PROSES|METODE|AKAD|PEMBIAYAAN)\b/i;
const PROPOSAL_TOP_RE = /^(?:[A-P]\.?\s+)?(JUDUL(?:\s+PENELITIAN)?|KONTEKS\s+PENELITIAN|LATAR\s+BELAKANG(?:\s+MASALAH)?|FOKUS\s+PENELITIAN|RUMUSAN\s+MASALAH|TUJUAN(?:\s+PENELITIAN|\s+PENULISAN|\s+KEGIATAN)?|MANFAAT\s+PENELITIAN|KEGUNAAN\s+PENELITIAN|DEFINISI\s+(?:ISTILAH|OPERASIONAL)|PENELITIAN\s+TERDAHULU|KAJIAN\s+PENELITIAN\s+TERDAHULU|KERANGKA\s+TEORI|KAJIAN\s+TEORI|METODE\s+PENELITIAN|SISTEMATIKA\s+PENULISAN|TIME\s+SCHEDULE\s+PENELITIAN|JADWAL\s+(?:PENELITIAN|KEGIATAN)|DASAR\s+KEGIATAN|NAMA\s+KEGIATAN|TEMA\s+KEGIATAN|BENTUK\s+KEGIATAN|SASARAN(?:\s+KEGIATAN)?|PESERTA(?:\s+KEGIATAN)?|WAKTU\s+DAN\s+TEMPAT|SUSUNAN\s+(?:ACARA|PANITIA|KEPANITIAAN)|RENCANA\s+ANGGARAN(?:\s+BIAYA)?|ANGGARAN\s+BIAYA|RANCANGAN\s+ANGGARAN\s+BIAYA(?:\s*\(RAB\))?|PENUTUP|DAFTAR\s+PUSTAKA|LAMPIRAN(?:-LAMPIRAN)?)$/i;
const MAKALAH_TOP_RE = /^(?:[A-H]\.?\s+)?(JUDUL(?:\s+MAKALAH)?|LATAR\s+BELAKANG(?:\s+MASALAH)?|RUMUSAN(?:\/FOKUS)?\s+MASALAH|FOKUS\s+MASALAH|TUJUAN(?:\s+PENULISAN|\s+PEMBAHASAN)?|PENDAHULUAN|PEMBAHASAN|KESIMPULAN|PENUTUP|SARAN|DAFTAR\s+PUSTAKA)$/i;
const EARLY_PROPOSAL_RE = /\bPROPOSAL(?:\s+SKRIPSI|\s+PENELITIAN|\s+KEGIATAN)?\b/i;
const EARLY_MAKALAH_RE = /\bMAKALAH\b/i;

function parseXml(xml) {
    const doc = parser.parseFromString(xml, 'application/xml');
    const error = doc.querySelector('parsererror');
    if (error)
        throw new Error(`XML DOCX tidak valid: ${error.textContent || 'parsererror'}`);
    return doc;
}
function xmlString(doc) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${serializer.serializeToString(doc.documentElement)}`;
}
function wEl(doc, local) {
    return doc.createElementNS(W, `w:${local}`);
}
function wAttr(el, local, value) { el.setAttributeNS(W, `w:${local}`, value); }
function rAttr(el, local, value) { el.setAttributeNS(R, `r:${local}`, value); }
function directChild(parent, local) {
    return Array.from(parent.children).find((c) => c.namespaceURI === W && c.localName === local) || null;
}
function directChildren(parent, local) {
    return Array.from(parent.children).filter((c) => c.namespaceURI === W && c.localName === local);
}
function ensureDirectChild(parent, local, beforeLocal) {
    const existing = directChild(parent, local);
    if (existing)
        return existing;
    const el = wEl(parent.ownerDocument, local);
    if (beforeLocal) {
        const before = directChild(parent, beforeLocal);
        if (before)
            parent.insertBefore(el, before);
        else
            parent.appendChild(el);
    }
    else
        parent.appendChild(el);
    return el;
}
const PPR_ORDER = ['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange'];
function ensurePPrChild(pPr, local) {
    let el = directChild(pPr, local);
    if (!el) el = wEl(pPr.ownerDocument, local);
    else el.remove();
    const rank = PPR_ORDER.indexOf(local);
    const before = Array.from(pPr.children).find((c) => {
        if (c.namespaceURI !== W) return false;
        const cr = PPR_ORDER.indexOf(c.localName);
        return rank >= 0 && cr >= 0 && cr > rank;
    });
    if (before) pPr.insertBefore(el, before); else pPr.appendChild(el);
    return el;
}
function ensureSectChild(sectPr, local, beforeNames) {
    const existing = directChild(sectPr, local);
    if (existing)
        return existing;
    const el = wEl(sectPr.ownerDocument, local);
    const before = Array.from(sectPr.children).find((c) => beforeNames.includes(c.localName));
    if (before)
        sectPr.insertBefore(el, before);
    else
        sectPr.appendChild(el);
    return el;
}
function normalizeText(s) { return s.replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim(); }
function paragraphText(p) {
    return normalizeText(Array.from(p.getElementsByTagNameNS(W, 't')).map((n) => n.textContent || '').join(''));
}
function getBody(doc) {
    const body = doc.getElementsByTagNameNS(W, 'body')[0];
    if (!body)
        throw new Error('word/document.xml tidak memiliki w:body.');
    return body;
}
function bodyParagraphs(doc) {
    return Array.from(getBody(doc).children).filter((e) => e.namespaceURI === W && e.localName === 'p');
}
function bodySectionProperties(doc) {
    const body = getBody(doc);
    const out = [];
    for (const child of Array.from(body.children)) {
        if (child.namespaceURI === W && child.localName === 'p') {
            const pPr = directChild(child, 'pPr');
            const sectPr = pPr && directChild(pPr, 'sectPr');
            if (sectPr)
                out.push(sectPr);
        }
    }
    const finalSect = directChild(body, 'sectPr');
    if (finalSect)
        out.push(finalSect);
    return out;
}
function chapterMatch(text) {
    const t = normalizeText(text).toUpperCase();
    return t.match(/^BAB\s*(V|IV|III|II|I)(?=\s|[A-ZÀ-ÖØ-Ý]|$)/);
}
function isChapterText(text) { return Boolean(chapterMatch(text)); }
function chapterRoman(text) { return chapterMatch(text)?.[1] || null; }
function isTableCaption(text) { return /^TABEL\s+(?:[IVXLC]+|\d+)(?:[.\-]\d+)*(?:\s|[.:])/i.test(normalizeText(text)); }
function isFigureCaption(text) { return /^GAMBAR\s+(?:[IVXLC]+|\d+)(?:[.\-]\d+)*(?:\s|[.:])/i.test(normalizeText(text)); }
function getOutlineLevel(p) {
    const pPr = directChild(p, 'pPr');
    const outline = pPr && directChild(pPr, 'outlineLvl');
    const val = Number(outline?.getAttributeNS(W, 'val'));
    if (Number.isFinite(val) && val >= 0 && val <= 2)
        return (val + 1);
    return undefined;
}
function hasExplicitBodyOutline(p) {
    const pPr = directChild(p, 'pPr');
    const outline = pPr && directChild(pPr, 'outlineLvl');
    const val = Number(outline?.getAttributeNS(W, 'val'));
    return Number.isFinite(val) && val >= 3;
}
function paragraphStyle(p) {
    const pPr = directChild(p, 'pPr');
    const style = pPr && directChild(pPr, 'pStyle');
    return style?.getAttributeNS(W, 'val') || '';
}
function styleHeadingLevel(p) {
    const s = paragraphStyle(p);
    const m = s.match(/^Heading([123])$/i) || s.match(/^Judul([123])$/i);
    if (!m)
        return undefined;
    return Number(m[1]);
}
function paragraphBoldRatio(p) {
    let total = 0;
    let bold = 0;
    for (const r of Array.from(p.getElementsByTagNameNS(W, 'r'))) {
        const text = Array.from(r.getElementsByTagNameNS(W, 't')).map((x) => x.textContent || '').join('');
        const n = text.length;
        total += n;
        const rPr = directChild(r, 'rPr');
        const b = rPr && directChild(rPr, 'b');
        if (b && b.getAttributeNS(W, 'val') !== '0' && b.getAttributeNS(W, 'val') !== 'false')
            bold += n;
    }
    return total ? bold / total : 0;
}
function paragraphIndentTwips(p) {
    const pPr = directChild(p, 'pPr');
    const ind = pPr && directChild(pPr, 'ind');
    return {
        left: Number(ind?.getAttributeNS(W, 'left') || ind?.getAttributeNS(W, 'start') || 0),
        first: Number(ind?.getAttributeNS(W, 'firstLine') || 0),
    };
}
function paragraphNumInfo(p, numbering) {
    const pPr = directChild(p, 'pPr');
    const numPr = pPr && directChild(pPr, 'numPr');
    if (!numPr)
        return null;
    const numId = directChild(numPr, 'numId')?.getAttributeNS(W, 'val') || '';
    const ilvl = directChild(numPr, 'ilvl')?.getAttributeNS(W, 'val') || '0';
    const lvl = numbering.get(numId)?.get(ilvl);
    return { numId, ilvl, fmt: lvl?.fmt || '', marker: lvl?.text || '' };
}
function hasComplexContent(p) {
    return p.getElementsByTagNameNS(W, 'drawing').length > 0 || p.getElementsByTagNameNS(W, 'object').length > 0 || p.getElementsByTagNameNS(W, 'pict').length > 0;
}
function isLikelyTocParagraph(p) {
    const style = paragraphStyle(p);
    if (/^TOC\d+$/i.test(style))
        return true;
    const text = paragraphText(p);
    const hasTab = p.getElementsByTagNameNS(W, 'tab').length > 0;
    return (hasTab && /\d+\s*$/.test(text)) || /\.{3,}\s*(?:\d+|[ivxlcdm]+)\s*$/i.test(text);
}
function likelySentence(text) {
    const t = normalizeText(text);
    return t.length > 115 || (/[,;:]\s/.test(t) && t.length > 70) || (/[.!?]$/.test(t) && t.split(/\s+/).length > 10);
}
async function loadNumberingMap(zip) {
    const out = new Map();
    const file = zip.file('word/numbering.xml');
    if (!file)
        return out;
    const doc = parseXml(await file.async('string'));
    const abstracts = new Map();
    for (const abs of Array.from(doc.getElementsByTagNameNS(W, 'abstractNum'))) {
        const id = abs.getAttributeNS(W, 'abstractNumId') || '';
        const levels = new Map();
        for (const lvl of directChildren(abs, 'lvl')) {
            const ilvl = lvl.getAttributeNS(W, 'ilvl') || '0';
            levels.set(ilvl, {
                fmt: directChild(lvl, 'numFmt')?.getAttributeNS(W, 'val') || '',
                text: directChild(lvl, 'lvlText')?.getAttributeNS(W, 'val') || '',
            });
        }
        abstracts.set(id, levels);
    }
    for (const num of Array.from(doc.getElementsByTagNameNS(W, 'num'))) {
        const numId = num.getAttributeNS(W, 'numId') || '';
        const absId = directChild(num, 'abstractNumId')?.getAttributeNS(W, 'val') || '';
        const levels = abstracts.get(absId);
        if (levels)
            out.set(numId, levels);
    }
    return out;
}
function looksLikeManualTocEntry(p) {
    const text = paragraphText(p);
    // Judul BAB asli seperti “BAB I” berakhir dengan angka Romawi tetapi bukan entri TOC.
    if (isChapterText(text) || FRONT_MATTER_RE.test(text) || END_MATTER_RE.test(text)) return false;
    if (isLikelyTocParagraph(p)) return true;
    // Nomor halaman harus menjadi token akhir yang didahului leader/titik atau spasi; hindari kata biasa yang kebetulan berakhir huruf Romawi.
    return /(?:\.{2,}|\t|\s{2,})(?:\d+|[ivxlcdm]+)\s*$/i.test(text) && text.length < 220;
}
function detectDocumentType(doc) {
    const ps = bodyParagraphs(doc);
    const early = ps.slice(0, Math.min(ps.length, 45)).map(paragraphText).filter(Boolean);
    if (early.some((t) => EARLY_PROPOSAL_RE.test(t))) return { type: 'proposal', confidence: 0.98, reason: 'Kata PROPOSAL terdeteksi pada bagian awal dokumen.' };
    if (early.some((t) => EARLY_MAKALAH_RE.test(t))) return { type: 'makalah', confidence: 0.98, reason: 'Kata MAKALAH terdeteksi pada bagian awal dokumen.' };
    const chapters = ps.filter((p) => isChapterText(paragraphText(p)) && !looksLikeManualTocEntry(p)).length;
    if (chapters >= 4) return { type: 'skripsi', confidence: 0.95, reason: `${chapters} BAB utama terdeteksi.` };
    const proposalHeads = ps.filter((p) => PROPOSAL_TOP_RE.test(paragraphText(p)) && !looksLikeManualTocEntry(p)).length;
    const makalahHeads = ps.filter((p) => MAKALAH_TOP_RE.test(paragraphText(p)) && !looksLikeManualTocEntry(p)).length;
    if (proposalHeads >= 5) return { type: 'proposal', confidence: 0.82, reason: `${proposalHeads} bagian proposal umum terdeteksi.` };
    if (makalahHeads >= 3) return { type: 'makalah', confidence: 0.8, reason: `${makalahHeads} bagian makalah umum terdeteksi.` };
    if (chapters >= 2) return { type: 'proposal', confidence: 0.62, reason: `${chapters} BAB terdeteksi; kemungkinan proposal berbasis BAB.` };
    return { type: 'skripsi', confidence: 0.45, reason: 'Jenis dokumen belum meyakinkan; default aman menggunakan profil skripsi sampai pengguna memilih jenis dokumen.' };
}
function resolveProfile(doc, options = {}) {
    const auto = detectDocumentType(doc);
    const type = options.documentType && options.documentType !== 'auto' ? options.documentType : auto.type;
    const hasBab = bodyParagraphs(doc).some((p) => isChapterText(paragraphText(p)) && !looksLikeManualTocEntry(p));
    let structure = options.structureMode && options.structureMode !== 'auto' ? options.structureMode : (hasBab ? 'with-bab' : 'without-bab');
    if (type === 'skripsi') structure = 'with-bab';
    return { type, structure, confidence: options.documentType && options.documentType !== 'auto' ? 1 : auto.confidence, reason: options.documentType && options.documentType !== 'auto' ? 'Jenis dokumen dipilih pengguna.' : auto.reason };
}
function firstMainParagraphIndex(doc, options = {}) {
    const ps = bodyParagraphs(doc);
    const profile = resolveProfile(doc, options);
    if (profile.structure === 'with-bab') {
        const idx = ps.findIndex((p) => isChapterText(paragraphText(p)) && !looksLikeManualTocEntry(p));
        if (idx >= 0) return idx;
    }
    const tocIdx = ps.findIndex((p) => /^DAFTAR\s+ISI$/i.test(paragraphText(p)));
    for (let i = 0; i < ps.length; i++) {
        if (i <= tocIdx) continue;
        const p = ps[i];
        const text = paragraphText(p);
        if (!text || looksLikeManualTocEntry(p) || FRONT_MATTER_RE.test(text) || hasComplexContent(p)) continue;
        const letter = /^[A-P]\.?\s+\S+/i.test(text);
        const topKnown = profile.type === 'proposal' ? PROPOSAL_TOP_RE.test(text) : profile.type === 'makalah' ? MAKALAH_TOP_RE.test(text) : false;
        if (letter || topKnown) return i;
    }
    // fallback: paragraf isi pertama setelah bagian awal/TOC.
    for (let i = Math.max(0, tocIdx + 1); i < ps.length; i++) {
        const text = paragraphText(ps[i]);
        if (text && !FRONT_MATTER_RE.test(text) && !looksLikeManualTocEntry(ps[i])) return i;
    }
    return -1;
}
function firstMainParagraph(doc, options = {}) {
    const idx = firstMainParagraphIndex(doc, options);
    return idx >= 0 ? bodyParagraphs(doc)[idx] : null;
}
function isTopLevelNoBabHeading(text, profile) {
    const t = normalizeText(text);
    if (/^[A-P]\.?\s+\S+/i.test(t)) return true;
    return profile.type === 'proposal' ? PROPOSAL_TOP_RE.test(t) : profile.type === 'makalah' ? MAKALAH_TOP_RE.test(t) : false;
}
function isMainEndHeading(text) { return /^(DAFTAR\s+PUSTAKA|REFERENSI|BIBLIOGRAFI)$/i.test(normalizeText(text)); }
function detectHeadingCandidates(doc, numbering, options = {}) {
    const ps = bodyParagraphs(doc);
    const profile = resolveProfile(doc, options);
    const mainStart = firstMainParagraphIndex(doc, options);
    const out = [];
    let inMain = false;
    let currentChapter = '';
    let afterBibliography = false;
    const tocStart = ps.findIndex((p) => /^DAFTAR\s+ISI$/i.test(paragraphText(p)));
    for (let i = 0; i < ps.length; i++) {
        const p = ps[i];
        const text = paragraphText(p);
        if (!text || looksLikeManualTocEntry(p) || hasComplexContent(p)) continue;
        if (isTableCaption(text) || isFigureCaption(text)) continue;
        if (i > 0 && (isTableCaption(paragraphText(ps[i - 1])) || isFigureCaption(paragraphText(ps[i - 1]))) && text.length < 120) continue;
        const existing = getOutlineLevel(p) || styleHeadingLevel(p);
        if (!existing && hasExplicitBodyOutline(p)) continue;
        const chapter = chapterRoman(text);
        const majorFront = FRONT_MATTER_RE.test(text);
        const majorEnd = END_MATTER_RE.test(text);
        if (i === mainStart || chapter) inMain = true;
        if (chapter) { afterBibliography = false; currentChapter = chapter; }
        if (isMainEndHeading(text)) afterBibliography = true;
        let level = 0, confidence = 0, reason = '';
        if (existing) {
            level = existing; confidence = 1; reason = `Sudah memiliki Outline/Heading level ${existing}`;
        } else if (chapter) {
            level = 1; confidence = 1; reason = `Terdeteksi BAB ${chapter}`;
        } else if (majorFront || majorEnd) {
            level = 1; confidence = 0.98; reason = 'Bagian utama/bagian awal dokumen akademik';
        } else if (!inMain || (afterBibliography && !majorEnd)) {
            continue;
        } else if (profile.structure === 'without-bab' && isTopLevelNoBabHeading(text, profile)) {
            level = 1; confidence = /^[A-P]\.?\s+/i.test(text) ? 0.98 : 0.9; reason = `${profile.type === 'proposal' ? 'Bagian proposal' : 'Bagian makalah'} tanpa BAB`;
        } else {
            const previousText = i > 0 ? paragraphText(ps[i - 1]) : '';
            if (profile.structure === 'with-bab' && isChapterText(previousText) && text.length < 100 && text === text.toUpperCase()) continue;
            const num = paragraphNumInfo(p, numbering);
            const ind = paragraphIndentTwips(p);
            const bold = paragraphBoldRatio(p);
            const short = text.length <= 160;
            const noSentenceEnd = !/[.!?]$/.test(text);
            if (profile.structure === 'with-bab' && /^[A-Z]\.[\s\t]+\S+/.test(text)) {
                level = 2; confidence = 0.97; reason = 'Pola Sub BAB huruf kapital (A., B., C.)';
            } else if (profile.structure === 'without-bab' && /^\d+(?:\.\d+){0,3}[.)]?[\s\t]+\S+/.test(text)) {
                level = 2; confidence = 0.94; reason = 'Subbagian angka pada dokumen tanpa BAB';
            } else if (profile.structure === 'without-bab' && /^[a-z][.)][\s\t]+\S+/.test(text)) {
                level = 3; confidence = 0.9; reason = 'Sub-subbagian huruf kecil';
            } else if (/^\d+(?:\.\d+){1,3}[\s\t]+\S+/.test(text)) {
                level = 3; confidence = 0.97; reason = 'Pola penomoran bertingkat';
            } else if (/^\d+[.)][\s\t]+\S+/.test(text)) {
                level = profile.structure === 'without-bab' ? 2 : 3; confidence = 0.9; reason = 'Pola judul bernomor';
            } else if (profile.structure === 'with-bab' && KNOWN_LEVEL2_RE.test(text)) {
                level = 2; confidence = 0.96; reason = 'Nama Sub BAB akademik yang umum';
            } else if (num && num.fmt === 'upperLetter' && bold >= 0.7) {
                level = profile.structure === 'without-bab' ? 1 : 2; confidence = 0.92; reason = 'Judul tebal dengan nomor list Word huruf kapital';
            } else if (num && ['lowerLetter', 'lowerRoman'].includes(num.fmt) && bold >= 0.7 && ind.left < 1000) {
                level = profile.structure === 'without-bab' ? 2 : 2; confidence = 0.88; reason = 'Judul tebal dengan penomoran list Word';
            } else if (num && ['lowerLetter', 'lowerRoman'].includes(num.fmt) && ind.left >= 1000 && bold >= 0.7 && short && !likelySentence(text)) {
                level = 3; confidence = 0.84; reason = 'Judul turunan menjorok';
            } else if (LEVEL3_HINT_RE.test(text) && short && !likelySentence(text) && bold >= 0.55) {
                level = profile.structure === 'without-bab' ? 2 : 3; confidence = 0.78; reason = 'Judul turunan akademik terdeteksi';
            } else if (bold >= 0.82 && short && noSentenceEnd && ind.first === 0 && !/^(TABEL|GAMBAR)\b/i.test(text)) {
                level = profile.structure === 'without-bab' ? 2 : 2; confidence = 0.78; reason = 'Paragraf pendek, tebal, dan bukan kalimat isi';
            }
        }
        if (level > 0) out.push({ paragraphIndex: i, text, level, confidence, reason, existingLevel: existing, chapter: currentChapter || undefined });
    }
    return out;
}
function isMainToc(doc) {
    const instr = [
        ...Array.from(doc.getElementsByTagNameNS(W, 'instrText')).map((n) => n.textContent || ''),
        ...Array.from(doc.getElementsByTagNameNS(W, 'fldSimple')).map((n) => n.getAttributeNS(W, 'instr') || ''),
    ];
    return instr.some((x) => /\bTOC\b/i.test(x) && !/\\f\s+[TG](?:\s|$)/i.test(x) && !/\\c\s+/i.test(x));
}
function hasSpecialToc(doc, mark) {
    const instr = [
        ...Array.from(doc.getElementsByTagNameNS(W, 'instrText')).map((n) => n.textContent || ''),
        ...Array.from(doc.getElementsByTagNameNS(W, 'fldSimple')).map((n) => n.getAttributeNS(W, 'instr') || ''),
    ];
    return instr.some((x) => new RegExp(`\\\\f\\s+${mark}(?:\\s|$)`, 'i').test(x));
}
function hasPageField(doc) {
    return Array.from(doc.getElementsByTagNameNS(W, 'instrText')).some((n) => /\bPAGE\b/i.test(n.textContent || '')) ||
        Array.from(doc.getElementsByTagNameNS(W, 'fldSimple')).some((n) => /\bPAGE\b/i.test(n.getAttributeNS(W, 'instr') || ''));
}
function findHeading(doc, regex) {
    return bodyParagraphs(doc).find((p) => regex.test(normalizeText(paragraphText(p)))) || null;
}
function findTocHeading(doc) { return findHeading(doc, /^DAFTAR\s+ISI$/i); }
function collectMargins(doc) {
    return bodySectionProperties(doc).map((sectPr) => {
        const pgMar = directChild(sectPr, 'pgMar');
        return {
            top: twipsToCm(pgMar?.getAttributeNS(W, 'top') || null),
            right: twipsToCm(pgMar?.getAttributeNS(W, 'right') || null),
            bottom: twipsToCm(pgMar?.getAttributeNS(W, 'bottom') || null),
            left: twipsToCm(pgMar?.getAttributeNS(W, 'left') || null),
        };
    });
}
function marginKey(m) { return `${m.top}|${m.right}|${m.bottom}|${m.left}`; }
function firstChapterParagraphIndex(doc) {
    return bodyParagraphs(doc).findIndex((p) => isChapterText(paragraphText(p)) && !looksLikeManualTocEntry(p));
}
function paragraphHasHardPageBreak(p) {
    return Array.from(p.getElementsByTagNameNS(W, 'br')).some((b) => (b.getAttributeNS(W, 'type') || '').toLowerCase() === 'page');
}
function hasStandalonePageBreakImmediatelyBefore(p) {
    const prev = p?.previousElementSibling || null;
    return Boolean(prev && prev.namespaceURI === W && prev.localName === 'p' && !paragraphText(prev) && paragraphHasHardPageBreak(prev));
}
function paragraphHasSection(p) {
    const pPr = directChild(p, 'pPr');
    return Boolean(pPr && directChild(pPr, 'sectPr'));
}
function safeEmptyParagraph(p) {
    if (paragraphText(p) || hasComplexContent(p))
        return false;
    if (paragraphHasSection(p))
        return false;
    if (p.getElementsByTagNameNS(W, 'br').length || p.getElementsByTagNameNS(W, 'fldChar').length || p.getElementsByTagNameNS(W, 'bookmarkStart').length || p.getElementsByTagNameNS(W, 'bookmarkEnd').length)
        return false;
    return true;
}
function countLikelyBlankParagraphs(doc, options = {}) {
    const ps = bodyParagraphs(doc);
    const mainStart = firstMainParagraphIndex(doc, options);
    if (mainStart < 0) return 0;
    let run = 0, excess = 0;
    for (let i = mainStart; i < ps.length; i++) {
        const text = paragraphText(ps[i]);
        if (isMainEndHeading(text)) break;
        const empty = safeEmptyParagraph(ps[i]);
        if (empty) run++;
        else { if (run >= 3) excess += run - 1; run = 0; }
    }
    if (run >= 3) excess += run - 1;
    return excess;
}
function summarizeProtectedContent(doc) {
    const texts = [];
    for (const p of bodyParagraphs(doc)) {
        const t = paragraphText(p);
        if (!t || looksLikeManualTocEntry(p)) continue;
        if (/^(Klik kanan|Daftar (tabel|gambar) akan diperbarui)/i.test(t)) continue;
        texts.push(t);
    }
    const counts = new Map();
    // Integritas V4 melindungi keberadaan setiap teks unik. Duplikasi identik (mis. judul pada cover + halaman judul/TOC)
    // tidak dihitung berkali-kali karena penghapusan daftar isi manual memang dapat mengurangi duplikat yang sah.
    texts.forEach((t) => counts.set(t, 1));
    return { counts, tables: doc.getElementsByTagNameNS(W, 'tbl').length, drawings: doc.getElementsByTagNameNS(W, 'drawing').length };
}
function validateProtectedContent(before, doc) {
    const after = summarizeProtectedContent(doc);
    const missing = [];
    for (const [t, n] of before.counts) {
        const have = after.counts.get(t) || 0;
        if (have < n) missing.push({ text: t, missing: n - have });
    }
    return { ok: missing.length === 0 && before.tables === after.tables && before.drawings === after.drawings, missing, tablesBefore: before.tables, tablesAfter: after.tables, drawingsBefore: before.drawings, drawingsAfter: after.drawings };
}
function restoreDocumentXml(doc, xml) {
    const restored = parseXml(xml);
    doc.replaceChild(doc.importNode(restored.documentElement, true), doc.documentElement);
}
function runProtectedStage(doc, label, changes, warnings, fn) {
    const xmlBefore = xmlString(doc);
    const protectedBefore = summarizeProtectedContent(doc);
    const changeCount = changes.length;
    try {
        fn();
        const check = validateProtectedContent(protectedBefore, doc);
        if (!check.ok) {
            restoreDocumentXml(doc, xmlBefore);
            changes.splice(changeCount);
            warnings.push(`${label} dibatalkan otomatis karena pemeriksaan anti-rusak mendeteksi isi/tabel/gambar berpotensi berubah${check.missing.length ? `; contoh teks: “${check.missing[0].text.slice(0, 90)}”` : ''}.`);
            return false;
        }
        return true;
    } catch (e) {
        restoreDocumentXml(doc, xmlBefore);
        changes.splice(changeCount);
        warnings.push(`${label} dibatalkan otomatis: ${e instanceof Error ? e.message : 'terjadi kesalahan.'}`);
        return false;
    }
}

function analyzeXml(fileName, doc, numbering, pageFieldOverride, options = {}) {
    const paragraphs = bodyParagraphs(doc);
    const profile = resolveProfile(doc, options);
    const candidates = detectHeadingCandidates(doc, numbering, options);
    const chapters = Array.from(new Set(candidates.filter((x) => x.level === 1 && isChapterText(x.text)).map((x) => chapterRoman(x.text)).filter(Boolean)));
    const margins = collectMargins(doc);
    const nativeToc = isMainToc(doc);
    const tocHeading = Boolean(findTocHeading(doc));
    const pageFields = pageFieldOverride ?? hasPageField(doc);
    const issues = [];
    const marginVariants = new Set(margins.map(marginKey));
    if (marginVariants.size > 1) issues.push({ id: 'margin-inconsistent', title: 'Margin tidak konsisten', detail: `Ditemukan ${marginVariants.size} konfigurasi margin pada ${margins.length} section.`, severity: 'warning', count: marginVariants.size });
    if (profile.type === 'skripsi') {
        const expected = ['I', 'II', 'III', 'IV', 'V'];
        const missing = expected.filter((r) => !chapters.includes(r));
        if (missing.length) issues.push({ id: 'chapter-missing', title: 'BAB skripsi belum lengkap terdeteksi', detail: `Belum menemukan: ${missing.map((x) => `BAB ${x}`).join(', ')}.`, severity: 'warning', count: missing.length });
    } else if (profile.structure === 'with-bab' && chapters.length === 0) {
        issues.push({ id: 'chapter-missing', title: 'Struktur Dengan BAB dipilih tetapi BAB belum terdeteksi', detail: 'Pilih Tanpa BAB bila dokumen memang tidak menggunakan BAB.', severity: 'warning' });
    }
    if (profile.structure === 'with-bab') {
        const chaptersWithoutPageBreak = paragraphs.filter((p) => {
            if (looksLikeManualTocEntry(p) || !isChapterText(paragraphText(p))) return false;
            const pPr = directChild(p, 'pPr');
            const hasBreak = Boolean(pPr && directChild(pPr, 'pageBreakBefore'));
            const hasBoundary = bodySectionProperties(doc).some((s) => nextNonEmptyParagraphAfterSectPr(doc, s) === p);
            return !hasBreak && !hasBoundary;
        }).length;
        if (chaptersWithoutPageBreak) issues.push({ id: 'chapter-pagebreak', title: 'BAB belum dipastikan mulai halaman baru', detail: `${chaptersWithoutPageBreak} judul BAB belum mempunyai Page Break Before/Section Break tepat sebelum BAB.`, severity: 'warning', count: chaptersWithoutPageBreak });
    }
    const mainStart = firstMainParagraphIndex(doc, options);
    if (mainStart > 0) {
        const frontStarts = paragraphs.slice(0, mainStart).filter((p) => FRONT_PAGE_START_RE.test(paragraphText(p)));
        const missingFrontBreak = frontStarts.filter((p) => {
            const pPr = directChild(p, 'pPr');
            const hasBoundary = bodySectionProperties(doc).some((s) => nextNonEmptyParagraphAfterSectPr(doc, s) === p);
            return !(pPr && directChild(pPr, 'pageBreakBefore')) && !paragraphHasHardPageBreak(p) && !hasStandalonePageBreakImmediatelyBefore(p) && !hasBoundary;
        }).length;
        if (missingFrontBreak) issues.push({ id: 'front-pagebreak', title: 'Bagian awal masih bergantung pada Enter', detail: `${missingFrontBreak} judul bagian awal belum mempunyai Page Break Before sehingga mudah bergeser.`, severity: 'warning', count: missingFrontBreak });
    }
    if (!nativeToc && options.createToc !== false) issues.push({ id: 'toc-native', title: 'TOC Word belum ditemukan', detail: tocHeading ? 'Judul “DAFTAR ISI” ada, tetapi field TOC Word belum terdeteksi.' : 'Judul dan field TOC Word belum terdeteksi.', severity: 'warning' });
    if (!pageFields && options.repairPageNumbering !== false) issues.push({ id: 'page-field', title: 'Penomoran PAGE belum terdeteksi', detail: 'Nomor halaman mungkin manual atau belum dibuat sebagai field PAGE.', severity: 'info' });
    const headingGaps = candidates.filter((c) => !c.existingLevel && c.confidence >= 0.76 && c.level > 0).length;
    if (headingGaps) issues.push({ id: 'toc-level-gap', title: 'Judul belum memiliki level TOC', detail: `${headingGaps} kandidat judul terdeteksi tetapi belum memiliki Outline Level.`, severity: 'warning', count: headingGaps });
    const blankRisk = countLikelyBlankParagraphs(doc, options);
    if (blankRisk) issues.push({ id: 'blank-risk', title: 'Enter/paragraf kosong berlebih', detail: `${blankRisk} paragraf kosong berlebih berpotensi membuat layout bergeser.`, severity: 'info', count: blankRisk });
    if (profile.confidence < 0.7 && (!options.documentType || options.documentType === 'auto')) issues.push({ id: 'type-uncertain', title: 'Jenis dokumen belum yakin', detail: `${profile.reason} Sebaiknya pilih Skripsi/Proposal/Makalah secara manual sebelum memproses.`, severity: 'warning' });
    const frontMatter = candidates.filter((c) => c.level === 1 && FRONT_MATTER_RE.test(c.text)).map((c) => c.text);
    const structure = {
        documentType: profile.type, structureMode: profile.structure, detectionConfidence: profile.confidence, detectionReason: profile.reason,
        mainStartText: mainStart >= 0 ? paragraphText(paragraphs[mainStart]) : '',
        frontMatter, chapters,
        bibliographyFound: paragraphs.some((p) => isMainEndHeading(paragraphText(p))),
        appendixFound: paragraphs.some((p) => /^LAMPIRAN/i.test(paragraphText(p))),
        tableCaptions: paragraphs.filter((p) => isTableCaption(paragraphText(p))).length,
        figureCaptions: paragraphs.filter((p) => isFigureCaption(paragraphText(p))).length,
        likelyBlankParagraphs: blankRisk,
    };
    const validationScore = Math.max(0, Math.min(100, 100 - issues.reduce((sum, x) => sum + (x.severity === 'error' ? 14 : x.severity === 'warning' ? 7 : 2), 0)));
    return { fileName, sections: bodySectionProperties(doc).length, paragraphs: paragraphs.length, chapters, subchapters: candidates.filter((x) => x.level === 2).length, subSubchapters: candidates.filter((x) => x.level === 3).length, hasNativeToc: nativeToc, hasTocHeading: tocHeading, hasPageFields: pageFields, margins, issues, headingCandidates: candidates, structure, validationScore };
}
async function loadDocumentXml(zip) {
    const file = zip.file('word/document.xml');
    if (!file)
        throw new Error('Bukan DOCX yang valid: word/document.xml tidak ditemukan.');
    return parseXml(await file.async('string'));
}
function nextNonEmptyParagraphAfterSectPr(doc, sectPr) {
    const pPr = sectPr.parentElement;
    const p = pPr?.parentElement;
    if (!p || p.localName !== 'p')
        return null;
    let n = p.nextElementSibling;
    while (n) {
        if (n.namespaceURI === W && n.localName === 'p' && paragraphText(n))
            return n;
        n = n.nextElementSibling;
    }
    return null;
}
function firstNonEmptyBodyParagraph(doc) { return bodyParagraphs(doc).find((p) => Boolean(paragraphText(p))) || null; }
function sectionOwnerParagraph(sectPr) {
    const pPr = sectPr.parentElement;
    const p = pPr?.parentElement;
    return p && p.namespaceURI === W && p.localName === 'p' ? p : null;
}
function dedupeSectionBoundariesBeforeChapters(doc, changes) {
    const chapters = bodyParagraphs(doc).filter((p) => isChapterText(paragraphText(p)) && !isLikelyTocParagraph(p));
    let removed = 0;
    for (const chapter of chapters) {
        const boundaries = bodySectionProperties(doc).filter((s) => nextNonEmptyParagraphAfterSectPr(doc, s) === chapter);
        if (boundaries.length <= 1)
            continue;
        // Pertahankan section break yang paling dekat dengan BAB. Hapus hanya duplikat pada paragraf kosong.
        for (const sect of boundaries.slice(0, -1)) {
            const owner = sectionOwnerParagraph(sect);
            if (!owner || paragraphText(owner))
                continue;
            sect.remove();
            removed++;
        }
    }
    if (removed)
        changes.push(`${removed} Section Break duplikat tepat sebelum BAB dibersihkan agar tidak membuat halaman/nomor lompat.`);
}
function removeTrailingEmptyParagraphsBefore(target) {
    let removed = 0;
    let prev = target.previousElementSibling;
    while (prev && prev.namespaceURI === W && prev.localName === 'p' && safeEmptyParagraph(prev)) {
        const next = prev.previousElementSibling;
        prev.remove();
        removed++;
        prev = next;
    }
    return removed;
}
function removeLeadingEmptyParagraphsAfter(target) {
    let removed = 0;
    let next = target.nextElementSibling;
    while (next && next.namespaceURI === W && next.localName === 'p' && safeEmptyParagraph(next)) {
        const after = next.nextElementSibling;
        next.remove();
        removed++;
        next = after;
    }
    return removed;
}
function flattenManualBreaksInParagraph(p) {
    const breaks = Array.from(p.getElementsByTagNameNS(W, 'br'));
    let changed = 0;
    for (const br of breaks) {
        const type = (br.getAttributeNS(W, 'type') || '').toLowerCase();
        if (type && type !== 'textwrapping' && type !== 'page')
            continue;
        const run = br.parentElement;
        if (run) {
            const texts = Array.from(run.getElementsByTagNameNS(W, 't'));
            const last = texts[texts.length - 1];
            if (last && last.textContent && !/\s$/.test(last.textContent))
                last.textContent += ' ';
        }
        br.remove();
        changed++;
    }
    return changed;
}
function ensurePageBreakBefore(p) {
    const pPr = ensureDirectChild(p, 'pPr');
    ensurePPrChild(pPr, 'pageBreakBefore');
}
function ensureHardPageBreakBefore(p) {
    if (paragraphHasHardPageBreak(p)) return false;
    const doc = p.ownerDocument;
    const run = wEl(doc, 'r'); const br = wEl(doc, 'br'); wAttr(br, 'type', 'page'); run.appendChild(br);
    const pPr = directChild(p, 'pPr');
    const before = pPr ? pPr.nextSibling : p.firstChild;
    if (before) p.insertBefore(run, before); else p.appendChild(run);
    return true;
}
function ensureStandalonePageBreakBefore(p) {
    const prev = p.previousElementSibling;
    if (prev && prev.namespaceURI === W && prev.localName === 'p' && !paragraphText(prev) && paragraphHasHardPageBreak(prev)) return false;
    const doc = p.ownerDocument; const breakP = wEl(doc, 'p'); const run = wEl(doc, 'r'); const br = wEl(doc, 'br');
    wAttr(br, 'type', 'page'); run.appendChild(br); breakP.appendChild(run); p.parentNode.insertBefore(breakP, p); return true;
}
function repeatedCoverTitleParagraph(doc, options = {}) {
    const ps = bodyParagraphs(doc);
    const mainStart = firstMainParagraphIndex(doc, options);
    if (mainStart <= 0) return null;
    const firstNonEmpty = ps.find((p, i) => i < mainStart && Boolean(paragraphText(p)));
    if (!firstNonEmpty) return null;
    const title = paragraphText(firstNonEmpty);
    if (title.length < 12) return null;
    const firstIndex = ps.indexOf(firstNonEmpty);
    return ps.find((p, i) => i > firstIndex && i < mainStart && paragraphText(p) === title) || null;
}
function applyFrontMatterPagination(doc, options, changes) {
    const ps = bodyParagraphs(doc);
    const mainStart = firstMainParagraphIndex(doc, options);
    if (mainStart <= 0) return;
    const profile = resolveProfile(doc, options);
    const starts = [];
    const cover2 = profile.type === 'skripsi' ? repeatedCoverTitleParagraph(doc, options) : null;
    if (cover2) starts.push(cover2);
    for (let i = 0; i < mainStart; i++) {
        const p = ps[i];
        if (FRONT_PAGE_START_RE.test(paragraphText(p))) starts.push(p);
    }
    let pages = 0, removed = 0, flattened = 0;
    const frontNumberStart = options.repairPageNumbering ? findFrontNumberStartParagraph(doc, options) : null;
    for (const p of Array.from(new Set(starts))) {
        removed += removeTrailingEmptyParagraphsBefore(p);
        flattened += flattenManualBreaksInParagraph(p);
        const alreadySectionStart = bodySectionProperties(doc).some((s) => nextNonEmptyParagraphAfterSectPr(doc, s) === p);
        if (p === frontNumberStart || alreadySectionStart) ensurePageBreakBefore(p);
        else ensureStandalonePageBreakBefore(p);
        pages++;
    }
    if (pages) changes.push(`${pages} awal halaman bagian depan dipastikan mulai halaman baru tanpa mengandalkan Enter kosong.`);
    if (removed) changes.push(`${removed} Enter kosong tepat sebelum awal halaman bagian depan dihapus; ruang vertikal desain cover tetap dilindungi.`);
    if (flattened) changes.push(`${flattened} manual page/line break pada judul bagian depan dinormalisasi.`);
}
function ensureSectionBoundaryBefore(doc, target, changes, label = 'bagian utama') {
    if (!target) return false;
    const exists = bodySectionProperties(doc).some((s) => nextNonEmptyParagraphAfterSectPr(doc, s) === target);
    if (exists) return false;
    const body = getBody(doc);
    const finalSect = directChild(body, 'sectPr');
    if (!finalSect) return false;
    let prev = target.previousElementSibling;
    while (prev && !(prev.namespaceURI === W && prev.localName === 'p')) prev = prev.previousElementSibling;
    if (!prev) return false;
    const pPr = ensureDirectChild(prev, 'pPr');
    if (directChild(pPr, 'sectPr')) return false;
    const clone = finalSect.cloneNode(true);
    directChildren(clone, 'headerReference').forEach((x) => x.remove());
    directChildren(clone, 'footerReference').forEach((x) => x.remove());
    const type = ensureSectChild(clone, 'type', ['pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid']);
    wAttr(type, 'val', 'nextPage');
    pPr.appendChild(clone);
    changes.push(`Section Break (Next Page) ditambahkan tepat sebelum ${label}.`);
    return true;
}
function ensureChapterSectionBoundaries(doc, changes) {
    let added = 0;
    const chapters = bodyParagraphs(doc).filter((p) => isChapterText(paragraphText(p)) && !looksLikeManualTocEntry(p));
    for (const chapter of chapters) if (ensureSectionBoundaryBefore(doc, chapter, [], `BAB ${chapterRoman(paragraphText(chapter)) || ''}`)) added++;
    if (added) changes.push(`${added} Section Break (Next Page) ditambahkan tepat sebelum BAB agar pembagian BAB dan penomoran stabil.`);
}
function ensureNumberingBoundaries(doc, options, changes) {
    const profile = resolveProfile(doc, options);
    const main = firstMainParagraph(doc, options);
    if (!main) return;
    if (profile.structure === 'with-bab') ensureChapterSectionBoundaries(doc, changes);
    else ensureSectionBoundaryBefore(doc, main, changes, 'bagian isi utama');
}

function applyMarginsAndPaper(doc, options, changes) {
    if (!options.normalizeMargins) return;
    const sizes = options.paper === 'A4' ? { w: '11907', h: '16840', code: '9' } : { w: '11907', h: '18709', code: undefined };
    let changed = 0, preservedLandscape = 0;
    for (const sectPr of bodySectionProperties(doc)) {
        const existingSize = directChild(sectPr, 'pgSz');
        const oldW = Number(existingSize?.getAttributeNS(W, 'w') || 0), oldH = Number(existingSize?.getAttributeNS(W, 'h') || 0);
        const landscape = (existingSize?.getAttributeNS(W, 'orient') || '').toLowerCase() === 'landscape' || (oldW > 0 && oldH > 0 && oldW > oldH);
        if (options.mode === 'safe' && landscape) { preservedLandscape++; continue; }
        const pgSz = ensureSectChild(sectPr, 'pgSz', ['pgMar', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid']);
        wAttr(pgSz, 'w', sizes.w); wAttr(pgSz, 'h', sizes.h); pgSz.removeAttributeNS(W, 'orient');
        if (sizes.code) wAttr(pgSz, 'code', sizes.code); else pgSz.removeAttributeNS(W, 'code');
        const pgMar = ensureSectChild(sectPr, 'pgMar', ['pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid']);
        wAttr(pgMar, 'top', cmToTwips(options.marginTopCm)); wAttr(pgMar, 'right', cmToTwips(options.marginRightCm)); wAttr(pgMar, 'bottom', cmToTwips(options.marginBottomCm)); wAttr(pgMar, 'left', cmToTwips(options.marginLeftCm));
        if (!pgMar.hasAttributeNS(W, 'header')) wAttr(pgMar, 'header', '720');
        if (!pgMar.hasAttributeNS(W, 'footer')) wAttr(pgMar, 'footer', '720');
        if (!pgMar.hasAttributeNS(W, 'gutter')) wAttr(pgMar, 'gutter', '0');
        changed++;
    }
    changes.push(`Margin diterapkan pada ${changed} section: kiri ${options.marginLeftCm} cm, atas ${options.marginTopCm} cm, kanan ${options.marginRightCm} cm, bawah ${options.marginBottomCm} cm; kertas ${options.paper}.`);
    if (preservedLandscape) changes.push(`${preservedLandscape} section landscape dipertahankan pada Mode Aman agar tabel/lampiran lebar tidak rusak.`);
}
function setOutlineLevel(p, level) {
    const pPr = ensureDirectChild(p, 'pPr');
    const existing = directChild(pPr, 'outlineLvl');
    const outline = existing ? ensurePPrChild(pPr, 'outlineLvl') : ensurePPrChild(pPr, 'outlineLvl');
    if (level === 0) {
        wAttr(outline, 'val', '9');
        return;
    }
    wAttr(outline, 'val', String(level - 1));
}
function applyHeadingLevels(doc, numbering, options, changes) {
    if (!options.createToc && !options.repairHeadings) return;
    const ps = bodyParagraphs(doc);
    const candidates = detectHeadingCandidates(doc, numbering, options);
    const mainStart = firstMainParagraphIndex(doc, options);
    const mainHeadingElements = [];
    let l1 = 0, l2 = 0, l3 = 0, excluded = 0;
    for (const c of candidates) {
        const p = ps[c.paragraphIndex];
        if (!p) continue;
        let level = Object.prototype.hasOwnProperty.call(options.headingOverrides || {}, String(c.paragraphIndex)) ? options.headingOverrides[String(c.paragraphIndex)] : c.level;
        if (!options.includeFrontMatterInToc && FRONT_MATTER_RE.test(c.text)) level = 0;
        setOutlineLevel(p, level);
        if (level === 0) { excluded++; continue; }
        if (options.repairHeadings) {
            const pPr = ensureDirectChild(p, 'pPr');
            ensurePPrChild(pPr, 'keepNext'); ensurePPrChild(pPr, 'keepLines');
            if (c.paragraphIndex >= mainStart || c.chapter || isChapterText(c.text)) mainHeadingElements.push(p);
        }
        if (level === 1) l1++; else if (level === 2) l2++; else if (level === 3) l3++;
    }
    changes.push(`Level TOC diterapkan: ${l1} Level 1, ${l2} Level 2, ${l3} Level 3${excluded ? `, ${excluded} dikeluarkan` : ''}.`);
    if (options.repairHeadings) {
        let removed = 0, flattened = 0, spacingFixed = 0;
        for (const p of mainHeadingElements) {
            flattened += flattenManualBreaksInParagraph(p);
            removed += removeTrailingEmptyParagraphsBefore(p);
            removed += removeLeadingEmptyParagraphsAfter(p);
            const pPr = ensureDirectChild(p, 'pPr');
            const spacing = ensurePPrChild(pPr, 'spacing');
            const before = Number(spacing.getAttributeNS(W, 'before') || 0);
            const after = Number(spacing.getAttributeNS(W, 'after') || 0);
            // Jarak ekstrem akibat Spacing Before/After sering tampak seperti banyak Enter.
            if (before > 720) { wAttr(spacing, 'before', options.mode === 'safe' ? '240' : '0'); spacingFixed++; }
            if (after > 480) { wAttr(spacing, 'after', options.mode === 'safe' ? '120' : '0'); spacingFixed++; }
        }
        changes.push('Judul utama/Subjudul/Sub-subjudul diberi Keep with next/Keep lines agar tidak tertinggal sendirian di bawah halaman.');
        if (removed) changes.push(`${removed} Enter kosong yang menempel tepat sebelum/sesudah judul di bagian isi dihapus.`);
        if (flattened) changes.push(`${flattened} manual line break di dalam judul dirapikan.`);
        if (spacingFixed) changes.push(`${spacingFixed} jarak Spacing Before/After yang terlalu besar pada judul dinormalisasi.`);
    }
}
function applyDocumentPagination(doc, options, changes) {
    if (!options.repairChapterPagination) return;
    const profile = resolveProfile(doc, options);
    if (profile.structure === 'with-bab') dedupeSectionBoundariesBeforeChapters(doc, changes);
    applyFrontMatterPagination(doc, options, changes);
    const ps = bodyParagraphs(doc);
    const mainStart = firstMainParagraphIndex(doc, options);
    if (mainStart < 0) return;
    if (profile.structure === 'without-bab') {
        const first = ps[mainStart];
        let removed = removeTrailingEmptyParagraphsBefore(first) + removeLeadingEmptyParagraphsAfter(first);
        const flattened = flattenManualBreaksInParagraph(first);
        ensurePageBreakBefore(first);
        const pPr = ensureDirectChild(first, 'pPr'); ensurePPrChild(pPr, 'keepNext'); ensurePPrChild(pPr, 'keepLines');
        changes.push(`Bagian isi utama “${paragraphText(first).slice(0, 80)}” dipastikan mulai halaman baru tanpa memaksa setiap bagian A/B/C ke halaman baru.`);
        if (removed) changes.push(`${removed} Enter kosong di sekitar awal isi utama dibersihkan.`);
        if (flattened) changes.push(`${flattened} manual break pada awal isi utama dinormalisasi.`);
        return;
    }
    let changed = 0, removed = 0, flattened = 0;
    for (const p of ps) {
        if (!isChapterText(paragraphText(p)) || looksLikeManualTocEntry(p)) continue;
        removed += removeTrailingEmptyParagraphsBefore(p); removed += removeLeadingEmptyParagraphsAfter(p); flattened += flattenManualBreaksInParagraph(p);
        const isSectionStart = bodySectionProperties(doc).some((s) => nextNonEmptyParagraphAfterSectPr(doc, s) === p);
        const pPr = ensureDirectChild(p, 'pPr'); ensurePPrChild(pPr, 'keepNext'); ensurePPrChild(pPr, 'keepLines');
        if (!isSectionStart) ensurePPrChild(pPr, 'pageBreakBefore');
        changed++;
    }
    if (changed) changes.push(`${changed} BAB dipastikan mulai halaman baru tanpa menambah Enter kosong.`);
    if (removed) changes.push(`${removed} Enter kosong yang menumpuk tepat di sekitar BAB dibersihkan.`);
    if (flattened) changes.push(`${flattened} manual break di judul BAB dinormalisasi.`);
}
function createFieldParagraph(doc, instruction, placeholder, style = 'TOC1') {
    const p = wEl(doc, 'p');
    const pPr = wEl(doc, 'pPr');
    if (style) {
        const pStyle = wEl(doc, 'pStyle');
        wAttr(pStyle, 'val', style);
        pPr.appendChild(pStyle);
    }
    p.appendChild(pPr);
    const beginRun = wEl(doc, 'r');
    const begin = wEl(doc, 'fldChar');
    wAttr(begin, 'fldCharType', 'begin');
    wAttr(begin, 'dirty', 'true');
    beginRun.appendChild(begin);
    p.appendChild(beginRun);
    const instrRun = wEl(doc, 'r');
    const instr = wEl(doc, 'instrText');
    instr.setAttribute('xml:space', 'preserve');
    instr.textContent = ` ${instruction} `;
    instrRun.appendChild(instr);
    p.appendChild(instrRun);
    const sepRun = wEl(doc, 'r');
    const sep = wEl(doc, 'fldChar');
    wAttr(sep, 'fldCharType', 'separate');
    sepRun.appendChild(sep);
    p.appendChild(sepRun);
    const hintRun = wEl(doc, 'r');
    const hint = wEl(doc, 't');
    hint.textContent = placeholder;
    hintRun.appendChild(hint);
    p.appendChild(hintRun);
    const endRun = wEl(doc, 'r');
    const end = wEl(doc, 'fldChar');
    wAttr(end, 'fldCharType', 'end');
    endRun.appendChild(end);
    p.appendChild(endRun);
    return p;
}
function createNativeTocBlock(doc, instruction, placeholder) {
    const sdt = wEl(doc, 'sdt');
    const pr = wEl(doc, 'sdtPr');
    const id = wEl(doc, 'id');
    wAttr(id, 'val', String(-Math.floor(100000000 + Math.random() * 1900000000)));
    pr.appendChild(id);
    const obj = wEl(doc, 'docPartObj');
    const gal = wEl(doc, 'docPartGallery');
    wAttr(gal, 'val', 'Table of Contents');
    obj.appendChild(gal);
    obj.appendChild(wEl(doc, 'docPartUnique'));
    pr.appendChild(obj);
    sdt.appendChild(pr);
    sdt.appendChild(wEl(doc, 'sdtEndPr'));
    const content = wEl(doc, 'sdtContent');
    content.appendChild(createFieldParagraph(doc, instruction, placeholder));
    sdt.appendChild(content);
    return sdt;
}
function createHeadingParagraph(doc, text) {
    const p = wEl(doc, 'p');
    const pPr = wEl(doc, 'pPr');
    const keep = wEl(doc, 'keepNext'); pPr.appendChild(keep);
    const br = wEl(doc, 'pageBreakBefore'); pPr.appendChild(br);
    const jc = wEl(doc, 'jc'); wAttr(jc, 'val', 'center'); pPr.appendChild(jc);
    const outline = wEl(doc, 'outlineLvl'); wAttr(outline, 'val', '0'); pPr.appendChild(outline);
    p.appendChild(pPr);
    const r = wEl(doc, 'r');
    const rPr = wEl(doc, 'rPr');
    rPr.appendChild(wEl(doc, 'b'));
    r.appendChild(rPr);
    const t = wEl(doc, 't');
    t.textContent = text;
    r.appendChild(t);
    p.appendChild(r);
    return p;
}
function removeLikelyManualTocRegion(doc, heading) {
    let removed = 0;
    let node = heading.nextElementSibling;
    while (node) {
        if (node.namespaceURI === W && node.localName === 'sdt')
            break;
        if (!(node.namespaceURI === W && node.localName === 'p'))
            break;
        const p = node;
        const text = paragraphText(p);
        if (/^(DAFTAR\s+(TABEL|GAMBAR|LAMPIRAN)|ABSTRAK|ABSTRACT|KATA\s+PENGANTAR)$/i.test(text))
            break;
        const likelyEntry = looksLikeManualTocEntry(p);
        const empty = !text;
        if (!likelyEntry && !empty)
            break;
        const next = node.nextElementSibling;
        node.remove();
        removed++;
        node = next;
    }
    return removed;
}
function findOrCreateHeadingBeforeMain(doc, regex, title, changes, options = {}) {
    let heading = findHeading(doc, regex);
    if (heading) return heading;
    const main = firstMainParagraph(doc, options);
    if (!main) return null;
    heading = createHeadingParagraph(doc, title);
    main.parentNode.insertBefore(heading, main);
    changes.push(`Judul “${title}” dibuat karena belum ditemukan.`);
    return heading;
}
function refreshMainTocInstruction(doc) {
    for (const n of Array.from(doc.getElementsByTagNameNS(W, 'instrText'))) {
        const text = n.textContent || '';
        if (/\bTOC\b/i.test(text) && !/\\f\s+[TG](?:\s|$)/i.test(text) && !/\\c\s+/i.test(text)) {
            n.textContent = ' TOC \\o "1-3" \\h \\z \\u ';
            return true;
        }
    }
    return false;
}
function applyToc(doc, options, changes, warnings) {
    if (!options.createToc)
        return;
    if (isMainToc(doc)) {
        refreshMainTocInstruction(doc);
        changes.push('TOC Word yang sudah ada dipertahankan, dipastikan memakai Level 1–3 + Outline Level (\\u), lalu ditandai untuk diperbarui.');
        return;
    }
    const heading = findOrCreateHeadingBeforeMain(doc, /^DAFTAR\s+ISI$/i, 'DAFTAR ISI', changes, options);
    if (!heading) {
        warnings.push('Awal bagian isi dan posisi “DAFTAR ISI” tidak dapat ditentukan; TOC tidak disisipkan.');
        return;
    }
    if (options.removeLikelyManualToc) {
        const removed = removeLikelyManualTocRegion(doc, heading);
        if (removed)
            changes.push(`${removed} paragraf daftar isi manual yang terdeteksi dihapus sebelum TOC Word dibuat.`);
    }
    const toc = createNativeTocBlock(doc, 'TOC \\o "1-3" \\h \\z \\u', 'Klik kanan → Update Field / Ctrl+A → F9 untuk menghitung TOC.');
    heading.parentNode.insertBefore(toc, heading.nextSibling);
    changes.push('TOC Microsoft Word native dibuat dengan Level 1, Level 2, dan Level 3 — bukan hanya BAB.');
}
function hasTcField(p, mark) {
    return Array.from(p.getElementsByTagNameNS(W, 'fldSimple')).some((x) => new RegExp(`\\bTC\\b.*\\\\f\\s+${mark}`, 'i').test(x.getAttributeNS(W, 'instr') || ''));
}
function addTcField(p, text, mark) {
    if (hasTcField(p, mark))
        return;
    const fld = wEl(p.ownerDocument, 'fldSimple');
    const safeText = text.replace(/"/g, "'");
    wAttr(fld, 'instr', ` TC "${safeText}" \\f ${mark} \\l 1 `);
    p.appendChild(fld);
}
function removeLikelyManualSpecialListRegion(heading, kind) {
    let removed = 0;
    let node = heading.nextElementSibling;
    const entryRe = kind === 'Tabel' ? /^TABEL\s+(?:[IVXLC]+|\d+)/i : /^GAMBAR\s+(?:[IVXLC]+|\d+)/i;
    while (node) {
        if (node.namespaceURI === W && node.localName === 'sdt')
            break;
        if (!(node.namespaceURI === W && node.localName === 'p'))
            break;
        const text = paragraphText(node);
        if (text && (FRONT_PAGE_START_RE.test(text) || isChapterText(text)))
            break;
        const removable = !text || entryRe.test(text) || isLikelyTocParagraph(node);
        if (!removable)
            break;
        const next = node.nextElementSibling;
        node.remove();
        removed++;
        node = next;
    }
    return removed;
}
function applyTableFigureList(doc, options, changes, warnings) {
    const ps = bodyParagraphs(doc);
    if (options.createTableList) {
        const caps = ps.filter((p) => isTableCaption(paragraphText(p)));
        caps.forEach((p) => addTcField(p, paragraphText(p), 'T'));
        if (caps.length) {
            if (!hasSpecialToc(doc, 'T')) {
                const heading = findOrCreateHeadingBeforeMain(doc, /^DAFTAR\s+TABEL$/i, 'DAFTAR TABEL', changes, options);
                if (heading) {
                    if (options.removeLikelyManualToc) {
                        const removed = removeLikelyManualSpecialListRegion(heading, 'Tabel');
                        if (removed)
                            changes.push(`${removed} baris Daftar Tabel manual diganti dengan field dinamis Word agar tidak dobel.`);
                    }
                    heading.parentNode.insertBefore(createFieldParagraph(doc, 'TOC \\h \\z \\f T', 'Daftar tabel akan diperbarui oleh Microsoft Word.'), heading.nextSibling);
                }
            }
            changes.push(`${caps.length} caption tabel ditandai dan Daftar Tabel native/dinamis disiapkan.`);
        }
        else
            warnings.push('Tidak menemukan caption yang diawali “Tabel …”; Daftar Tabel tidak dibuat agar tidak ngasal.');
    }
    if (options.createFigureList) {
        const caps = ps.filter((p) => isFigureCaption(paragraphText(p)));
        caps.forEach((p) => addTcField(p, paragraphText(p), 'G'));
        if (caps.length) {
            if (!hasSpecialToc(doc, 'G')) {
                const heading = findOrCreateHeadingBeforeMain(doc, /^DAFTAR\s+GAMBAR$/i, 'DAFTAR GAMBAR', changes, options);
                if (heading) {
                    if (options.removeLikelyManualToc) {
                        const removed = removeLikelyManualSpecialListRegion(heading, 'Gambar');
                        if (removed)
                            changes.push(`${removed} baris Daftar Gambar manual diganti dengan field dinamis Word agar tidak dobel.`);
                    }
                    heading.parentNode.insertBefore(createFieldParagraph(doc, 'TOC \\h \\z \\f G', 'Daftar gambar akan diperbarui oleh Microsoft Word.'), heading.nextSibling);
                }
            }
            changes.push(`${caps.length} caption gambar ditandai dan Daftar Gambar native/dinamis disiapkan.`);
        }
        else
            warnings.push('Tidak menemukan caption yang diawali “Gambar …”; Daftar Gambar tidak dibuat agar tidak ngasal.');
    }
}
function applyBodyFormatting(doc, numbering, options, changes) {
    if (!options.normalizeBodyFormatting) return;
    const ps = bodyParagraphs(doc);
    const headingIdx = new Set(detectHeadingCandidates(doc, numbering, options).map((x) => x.paragraphIndex));
    const mainStart = firstMainParagraphIndex(doc, options);
    if (mainStart < 0) return;
    let changed = 0;
    const minLen = options.mode === 'safe' ? 70 : 25;
    for (let i = mainStart; i < ps.length; i++) {
        const p = ps[i]; const text = paragraphText(p);
        if (isMainEndHeading(text)) break;
        if (!text || headingIdx.has(i) || isTableCaption(text) || isFigureCaption(text) || END_MATTER_RE.test(text) || text.length < minLen || hasComplexContent(p)) continue;
        const pPr = ensureDirectChild(p, 'pPr');
        if (directChild(pPr, 'numPr') || paragraphHasSection(p) || paragraphHasHardPageBreak(p) || p.getElementsByTagNameNS(W, 'tab').length) continue;
        const oldJc = directChild(pPr, 'jc')?.getAttributeNS(W, 'val') || '';
        const indExisting = directChild(pPr, 'ind');
        const left = Number(indExisting?.getAttributeNS(W, 'left') || indExisting?.getAttributeNS(W, 'start') || 0);
        const hanging = Number(indExisting?.getAttributeNS(W, 'hanging') || 0);
        if (options.mode === 'safe') {
            if (oldJc && !['left', 'both', 'distribute'].includes(oldJc)) continue;
            if (left > 0 || hanging > 0 || paragraphBoldRatio(p) > 0.55) continue;
        }
        const jc = ensurePPrChild(pPr, 'jc'); wAttr(jc, 'val', 'both');
        const spacing = ensurePPrChild(pPr, 'spacing'); wAttr(spacing, 'line', String(Math.round(options.lineSpacing * 240))); wAttr(spacing, 'lineRule', 'auto'); wAttr(spacing, 'after', '0');
        const ind = ensurePPrChild(pPr, 'ind'); if (!ind.hasAttributeNS(W, 'hanging')) wAttr(ind, 'firstLine', cmToTwips(options.firstLineIndentCm));
        if (options.mode === 'total') {
            for (const r of Array.from(p.getElementsByTagNameNS(W, 'r'))) {
                if (!Array.from(r.getElementsByTagNameNS(W, 't')).length) continue;
                const rPr = directChild(r, 'rPr') || r.insertBefore(wEl(doc, 'rPr'), r.firstChild);
                const fonts = ensureDirectChild(rPr, 'rFonts'); wAttr(fonts, 'ascii', options.fontFamily); wAttr(fonts, 'hAnsi', options.fontFamily); wAttr(fonts, 'eastAsia', options.fontFamily);
                const sz = ensureDirectChild(rPr, 'sz'); wAttr(sz, 'val', String(Math.round(options.fontSizePt * 2)));
                const szCs = ensureDirectChild(rPr, 'szCs'); wAttr(szCs, 'val', String(Math.round(options.fontSizePt * 2)));
            }
        }
        changed++;
    }
    changes.push(`${changed} paragraf isi dinormalisasi${options.mode === 'safe' ? ' secara konservatif (cover, halaman awal, heading, list, tabel dan format khusus dilindungi)' : `: ${options.fontFamily} ${options.fontSizePt} pt, spasi ${options.lineSpacing}, justify, first-line ${options.firstLineIndentCm} cm`}.`);
}
function cleanBlankParagraphs(doc, options, changes) {
    if (!options.cleanBlankParagraphs) return;
    const ps = bodyParagraphs(doc);
    const mainStart = firstMainParagraphIndex(doc, options);
    if (mainStart < 0) return;
    const threshold = options.mode === 'safe' ? 3 : 2;
    const remove = []; let run = [];
    const flush = () => { if (run.length >= threshold) remove.push(...run.slice(1)); run = []; };
    for (let i = mainStart; i < ps.length; i++) {
        const p = ps[i]; const text = paragraphText(p);
        if (isMainEndHeading(text)) { flush(); break; }
        if (safeEmptyParagraph(p)) run.push(p); else flush();
    }
    flush(); remove.forEach((p) => p.remove());
    if (remove.length) changes.push(`${remove.length} paragraf kosong berlebih di bagian isi dihapus. Cover dan bagian awal dokumen sengaja tidak disentuh.`);
}
async function loadRels(zip) {
    const file = zip.file('word/_rels/document.xml.rels');
    if (!file)
        throw new Error('document.xml.rels tidak ditemukan.');
    return parseXml(await file.async('string'));
}
async function loadContentTypes(zip) {
    const file = zip.file('[Content_Types].xml');
    if (!file)
        throw new Error('[Content_Types].xml tidak ditemukan.');
    return parseXml(await file.async('string'));
}
function nextRid(rels) {
    const ids = Array.from(rels.documentElement.children).map((e) => e.getAttribute('Id') || '').map((x) => Number(x.replace(/^rId/i, ''))).filter(Number.isFinite);
    return `rId${Math.max(0, ...ids) + 1}`;
}
function nextPartNumber(zip, prefix) {
    const re = new RegExp(`^word/${prefix}(\\d+)\\.xml$`);
    const nums = Object.keys(zip.files).map((name) => name.match(re)?.[1]).filter(Boolean).map(Number);
    return Math.max(0, ...nums) + 1;
}
function addRel(rels, id, type, target) {
    const rel = rels.createElementNS(RELS, 'Relationship');
    rel.setAttribute('Id', id);
    rel.setAttribute('Type', type);
    rel.setAttribute('Target', target);
    rels.documentElement.appendChild(rel);
}
function addContentType(ct, partName, contentType) {
    if (Array.from(ct.documentElement.children).some((e) => e.getAttribute('PartName') === partName))
        return;
    const ov = ct.createElementNS(CT, 'Override');
    ov.setAttribute('PartName', partName);
    ov.setAttribute('ContentType', contentType);
    ct.documentElement.appendChild(ov);
}
function makeHeaderFooterXml(kind, align, withPage) {
    const rootTag = kind === 'header' ? 'hdr' : 'ftr';
    const pageField = withPage ? '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>' : '<w:r><w:t></w:t></w:r>';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${rootTag} xmlns:w="${W}" xmlns:r="${R}"><w:p><w:pPr><w:jc w:val="${align}"/></w:pPr>${pageField}</w:p></w:${rootTag}>`;
}
function removeReferences(sectPr, local) { directChildren(sectPr, local).forEach((e) => e.remove()); }
async function addPart(zip, rels, ct, kind, align, withPage) {
    const n = nextPartNumber(zip, kind);
    const fileName = `${kind}${n}.xml`;
    const path = `word/${fileName}`;
    zip.file(path, makeHeaderFooterXml(kind, align, withPage));
    const rid = nextRid(rels);
    addRel(rels, rid, `${R}/${kind}`, fileName);
    addContentType(ct, `/${path}`, `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`);
    return rid;
}
function appendReference(sectPr, kind, type, rid) {
    const ref = wEl(sectPr.ownerDocument, `${kind}Reference`);
    wAttr(ref, 'type', type);
    rAttr(ref, 'id', rid);
    const firstNonRef = Array.from(sectPr.children).find((e) => !['headerReference', 'footerReference'].includes(e.localName));
    if (firstNonRef)
        sectPr.insertBefore(ref, firstNonRef);
    else
        sectPr.appendChild(ref);
}
function referenceByType(sectPr, kind, type) {
    return directChildren(sectPr, `${kind}Reference`).find((r) => (r.getAttributeNS(W, 'type') || 'default') === type) || null;
}
function relationshipTarget(rels, rid) {
    const rel = Array.from(rels.documentElement.children).find((e) => e.getAttribute('Id') === rid);
    return rel?.getAttribute('Target') || null;
}
function replaceReferenceType(sectPr, kind, type, rid) {
    const old = referenceByType(sectPr, kind, type);
    if (old) old.remove();
    appendReference(sectPr, kind, type, rid);
}
function partHasPageField(doc) {
    return Array.from(doc.getElementsByTagNameNS(W, 'fldSimple')).some((n) => /\bPAGE\b/i.test(n.getAttributeNS(W, 'instr') || '')) ||
        Array.from(doc.getElementsByTagNameNS(W, 'instrText')).some((n) => /\bPAGE\b/i.test(n.textContent || ''));
}
function appendPageParagraph(doc, align, numberFormat = 'Arabic') {
    const root = doc.documentElement;
    const p = wEl(doc, 'p'); const pPr = wEl(doc, 'pPr'); const jc = wEl(doc, 'jc'); wAttr(jc, 'val', align); pPr.appendChild(jc); p.appendChild(pPr);
    const fld = wEl(doc, 'fldSimple'); wAttr(fld, 'instr', ` PAGE \\* ${numberFormat} `); const r = wEl(doc, 'r'); const t = wEl(doc, 't'); t.textContent = '1'; r.appendChild(t); fld.appendChild(r); p.appendChild(fld); root.appendChild(p);
}
function ancestorByLocal(node, local) {
    let cur = node?.parentElement || null;
    while (cur) {
        if (cur.namespaceURI === W && cur.localName === local) return cur;
        cur = cur.parentElement;
    }
    return null;
}
function runHasFieldChar(run, type) {
    return Array.from(run?.getElementsByTagNameNS(W, 'fldChar') || []).some((n) => (n.getAttributeNS(W, 'fldCharType') || '') === type);
}
function stripPageFieldsFromPart(doc) {
    let removed = 0;
    // Field sederhana: <w:fldSimple w:instr=" PAGE ">...</w:fldSimple>
    for (const fld of Array.from(doc.getElementsByTagNameNS(W, 'fldSimple'))) {
        if (/\bPAGE\b/i.test(fld.getAttributeNS(W, 'instr') || '')) { fld.remove(); removed++; }
    }
    // Field PAGE kompleks Word biasanya berada dalam SDT dan menyimpan hasil lama sebagai <w:t>2</w:t>.
    // Hapus SATU field lengkap (begin -> instrText -> separate -> result -> end), bukan hanya instrText,
    // agar angka lama tidak tertinggal sebagai teks biasa pada cover/halaman awal.
    const handledContainers = new Set();
    for (const instr of Array.from(doc.getElementsByTagNameNS(W, 'instrText'))) {
        if (!/\bPAGE\b/i.test(instr.textContent || '')) continue;
        const sdt = ancestorByLocal(instr, 'sdt');
        if (sdt && !handledContainers.has(sdt)) {
            // Jika SDT ini memang container nomor halaman, buang container-nya secara utuh.
            handledContainers.add(sdt); sdt.remove(); removed++; continue;
        }
        const p = ancestorByLocal(instr, 'p');
        if (!p) { (instr.parentElement || instr).remove(); removed++; continue; }
        const runs = Array.from(p.children).filter((e) => e.namespaceURI === W && e.localName === 'r');
        let instrIndex = runs.findIndex((r) => r === instr.parentElement || r.contains(instr));
        if (instrIndex < 0) { (instr.parentElement || instr).remove(); removed++; continue; }
        let start = instrIndex;
        while (start >= 0 && !runHasFieldChar(runs[start], 'begin')) start--;
        if (start < 0) start = instrIndex;
        let end = instrIndex;
        while (end < runs.length && !runHasFieldChar(runs[end], 'end')) end++;
        if (end >= runs.length) end = instrIndex;
        for (let i = end; i >= start; i--) runs[i]?.remove();
        removed++;
    }
    return removed;
}
async function cloneAndAssignPart(zip, rels, ct, sect, kind, type, part) {
    const n = nextPartNumber(zip, kind); const fileName = `${kind}${n}.xml`; const newPath = `word/${fileName}`;
    zip.file(newPath, xmlString(part));
    const newRid = nextRid(rels); addRel(rels, newRid, `${R}/${kind}`, fileName);
    addContentType(ct, `/${newPath}`, `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`);
    replaceReferenceType(sect, kind, type, newRid);
}
async function setSectionPart(zip, rels, ct, sect, kind, type, align, withPage, options, numberFormat = 'Arabic') {
    const existing = referenceByType(sect, kind, type);
    if (options.mode === 'safe' && existing) {
        const rid = existing.getAttributeNS(R, 'id') || '';
        const target = relationshipTarget(rels, rid);
        const path = target ? `word/${target.replace(/^\.\//, '')}` : '';
        const file = path ? zip.file(path) : null;
        if (file) {
            // Selalu kloning part: section lain yang berbagi header/footer tidak ikut berubah.
            // PAGE lama dihapus lalu dibuat ulang agar posisi dan format section saat ini benar,
            // sekaligus membuang cached result angka lama dari complex field Word.
            const part = parseXml(await file.async('string'));
            stripPageFieldsFromPart(part);
            if (withPage) appendPageParagraph(part, align, numberFormat);
            await cloneAndAssignPart(zip, rels, ct, sect, kind, type, part);
            return;
        }
    }
    // Total mode atau tidak ada referensi: buat part baru hanya untuk section ini.
    const rid = await addPart(zip, rels, ct, kind, align, withPage);
    replaceReferenceType(sect, kind, type, rid);
}
async function configureUnnumberedSection(zip, rels, ct, sect, options) {
    const pgNum = ensureSectChild(sect, 'pgNumType', ['cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid']);
    wAttr(pgNum, 'fmt', 'decimal'); wAttr(pgNum, 'start', '1');
    directChild(sect, 'titlePg')?.remove();
    await setSectionPart(zip, rels, ct, sect, 'header', 'default', 'right', false, options);
    await setSectionPart(zip, rels, ct, sect, 'footer', 'default', 'center', false, options);
}
async function configureFrontSection(zip, rels, ct, sect, start, options) {
    const pgNum = ensureSectChild(sect, 'pgNumType', ['cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid']);
    wAttr(pgNum, 'fmt', 'lowerRoman'); if (start !== null) wAttr(pgNum, 'start', String(start)); else pgNum.removeAttributeNS(W, 'start');
    directChild(sect, 'titlePg')?.remove();
    await setSectionPart(zip, rels, ct, sect, 'header', 'default', 'right', false, options);
    await setSectionPart(zip, rels, ct, sect, 'footer', 'default', 'center', true, options, 'roman');
}
async function configureMainSection(zip, rels, ct, sect, isChapterStart, restartAtOne, options) {
    const pgNum = ensureSectChild(sect, 'pgNumType', ['cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid']);
    wAttr(pgNum, 'fmt', 'decimal'); if (restartAtOne) wAttr(pgNum, 'start', '1'); else pgNum.removeAttributeNS(W, 'start');
    if (isChapterStart) ensureSectChild(sect, 'titlePg', ['textDirection', 'bidi', 'rtlGutter', 'docGrid']); else directChild(sect, 'titlePg')?.remove();
    await setSectionPart(zip, rels, ct, sect, 'header', 'default', 'right', true, options, 'Arabic');
    await setSectionPart(zip, rels, ct, sect, 'footer', 'default', 'center', false, options);
    if (isChapterStart) {
        await setSectionPart(zip, rels, ct, sect, 'header', 'first', 'right', false, options);
        await setSectionPart(zip, rels, ct, sect, 'footer', 'first', 'center', true, options, 'Arabic');
    }
}
function findFrontNumberStartParagraph(doc, options) {
    const ps = bodyParagraphs(doc); const mainIdx = firstMainParagraphIndex(doc, options); if (mainIdx <= 0) return null;
    const profile = resolveProfile(doc, options); const numberingProfile = options.numberingProfile || 'uin-madura';
    if (numberingProfile === 'simple') return null;
    if (profile.type === 'skripsi') {
        const cover2 = repeatedCoverTitleParagraph(doc, options);
        if (cover2) return cover2;
        return ps.slice(0, mainIdx).find((p) => /^(HALAMAN\s+JUDUL|HALAMAN\s+PERSETUJUAN|LEMBAR\s+PERSETUJUAN|ABSTRAK|KATA\s+PENGANTAR|DAFTAR\s+ISI)$/i.test(paragraphText(p))) || null;
    }
    if (profile.type === 'proposal') {
        return ps.slice(0, mainIdx).find((p) => /^(?:HALAMAN|LEMBAR)\s+(?:PERSETUJUAN|PENGESAHAN)$|^KATA\s+PENGANTAR$|^DAFTAR\s+ISI$/i.test(paragraphText(p))) || null;
    }
    if (profile.type === 'makalah') {
        return ps.slice(0, mainIdx).find((p) => /^KATA\s+PENGANTAR$|^DAFTAR\s+ISI$/i.test(paragraphText(p))) || null;
    }
    return null;
}
function sectionStartParagraphs(doc) {
    const sects = bodySectionProperties(doc); const starts = [];
    starts.push(firstNonEmptyBodyParagraph(doc));
    for (let i = 1; i < sects.length; i++) starts.push(nextNonEmptyParagraphAfterSectPr(doc, sects[i - 1]));
    return starts;
}
async function applyPageNumbering(zip, doc, options, changes, warnings) {
    if (!options.repairPageNumbering) return;
    const profile = resolveProfile(doc, options);
    if (profile.structure === 'with-bab') dedupeSectionBoundariesBeforeChapters(doc, changes);
    const main = firstMainParagraph(doc, options);
    if (!main) { warnings.push('Awal bagian isi tidak dapat ditentukan; penomoran otomatis dilewati agar file tidak rusak.'); return; }
    const frontStart = findFrontNumberStartParagraph(doc, options);
    if (frontStart) ensureSectionBoundaryBefore(doc, frontStart, changes, 'bagian awal bernomor Romawi');
    ensureNumberingBoundaries(doc, options, changes);
    const sects = bodySectionProperties(doc); const starts = sectionStartParagraphs(doc);
    const mainIndex = starts.findIndex((p) => p === main);
    if (mainIndex < 0) { warnings.push('Section awal isi tidak berhasil dipetakan; penomoran otomatis dilewati.'); return; }
    const frontIndex = frontStart ? starts.findIndex((p) => p === frontStart) : -1;
    const rels = await loadRels(zip); const ct = await loadContentTypes(zip);
    const numberingProfile = options.numberingProfile || 'uin-madura';
    for (let i = 0; i < sects.length; i++) {
        if (i < mainIndex && (frontIndex < 0 || i < frontIndex)) {
            await configureUnnumberedSection(zip, rels, ct, sects[i], options);
        } else if (i < mainIndex) {
            const secondCoverStart = profile.type === 'skripsi' && repeatedCoverTitleParagraph(doc, options) === frontStart;
            const startRoman = i === frontIndex ? (secondCoverStart || (numberingProfile === 'zain-legacy' && profile.type === 'skripsi') ? 2 : 1) : null;
            await configureFrontSection(zip, rels, ct, sects[i], startRoman, options);
        } else {
            const startP = starts[i]; const chap = startP ? chapterRoman(paragraphText(startP)) : null;
            const chapterFirst = profile.structure === 'with-bab' && Boolean(chap);
            await configureMainSection(zip, rels, ct, sects[i], chapterFirst, i === mainIndex, options);
        }
    }
    zip.file('word/_rels/document.xml.rels', xmlString(rels)); zip.file('[Content_Types].xml', xmlString(ct));
    if (profile.type === 'skripsi') changes.push('Penomoran skripsi dibuat sebagai field PAGE: sampul tanpa nomor tampak, bagian awal Romawi, isi mulai 1; halaman awal BAB bawah-tengah dan halaman isi lain kanan-atas.');
    else if (profile.type === 'proposal') changes.push(`Penomoran proposal dibuat terpisah: halaman judul/sampul tanpa nomor tampak, bagian awal yang tersedia Romawi, ${profile.structure === 'with-bab' ? 'BAB I' : 'bagian isi pertama'} mulai angka 1.`);
    else changes.push(`Penomoran makalah dibuat terpisah: sampul tanpa nomor tampak, Kata Pengantar/Daftar Isi memakai Romawi bila ada, bagian isi mulai angka 1${profile.structure === 'with-bab' ? '; halaman awal BAB bawah-tengah' : ' di kanan-atas'}.`);
}
async function ensureUpdateFields(zip, changes) {
    const path = 'word/settings.xml';
    const file = zip.file(path);
    if (!file)
        return;
    const doc = parseXml(await file.async('string'));
    const root = doc.documentElement;
    let update = directChild(root, 'updateFields');
    if (!update) {
        update = wEl(doc, 'updateFields');
        root.appendChild(update);
    }
    wAttr(update, 'val', 'true');
    zip.file(path, xmlString(doc));
    if (!changes.some((x) => x.includes('update field')))
        changes.push('Microsoft Word disetel untuk update field otomatis (TOC/PAGE/Daftar Tabel/Daftar Gambar) saat dibuka.');
}
async function zipHasPageField(zip, doc) {
    if (hasPageField(doc))
        return true;
    const names = Object.keys(zip.files).filter((name) => /^word\/(header|footer)\d+\.xml$/i.test(name));
    for (const name of names) {
        const xml = await zip.file(name)?.async('string');
        if (xml && /(?:w:instr="[^"]*\bPAGE\b|<w:instrText[^>]*>[^<]*\bPAGE\b)/i.test(xml))
            return true;
    }
    return false;
}
function packageAssetSnapshot(zip) {
    const names = Object.keys(zip.files).filter((n) => /^word\/(media\/|footnotes\.xml$|endnotes\.xml$|comments\.xml$|embeddings\/)/i.test(n)).sort();
    return names;
}
function packageAssetsEqual(before, zip) {
    const after = packageAssetSnapshot(zip);
    return before.length === after.length && before.every((n, i) => n === after[i]);
}
function normalizedOptions(options = {}) {
    return {
        documentType: 'auto', structureMode: 'auto', numberingProfile: 'uin-madura', mode: 'safe', paper: 'A4',
        marginTopCm: 4, marginRightCm: 3, marginBottomCm: 3, marginLeftCm: 4,
        normalizeMargins: true, repairPageNumbering: true, createToc: true, repairChapterPagination: true, repairHeadings: true,
        normalizeBodyFormatting: false, createTableList: false, createFigureList: false, cleanBlankParagraphs: true, validateDocument: true,
        removeLikelyManualToc: true, includeFrontMatterInToc: true, fontFamily: 'Times New Roman', fontSizePt: 12, lineSpacing: 1.5, firstLineIndentCm: 1.27,
        headingOverrides: {}, ...options,
    };
}
export async function analyzeDocx(file, rawOptions = {}) {
    const options = normalizedOptions(rawOptions);
    const zip = await JSZip.loadAsync(file);
    const doc = await loadDocumentXml(zip);
    const numbering = await loadNumberingMap(zip);
    return analyzeXml(file.name, doc, numbering, await zipHasPageField(zip, doc), options);
}
export async function repairDocx(file, rawOptions = {}) {
    const options = normalizedOptions(rawOptions);
    const zip = await JSZip.loadAsync(file);
    const doc = await loadDocumentXml(zip);
    const numbering = await loadNumberingMap(zip);
    const changes = [];
    const warnings = [];
    const packageBefore = packageAssetSnapshot(zip);
    const contentBefore = summarizeProtectedContent(doc);
    const profile = resolveProfile(doc, options);
    changes.push(`Profil V4: ${profile.type.toUpperCase()} • ${profile.structure === 'with-bab' ? 'Dengan BAB' : 'Tanpa BAB'}${options.documentType === 'auto' ? ` • deteksi ${Math.round(profile.confidence * 100)}%` : ' • dipilih pengguna'}.`);
    if (options.repairPageNumbering || options.repairChapterPagination) {
        if (profile.structure === 'with-bab') runProtectedStage(doc, 'Pembersihan Section Break', changes, warnings, () => dedupeSectionBoundariesBeforeChapters(doc, changes));
    }
    runProtectedStage(doc, 'Margin & ukuran kertas', changes, warnings, () => applyMarginsAndPaper(doc, options, changes));
    runProtectedStage(doc, 'Level TOC & heading', changes, warnings, () => applyHeadingLevels(doc, numbering, options, changes));
    runProtectedStage(doc, 'Halaman awal & struktur utama', changes, warnings, () => applyDocumentPagination(doc, options, changes));
    runProtectedStage(doc, 'Daftar Isi TOC', changes, warnings, () => applyToc(doc, options, changes, warnings));
    runProtectedStage(doc, 'Daftar Tabel/Gambar', changes, warnings, () => applyTableFigureList(doc, options, changes, warnings));
    runProtectedStage(doc, 'Format isi', changes, warnings, () => applyBodyFormatting(doc, numbering, options, changes));
    runProtectedStage(doc, 'Pembersihan Enter', changes, warnings, () => cleanBlankParagraphs(doc, options, changes));
    if (options.repairPageNumbering) {
        try { await applyPageNumbering(zip, doc, options, changes, warnings); }
        catch (e) { warnings.push(`Penomoran otomatis dihentikan karena terjadi kesalahan; modul lain tetap dipertahankan: ${e instanceof Error ? e.message : 'error tidak diketahui'}`); }
    }
    const integrity = validateProtectedContent(contentBefore, doc);
    if (!integrity.ok) {
        throw new Error(`V4 memblokir hasil karena pemeriksaan anti-rusak menemukan konten utama berpotensi hilang (${integrity.missing.length} teks, tabel ${integrity.tablesBefore}→${integrity.tablesAfter}, gambar ${integrity.drawingsBefore}→${integrity.drawingsAfter}). File asli tidak diubah.`);
    }
    if (!packageAssetsEqual(packageBefore, zip)) throw new Error('V4 memblokir hasil karena daftar media/footnote/endnote/embedding berubah secara tidak terduga. File asli tidak diubah.');
    changes.push('Pemeriksaan anti-rusak lulus: teks utama, tabel, drawing, media, footnote/endnote/embedding tetap terjaga.');
    zip.file('word/document.xml', xmlString(doc));
    if (options.createToc || options.repairPageNumbering || options.createTableList || options.createFigureList) await ensureUpdateFields(zip, changes);
    const output = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    const suffix = profile.type === 'skripsi' ? 'SKRIPSI' : profile.type === 'proposal' ? 'PROPOSAL' : 'MAKALAH';
    const outputName = file.name.replace(/\.docx$/i, '') + `_RAPI_${suffix}_ZAINNET_V4.docx`;
    const reloaded = await JSZip.loadAsync(output);
    const afterDoc = await loadDocumentXml(reloaded);
    const afterNumbering = await loadNumberingMap(reloaded);
    const analysisAfter = analyzeXml(outputName, afterDoc, afterNumbering, await zipHasPageField(reloaded, afterDoc), options);
    if (options.validateDocument) changes.push(`Validasi akhir selesai. Skor struktur: ${analysisAfter.validationScore}/100 dengan ${analysisAfter.issues.length} catatan tersisa.`);
    return { blob: output, outputName, changes, warnings, analysisAfter, integrity: { ok: true, protectedText: contentBefore.counts.size, tables: contentBefore.tables, drawings: contentBefore.drawings } };
}
