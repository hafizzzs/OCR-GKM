/**
 * ============================================================================
 *  VERIFIKASI STRUK TIMBANGAN (OCR) vs GOOGLE SPREADSHEET
 *  Rekapitulasi Pembongkaran Raw Sugar – audit No Polisi, Gross, Tare, Netto
 * ============================================================================
 *  Cara kerja singkat:
 *   1. Gambar struk diupload dari index.html (JPG, PNG, HEIC, WEBP, BMP, GIF, PDF).
 *      HEIC & format lain dikonversi ke JPEG di browser, lalu dikirim ke server.
 *   2. Server melakukan OCR memakai Google Drive OCR (tanpa API key, tanpa
 *      Advanced Service) lalu file sementara langsung dihapus.
 *   3. Teks OCR diurai: NO MOBIL, BERAT IN, BERAT OUT, NETTO, NO URUT.
 *   4. No. Surat Jalan diambil dari NAMA FILE: "90_018137.jpg" -> "90/018137".
 *   5. Baris dicari di sheet berdasarkan No. Surat Jalan lalu dibandingkan:
 *        No Polisi  <-> NO MOBIL
 *        Gross      <-> BERAT OUT
 *        Tare       <-> BERAT IN
 *        Netto      <-> NETTO
 *   6. Hasil bisa disimpan ke sheet log, dan (opsional) sel yang selisih
 *      di sheet utama diberi warna + catatan (note).
 * ============================================================================
 */

/***************************** KONFIGURASI ***********************************/
const CONFIG = {
  // Kosongkan jika script dibuat dari Extensions > Apps Script di spreadsheet itu
  // sendiri. Isi dengan ID spreadsheet jika script dibuat terpisah (standalone).
  SPREADSHEET_ID: '',

  // Nama tab yang diperiksa
  SHEET_NAME: 'PEMBONGKARAN GULA (PORT)',

  // Nama tab log hasil verifikasi (dibuat otomatis)
  LOG_SHEET_NAME: 'LOG VERIFIKASI OCR',

  // Baris header dicari otomatis di N baris pertama
  HEADER_SCAN_ROWS: 40,

  // Bahasa OCR Google Drive
  OCR_LANGUAGE: 'id',

  // Toleransi selisih berat (kg). 0 = harus sama persis
  TOLERANSI_KG: 0,

  // Warna penanda sel selisih di sheet utama
  WARNA_SELISIH: '#F4CCCC',
  WARNA_MIRIP: '#FFF2CC',

  // Kolom status yang ditambahkan otomatis di kanan tabel (sekali saja)
  KOLOM_STATUS: ['STATUS OCR', 'KETERANGAN OCR', 'CEK OCR (WAKTU | FILE)'],
  WARNA_STATUS: { 'VALID': '#D9EAD3', 'TIDAK VALID': '#F4CCCC', 'PERLU CEK': '#FFF2CC' },

  // Warnai juga sel kolom "No." agar baris bermasalah terlihat dari kiri
  TANDAI_KOLOM_NO: true
};

// Penanda internal di catatan (note) sel, untuk mengembalikan warna & catatan asli
const OCR_TAG = '[OCR';
const MARK_BG = '⟦bg:';
const MARK_END = '⟧';
const MARK_NOTE = '⟦catatan asli⟧';
const FIELD_LABEL = { polisi: 'No Polisi', gross: 'Gross', tare: 'Tare', netto: 'Netto' };

/* Aturan pengenalan header (huruf kecil). Header tidak perlu di kolom tetap. */
const HEADER_RULES = {
  polisi:     h => h.indexOf('polisi') > -1,
  suratJalan: h => h.indexOf('surat jalan') > -1,
  gross:      h => h.indexOf('gross') === 0,
  tare:       h => h.indexOf('tare') === 0,
  netto:      h => h.indexOf('netto') === 0,
  no:         h => h === 'no.' || h === 'no',
  tanggal:    h => h.indexOf('tanggal') === 0,
  driver:     h => h.indexOf('driver') > -1 || h.indexOf('sopir') > -1
};
const HEADER_WAJIB = ['polisi', 'suratJalan', 'gross', 'tare', 'netto'];

const LABEL_RE = /(BERAT|NETT|MOBIL|URUT|TIME|GUDANG|VESSEL|KEGIATAN|STATUS|JUMLAH|GROUP)/;


/***************************** ENTRY POINT ***********************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔍 Verifikasi OCR')
    .addItem('Buka Verifikasi Struk Timbangan', 'bukaDialog')
    .addSeparator()
    .addItem('Hapus semua tanda OCR di sheet', 'clearAllMarks')
    .addItem('Izinkan Akses (jalankan sekali)', 'authorize')
    .addToUi();
}

function bukaDialog() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(1500)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'Verifikasi Struk Timbangan (OCR)');
}

/** Untuk Deploy > New deployment > Web app */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Verifikasi Struk Timbangan (OCR)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Jalankan SEKALI dari editor agar izin Drive (baca+tulis), Spreadsheet & UrlFetch diberikan.
 * Catatan: Apps Script hanya meminta izin sesuai method yang dipakai di kode.
 * createFile() dipakai di sini supaya izin Drive yang diminta adalah PENUH
 * (bukan read-only), karena OCR perlu membuat file sementara di Drive.
 */
function authorize() {
  // 1) Paksa izin Drive penuh: buat file uji lalu langsung buang ke trash
  const tmp = DriveApp.createFile('TMP_CEK_IZIN_OCR.txt', 'cek izin', MimeType.PLAIN_TEXT);
  tmp.setTrashed(true);

  // 2) Cek akses spreadsheet
  getSheet_().getName();

  // 3) Uji upload + OCR kecil via Drive API (sama persis dengan proses sebenarnya)
  const auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  const r = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=media&fields=id',
    { method: 'post', contentType: 'text/plain', payload: 'tes', headers: auth, muteHttpExceptions: true });
  if (r.getResponseCode() >= 300) {
    throw new Error('Izin Drive masih kurang (' + r.getResponseCode() + '). ' +
      'Hapus akses lama di https://myaccount.google.com/permissions lalu jalankan authorize() lagi. Detail: ' +
      r.getContentText().slice(0, 200));
  }
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + JSON.parse(r.getContentText()).id,
    { method: 'delete', headers: auth, muteHttpExceptions: true });

  Logger.log('✅ Akses OK. Tool siap dipakai.');
  return 'OK';
}


/***************************** API UNTUK index.html **************************/

/** Info sheet yang terdeteksi (ditampilkan di header aplikasi). */
function getInfo() {
  const idx = loadIndex_();
  const names = {};
  Object.keys(idx.cols).forEach(k => {
    names[k] = idx.headerNames[k] + ' (kolom ' + colLetter_(idx.cols[k] + 1) + ')';
  });
  return {
    spreadsheet: idx.sheet.getParent().getName(),
    sheet: idx.sheet.getName(),
    headerRow: idx.headerRow,
    totalBaris: idx.total,
    kolom: names,
    toleransi: CONFIG.TOLERANSI_KG
  };
}

/**
 * Proses 1 gambar: OCR -> urai -> verifikasi.
 * @param {{fileName:string, base64:string, mimeType:string}} payload
 */
function processImage(payload, opts) {
  const t0 = Date.now();
  if (!payload || !payload.base64) throw new Error('Data gambar kosong.');

  const sj = sjFromFileName_(payload.fileName);
  const raw = ocr_(payload.base64, payload.mimeType, payload.fileName);
  const ocr = parseReceipt_(raw);

  const idx = loadIndex_();
  const res = verify_(Object.assign({ fileName: payload.fileName, noSuratJalan: sj }, ocr), idx);
  res.ocr.rawText = raw;
  autoMark_(res, idx, opts);
  res.durasiMs = Date.now() - t0;
  return res;
}

/** Tandai langsung ke sheet bila opsi autoMark aktif. */
function autoMark_(res, idx, opts) {
  if (!opts || !opts.autoMark || !res.sheet) return;
  try { markResults_([res], idx); res.marked = true; }
  catch (e) { res.catatan.push('Gagal menandai sheet: ' + e.message); }
}

/** Tandai hasil ke sheet secara manual (dipanggil tombol "Tandai ke Sheet"). */
function markResults(results) {
  const idx = loadIndex_();
  return { marked: markResults_(results || [], idx) };
}

/** Verifikasi ulang setelah koreksi manual di aplikasi (tanpa OCR ulang). */
function reverify(rec, opts) {
  const r = Object.assign({}, rec);
  ['beratIn', 'beratOut', 'netto'].forEach(k => { r[k] = toNumSheet_(r[k]); });
  r.noMobil = cleanPlate_(r.noMobil) || String(r.noMobil || '').toUpperCase().trim();
  r.konsisten = (r.beratIn != null && r.beratOut != null && r.netto != null)
    ? (r.beratOut - r.beratIn === r.netto) : null;
  r.metode = 'koreksi manual';
  const idx = loadIndex_();
  const res = verify_(r, idx);
  res.manual = true;
  autoMark_(res, idx, opts);
  return res;
}

/**
 * Simpan hasil ke tab log (jejak audit). Penandaan sheet dilakukan terpisah
 * (otomatis saat proses, atau lewat markResults).
 * @param {Object[]} results  hasil dari processImage / reverify
 */
function saveResults(results) {
  if (!results || !results.length) return { saved: 0 };
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const idx = loadIndex_();
    const ss = idx.sheet.getParent();
    const log = getLogSheet_(ss);
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || '-';

    const rows = results.map(r => {
      const c = r.cek || {};
      const g = k => c[k] || {};
      return [
        now, r.fileName, r.noSuratJalan || '', r.sheet ? r.sheet.row : '',
        r.status,
        r.sheet ? r.sheet.polisi : '', r.ocr ? r.ocr.noMobil : '', g('polisi').status || '',
        g('gross').sheet, g('gross').ocr, g('gross').selisih, g('gross').status || '',
        g('tare').sheet, g('tare').ocr, g('tare').selisih, g('tare').status || '',
        g('netto').sheet, g('netto').ocr, g('netto').selisih, g('netto').status || '',
        r.ocr ? r.ocr.noUrut : '',
        r.manual ? 'Ya' : 'Tidak',
        (r.catatan || []).join(' | '),
        r.ocr && r.ocr.rawText ? String(r.ocr.rawText).slice(0, 1000) : '',
        user
      ].map(v => (v === null || v === undefined) ? '' : v);
    });
    log.getRange(log.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { saved: rows.length, logUrl: ss.getUrl() + '#gid=' + log.getSheetId() };
  } finally {
    lock.releaseLock();
  }
}


/***************************** PENANDAAN LANGSUNG DI SHEET *******************/
/**
 * Untuk setiap hasil:
 *  - Sel No Polisi / Gross / Tare / Netto yang SELISIH -> merah, yang MIRIP -> kuning,
 *    + catatan (note) berisi nilai struk vs sheet.
 *  - Sel kolom "No." diwarnai sesuai status (merah/kuning) agar terlihat dari kiri.
 *  - Kolom STATUS OCR | KETERANGAN OCR | CEK OCR diisi (bisa difilter).
 *  - Jika baris sudah VALID (mis. setelah diperbaiki & dicek ulang), tanda lama
 *    dihapus dan warna + catatan asli sel dikembalikan.
 * Baris dicari ulang berdasarkan No. Surat Jalan, jadi aman walau sheet sudah di-sort.
 */
function markResults_(results, idx) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = idx.sheet;
    const sc = ensureStatusCols_(idx);
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    let n = 0;

    results.forEach(r => {
      if (!r || !r.cek || !r.sheet) return;
      const cur = findRow_(idx, r);
      if (!cur) return;
      const row = cur.row;
      const ket = [];

      Object.keys(FIELD_LABEL).forEach(f => {
        const ck = r.cek[f];
        const cell = sh.getRange(row, idx.cols[f] + 1);
        if (ck && (ck.status === 'TIDAK COCOK' || ck.status === 'MIRIP')) {
          const sel = ck.selisih != null && ck.selisih !== 0
            ? ' (Δ ' + (ck.selisih > 0 ? '+' : '') + fmtVal_(ck.selisih) + ' kg)' : '';
          ket.push(FIELD_LABEL[f] + ': sheet ' + fmtVal_(ck.sheet) + ' ≠ struk ' + fmtVal_(ck.ocr) + sel +
            (ck.status === 'MIRIP' ? ' [mirip, cek foto]' : ''));
          markCell_(cell, ck.status === 'MIRIP' ? CONFIG.WARNA_MIRIP : CONFIG.WARNA_SELISIH,
            '[OCR ' + stamp + '] ' + r.fileName +
            '\nStruk : ' + fmtVal_(ck.ocr) + '\nSheet : ' + fmtVal_(ck.sheet) +
            (sel ? '\nSelisih: ' + sel.replace(/[()]/g, '').trim() : '') +
            '\nSeharusnya (sesuai struk): ' + fmtVal_(ck.ocr));
        } else {
          unmarkCell_(cell);
          if (ck && ck.status !== 'COCOK') ket.push(FIELD_LABEL[f] + ': ' + ck.status);
        }
      });
      (r.catatan || []).forEach(c => ket.push(c));

      const color = CONFIG.WARNA_STATUS[r.status] || CONFIG.WARNA_MIRIP;
      sh.getRange(row, sc.status, 1, 3).setValues([[
        r.status + (r.manual ? ' (manual)' : ''),
        ket.length ? ket.join(' | ') : 'Semua sesuai struk',
        stamp + ' | ' + r.fileName
      ]]);
      sh.getRange(row, sc.status).setBackground(color).setFontWeight('bold');

      if (CONFIG.TANDAI_KOLOM_NO && idx.cols.no !== undefined) {
        const noCell = sh.getRange(row, idx.cols.no + 1);
        if (r.status === 'VALID') unmarkCell_(noCell);
        else markCell_(noCell, color, '[OCR ' + stamp + '] ' + r.status + '\n' + ket.join('\n'));
      }
      n++;
    });
    SpreadsheetApp.flush();
    return n;
  } finally {
    lock.releaseLock();
  }
}

/** Cari baris terkini (sheet bisa saja sudah di-sort / disisipi baris). */
function findRow_(idx, r) {
  const rows = idx.map[normSJ_(r.noSuratJalan)] || [];
  if (!rows.length) return null;
  if (rows.length === 1) return rows[0];
  const prev = r.sheet ? plateKey_(r.sheet.polisi) : '';
  return rows.find(x => plateKey_(x.polisi) === prev) ||
         rows.find(x => plateKey_(x.polisi) === plateKey_(r.ocr && r.ocr.noMobil)) || rows[0];
}

/** Warnai sel + catatan, simpan warna & catatan asli agar bisa dikembalikan. */
function markCell_(cell, color, text) {
  const old = String(cell.getNote() || '');
  let bg, orig;
  if (old.indexOf(OCR_TAG) === 0) { const p = parseMark_(old); bg = p.bg; orig = p.note; }
  else { bg = cell.getBackground(); orig = old; }
  cell.setBackground(color)
      .setNote(text + '\n' + MARK_BG + bg + MARK_END + (orig ? '\n' + MARK_NOTE + '\n' + orig : ''));
}

/** Kembalikan warna & catatan asli jika sel pernah ditandai tool ini. */
function unmarkCell_(cell) {
  const old = String(cell.getNote() || '');
  if (old.indexOf(OCR_TAG) !== 0) return;
  const p = parseMark_(old);
  cell.setBackground(p.bg && p.bg.toLowerCase() !== '#ffffff' ? p.bg : null).setNote(p.note || '');
}

function parseMark_(note) {
  const m = note.match(/⟦bg:([^⟧]*)⟧/);
  const i = note.indexOf(MARK_NOTE);
  return {
    bg: m ? m[1] : '#ffffff',
    note: i > -1 ? note.slice(i + MARK_NOTE.length).replace(/^\n/, '') : ''
  };
}

/** Pastikan kolom STATUS OCR ada di baris header (dibuat sekali di kanan tabel). */
function ensureStatusCols_(idx) {
  const sh = idx.sheet, hr = idx.headerRow, names = CONFIG.KOLOM_STATUS;
  const width = sh.getLastColumn();
  const hv = sh.getRange(hr, 1, 1, width).getDisplayValues()[0].map(v => String(v).trim().toUpperCase());
  const found = hv.indexOf(names[0].toUpperCase());
  if (found > -1) return { status: found + 1, ket: found + 2, waktu: found + 3 };

  if (sh.getMaxColumns() < width + 3) sh.insertColumnsAfter(sh.getMaxColumns(), width + 3 - sh.getMaxColumns());
  sh.getRange(hr, width + 1, 1, 3).setValues([names]).setFontWeight('bold')
    .setBackground('#1F3A5F').setFontColor('#FFFFFF').setWrap(true)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setColumnWidth(width + 1, 130);
  sh.setColumnWidth(width + 2, 420);
  sh.setColumnWidth(width + 3, 220);
  extendFilter_(sh, width + 3);
  return { status: width + 1, ket: width + 2, waktu: width + 3 };
}

/** Perluas filter yang ada agar kolom STATUS OCR ikut bisa difilter (kriteria lama dipertahankan). */
function extendFilter_(sh, lastCol) {
  try {
    const f = sh.getFilter();
    if (!f) return;
    const r = f.getRange();
    if (r.getLastColumn() >= lastCol) return;
    const crit = {};
    for (let c = r.getColumn(); c <= r.getLastColumn(); c++) {
      const cr = f.getColumnFilterCriteria(c);
      if (cr) crit[c] = cr;
    }
    const row = r.getRow();
    const nRows = Math.max(r.getNumRows(), sh.getLastRow() - row + 1);
    f.remove();
    const nf = sh.getRange(row, r.getColumn(), nRows, lastCol - r.getColumn() + 1).createFilter();
    Object.keys(crit).forEach(c => nf.setColumnFilterCriteria(Number(c), crit[c]));
  } catch (e) {
    console.warn('Filter tidak diperluas: ' + e.message);
  }
}

/** Menu: hapus semua tanda dari tool ini dan kembalikan tampilan asli. */
function clearAllMarks() {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert('Hapus semua tanda OCR?',
    'Warna & catatan dari tool OCR dikembalikan seperti semula, dan kolom STATUS OCR dikosongkan.',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  const idx = loadIndex_(), sh = idx.sheet;
  const first = idx.headerRow + 1, n = sh.getLastRow() - idx.headerRow;
  if (n < 1) return;
  let cnt = 0;
  ['no', 'polisi', 'gross', 'tare', 'netto'].filter(k => idx.cols[k] !== undefined).forEach(k => {
    const rng = sh.getRange(first, idx.cols[k] + 1, n, 1);
    const notes = rng.getNotes(), bgs = rng.getBackgrounds();
    let changed = false;
    notes.forEach((nt, i) => {
      if (String(nt[0]).indexOf(OCR_TAG) === 0) {
        const p = parseMark_(String(nt[0]));
        bgs[i][0] = p.bg; notes[i][0] = p.note; changed = true; cnt++;
      }
    });
    if (changed) { rng.setBackgrounds(bgs); rng.setNotes(notes); }
  });

  const hv = sh.getRange(idx.headerRow, 1, 1, sh.getLastColumn()).getDisplayValues()[0]
    .map(v => String(v).trim().toUpperCase());
  const s = hv.indexOf(CONFIG.KOLOM_STATUS[0].toUpperCase());
  if (s > -1) sh.getRange(first, s + 1, n, 3).clearContent().setBackground(null).setFontWeight('normal');
  ui.alert('Selesai: ' + cnt + ' sel dikembalikan seperti semula.');
}


/***************************** OCR (GOOGLE DRIVE) ****************************/
function ocr_(base64, mimeType, fileName) {
  const bytes = Utilities.base64Decode(base64);
  const boundary = 'ocrBoundary' + Utilities.getUuid().replace(/-/g, '');
  const meta = { name: 'TMP_OCR_' + (fileName || 'struk'), mimeType: 'application/vnd.google-apps.document' };
  const head = Utilities.newBlob(
    '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(meta) + '\r\n--' + boundary + '\r\nContent-Type: ' + (mimeType || 'image/jpeg') + '\r\n\r\n'
  ).getBytes();
  const tail = Utilities.newBlob('\r\n--' + boundary + '--').getBytes();
  const auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };

  const up = fetchRetry_(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id&ocrLanguage=' +
      encodeURIComponent(CONFIG.OCR_LANGUAGE),
    { method: 'post', contentType: 'multipart/related; boundary=' + boundary,
      payload: head.concat(bytes, tail), headers: auth, muteHttpExceptions: true }
  );
  if (up.getResponseCode() >= 300) {
    throw new Error('OCR gagal (' + up.getResponseCode() + '): ' + up.getContentText().slice(0, 250) +
      ' — Pastikan fungsi authorize() sudah dijalankan.');
  }
  const id = JSON.parse(up.getContentText()).id;
  try {
    const ex = fetchRetry_('https://www.googleapis.com/drive/v3/files/' + id + '/export?mimeType=text%2Fplain',
      { headers: auth, muteHttpExceptions: true });
    if (ex.getResponseCode() >= 300) throw new Error('Gagal membaca hasil OCR (' + ex.getResponseCode() + ')');
    return ex.getContentText('UTF-8').replace(/^\uFEFF/, '');
  } finally {
    // Hapus permanen file sementara agar Drive tetap bersih
    UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id,
      { method: 'delete', headers: auth, muteHttpExceptions: true });
  }
}

function fetchRetry_(url, opt) {
  let resp;
  for (let i = 0; i < 4; i++) {
    resp = UrlFetchApp.fetch(url, opt);
    const code = resp.getResponseCode();
    if (code !== 429 && code < 500) return resp;
    Utilities.sleep(800 * Math.pow(2, i));
  }
  return resp;
}


/***************************** PENGURAI STRUK ********************************/
/**
 * Mengurai teks OCR struk "SURAT JALAN GATE PTP (GATE OUT)".
 * Strategi:
 *  1) Ambil nilai setelah label (NO MOBIL, BERAT IN, BERAT OUT, NETTO).
 *  2) Jika tidak lengkap / tidak konsisten (OUT − IN ≠ NETTO), cari kombinasi
 *     angka "xx.xxx KG" yang memenuhi OUT − IN = NETTO (fallback aritmetika).
 */
function parseReceipt_(raw) {
  const text = String(raw || '').toUpperCase().replace(/\r/g, '');
  const lines = text.split('\n').map(s => s.replace(/\s+/g, ' ').trim()).filter(String);
  const res = { noMobil: '', beratIn: null, beratOut: null, netto: null, noUrut: '', konsisten: null, metode: 'label' };

  // No Mobil
  res.noMobil = cleanPlate_(valueAfter_(lines, /N[O0]\.?\s*M[O0]B[I1L]L/));
  if (!res.noMobil) {
    // OCR kadang memisah label & nilai per kolom: cari baris yang berbentuk plat
    for (const ln of lines) {
      const s = ln.replace(/^[\s:;.\-]+/, '');
      if ((s.match(/\d/g) || []).length < 2) continue;
      const p = cleanPlate_(s);
      const parts = p.split(' ');
      if (p && parts.length === 3 && parts[1].length >= 3) { res.noMobil = p; res.metode = 'pola plat'; break; }
    }
  }

  // Berat
  res.beratIn  = toKg_(valueAfter_(lines, /BERAT\s*[I1L]N\b/));
  res.beratOut = toKg_(valueAfter_(lines, /BERAT\s*[O0]UT\b/));
  res.netto    = toKg_(valueAfter_(lines, /NETT[O0]\b/));
  const hint = { in: res.beratIn, out: res.beratOut, netto: res.netto };

  // No Urut
  const u = fixDigits_(text).match(/N[O0]\.?\s*URUT\s*[:;]?\s*(\d{1,5})/);
  if (u) res.noUrut = String(parseInt(u[1], 10));

  const ok = (o, i, n) => o != null && i != null && n != null && o - i === n;
  if (!ok(res.beratOut, res.beratIn, res.netto)) {
    const t = findTriple_(allKg_(text), hint);
    if (t) {
      res.beratOut = t.out; res.beratIn = t.in; res.netto = t.netto;
      res.metode = res.metode === 'label' ? 'aritmetika' : res.metode + ' + aritmetika';
    }
  }
  if (res.beratOut != null && res.beratIn != null && res.netto != null) {
    res.konsisten = ok(res.beratOut, res.beratIn, res.netto);
  }
  return res;
}

function valueAfter_(lines, re) {
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(re);
    if (!m) continue;
    const rest = lines[i].slice(m.index + m[0].length).replace(/^[\s:;.,=\-]+/, '').trim();
    if (/[A-Z0-9]/.test(rest)) return rest;
    const nx = lines[i + 1] || '';
    if (nx && !LABEL_RE.test(nx)) return nx.replace(/^[\s:;.,=\-]+/, '').trim();
  }
  return '';
}

/** Perbaiki huruf O yang terbaca di tengah angka (mis. 15.78O -> 15.780). */
function fixDigits_(s) {
  return String(s || '').replace(/(?<=\d[.,]?)O/g, '0').replace(/O(?=[.,]?\d)/g, '0');
}

function toKg_(s) {
  if (!s) return null;
  const m = fixDigits_(s).match(/\d{1,3}(?:[.,\s]\d{3})+|\d{4,6}/);
  if (!m) return null;
  const v = parseInt(m[0].replace(/\D/g, ''), 10);
  return (v >= 100 && v <= 200000) ? v : null;
}

function allKg_(text) {
  const out = [];
  const re = /(\d{1,3}(?:[.,]\d{3})+|\d{4,6})\s*K\s*[G6]/g;
  const t = fixDigits_(text);
  let m;
  while ((m = re.exec(t)) !== null) {
    const v = parseInt(m[1].replace(/\D/g, ''), 10);
    if (v >= 1000 && v <= 200000) out.push(v);
  }
  return out;
}

/**
 * Cari kombinasi OUT − IN = NETTO dari daftar angka "KG".
 * Karena OUT − A = B dan OUT − B = A sama-sama valid, pilih yang paling
 * cocok dengan nilai berlabel (hint), lalu urutan cetak di struk (IN sebelum NETTO).
 */
function findTriple_(vals, hint) {
  hint = hint || {};
  const u = Array.from(new Set(vals));
  const pos = v => vals.indexOf(v);
  let best = null, bestScore = -1;
  for (const o of u) {
    for (const i of u) {
      if (i >= o) continue;
      const n = o - i;
      if (n === i || u.indexOf(n) < 0) continue;
      let sc = 0;
      if (hint.in === i) sc += 4;
      if (hint.netto === n) sc += 4;
      if (hint.out === o) sc += 2;
      if (pos(i) < pos(n)) sc += 1;
      if (pos(i) < pos(o) && pos(o) < pos(n)) sc += 1;
      if (sc > bestScore || (sc === bestScore && o > best.out)) { best = { out: o, in: i, netto: n }; bestScore = sc; }
    }
  }
  return best;
}

/** Normalisasi plat Indonesia: "b9018tyz" / "B 9O18 TYZ" -> "B 9018 TYZ" */
function cleanPlate_(s) {
  if (!s) return '';
  let t = String(s).toUpperCase().replace(/[^A-Z0-9]/g, '');
  t = t.replace(/^8(?=\d)/, 'B');
  const m = t.match(/^([A-Z]{1,2})([0-9OILSBZDQG]{1,4})([A-Z0-9]{0,3})/);
  if (!m) return '';
  const mid = m[2].replace(/[ODQ]/g, '0').replace(/[IL]/g, '1').replace(/S/g, '5')
    .replace(/B/g, '8').replace(/Z/g, '2').replace(/G/g, '6');
  if (!/^\d+$/.test(mid)) return '';
  const suf = m[3].replace(/0/g, 'O').replace(/1/g, 'I').replace(/5/g, 'S')
    .replace(/8/g, 'B').replace(/2/g, 'Z').replace(/6/g, 'G').replace(/[^A-Z]/g, '');
  return (m[1] + ' ' + mid + ' ' + suf).trim();
}


/***************************** DATA SPREADSHEET ******************************/
function getSheet_() {
  const ss = CONFIG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet tidak ditemukan. Isi CONFIG.SPREADSHEET_ID di Code.gs.');
  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) {
    throw new Error('Tab "' + CONFIG.SHEET_NAME + '" tidak ada. Tab tersedia: ' +
      ss.getSheets().map(s => s.getName()).join(', '));
  }
  return sh;
}

function loadIndex_() {
  const sh = getSheet_();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 1) throw new Error('Sheet kosong.');

  const scan = Math.min(CONFIG.HEADER_SCAN_ROWS, lastRow);
  const head = sh.getRange(1, 1, scan, lastCol).getDisplayValues();
  let headerRow = -1, cols = {}, headerNames = {};
  for (let r = 0; r < head.length; r++) {
    const c = {}, n = {};
    head[r].forEach((v, i) => {
      const h = String(v).toLowerCase().replace(/\s+/g, ' ').trim();
      if (!h) return;
      Object.keys(HEADER_RULES).forEach(k => {
        if (c[k] === undefined && HEADER_RULES[k](h)) { c[k] = i; n[k] = String(v).replace(/\s+/g, ' ').trim(); }
      });
    });
    if (HEADER_WAJIB.every(k => c[k] !== undefined)) { headerRow = r + 1; cols = c; headerNames = n; break; }
  }
  if (headerRow < 0) {
    throw new Error('Header tidak ditemukan. Pastikan ada kolom: No Polisi, No. Surat Jalan, Gross, Tare, Netto dalam satu baris.');
  }

  const map = {};
  let total = 0;
  const n = lastRow - headerRow;
  if (n > 0) {
    const rng = sh.getRange(headerRow + 1, 1, n, lastCol);
    const disp = rng.getDisplayValues();
    const vals = rng.getValues();
    for (let i = 0; i < n; i++) {
      const key = normSJ_(disp[i][cols.suratJalan]);
      if (!key) continue;
      total++;
      const rec = {
        row: headerRow + 1 + i,
        no: cols.no !== undefined ? disp[i][cols.no] : '',
        tanggal: cols.tanggal !== undefined ? disp[i][cols.tanggal] : '',
        driver: cols.driver !== undefined ? disp[i][cols.driver] : '',
        polisi: String(disp[i][cols.polisi] || '').trim(),
        suratJalan: String(disp[i][cols.suratJalan] || '').trim(),
        gross: toNumSheet_(vals[i][cols.gross]),
        tare: toNumSheet_(vals[i][cols.tare]),
        netto: toNumSheet_(vals[i][cols.netto])
      };
      (map[key] = map[key] || []).push(rec);
    }
  }
  return { sheet: sh, headerRow: headerRow, cols: cols, headerNames: headerNames, map: map, total: total };
}

function getLogSheet_(ss) {
  let log = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!log) {
    log = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    const h = ['Waktu Verifikasi', 'Nama File', 'No. Surat Jalan', 'Baris Sheet', 'Status Akhir',
      'No Polisi (Sheet)', 'No Mobil (Struk)', 'Status Polisi',
      'Gross (Sheet)', 'Berat Out (Struk)', 'Selisih Gross', 'Status Gross',
      'Tare (Sheet)', 'Berat In (Struk)', 'Selisih Tare', 'Status Tare',
      'Netto (Sheet)', 'Netto (Struk)', 'Selisih Netto', 'Status Netto',
      'No Urut Struk', 'Koreksi Manual', 'Catatan', 'Teks OCR (ringkas)', 'Diverifikasi Oleh'];
    log.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold')
      .setBackground('#1F3A5F').setFontColor('#FFFFFF').setWrap(true);
    log.setFrozenRows(1);
    log.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
    ['I:K', 'M:O', 'Q:S'].forEach(a => log.getRange(a).setNumberFormat('#,##0'));
  }
  return log;
}


/***************************** VERIFIKASI ************************************/
function verify_(rec, idx) {
  const out = {
    fileName: rec.fileName,
    noSuratJalan: rec.noSuratJalan || '',
    ocr: {
      noMobil: rec.noMobil || '', beratIn: rec.beratIn, beratOut: rec.beratOut, netto: rec.netto,
      noUrut: rec.noUrut || '', konsisten: rec.konsisten, metode: rec.metode || ''
    },
    sheet: null, cek: {}, status: '', catatan: []
  };

  const key = normSJ_(rec.noSuratJalan);
  if (!key) {
    out.status = 'NAMA FILE TIDAK VALID';
    out.catatan.push('Nama file harus berformat <no>_<nomor> atau <no>/<nomor>, contoh: 90_018137.jpg');
    return out;
  }

  const rows = idx.map[key] || [];
  if (!rows.length) {
    out.status = 'TIDAK DITEMUKAN';
    out.catatan.push('No. Surat Jalan ' + rec.noSuratJalan + ' tidak ada di sheet "' + CONFIG.SHEET_NAME + '"');
    return out;
  }

  let row = rows[0];
  if (rows.length > 1) {
    const byPlate = rows.find(r => plateKey_(r.polisi) === plateKey_(rec.noMobil));
    if (byPlate) row = byPlate;
    out.catatan.push('No. Surat Jalan duplikat di sheet (baris ' + rows.map(r => r.row).join(', ') + ')');
  }
  out.sheet = row;

  out.cek.polisi = cmpPlate_(row.polisi, rec.noMobil);
  out.cek.gross = cmpKg_(row.gross, rec.beratOut);
  out.cek.tare = cmpKg_(row.tare, rec.beratIn);
  out.cek.netto = cmpKg_(row.netto, rec.netto);

  if (row.gross != null && row.tare != null && row.netto != null && row.gross - row.tare !== row.netto) {
    out.catatan.push('Sheet: Gross − Tare ≠ Netto (beda ' + (row.gross - row.tare - row.netto) + ' kg)');
  }
  if (rec.konsisten === false) {
    out.catatan.push('Struk: Berat Out − Berat In ≠ Netto (kemungkinan salah baca OCR, cek manual)');
  }
  if (rec.noUrut) {
    const pref = String(rec.noSuratJalan).split('/')[0];
    if (parseInt(pref, 10) !== parseInt(rec.noUrut, 10)) {
      out.catatan.push('Info: No Urut struk (' + rec.noUrut + ') berbeda dengan awalan nama file (' + pref + ')');
    }
  }

  const st = Object.keys(out.cek).map(k => out.cek[k].status);
  if (st.indexOf('TIDAK COCOK') > -1) out.status = 'TIDAK VALID';
  else if (st.every(s => s === 'COCOK')) out.status = 'VALID';
  else out.status = 'PERLU CEK';
  return out;
}

function cmpKg_(sheetVal, ocrVal) {
  if (ocrVal == null) return { status: 'TIDAK TERBACA', sheet: sheetVal, ocr: null, selisih: null };
  if (sheetVal == null) return { status: 'KOSONG DI SHEET', sheet: null, ocr: ocrVal, selisih: null };
  const d = ocrVal - sheetVal;
  return { status: Math.abs(d) <= CONFIG.TOLERANSI_KG ? 'COCOK' : 'TIDAK COCOK', sheet: sheetVal, ocr: ocrVal, selisih: d };
}

function cmpPlate_(sheetVal, ocrVal) {
  if (!ocrVal) return { status: 'TIDAK TERBACA', sheet: sheetVal, ocr: '' };
  if (!sheetVal) return { status: 'KOSONG DI SHEET', sheet: '', ocr: ocrVal };
  const a = plateKey_(sheetVal), b = plateKey_(ocrVal);
  if (a === b) return { status: 'COCOK', sheet: sheetVal, ocr: ocrVal };
  if (plateCanon_(a) === plateCanon_(b)) return { status: 'MIRIP', sheet: sheetVal, ocr: ocrVal };
  return { status: 'TIDAK COCOK', sheet: sheetVal, ocr: ocrVal };
}


/***************************** UTILITAS **************************************/
/**
 * Pemisah No. Surat Jalan yang diterima di nama file:
 *  "_"  -> 90_018137.jpg
 *  "/"  -> 90/018137.jpg (nama file di Google Drive)
 *  ":"  -> macOS menyimpan "/" dari Finder sebagai ":"
 *  "%2F"-> nama file hasil unduhan yang ter-encode
 *  "∕", "／", "⁄" (slash Unicode), "\", "-", spasi
 */
const SJ_SEP = '[_\\/\\\\:∕／⁄\\-\\s]';
const SJ_RE = new RegExp('(\\d+)\\s*' + SJ_SEP + '+\\s*(\\d+)');

/** "90_018137.jpg" / "90/018137.jpg" / "90:018137.jpg" -> "90/018137" */
function sjFromFileName_(name) {
  let base = String(name || '').replace(/%2F/gi, '/');
  base = base.replace(/\.(jpe?g|jfif|png|hei[cf]|webp|bmp|gif|tiff?|pdf)$/i, '');
  const m = base.match(SJ_RE);
  return m ? m[1] + '/' + m[2] : '';
}

/** Kunci pencocokan SJ: "16/18063" dan "16/018063" dianggap sama -> "16/18063" */
function normSJ_(s) {
  const m = String(s || '').replace(/%2F/gi, '/').match(SJ_RE);
  return m ? parseInt(m[1], 10) + '/' + parseInt(m[2], 10) : '';
}

function toNumSheet_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return Math.round(v);
  const d = String(v).replace(/\D/g, '');
  return d ? parseInt(d, 10) : null;
}

function plateKey_(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }

function plateCanon_(s) {
  return plateKey_(s).replace(/0/g, 'O').replace(/1/g, 'I').replace(/5/g, 'S')
    .replace(/8/g, 'B').replace(/2/g, 'Z').replace(/6/g, 'G');
}

function fmtVal_(v) { return (typeof v === 'number') ? v.toLocaleString('id-ID') : String(v || '-'); }

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
