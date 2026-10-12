/**
 * BACKEND REKAP — Kuis Sifat Koligatif Larutan
 * Cara pakai:
 * 1. Buat Google Sheet baru → menu Ekstensi → Apps Script → tempel seluruh kode ini.
 * 2. Ganti ADMIN_PIN di bawah (samakan dengan CONFIG.ADMIN_PIN di file HTML).
 * 3. Deploy → Deployment baru → Jenis: Aplikasi web
 *    Jalankan sebagai: Saya  |  Yang memiliki akses: Siapa saja → Deploy.
 * 4. Salin URL Web App (berakhiran /exec) ke CONFIG.SHEET_URL pada file HTML.
 * Setiap kali kode ini diubah, buat "Deployment baru" lagi agar perubahan aktif.
 */
const ADMIN_PIN = 'guru2026';
const NAMA_SHEET = 'Hasil';
const HEAD = ['ID', 'Waktu', 'Nama', 'Kelas', 'Absen', 'Benar Konsep', 'Benar Hitung', 'Nilai', 'Durasi (detik)', 'Pelanggaran', 'Keterangan', 'JSON'];

function doGet() {
  return out_({ ok: true, info: 'Backend kuis aktif' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = JSON.parse((e.postData && e.postData.contents) || '{}');
    switch (p.action) {
      case 'settings':
        return out_({ ok: true, settings: getSet_() });
      case 'save':
        save_(p.result);
        return out_({ ok: true });
      case 'list':
        auth_(p);
        return out_({ ok: true, results: list_() });
      case 'delete':
        auth_(p);
        del_(p.id);
        return out_({ ok: true });
      case 'clear':
        auth_(p);
        clear_();
        return out_({ ok: true });
      case 'saveSettings':
        auth_(p);
        PropertiesService.getScriptProperties().setProperty('settings', JSON.stringify(p.settings || {}));
        return out_({ ok: true });
      default:
        return out_({ ok: false, error: 'Aksi tidak dikenal' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function auth_(p) {
  if (String(p.pin) !== ADMIN_PIN) throw new Error('PIN salah');
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(NAMA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(NAMA_SHEET);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Cegah teks siswa terbaca sebagai rumus Sheet (=, +, -, @)
function safe_(v) {
  v = String(v == null ? '' : v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function getSet_() {
  const s = PropertiesService.getScriptProperties().getProperty('settings');
  return s ? JSON.parse(s) : {};
}

function save_(r) {
  if (!r || !r.id) throw new Error('Data kosong');
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    const ids = sh.getRange(2, 1, n, 1).getValues().map(function (x) { return x[0]; });
    if (ids.indexOf(r.id) > -1) return; // sudah tersimpan (kirim ulang)
  }
  sh.appendRow([
    r.id, new Date(r.ts), safe_(r.nama), safe_(r.kelas), safe_(r.absen),
    r.bK, r.bH, r.nilai, r.dur, r.viol, safe_(r.auto), JSON.stringify(r)
  ]);
}

function list_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 12, n, 1).getValues().map(function (x) {
    try { return JSON.parse(x[0]); } catch (e) { return null; }
  }).filter(Boolean);
}

function del_(id) {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const ids = sh.getRange(2, 1, n, 1).getValues().map(function (x) { return x[0]; });
  const i = ids.indexOf(id);
  if (i > -1) sh.deleteRow(i + 2);
}

function clear_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n > 0) sh.deleteRows(2, n);
}
