/**
 * =========================================================
 *  VOLI-AI — Backend API (Google Apps Script + Google Sheets + Drive)
 *  File ini TIDAK lagi menyajikan index.html.
 *  index.html di-hosting terpisah di Vercel (supaya kamera
 *  bisa diakses langsung, tidak lewat iframe sandbox Apps Script).
 *
 *  Frontend memanggil backend ini lewat fetch() ke URL deploy
 *  Web App (…/exec), dengan parameter "action".
 *
 *  VIDEO CONTOH: guru mengunggah video dari galeri/kamera HP,
 *  frontend mengirim file itu sebagai base64 ke backend ini.
 *  Video disimpan sebagai file di Google Drive (bukan di sel
 *  Sheet — sel teks tidak muat untuk file video), lalu link-nya
 *  disimpan di sheet "Config" supaya semua siswa (dari HP mana
 *  pun) memutar video contoh yang sama.
 *
 *  Sheets yang dipakai (dibuat otomatis kalau belum ada):
 *   - "Riwayat" : ID | Timestamp | Nama | Kelas | Teknik | Skor | Status | Detail(JSON)
 *   - "Config"  : Teknik | FileId | VideoURL
 * =========================================================
 */

const SHEET_RIWAYAT = 'Riwayat';
const SHEET_CONFIG = 'Config';

// Nama folder Drive tempat video contoh disimpan (dibuat otomatis)
const DRIVE_FOLDER_NAME = 'VOLI-AI Video Contoh';

// Batas ukuran file video yang diterima (MB). Apps Script punya batas
// ukuran request; jaga video tetap pendek & terkompresi.
const MAX_VIDEO_MB = 20;

// GANTI password admin di sini
const ADMIN_PASSWORD = 'admin123';

const TEKNIK_LIST = ['Passing Bawah', 'Passing Atas', 'Servis Bawah'];

/* ============================================================
   ENTRY POINTS (dipanggil lewat fetch dari index.html di Vercel)
============================================================ */

// Semua request "baca" (GET): ?action=config | attempts | stats
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || '';
    let result;
    if (action === 'config') {
      result = getConfig();
    } else if (action === 'attempts') {
      checkAdmin(e.parameter.password);
      result = getAttempts({ kelas: e.parameter.kelas, teknik: e.parameter.teknik });
    } else if (action === 'stats') {
      checkAdmin(e.parameter.password);
      result = getStats();
    } else {
      result = { info: 'VOLI-AI API aktif. Gunakan ?action=config / attempts / stats.' };
    }
    return jsonOutput(result);
  } catch (err) {
    return jsonOutput({ error: err.message });
  }
}

// Semua request "tulis" (POST), body berupa JSON text/plain:
// { action: 'saveAttempt' | 'uploadVideo' | 'deleteVideo' | 'adminLogin' | 'deleteAttempt', ... }
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;
    let result;
    if (action === 'saveAttempt') {
      result = saveAttempt(body.payload || {});
    } else if (action === 'uploadVideo') {
      checkAdmin(body.password);
      result = uploadVideo(body.teknik, body.data, body.mimeType, body.filename);
    } else if (action === 'deleteVideo') {
      checkAdmin(body.password);
      result = deleteVideo(body.teknik);
    } else if (action === 'adminLogin') {
      result = { ok: body.password === ADMIN_PASSWORD };
    } else if (action === 'deleteAttempt') {
      checkAdmin(body.password);
      result = deleteAttemptById(body.id);
    } else {
      result = { error: 'Aksi tidak dikenal: ' + action };
    }
    return jsonOutput(result);
  } catch (err) {
    return jsonOutput({ error: err.message });
  }
}

function checkAdmin(password) {
  if (password !== ADMIN_PASSWORD) {
    throw new Error('Password admin salah.');
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- Helper Sheet ---------------- */

function _ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _getOrCreateSheet(name, headers) {
  const ss = _ss();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _riwayatSheet() {
  return _getOrCreateSheet(SHEET_RIWAYAT, ['ID', 'Timestamp', 'Nama', 'Kelas', 'Teknik', 'Skor', 'Status', 'Detail']);
}

function _configSheet() {
  return _getOrCreateSheet(SHEET_CONFIG, ['Teknik', 'FileId', 'VideoURL']);
}

/* ---------------- Config & Video (Google Drive) ---------------- */

function getConfig() {
  const sh = _configSheet();
  const data = sh.getDataRange().getValues();
  const cfg = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) cfg[data[i][0]] = data[i][2] || '';
  }
  TEKNIK_LIST.forEach(function (t) {
    if (!cfg[t]) cfg[t] = '';
  });
  return cfg;
}

function _findConfigRow(sh, teknik) {
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === teknik) return { rowIndex: i + 1, fileId: data[i][1] };
  }
  return null;
}

function _getVideoFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

// teknik: nama teknik: base64Data: isi file video (base64, tanpa prefix data:...)
// mimeType: mis. 'video/mp4', filename: nama file asli (opsional)
function uploadVideo(teknik, base64Data, mimeType, filename) {
  if (TEKNIK_LIST.indexOf(teknik) === -1) throw new Error('Teknik tidak dikenal: ' + teknik);
  if (!base64Data) throw new Error('Data video kosong.');

  const approxBytes = Math.floor(base64Data.length * 0.75);
  if (approxBytes > MAX_VIDEO_MB * 1024 * 1024) {
    throw new Error('Ukuran video melebihi batas ' + MAX_VIDEO_MB + ' MB. Kompres/perpendek videonya dulu.');
  }

  const sh = _configSheet();
  const existing = _findConfigRow(sh, teknik);

  // hapus file lama di Drive supaya tidak menumpuk
  if (existing && existing.fileId) {
    try { DriveApp.getFileById(existing.fileId).setTrashed(true); } catch (e) { /* file lama sudah tidak ada, abaikan */ }
  }

  const folder = _getVideoFolder();
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType || 'video/mp4', (filename || teknik) + '.mp4');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const url = 'https://drive.google.com/file/d/' + fileId + '/preview';

  if (existing) {
    sh.getRange(existing.rowIndex, 2, 1, 2).setValues([[fileId, url]]);
  } else {
    sh.appendRow([teknik, fileId, url]);
  }

  return { ok: true, config: getConfig() };
}

function deleteVideo(teknik) {
  const sh = _configSheet();
  const existing = _findConfigRow(sh, teknik);
  if (existing && existing.fileId) {
    try { DriveApp.getFileById(existing.fileId).setTrashed(true); } catch (e) { /* abaikan */ }
    sh.getRange(existing.rowIndex, 2, 1, 2).setValues([['', '']]);
  }
  return { ok: true, config: getConfig() };
}

/* ---------------- Riwayat percobaan siswa ---------------- */

function saveAttempt(payload) {
  const sh = _riwayatSheet();
  const id = Utilities.getUuid();
  const ts = new Date();
  sh.appendRow([
    id,
    ts,
    payload.nama || '-',
    payload.kelas || '-',
    payload.teknik || '-',
    payload.skor || 0,
    payload.status || '-',
    JSON.stringify(payload.detail || [])
  ]);
  return { ok: true, id: id };
}

function getAttempts(filter) {
  const sh = _riwayatSheet();
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const item = {
      id: row[0],
      timestamp: row[1] instanceof Date ? row[1].toISOString() : row[1],
      nama: row[2],
      kelas: row[3],
      teknik: row[4],
      skor: row[5],
      status: row[6],
      detail: (function () {
        try { return JSON.parse(row[7]); } catch (e) { return []; }
      })()
    };
    if (filter) {
      if (filter.kelas && filter.kelas !== 'Semua' && item.kelas !== filter.kelas) continue;
      if (filter.teknik && filter.teknik !== 'Semua' && item.teknik !== filter.teknik) continue;
    }
    out.push(item);
  }
  out.sort(function (a, b) { return new Date(b.timestamp) - new Date(a.timestamp); });
  return out;
}

function deleteAttemptById(id) {
  const sh = _riwayatSheet();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false };
}

function getStats() {
  const sh = _riwayatSheet();
  const data = sh.getDataRange().getValues();
  const stats = {};
  TEKNIK_LIST.forEach(function (t) {
    stats[t] = { jumlah: 0, totalSkor: 0, sesuai: 0 };
  });
  for (let i = 1; i < data.length; i++) {
    const t = data[i][4];
    if (!stats[t]) stats[t] = { jumlah: 0, totalSkor: 0, sesuai: 0 };
    stats[t].jumlah++;
    stats[t].totalSkor += Number(data[i][5]) || 0;
    if (String(data[i][6]).toLowerCase().indexOf('sesuai') === 0) stats[t].sesuai++;
  }
  const result = {};
  Object.keys(stats).forEach(function (t) {
    const s = stats[t];
    result[t] = {
      jumlah: s.jumlah,
      rataSkor: s.jumlah ? Math.round((s.totalSkor / s.jumlah) * 10) / 10 : 0,
      sesuai: s.sesuai
    };
  });
  return result;
}
