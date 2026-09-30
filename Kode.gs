/**
 * ============================================================
 * SIM HALAQAH — Sistem Informasi Manajemen Halaqah Al-Qur'an
 * Backend (Kode.gs)
 * STIS Al Wafa — Versi 1.0
 * ============================================================
 */

// ── Konstanta Global ──
const APP_NAME = 'SIM Halaqah';
const FOLDER_NAME = 'SIM-Halaqah';
const SS_NAME = 'DB_SIM_Halaqah';

// Daftar lengkap 114 nama Surah (urutan mushaf) — data awal untuk modul Data Master
const DAFTAR_SURAH_LENGKAP = ['Al-Fatihah','Al-Baqarah',"Ali 'Imran",'An-Nisa',"Al-Ma'idah","Al-An'am","Al-A'raf",'Al-Anfal','At-Taubah','Yunus','Hud','Yusuf',"Ar-Ra'd",'Ibrahim','Al-Hijr','An-Nahl','Al-Isra','Al-Kahf','Maryam','Ta-Ha','Al-Anbiya','Al-Hajj',"Al-Mu'minun",'An-Nur','Al-Furqan',"Asy-Syu'ara",'An-Naml','Al-Qasas',"Al-'Ankabut",'Ar-Rum','Luqman','As-Sajdah','Al-Ahzab','Saba','Fatir','Ya-Sin','As-Saffat','Sad','Az-Zumar','Ghafir','Fussilat','Asy-Syura','Az-Zukhruf','Ad-Dukhan','Al-Jasiyah','Al-Ahqaf','Muhammad','Al-Fath','Al-Hujurat','Qaf','Az-Zariyat','At-Tur','An-Najm','Al-Qamar','Ar-Rahman',"Al-Waqi'ah",'Al-Hadid','Al-Mujadalah','Al-Hasyr','Al-Mumtahanah','As-Saff',"Al-Jumu'ah",'Al-Munafiqun','At-Taghabun','At-Talaq','At-Tahrim','Al-Mulk','Al-Qalam','Al-Haqqah',"Al-Ma'arij",'Nuh','Al-Jinn','Al-Muzzammil','Al-Muddassir','Al-Qiyamah','Al-Insan','Al-Mursalat','An-Naba',"An-Nazi'at",'Abasa','At-Takwir','Al-Infitar','Al-Mutaffifin','Al-Insyiqaq','Al-Buruj','At-Tariq',"Al-A'la",'Al-Ghasyiyah','Al-Fajr','Al-Balad','Asy-Syams','Al-Lail','Ad-Duha','Asy-Syarh','At-Tin',"Al-'Alaq",'Al-Qadr','Al-Bayyinah','Az-Zalzalah',"Al-'Adiyat","Al-Qari'ah",'At-Takasur',"Al-'Asr",'Al-Humazah','Al-Fil','Quraisy',"Al-Ma'un",'Al-Kausar','Al-Kafirun','An-Nasr','Al-Lahab','Al-Ikhlas','Al-Falaq','An-Nas'];

// ── Helper: Include HTML Partial ──
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Helper: Response Builder ──
function createResponse(success, data, message) {
  return { success: success, data: sanitizeForClient(data), message: message };
}

// ── Helper: Bersihkan nilai Date/exotic menjadi string sebelum dikirim ke client.
// Google Sheets sering auto-konversi teks tanggal ("16/08/2026") menjadi tipe
// Date asli saat disimpan. Jika lolos apa adanya ke google.script.run, ini bisa
// menyebabkan payload gagal terkirim (client menerima null) pada beberapa kasus.
function sanitizeForClient(value) {
  if (value === null || value === undefined) return value;
  if (value instanceof Date) return formatTanggalID(value);
  if (Array.isArray(value)) return value.map(sanitizeForClient);
  if (typeof value === 'object') {
    const out = {};
    Object.keys(value).forEach(k => out[k] = sanitizeForClient(value[k]));
    return out;
  }
  return value;
}

// ── Helper: Tes koneksi sederhana tanpa akses Sheet/Drive sama sekali ──
function pingTest() {
  return createResponse(true, { time: new Date().toISOString(), pong: true }, 'Koneksi server normal.');
}

// ── Helper: Generate UUID ──
function generateUUID() {
  return Utilities.getUuid();
}

// ── Helper: Format Tanggal Indonesia ──
function formatTanggalID(date) {
  if (!date) return '';
  const d = (date instanceof Date) ? date : new Date(date);
  if (isNaN(d.getTime())) return String(date);
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Jakarta', 'dd/MM/yyyy');
}

// ════════════════════════════════════════════════════════
// ENTRY POINT — SPA MURNI (TANPA ?page=)
// ════════════════════════════════════════════════════════

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ════════════════════════════════════════════════════════
// AUTO-SETUP — dijalankan sekali oleh Ustadz
// ════════════════════════════════════════════════════════

function setupAppEnvironment() {
  try {
    // 1. Folder utama
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    const mainFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
    Logger.log('📁 Folder utama: ' + mainFolder.getUrl());

    // 2. Sub-folder
    const fotoFolder = getOrCreateFolder(mainFolder, 'Foto-Profil');
    const murojaahFolder = getOrCreateFolder(mainFolder, 'Murojaah-Liburan');

    // 3. Spreadsheet
    let ss;
    const files = mainFolder.getFilesByName(SS_NAME);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create(SS_NAME);
      DriveApp.getFileById(ss.getId()).moveTo(mainFolder);
    }
    Logger.log('📊 Spreadsheet: ' + ss.getUrl());

    const todayStr = formatTanggalID(new Date());
    const dummySantriId = generateUUID();

    // 4. Sheet-sheet
    createSheetIfNotExists(ss, 'DataSantri',
      ['ID', 'NamaLengkap', 'Nickname', 'TanggalLahir', 'EmailOrtu', 'NamaOrtu', 'NoWhatsappOrtu', 'FotoProfilURL', 'TanggalBergabung', 'StatusAktif'],
      [dummySantriId, 'Ahmad Faris', 'Faris', '01/01/2015', 'contoh.ortu@gmail.com', 'Bapak Contoh', '628123456789', '', todayStr, 'Aktif']
    );

    createSheetIfNotExists(ss, 'Setoran',
      ['ID', 'IDSantri', 'Tanggal', 'Sesi', 'JenisSetoran', 'Surah', 'AyatMulai', 'AyatSelesai', 'JumlahAyat', 'NilaiKelancaran', 'NilaiTajwid', 'NilaiAdab', 'NilaiTotal', 'CatatanUstadz', 'LinkYoutube'],
      [generateUUID(), dummySantriId, todayStr, 'Sesi 1', 'Hafalan Baru', 'Al-Mulk', 1, 15, 15, 85, 90, 95, 90, 'Contoh catatan perkembangan', '']
    );

    createSheetIfNotExists(ss, 'Absensi',
      ['ID', 'IDSantri', 'Tanggal', 'Sesi', 'Status', 'Keterangan'],
      [generateUUID(), dummySantriId, todayStr, 'Sesi 1', 'Hadir', '']
    );

    createSheetIfNotExists(ss, 'MurojaahLiburan',
      ['ID', 'IDSantri', 'TanggalUpload', 'Surah', 'CatatanOrtu', 'LinkLampiran', 'StatusReview', 'CatatanUstadz']
    );

    createSheetIfNotExists(ss, 'PeriodeLiburan',
      ['ID', 'NamaPeriode', 'TanggalMulai', 'TanggalSelesai', 'Status']
    );

    createSheetIfNotExists(ss, 'AkunPengguna',
      ['Email', 'Peran', 'IDSantriTerkait', 'Status', 'Password', 'Username'],
      [Session.getActiveUser().getEmail(), 'Koordinator', '', 'Aktif', '', ''],
      ['contoh.ortu@gmail.com', 'Orang Tua', dummySantriId, 'Aktif', '123456', 'ortu.contoh']
    );

    createSheetIfNotExists(ss, 'DataGuru',
      ['ID', 'NamaLengkap', 'Email', 'NoWhatsapp', 'Status', 'TanggalBergabung']
    );

    createSheetIfNotExists(ss, 'DataMaster',
      ['ID', 'Kategori', 'Nama', 'Urutan'],
      ...DAFTAR_SURAH_LENGKAP.map((nama, i) => [generateUUID(), 'Surah', nama, i + 1]),
      [generateUUID(), 'Sesi', 'Sesi 1', 1],
      [generateUUID(), 'Sesi', 'Sesi 2', 2],
      [generateUUID(), 'Sesi', 'Sesi 3', 3]
    );

    createSheetIfNotExists(ss, 'AppConfig',
      ['Key', 'Value'],
      ['appName', APP_NAME],
      ['folderId', mainFolder.getId()],
      ['fotoFolderId', fotoFolder.getId()],
      ['murojaahFolderId', murojaahFolder.getId()],
      ['spreadsheetId', ss.getId()],
      ['logoUrl', ''],
      ['adminEmail', Session.getActiveUser().getEmail()],
      ['templateWaHarian', ''],
      ['templateWaSantri', '']
    );

    // Simpan ID Spreadsheet ke Script Properties (akses cepat)
    PropertiesService.getScriptProperties().setProperty('spreadsheetId', ss.getId());

    // Hapus Sheet1 default
    const defaultSheet = ss.getSheetByName('Sheet1');
    if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

    Logger.log('✅ Setup selesai!');
    Logger.log('📊 Spreadsheet ID: ' + ss.getId());
    Logger.log('👤 Super Admin: ' + Session.getActiveUser().getEmail());

    return createResponse(true, { spreadsheetId: ss.getId(), folderId: mainFolder.getId() }, 'Setup berhasil! Silakan deploy sebagai Web App.');
  } catch (error) {
    Logger.log('❌ Error setup: ' + error.message);
    return createResponse(false, null, error.message);
  }
}

function getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// ── Self-healing migration: tambah kolom ke sheet lama tanpa perlu setup ulang ──
function ensureColumn(sheet, columnName) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf(columnName);
  if (idx !== -1) return idx + 1;
  sheet.getRange(1, lastCol + 1).setValue(columnName);
  return lastCol + 1;
}

// ── Self-healing: pastikan sheet DataMaster ada (untuk user yang sudah setup sebelum fitur ini ada) ──
function ensureDataMasterSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('DataMaster');
  if (!sheet) {
    sheet = createSheetIfNotExists(ss, 'DataMaster',
      ['ID', 'Kategori', 'Nama', 'Urutan'],
      ...DAFTAR_SURAH_LENGKAP.map((nama, i) => [generateUUID(), 'Surah', nama, i + 1])
    );
  }
  return sheet;
}

function createSheetIfNotExists(ss, sheetName, headers, ...dataRows) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    dataRows.forEach(row => sheet.appendRow(row));
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0f5238').setFontColor('#ffffff');
    headers.forEach((_, i) => sheet.autoResizeColumn(i + 1));
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSpreadsheet() {
  let ssId = PropertiesService.getScriptProperties().getProperty('spreadsheetId');
  if (!ssId) {
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    if (!folders.hasNext()) throw new Error('Aplikasi belum di-setup. Jalankan setupAppEnvironment() terlebih dahulu.');
    const files = folders.next().getFilesByName(SS_NAME);
    if (!files.hasNext()) throw new Error('Spreadsheet belum ditemukan. Jalankan setupAppEnvironment() terlebih dahulu.');
    ssId = files.next().getId();
    PropertiesService.getScriptProperties().setProperty('spreadsheetId', ssId);
  }
  return SpreadsheetApp.openById(ssId);
}

// ════════════════════════════════════════════════════════
// AppConfig
// ════════════════════════════════════════════════════════

function getConfig(key) {
  try {
    const sheet = getSpreadsheet().getSheetByName('AppConfig');
    const data = sheet.getDataRange().getValues();
    const row = data.find(r => r[0] === key);
    return row ? row[1] : null;
  } catch (e) { return null; }
}

function setConfig(key, value) {
  try {
    const sheet = getSpreadsheet().getSheetByName('AppConfig');
    const data = sheet.getDataRange().getValues();
    const idx = data.findIndex(r => r[0] === key);
    if (idx >= 0) sheet.getRange(idx + 1, 2).setValue(value);
    else sheet.appendRow([key, value]);
    return createResponse(true, null, 'Konfigurasi disimpan.');
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function getAllConfig() {
  try {
    const data = getSpreadsheet().getSheetByName('AppConfig').getDataRange().getValues();
    const config = {};
    data.slice(1).forEach(row => config[row[0]] = row[1]);
    return createResponse(true, config, 'OK');
  } catch (error) {
    Logger.log('[getAllConfig] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

// ── Pengaturan Aplikasi (Logo & Tema Warna) ──
function getAppSettings() {
  try {
    const cfg = getAllConfig();
    if (!cfg.success) throw new Error(cfg.message);
    return createResponse(true, {
      appName: cfg.data.appName || 'SIM Halaqah',
      logoUrl: cfg.data.logoUrl || '',
      primaryColor: cfg.data.primaryColor || '#0f5238',
      secondaryColor: cfg.data.secondaryColor || '#C5A059'
    }, 'OK');
  } catch (error) {
    Logger.log('[getAppSettings] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

function saveAppSettings(settings) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSpreadsheet().getSheetByName('AppConfig');
    const data = sheet.getDataRange().getValues();
    Object.keys(settings).forEach(key => {
      const idx = data.findIndex(r => r[0] === key);
      if (idx >= 0) sheet.getRange(idx + 1, 2).setValue(settings[key]);
      else sheet.appendRow([key, settings[key]]);
    });
    return createResponse(true, settings, 'Pengaturan aplikasi berhasil disimpan.');
  } catch (error) {
    Logger.log('[saveAppSettings] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════
// AUTENTIKASI & SESI PENGGUNA
// ════════════════════════════════════════════════════════

// Catatan: Login otomatis via Google untuk staf (Koordinator & Super Admin/Guru).
// Orang Tua login manual dengan Username + Password (lihat loginOrtu) karena
// Session.getActiveUser() tidak selalu bisa membaca email akun Google pribadi
// (Gmail) di luar domain organisasi — ini yang menyebabkan login Orang Tua
// via Google sebelumnya sering gagal.
function getCurrentUserInfo() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) {
      return createResponse(false, null, 'AUTO_LOGIN_GAGAL');
    }
    const sheet = getSpreadsheet().getSheetByName('AkunPengguna');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const roleCol = headers.indexOf('Peran');
    const statusCol = headers.indexOf('Status');

    const row = data.slice(1).find(r => String(r[emailCol]).toLowerCase() === email.toLowerCase() && (r[roleCol] === 'Super Admin' || r[roleCol] === 'Koordinator'));
    if (!row) {
      return createResponse(false, null, 'AUTO_LOGIN_GAGAL');
    }
    if (row[statusCol] !== 'Aktif') {
      return createResponse(false, null, 'Akun Anda tidak aktif. Silakan hubungi Koordinator/Ustadz.');
    }

    return createResponse(true, {
      email: email,
      role: row[roleCol],
      idSantriTerkait: null,
      santri: null,
      liburanAktif: isLiburanAktif()
    }, 'OK');
  } catch (error) {
    Logger.log('[getCurrentUserInfo] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

// Cari role staf (Koordinator/Super Admin) dari akun Google yang sedang login.
// Dipakai untuk menentukan target audiens Pengumuman secara otomatis & aman
// (tidak bisa dipalsukan dari client).
function getCallingStaffRole() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return null;
    const sheet = getSpreadsheet().getSheetByName('AkunPengguna');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const roleCol = headers.indexOf('Peran');
    const row = data.slice(1).find(r => String(r[emailCol]).toLowerCase() === email.toLowerCase() && (r[roleCol] === 'Super Admin' || r[roleCol] === 'Koordinator'));
    return row ? row[roleCol] : null;
  } catch (e) {
    return null;
  }
}

// ── Login manual Orang Tua (Email + Password dibuat oleh Admin) ──
function loginOrtu(username, password) {
  try {
    if (!username || !password) return createResponse(false, null, 'Username dan password wajib diisi.');
    const sheet = getSpreadsheet().getSheetByName('AkunPengguna');
    ensureColumn(sheet, 'Password');
    ensureColumn(sheet, 'Username');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const usernameCol = headers.indexOf('Username');
    const roleCol = headers.indexOf('Peran');
    const santriCol = headers.indexOf('IDSantriTerkait');
    const statusCol = headers.indexOf('Status');
    const passCol = headers.indexOf('Password');

    // Cocokkan berdasar Username. Untuk akun lama yang belum punya Username
    // (dibuat sebelum fitur ini ada), izinkan login pakai Email seperti biasa
    // agar akun lama tidak mendadak terkunci.
    const row = data.slice(1).find(r => {
      if (r[roleCol] !== 'Orang Tua') return false;
      const uname = String(r[usernameCol] || '').toLowerCase();
      if (uname) return uname === String(username).toLowerCase();
      return String(r[emailCol]).toLowerCase() === String(username).toLowerCase();
    });
    if (!row) return createResponse(false, null, 'Username tidak terdaftar. Hubungi Guru/Ustadz untuk didaftarkan.');
    if (row[statusCol] !== 'Aktif') return createResponse(false, null, 'Akun Anda tidak aktif. Silakan hubungi Guru/Ustadz.');
    if (!row[passCol]) return createResponse(false, null, 'Password untuk akun ini belum diatur. Silakan hubungi Guru/Ustadz.');
    if (String(row[passCol]) !== String(password)) return createResponse(false, null, 'Password salah. Silakan coba lagi atau hubungi Guru/Ustadz.');

    let santriInfo = null;
    if (row[santriCol]) {
      const santriRes = getSantriById(row[santriCol]);
      if (santriRes.success) santriInfo = santriRes.data;
    }

    return createResponse(true, {
      email: row[emailCol],
      username: row[usernameCol] || row[emailCol],
      role: 'Orang Tua',
      idSantriTerkait: row[santriCol] || null,
      santri: santriInfo,
      liburanAktif: isLiburanAktif()
    }, 'Login berhasil.');
  } catch (error) {
    Logger.log('[loginOrtu] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

// ════════════════════════════════════════════════════════
// MODUL — DATA SANTRI (CRUD)
// ════════════════════════════════════════════════════════

function getAllSantri() {
  try {
    const sheet = getSpreadsheet().getSheetByName('DataSantri');
    ensureColumn(sheet, 'IDGuru'); // afiliasi santri ke Guru pengampu (fitur multi-guru)
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse(true, [], 'Belum ada data santri.');
    const headers = data[0];
    const records = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
    return createResponse(true, records, 'OK');
  } catch (error) {
    Logger.log('[getAllSantri] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

function getSantriById(id) {
  try {
    const res = getAllSantri();
    if (!res.success) return res;
    const found = res.data.find(s => s.ID === id);
    if (!found) return createResponse(false, null, 'Santri tidak ditemukan.');
    return createResponse(true, found, 'OK');
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function addSantri(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSpreadsheet().getSheetByName('DataSantri');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    recordObj.ID = generateUUID();
    if (!recordObj.TanggalBergabung) recordObj.TanggalBergabung = formatTanggalID(new Date());
    if (!recordObj.StatusAktif) recordObj.StatusAktif = 'Aktif';
    const newRow = headers.map(h => recordObj[h] !== undefined ? recordObj[h] : '');
    sheet.appendRow(newRow);
    return createResponse(true, recordObj, 'Santri berhasil ditambahkan.');
  } catch (error) {
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function updateSantri(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSpreadsheet().getSheetByName('DataSantri');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');
    const rowIndex = data.findIndex((row, i) => i > 0 && row[idCol] === recordObj.ID);
    if (rowIndex === -1) return createResponse(false, null, 'Santri tidak ditemukan.');
    const updatedRow = headers.map((h, i) => recordObj[h] !== undefined ? recordObj[h] : data[rowIndex][i]);
    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([updatedRow]);
    return createResponse(true, recordObj, 'Data santri berhasil diperbarui.');
  } catch (error) {
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function deleteSantri(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSpreadsheet().getSheetByName('DataSantri');
    const data = sheet.getDataRange().getValues();
    const idCol = data[0].indexOf('ID');
    const rowIndex = data.findIndex((row, i) => i > 0 && row[idCol] === id);
    if (rowIndex === -1) return createResponse(false, null, 'Santri tidak ditemukan.');
    sheet.deleteRow(rowIndex + 1);
    return createResponse(true, null, 'Santri berhasil dihapus.');
  } catch (error) {
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════
// UPLOAD FILE KE DRIVE (Foto Profil / Lampiran Murojaah)
// ════════════════════════════════════════════════════════

function uploadFileToDrive(fileData, fileName, mimeType, folderKey) {
  try {
    const folderId = getConfig(folderKey || 'fotoFolderId');
    if (!folderId) throw new Error('Folder upload belum dikonfigurasi. Jalankan setupAppEnvironment().');
    const folder = DriveApp.getFolderById(folderId);
    const decoded = Utilities.base64Decode(fileData);
    const blob = Utilities.newBlob(decoded, mimeType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return createResponse(true, {
      fileId: file.getId(),
      // Format thumbnail jauh lebih andal untuk ditampilkan langsung via <img>
      // dibanding drive.google.com/uc?id=... yang sering diblokir untuk hotlink.
      fileUrl: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000',
      fileName: file.getName()
    }, 'File berhasil diupload.');
  } catch (error) {
    Logger.log('[uploadFileToDrive] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

// ════════════════════════════════════════════════════════
// MODUL — SETORAN HAFALAN
// ════════════════════════════════════════════════════════

function addSetoran(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSpreadsheet().getSheetByName('Setoran');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    recordObj.ID = generateUUID();
    const ayatMulai = Number(recordObj.AyatMulai) || 0;
    const ayatSelesai = Number(recordObj.AyatSelesai) || 0;
    recordObj.JumlahAyat = Math.max(0, ayatSelesai - ayatMulai + 1);

    const nk = Number(recordObj.NilaiKelancaran) || 0;
    const nt = Number(recordObj.NilaiTajwid) || 0;
    const na = Number(recordObj.NilaiAdab) || 0;
    recordObj.NilaiTotal = Math.round(((nk + nt + na) / 3) * 10) / 10;

    if (!recordObj.Tanggal) recordObj.Tanggal = formatTanggalID(new Date());

    const newRow = headers.map(h => recordObj[h] !== undefined ? recordObj[h] : '');
    sheet.appendRow(newRow);

    // Kirim notifikasi email ke orang tua (non-blocking terhadap kegagalan)
    try {
      const santriRes = getSantriById(recordObj.IDSantri);
      if (santriRes.success) sendSetoranEmail(recordObj, santriRes.data);
    } catch (emailErr) {
      Logger.log('⚠️ Gagal kirim email: ' + emailErr.message);
    }

    return createResponse(true, recordObj, 'Setoran berhasil disimpan.');
  } catch (error) {
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function getAllSetoran() {
  try {
    const sheet = getSpreadsheet().getSheetByName('Setoran');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse(true, [], 'Belum ada data setoran.');
    const headers = data[0];
    const records = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).reverse(); // terbaru dulu
    return createResponse(true, records, 'OK');
  } catch (error) {
    Logger.log('[getAllSetoran] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

function getSetoranBySantri(idSantri) {
  try {
    const res = getAllSetoran();
    if (!res.success) return res;
    return createResponse(true, res.data.filter(s => s.IDSantri === idSantri), 'OK');
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function sendSetoranEmail(setoran, santri) {
  if (!santri.EmailOrtu) return;
  const subject = `[${APP_NAME}] Update Setoran Hafalan — ${santri.NamaLengkap}`;
  const body = `Assalamu'alaikum ${santri.NamaOrtu || 'Bapak/Ibu'},\n\n`
    + `Berikut update setoran hafalan ananda ${santri.NamaLengkap} hari ini:\n\n`
    + `Tanggal      : ${setoran.Tanggal}\n`
    + `Sesi         : ${setoran.Sesi}\n`
    + `Jenis        : ${setoran.JenisSetoran}\n`
    + `Surah        : ${setoran.Surah} (ayat ${setoran.AyatMulai}-${setoran.AyatSelesai})\n`
    + `Nilai Total  : ${setoran.NilaiTotal}\n`
    + `Catatan Ustadz: ${setoran.CatatanUstadz || '-'}\n\n`
    + `Jazakumullahu khairan atas dukungannya.\n\n${APP_NAME} — STIS Al Wafa`;
  try {
    MailApp.sendEmail(santri.EmailOrtu, subject, body);
  } catch (e) {
    Logger.log('Gagal kirim email ke ' + santri.EmailOrtu + ': ' + e.message);
  }
}

// ════════════════════════════════════════════════════════
// MODUL — ABSENSI
// ════════════════════════════════════════════════════════

function addAbsensiBulk(tanggal, sesi, entries) {
  // entries: [{ idSantri, status, keterangan }]
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = getSpreadsheet().getSheetByName('Absensi');
    const rows = entries.map(e => [generateUUID(), e.idSantri, tanggal, sesi, e.status, e.keterangan || '']);
    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    }
    return createResponse(true, null, `Absensi ${rows.length} santri berhasil disimpan.`);
  } catch (error) {
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function getAllAbsensi() {
  try {
    const sheet = getSpreadsheet().getSheetByName('Absensi');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse(true, [], 'Belum ada data absensi.');
    const headers = data[0];
    const records = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
    return createResponse(true, records, 'OK');
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function getAbsensiBySantri(idSantri) {
  try {
    const res = getAllAbsensi();
    if (!res.success) return res;
    return createResponse(true, res.data.filter(a => a.IDSantri === idSantri), 'OK');
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function hitungPersenKehadiran(absensiList) {
  if (!absensiList || absensiList.length === 0) return 0;
  const hadir = absensiList.filter(a => a.Status === 'Hadir').length;
  return Math.round((hadir / absensiList.length) * 100);
}

// ════════════════════════════════════════════════════════
// DASHBOARD — SUPER ADMIN
// ════════════════════════════════════════════════════════

function getDashboardAdmin() {
  try {
    const santriRes = getAllSantri();
    const setoranRes = getAllSetoran();
    const absensiRes = getAllAbsensi();
    if (!santriRes.success) return santriRes;

    const santriAktif = santriRes.data.filter(s => s.StatusAktif === 'Aktif');
    const setoran = setoranRes.success ? setoranRes.data : [];
    const absensi = absensiRes.success ? absensiRes.data : [];

    const totalAyat = setoran.reduce((sum, s) => sum + (Number(s.JumlahAyat) || 0), 0);
    const avgNilai = setoran.length ? (setoran.reduce((sum, s) => sum + (Number(s.NilaiTotal) || 0), 0) / setoran.length) : 0;

    const todayStr = formatTanggalID(new Date());
    const activeSessionsToday = new Set(setoran.filter(s => s.Tanggal === todayStr).map(s => s.IDSantri)).size;

    // Grafik 7 hari terakhir — total ayat per hari
    const growthMap = {};
    const days = [];
    for (let i = 6; i >= 0; i--) {
      const d = new Date();
      d.setDate(d.getDate() - i);
      const key = formatTanggalID(d);
      growthMap[key] = 0;
      days.push(key);
    }
    setoran.forEach(s => {
      if (growthMap.hasOwnProperty(s.Tanggal)) growthMap[s.Tanggal] += Number(s.JumlahAyat) || 0;
    });

    const insights = [];
    if (setoran.length > 0) {
      insights.push(`Total ${totalAyat} ayat telah disetorkan dari ${santriAktif.length} santri aktif.`);
      insights.push(`Rata-rata nilai setoran saat ini adalah ${avgNilai.toFixed(1)}.`);
    } else {
      insights.push('Belum ada data setoran. Mulai catat setoran hafalan santri hari ini.');
    }

    return createResponse(true, {
      totalSantri: santriAktif.length,
      totalAyat: totalAyat,
      activeSessions: activeSessionsToday,
      avgNilai: avgNilai.toFixed(1),
      growthLabels: days,
      growthValues: days.map(d => growthMap[d]),
      insights: insights
    }, 'OK');
  } catch (error) {
    Logger.log('[getDashboardAdmin] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

// ════════════════════════════════════════════════════════
// DASHBOARD — ORANG TUA
// ════════════════════════════════════════════════════════

function getDashboardOrtu(idSantri) {
  try {
    const santriRes = getSantriById(idSantri);
    if (!santriRes.success) return santriRes;

    const setoranRes = getSetoranBySantri(idSantri);
    const absensiRes = getAbsensiBySantri(idSantri);
    const setoran = setoranRes.success ? setoranRes.data : [];
    const absensi = absensiRes.success ? absensiRes.data : [];

    const totalAyat = setoran.reduce((sum, s) => sum + (Number(s.JumlahAyat) || 0), 0);
    const avgNilai = setoran.length ? (setoran.reduce((sum, s) => sum + (Number(s.NilaiTotal) || 0), 0) / setoran.length) : 0;
    const persenHadir = hitungPersenKehadiran(absensi);

    // Progress mingguan (5 hari terakhir, Sen-Jum sederhana: ambil 7 hari terakhir)
    const days = [];
    const growthMap = {};
    for (let i = 6; i >= 0; i--) {
      const d = new Date();
      d.setDate(d.getDate() - i);
      const key = formatTanggalID(d);
      growthMap[key] = 0;
      days.push(key);
    }
    setoran.forEach(s => {
      if (growthMap.hasOwnProperty(s.Tanggal)) growthMap[s.Tanggal] += Number(s.JumlahAyat) || 0;
    });

    return createResponse(true, {
      santri: santriRes.data,
      totalAyat: totalAyat,
      avgNilai: avgNilai.toFixed(1),
      persenHadir: persenHadir,
      setoranTerakhir: setoran.slice(0, 5),
      growthLabels: days,
      growthValues: days.map(d => growthMap[d])
    }, 'OK');
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

// ════════════════════════════════════════════════════════
// MODUL — PAPAN PERINGKAT
// ════════════════════════════════════════════════════════

function getLeaderboard(forOrtu) {
  try {
    const santriRes = getAllSantri();
    const setoranRes = getAllSetoran();
    const absensiRes = getAllAbsensi();
    if (!santriRes.success) return santriRes;

    const setoran = setoranRes.success ? setoranRes.data : [];
    const absensi = absensiRes.success ? absensiRes.data : [];

    const ranking = santriRes.data.filter(s => s.StatusAktif === 'Aktif').map(s => {
      const setoranSantri = setoran.filter(x => x.IDSantri === s.ID);
      const absensiSantri = absensi.filter(x => x.IDSantri === s.ID);
      const totalAyat = setoranSantri.reduce((sum, x) => sum + (Number(x.JumlahAyat) || 0), 0);
      const avgNilai = setoranSantri.length ? (setoranSantri.reduce((sum, x) => sum + (Number(x.NilaiTotal) || 0), 0) / setoranSantri.length) : 0;
      return {
        id: s.ID,
        nama: s.NamaLengkap,
        foto: s.FotoProfilURL,
        totalAyat: totalAyat,
        avgNilai: Math.round(avgNilai * 10) / 10,
        persenHadir: hitungPersenKehadiran(absensiSantri)
      };
    }).sort((a, b) => b.totalAyat - a.totalAyat || b.avgNilai - a.avgNilai);

    ranking.forEach((r, i) => r.rank = i + 1);

    if (forOrtu) {
      // Orang tua hanya lihat nama & posisi, tanpa nilai detail santri lain
      return createResponse(true, ranking.map(r => ({ rank: r.rank, nama: r.nama, foto: r.foto, totalAyat: r.totalAyat })), 'OK');
    }
    return createResponse(true, ranking, 'OK');
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

// ════════════════════════════════════════════════════════
// MODUL — GENERATOR LAPORAN WHATSAPP
// ════════════════════════════════════════════════════════

// Ganti {placeholder} pada template dengan nilai dari dataObj
function replacePlaceholders(template, dataObj) {
  let text = template;
  Object.keys(dataObj).forEach(key => {
    text = text.split('{' + key + '}').join(dataObj[key] !== undefined && dataObj[key] !== null ? dataObj[key] : '');
  });
  return text;
}

function generateLaporanWA(mode, params) {
  try {
    const todayStr = formatTanggalID(new Date());
    const templates = getConfig('templateWaHarian') ? { harian: getConfig('templateWaHarian') } : {};
    const templateSantri = getConfig('templateWaSantri');

    if (mode === 'harian') {
      const setoranRes = getAllSetoran();
      const setoranHariIni = setoranRes.success ? setoranRes.data.filter(s => s.Tanggal === todayStr) : [];
      if (setoranHariIni.length === 0) {
        return createResponse(true, `📋 *Laporan Halaqah — ${todayStr}*\n\nBelum ada setoran tercatat hari ini.`, 'OK');
      }
      const santriRes = getAllSantri();
      const santriMap = {};
      if (santriRes.success) santriRes.data.forEach(s => santriMap[s.ID] = s.NamaLengkap);

      let daftarSetoran = '';
      setoranHariIni.forEach((s, i) => {
        daftarSetoran += `${i + 1}. *${santriMap[s.IDSantri] || 'Santri'}*\n`;
        daftarSetoran += `   ${s.JenisSetoran} — ${s.Surah} (${s.AyatMulai}-${s.AyatSelesai})\n`;
        daftarSetoran += `   Nilai: ${s.NilaiTotal} | ${s.Sesi}\n\n`;
      });

      const templateHarian = getConfig('templateWaHarian');
      let text;
      if (templateHarian) {
        text = replacePlaceholders(templateHarian, { tanggal: todayStr, daftar_setoran: daftarSetoran.trim(), nama_aplikasi: APP_NAME });
      } else {
        text = `📋 *Laporan Halaqah — ${todayStr}*\n\n${daftarSetoran}Jazakumullahu khairan.\n${APP_NAME} — STIS Al Wafa`;
      }
      return createResponse(true, text, 'OK');

    } else if (mode === 'santri') {
      const santriRes = getSantriById(params.idSantri);
      if (!santriRes.success) return santriRes;
      const setoranRes = getSetoranBySantri(params.idSantri);
      const setoran = setoranRes.success ? setoranRes.data.slice(0, 5) : [];

      let daftarSetoran = '';
      setoran.forEach((s, i) => {
        daftarSetoran += `${i + 1}. ${s.Tanggal} — ${s.Surah} (${s.AyatMulai}-${s.AyatSelesai}) — Nilai ${s.NilaiTotal}\n`;
      });

      let text;
      if (templateSantri) {
        text = replacePlaceholders(templateSantri, { nama_santri: santriRes.data.NamaLengkap, tanggal: todayStr, daftar_setoran: daftarSetoran.trim(), nama_aplikasi: APP_NAME });
      } else {
        text = `📋 *Laporan Perkembangan — ${santriRes.data.NamaLengkap}*\nTanggal: ${todayStr}\n\n*5 Setoran Terakhir:*\n${daftarSetoran}\nJazakumullahu khairan atas dukungannya.\n${APP_NAME} — STIS Al Wafa`;
      }
      return createResponse(true, text, 'OK');
    }
    return createResponse(false, null, 'Mode laporan tidak dikenali.');
  } catch (error) {
    Logger.log('[generateLaporanWA] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

// ── Template Pesan WhatsApp (Pengaturan) ──
function getWaTemplates() {
  try {
    return createResponse(true, {
      templateWaHarian: getConfig('templateWaHarian') || '',
      templateWaSantri: getConfig('templateWaSantri') || ''
    }, 'OK');
  } catch (error) {
    Logger.log('[getWaTemplates] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

function saveWaTemplates(templates) {
  try {
    setConfig('templateWaHarian', templates.templateWaHarian || '');
    setConfig('templateWaSantri', templates.templateWaSantri || '');
    return createResponse(true, null, 'Template WhatsApp berhasil disimpan.');
  } catch (error) {
    Logger.log('[saveWaTemplates] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

function kirimEmailManual(idSantri, subjek, pesan) {
  try {
    const santriRes = getSantriById(idSantri);
    if (!santriRes.success) return santriRes;
    if (!santriRes.data.EmailOrtu) return createResponse(false, null, 'Email orang tua tidak ditemukan.');
    MailApp.sendEmail(santriRes.data.EmailOrtu, subjek, pesan);
    return createResponse(true, null, 'Email berhasil dikirim ke ' + santriRes.data.EmailOrtu);
  } catch (error) {
    Logger.log('[kirimEmailManual] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

// ── Laporan Lengkap (Rapor + Kehadiran) untuk dikirim langsung ke Orang Tua ──
function generateLaporanLengkapText(idSantri) {
  try {
    const dashRes = getDashboardOrtu(idSantri);
    if (!dashRes.success) return dashRes;
    const d = dashRes.data;
    const todayStr = formatTanggalID(new Date());

    let text = `LAPORAN PERKEMBANGAN SANTRI\n`;
    text += `${APP_NAME} — STIS Al Wafa\n`;
    text += `Tanggal Laporan: ${todayStr}\n`;
    text += `============================\n\n`;
    text += `Nama Santri : ${d.santri.NamaLengkap}\n`;
    text += `Nickname    : ${d.santri.Nickname || '-'}\n\n`;
    text += `RINGKASAN HAFALAN\n`;
    text += `- Total Ayat Terhafal : ${d.totalAyat}\n`;
    text += `- Rata-rata Nilai     : ${d.avgNilai}\n`;
    text += `- Persentase Kehadiran: ${d.persenHadir}%\n\n`;
    text += `RIWAYAT SETORAN TERAKHIR\n`;
    if (d.setoranTerakhir.length === 0) {
      text += `Belum ada setoran tercatat.\n`;
    } else {
      d.setoranTerakhir.forEach((s, i) => {
        text += `${i + 1}. ${s.Tanggal} — ${s.JenisSetoran} — ${s.Surah} (${s.AyatMulai}-${s.AyatSelesai}) — Nilai ${s.NilaiTotal}\n`;
        if (s.CatatanUstadz) text += `   Catatan Ustadz: ${s.CatatanUstadz}\n`;
      });
    }
    text += `\nJazakumullahu khairan atas kerjasama dan dukungannya dalam mendampingi ananda menghafal Al-Qur'an.\n\n${APP_NAME} — STIS Al Wafa`;
    return createResponse(true, text, 'OK');
  } catch (error) {
    Logger.log('[generateLaporanLengkapText] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

function kirimLaporanLengkapEmail(idSantri) {
  try {
    const textRes = generateLaporanLengkapText(idSantri);
    if (!textRes.success) return textRes;
    const santriRes = getSantriById(idSantri);
    if (!santriRes.success) return santriRes;
    if (!santriRes.data.EmailOrtu) return createResponse(false, null, 'Email orang tua untuk santri ini belum diisi.');

    const subjek = `Laporan Perkembangan Santri — ${santriRes.data.NamaLengkap} (${formatTanggalID(new Date())})`;
    MailApp.sendEmail(santriRes.data.EmailOrtu, subjek, textRes.data);
    return createResponse(true, null, 'Laporan lengkap berhasil dikirim ke ' + santriRes.data.EmailOrtu);
  } catch (error) {
    Logger.log('[kirimLaporanLengkapEmail] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

// ════════════════════════════════════════════════════════
// MODUL — MODE LIBURAN & MUROJAAH MANDIRI
// ════════════════════════════════════════════════════════

function getPeriodeLiburanList() {
  try {
    const sheet = getSpreadsheet().getSheetByName('PeriodeLiburan');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse(true, [], 'Belum ada periode liburan.');
    const headers = data[0];
    const records = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
    return createResponse(true, records, 'OK');
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function addPeriodeLiburan(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSpreadsheet().getSheetByName('PeriodeLiburan');
    recordObj.ID = generateUUID();
    if (!recordObj.Status) recordObj.Status = 'Aktif';
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    sheet.appendRow(headers.map(h => recordObj[h] !== undefined ? recordObj[h] : ''));
    return createResponse(true, recordObj, 'Periode liburan berhasil ditambahkan.');
  } catch (error) {
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function updatePeriodeLiburan(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSpreadsheet().getSheetByName('PeriodeLiburan');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');
    const rowIndex = data.findIndex((row, i) => i > 0 && row[idCol] === recordObj.ID);
    if (rowIndex === -1) return createResponse(false, null, 'Periode tidak ditemukan.');
    const updatedRow = headers.map((h, i) => recordObj[h] !== undefined ? recordObj[h] : data[rowIndex][i]);
    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([updatedRow]);
    return createResponse(true, recordObj, 'Periode liburan diperbarui.');
  } catch (error) {
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

// PENTING: fungsi ini membaca sheet LANGSUNG (bukan lewat getPeriodeLiburanList/
// createResponse), karena createResponse mengonversi nilai Date menjadi teks
// "DD/MM/YYYY" untuk keamanan pengiriman ke client — format itu tidak bisa
// di-parse ulang oleh `new Date()` dengan benar dan menyebabkan status liburan
// selalu terbaca tidak aktif. Di sini kita butuh objek Date asli untuk perbandingan.
function isLiburanAktif() {
  try {
    const sheet = getSpreadsheet().getSheetByName('PeriodeLiburan');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return false;
    const headers = data[0];
    const statusCol = headers.indexOf('Status');
    const mulaiCol = headers.indexOf('TanggalMulai');
    const selesaiCol = headers.indexOf('TanggalSelesai');
    const now = new Date();
    return data.slice(1).some(row => {
      if (row[statusCol] !== 'Aktif') return false;
      const mulai = new Date(row[mulaiCol]);
      const selesai = new Date(row[selesaiCol]);
      selesai.setHours(23, 59, 59, 999); // agar hari terakhir periode tetap dihitung aktif
      return now >= mulai && now <= selesai;
    });
  } catch (e) {
    Logger.log('[isLiburanAktif] ' + e.message);
    return false;
  }
}

// Endpoint publik agar client bisa cek status liburan TERKINI kapan saja,
// tidak hanya dari data sesi login yang mungkin sudah usang.
function checkLiburanAktif() {
  return createResponse(true, isLiburanAktif(), 'OK');
}

function addMurojaahLiburan(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!isLiburanAktif()) return createResponse(false, null, 'Periode liburan tidak sedang aktif.');
    const sheet = getSpreadsheet().getSheetByName('MurojaahLiburan');
    recordObj.ID = generateUUID();
    recordObj.TanggalUpload = formatTanggalID(new Date());
    recordObj.StatusReview = 'Belum ditinjau';
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    sheet.appendRow(headers.map(h => recordObj[h] !== undefined ? recordObj[h] : ''));
    return createResponse(true, recordObj, 'Murojaah berhasil diunggah. Ustadz akan meninjau.');
  } catch (error) {
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function getMurojaahBySantri(idSantri) {
  try {
    const sheet = getSpreadsheet().getSheetByName('MurojaahLiburan');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse(true, [], 'OK');
    const headers = data[0];
    const records = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).filter(r => r.IDSantri === idSantri).reverse();
    return createResponse(true, records, 'OK');
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function getAllMurojaahForReview() {
  try {
    const sheet = getSpreadsheet().getSheetByName('MurojaahLiburan');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse(true, [], 'OK');
    const headers = data[0];
    const santriRes = getAllSantri();
    const santriMap = {};
    if (santriRes.success) santriRes.data.forEach(s => santriMap[s.ID] = s.NamaLengkap);
    const records = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      obj.NamaSantri = santriMap[obj.IDSantri] || 'Tidak diketahui';
      return obj;
    }).reverse();
    return createResponse(true, records, 'OK');
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function reviewMurojaah(id, statusReview, catatanUstadz) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSpreadsheet().getSheetByName('MurojaahLiburan');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');
    const statusCol = headers.indexOf('StatusReview');
    const catatanCol = headers.indexOf('CatatanUstadz');
    const rowIndex = data.findIndex((row, i) => i > 0 && row[idCol] === id);
    if (rowIndex === -1) return createResponse(false, null, 'Data tidak ditemukan.');
    sheet.getRange(rowIndex + 1, statusCol + 1).setValue(statusReview);
    sheet.getRange(rowIndex + 1, catatanCol + 1).setValue(catatanUstadz || '');
    return createResponse(true, null, 'Feedback berhasil disimpan.');
  } catch (error) {
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════
// MODUL — MANAJEMEN AKUN ORANG TUA
// ════════════════════════════════════════════════════════

function getAllAkunOrtu() {
  try {
    const sheet = getSpreadsheet().getSheetByName('AkunPengguna');
    ensureColumn(sheet, 'Password');
    ensureColumn(sheet, 'Username');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse(true, [], 'OK');
    const headers = data[0];
    const santriRes = getAllSantri();
    const santriMap = {};
    if (santriRes.success) santriRes.data.forEach(s => santriMap[s.ID] = s.NamaLengkap);
    const records = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      obj.NamaSantri = santriMap[obj.IDSantriTerkait] || '-';
      obj.PunyaPassword = !!obj.Password; // beri tahu client apakah password sudah diset, TANPA kirim passwordnya
      delete obj.Password; // jangan pernah kirim password ke client
      return obj;
    });
    return createResponse(true, records, 'OK');
  } catch (error) {
    Logger.log('[getAllAkunOrtu] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

function addAkunOrtu(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSpreadsheet().getSheetByName('AkunPengguna');
    ensureColumn(sheet, 'Password');
    ensureColumn(sheet, 'Username');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const usernameCol = headers.indexOf('Username');
    const emailExists = recordObj.Email && data.slice(1).some(r => String(r[emailCol]).toLowerCase() === String(recordObj.Email).toLowerCase());
    if (emailExists) return createResponse(false, null, 'Email sudah terdaftar.');
    if (!recordObj.Username) return createResponse(false, null, 'Username wajib diisi untuk akun Orang Tua.');
    const usernameExists = data.slice(1).some(r => String(r[usernameCol]).toLowerCase() === String(recordObj.Username).toLowerCase());
    if (usernameExists) return createResponse(false, null, 'Username sudah dipakai, silakan pilih username lain.');
    if (!recordObj.Password) return createResponse(false, null, 'Password awal wajib diisi untuk akun Orang Tua.');
    recordObj.Peran = recordObj.Peran || 'Orang Tua';
    recordObj.Status = recordObj.Status || 'Aktif';
    sheet.appendRow(headers.map(h => recordObj[h] !== undefined ? recordObj[h] : ''));
    return createResponse(true, { Email: recordObj.Email, Username: recordObj.Username }, 'Akun orang tua berhasil ditambahkan.');
  } catch (error) {
    Logger.log('[addAkunOrtu] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

// Reset/ubah password akun Orang Tua (dipanggil terpisah dari update status biasa)
function resetPasswordOrtu(email, newPassword) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!newPassword) return createResponse(false, null, 'Password baru wajib diisi.');
    const sheet = getSpreadsheet().getSheetByName('AkunPengguna');
    ensureColumn(sheet, 'Password');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const passCol = headers.indexOf('Password');
    const rowIndex = data.findIndex((row, i) => i > 0 && row[emailCol] === email);
    if (rowIndex === -1) return createResponse(false, null, 'Akun tidak ditemukan.');
    sheet.getRange(rowIndex + 1, passCol + 1).setValue(newPassword);
    return createResponse(true, null, 'Password berhasil diperbarui.');
  } catch (error) {
    Logger.log('[resetPasswordOrtu] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function updateAkunOrtu(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSpreadsheet().getSheetByName('AkunPengguna');
    ensureColumn(sheet, 'Password');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const rowIndex = data.findIndex((row, i) => i > 0 && row[emailCol] === recordObj.Email);
    if (rowIndex === -1) return createResponse(false, null, 'Akun tidak ditemukan.');
    const updatedRow = headers.map((h, i) => recordObj[h] !== undefined ? recordObj[h] : data[rowIndex][i]);
    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([updatedRow]);
    return createResponse(true, recordObj, 'Akun berhasil diperbarui.');
  } catch (error) {
    Logger.log('[updateAkunOrtu] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function deleteAkunOrtu(email) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSpreadsheet().getSheetByName('AkunPengguna');
    const data = sheet.getDataRange().getValues();
    const emailCol = data[0].indexOf('Email');
    const rowIndex = data.findIndex((row, i) => i > 0 && row[emailCol] === email);
    if (rowIndex === -1) return createResponse(false, null, 'Akun tidak ditemukan.');
    sheet.deleteRow(rowIndex + 1);
    return createResponse(true, null, 'Akun berhasil dihapus.');
  } catch (error) {
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════
// MODUL — DATA MASTER (Surah & data referensi lainnya)
// ════════════════════════════════════════════════════════

function getAllDataMaster() {
  try {
    const sheet = ensureDataMasterSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse(true, [], 'OK');
    const headers = data[0];
    const records = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).sort((a, b) => (Number(a.Urutan) || 0) - (Number(b.Urutan) || 0));
    return createResponse(true, records, 'OK');
  } catch (error) {
    Logger.log('[getAllDataMaster] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

function getDataMasterByKategori(kategori) {
  try {
    const res = getAllDataMaster();
    if (!res.success) return res;
    return createResponse(true, res.data.filter(d => d.Kategori === kategori), 'OK');
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function getAllKategoriMaster() {
  try {
    const res = getAllDataMaster();
    if (!res.success) return res;
    const kategoriSet = [...new Set(res.data.map(d => d.Kategori))];
    return createResponse(true, kategoriSet, 'OK');
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

// Dipakai Input Setoran & Murojaah Liburan untuk mengisi dropdown Surah.
// Jika sheet DataMaster kosong/gagal dibaca, kembalikan daftar bawaan sebagai cadangan.
function getSurahList() {
  try {
    const res = getDataMasterByKategori('Surah');
    if (res.success && res.data.length > 0) {
      return createResponse(true, res.data.map(d => d.Nama), 'OK');
    }
    return createResponse(true, DAFTAR_SURAH_LENGKAP, 'OK (data cadangan)');
  } catch (error) {
    Logger.log('[getSurahList] ' + error.message);
    return createResponse(true, DAFTAR_SURAH_LENGKAP, 'OK (data cadangan setelah error)');
  }
}

function addDataMaster(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = ensureDataMasterSheet();
    recordObj.ID = generateUUID();
    if (!recordObj.Urutan) {
      const data = sheet.getDataRange().getValues();
      const kategoriRows = data.slice(1).filter(r => r[1] === recordObj.Kategori);
      recordObj.Urutan = kategoriRows.length + 1;
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    sheet.appendRow(headers.map(h => recordObj[h] !== undefined ? recordObj[h] : ''));
    return createResponse(true, recordObj, 'Data master berhasil ditambahkan.');
  } catch (error) {
    Logger.log('[addDataMaster] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function updateDataMaster(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = ensureDataMasterSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');
    const rowIndex = data.findIndex((row, i) => i > 0 && row[idCol] === recordObj.ID);
    if (rowIndex === -1) return createResponse(false, null, 'Data tidak ditemukan.');
    const updatedRow = headers.map((h, i) => recordObj[h] !== undefined ? recordObj[h] : data[rowIndex][i]);
    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([updatedRow]);
    return createResponse(true, recordObj, 'Data master berhasil diperbarui.');
  } catch (error) {
    Logger.log('[updateDataMaster] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function deleteDataMaster(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = ensureDataMasterSheet();
    const data = sheet.getDataRange().getValues();
    const idCol = data[0].indexOf('ID');
    const rowIndex = data.findIndex((row, i) => i > 0 && row[idCol] === id);
    if (rowIndex === -1) return createResponse(false, null, 'Data tidak ditemukan.');
    sheet.deleteRow(rowIndex + 1);
    return createResponse(true, null, 'Data master berhasil dihapus.');
  } catch (error) {
    Logger.log('[deleteDataMaster] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════
// MODUL — PENGUMUMAN (ditampilkan di lonceng notifikasi)
// ════════════════════════════════════════════════════════

function ensurePengumumanSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('Pengumuman');
  if (!sheet) {
    sheet = createSheetIfNotExists(ss, 'Pengumuman',
      ['ID', 'Judul', 'Isi', 'LampiranURL', 'LampiranNama', 'LinkEksternal', 'TanggalDibuat', 'Status', 'TargetRole', 'DibuatOleh']
    );
  }
  ensureColumn(sheet, 'TargetRole');
  ensureColumn(sheet, 'DibuatOleh');
  return sheet;
}

// Menentukan audiens yang DIKELOLA oleh suatu peran (siapa yang mereka publikasikan
// pengumumannya). Koordinator -> ditujukan ke Super Admin (Guru). Super Admin (Guru)
// -> ditujukan ke Orang Tua. Alurnya searah ke bawah saja, sesuai permintaan.
function getManagedTargetRole(creatorRole) {
  if (creatorRole === 'Koordinator') return 'Super Admin';
  if (creatorRole === 'Super Admin') return 'Orang Tua';
  return null;
}

function getAllPengumuman(viewerRole) {
  try {
    const sheet = ensurePengumumanSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse(true, [], 'OK');
    const headers = data[0];
    let records = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).reverse(); // terbaru dulu
    const managedTarget = getManagedTargetRole(viewerRole);
    if (managedTarget) records = records.filter(p => p.TargetRole === managedTarget);
    return createResponse(true, records, 'OK');
  } catch (error) {
    Logger.log('[getAllPengumuman] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

// Dipanggil dari lonceng notifikasi — hanya yang berstatus Aktif DAN ditujukan
// untuk peran viewer ini (Koordinator tidak menerima pengumuman dari siapapun,
// sesuai permintaan "tidak berlaku ke atas").
function getPengumumanAktif(viewerRole) {
  try {
    const res = getAllPengumumanRaw();
    if (!res.success) return res;
    const filtered = res.data.filter(p => p.Status === 'Aktif' && p.TargetRole === viewerRole);
    return createResponse(true, filtered, 'OK');
  } catch (error) {
    Logger.log('[getPengumumanAktif] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

// Versi tanpa filter peran, untuk dipakai internal oleh getPengumumanAktif
function getAllPengumumanRaw() {
  try {
    const sheet = ensurePengumumanSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse(true, [], 'OK');
    const headers = data[0];
    const records = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).reverse();
    return createResponse(true, records, 'OK');
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function addPengumuman(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const creatorRole = getCallingStaffRole();
    const targetRole = getManagedTargetRole(creatorRole);
    if (!targetRole) return createResponse(false, null, 'Tidak dapat menentukan audiens pengumuman untuk akun Anda.');

    const sheet = ensurePengumumanSheet();
    recordObj.ID = generateUUID();
    recordObj.TanggalDibuat = formatTanggalID(new Date());
    recordObj.Status = recordObj.Status || 'Aktif';
    recordObj.TargetRole = targetRole;
    recordObj.DibuatOleh = Session.getActiveUser().getEmail() || '-';
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    sheet.appendRow(headers.map(h => recordObj[h] !== undefined ? recordObj[h] : ''));
    return createResponse(true, recordObj, 'Pengumuman berhasil ditambahkan.');
  } catch (error) {
    Logger.log('[addPengumuman] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function updatePengumuman(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = ensurePengumumanSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');
    const rowIndex = data.findIndex((row, i) => i > 0 && row[idCol] === recordObj.ID);
    if (rowIndex === -1) return createResponse(false, null, 'Pengumuman tidak ditemukan.');
    const updatedRow = headers.map((h, i) => recordObj[h] !== undefined ? recordObj[h] : data[rowIndex][i]);
    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([updatedRow]);
    return createResponse(true, recordObj, 'Pengumuman berhasil diperbarui.');
  } catch (error) {
    Logger.log('[updatePengumuman] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function deletePengumuman(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = ensurePengumumanSheet();
    const data = sheet.getDataRange().getValues();
    const idCol = data[0].indexOf('ID');
    const rowIndex = data.findIndex((row, i) => i > 0 && row[idCol] === id);
    if (rowIndex === -1) return createResponse(false, null, 'Pengumuman tidak ditemukan.');
    sheet.deleteRow(rowIndex + 1);
    return createResponse(true, null, 'Pengumuman berhasil dihapus.');
  } catch (error) {
    Logger.log('[deletePengumuman] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════
// MODUL — DATA GURU (dikelola oleh Koordinator)
// ════════════════════════════════════════════════════════

function ensureDataGuruSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('DataGuru');
  if (!sheet) {
    sheet = createSheetIfNotExists(ss, 'DataGuru', ['ID', 'NamaLengkap', 'Email', 'NoWhatsapp', 'Status', 'TanggalBergabung']);
  }
  return sheet;
}

function getAllDataGuru() {
  try {
    const sheet = ensureDataGuruSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse(true, [], 'OK');
    const headers = data[0];
    const santriRes = getAllSantri();
    const countMap = {};
    if (santriRes.success) {
      santriRes.data.forEach(s => { if (s.IDGuru) countMap[s.IDGuru] = (countMap[s.IDGuru] || 0) + 1; });
    }
    const records = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      obj.JumlahSantri = countMap[obj.ID] || 0;
      return obj;
    });
    return createResponse(true, records, 'OK');
  } catch (error) {
    Logger.log('[getAllDataGuru] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

// Menambah Guru = mendaftarkan profil DAN otomatis memberi akses login
// (membuat baris AkunPengguna dengan Peran Super Admin, login via Google).
function addDataGuru(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = ensureDataGuruSheet();
    const data = sheet.getDataRange().getValues();
    const emailCol = data[0].indexOf('Email');
    const exists = data.slice(1).some(r => String(r[emailCol]).toLowerCase() === String(recordObj.Email).toLowerCase());
    if (exists) return createResponse(false, null, 'Email guru ini sudah terdaftar.');

    recordObj.ID = generateUUID();
    recordObj.Status = recordObj.Status || 'Aktif';
    recordObj.TanggalBergabung = formatTanggalID(new Date());
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    sheet.appendRow(headers.map(h => recordObj[h] !== undefined ? recordObj[h] : ''));

    // Buat akses login otomatis jika belum ada
    const akunSheet = getSpreadsheet().getSheetByName('AkunPengguna');
    const akunData = akunSheet.getDataRange().getValues();
    const akunEmailCol = akunData[0].indexOf('Email');
    const akunExists = akunData.slice(1).some(r => String(r[akunEmailCol]).toLowerCase() === String(recordObj.Email).toLowerCase());
    if (!akunExists) {
      const akunHeaders = akunSheet.getRange(1, 1, 1, akunSheet.getLastColumn()).getValues()[0];
      const newAkun = { Email: recordObj.Email, Peran: 'Super Admin', IDSantriTerkait: '', Status: 'Aktif' };
      akunSheet.appendRow(akunHeaders.map(h => newAkun[h] !== undefined ? newAkun[h] : ''));
    }

    return createResponse(true, recordObj, 'Guru berhasil ditambahkan dan diberi akses ke aplikasi (login via Google).');
  } catch (error) {
    Logger.log('[addDataGuru] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function updateDataGuru(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = ensureDataGuruSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');
    const rowIndex = data.findIndex((row, i) => i > 0 && row[idCol] === recordObj.ID);
    if (rowIndex === -1) return createResponse(false, null, 'Data guru tidak ditemukan.');
    const oldEmail = data[rowIndex][headers.indexOf('Email')];
    const updatedRow = headers.map((h, i) => recordObj[h] !== undefined ? recordObj[h] : data[rowIndex][i]);
    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([updatedRow]);

    // Sinkronkan status/email ke AkunPengguna jika berubah
    if (recordObj.Status || (recordObj.Email && recordObj.Email !== oldEmail)) {
      const akunSheet = getSpreadsheet().getSheetByName('AkunPengguna');
      const akunData = akunSheet.getDataRange().getValues();
      const akunEmailCol = akunData[0].indexOf('Email');
      const akunStatusCol = akunData[0].indexOf('Status');
      const akunRowIndex = akunData.findIndex((row, i) => i > 0 && row[akunEmailCol] === oldEmail);
      if (akunRowIndex !== -1) {
        if (recordObj.Email) akunSheet.getRange(akunRowIndex + 1, akunEmailCol + 1).setValue(recordObj.Email);
        if (recordObj.Status) akunSheet.getRange(akunRowIndex + 1, akunStatusCol + 1).setValue(recordObj.Status);
      }
    }
    return createResponse(true, recordObj, 'Data guru berhasil diperbarui.');
  } catch (error) {
    Logger.log('[updateDataGuru] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

// Mencabut akses Guru: hapus dari DataGuru + hapus baris AkunPengguna terkait
function deleteDataGuru(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = ensureDataGuruSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');
    const emailCol = headers.indexOf('Email');
    const rowIndex = data.findIndex((row, i) => i > 0 && row[idCol] === id);
    if (rowIndex === -1) return createResponse(false, null, 'Data guru tidak ditemukan.');
    const email = data[rowIndex][emailCol];
    sheet.deleteRow(rowIndex + 1);

    const akunSheet = getSpreadsheet().getSheetByName('AkunPengguna');
    const akunData = akunSheet.getDataRange().getValues();
    const akunEmailCol = akunData[0].indexOf('Email');
    const akunRowIndex = akunData.findIndex((row, i) => i > 0 && row[akunEmailCol] === email);
    if (akunRowIndex !== -1) akunSheet.deleteRow(akunRowIndex + 1);

    return createResponse(true, null, 'Guru dan akses aplikasinya berhasil dihapus.');
  } catch (error) {
    Logger.log('[deleteDataGuru] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════
// MODUL — JUMLAH HAFALAN SANTRI (update bulanan oleh Guru)
// ════════════════════════════════════════════════════════

function ensureJumlahHafalanSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('JumlahHafalan');
  if (!sheet) {
    sheet = createSheetIfNotExists(ss, 'JumlahHafalan', ['ID', 'IDSantri', 'Bulan', 'TotalJuz', 'TotalHalaman', 'JuzDihafal', 'Catatan', 'TanggalUpdate', 'DiupdateOleh']);
  }
  return sheet;
}

function getBulanIniKey() {
  const d = new Date();
  const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
  return Utilities.formatDate(d, tz, 'yyyy-MM');
}

function formatBulanLabel(bulanKey) {
  const namaBulan = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  if (!bulanKey || bulanKey.indexOf('-') === -1) return bulanKey || '-';
  const [y, m] = bulanKey.split('-');
  return namaBulan[Number(m) - 1] + ' ' + y;
}

function getAllJumlahHafalan() {
  try {
    const sheet = ensureJumlahHafalanSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return createResponse(true, [], 'OK');
    const headers = data[0];
    const records = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      obj.BulanLabel = formatBulanLabel(obj.Bulan);
      return obj;
    });
    return createResponse(true, records, 'OK');
  } catch (error) {
    Logger.log('[getAllJumlahHafalan] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

function getJumlahHafalanBySantri(idSantri) {
  try {
    const res = getAllJumlahHafalan();
    if (!res.success) return res;
    const list = res.data.filter(h => h.IDSantri === idSantri).sort((a, b) => b.Bulan.localeCompare(a.Bulan));
    return createResponse(true, list, 'OK');
  } catch (error) {
    return createResponse(false, null, error.message);
  }
}

function getLatestJumlahHafalanBySantri(idSantri) {
  const res = getJumlahHafalanBySantri(idSantri);
  if (!res.success || res.data.length === 0) return createResponse(true, null, 'Belum ada data.');
  return createResponse(true, res.data[0], 'OK');
}

// Simpan/perbarui data hafalan — jika bulan yang sama untuk santri yang sama
// sudah ada, TIMPA (update) baris tersebut alih-alih membuat duplikat.
function upsertJumlahHafalan(recordObj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = ensureJumlahHafalanSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const santriCol = headers.indexOf('IDSantri');
    const bulanCol = headers.indexOf('Bulan');

    recordObj.TanggalUpdate = formatTanggalID(new Date());
    recordObj.DiupdateOleh = Session.getActiveUser().getEmail() || '-';

    const rowIndex = data.findIndex((row, i) => i > 0 && row[santriCol] === recordObj.IDSantri && row[bulanCol] === recordObj.Bulan);
    if (rowIndex !== -1) {
      recordObj.ID = data[rowIndex][headers.indexOf('ID')];
      const updatedRow = headers.map((h, i) => recordObj[h] !== undefined ? recordObj[h] : data[rowIndex][i]);
      sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([updatedRow]);
      return createResponse(true, recordObj, 'Data hafalan bulan ' + formatBulanLabel(recordObj.Bulan) + ' berhasil diperbarui.');
    } else {
      recordObj.ID = generateUUID();
      sheet.appendRow(headers.map(h => recordObj[h] !== undefined ? recordObj[h] : ''));
      return createResponse(true, recordObj, 'Data hafalan bulan ' + formatBulanLabel(recordObj.Bulan) + ' berhasil disimpan.');
    }
  } catch (error) {
    Logger.log('[upsertJumlahHafalan] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

function deleteJumlahHafalan(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = ensureJumlahHafalanSheet();
    const data = sheet.getDataRange().getValues();
    const idCol = data[0].indexOf('ID');
    const rowIndex = data.findIndex((row, i) => i > 0 && row[idCol] === id);
    if (rowIndex === -1) return createResponse(false, null, 'Data tidak ditemukan.');
    sheet.deleteRow(rowIndex + 1);
    return createResponse(true, null, 'Data hafalan berhasil dihapus.');
  } catch (error) {
    Logger.log('[deleteJumlahHafalan] ' + error.message);
    return createResponse(false, null, error.message);
  } finally {
    lock.releaseLock();
  }
}

// Daftar santri yang BELUM diupdate hafalannya bulan ini — untuk pengingat Guru
function getReminderHafalanBulanIni() {
  try {
    const bulanIni = getBulanIniKey();
    const santriRes = getAllSantri();
    const hafalanRes = getAllJumlahHafalan();
    if (!santriRes.success) return santriRes;
    const santriAktif = santriRes.data.filter(s => s.StatusAktif === 'Aktif');
    const sudahUpdateSet = new Set((hafalanRes.success ? hafalanRes.data : []).filter(h => h.Bulan === bulanIni).map(h => h.IDSantri));
    const belumUpdate = santriAktif.filter(s => !sudahUpdateSet.has(s.ID));
    return createResponse(true, {
      bulanIni: bulanIni,
      bulanLabel: formatBulanLabel(bulanIni),
      totalSantri: santriAktif.length,
      totalSudahUpdate: santriAktif.length - belumUpdate.length,
      totalBelumUpdate: belumUpdate.length,
      daftarBelumUpdate: belumUpdate.map(s => ({ id: s.ID, nama: s.NamaLengkap }))
    }, 'OK');
  } catch (error) {
    Logger.log('[getReminderHafalanBulanIni] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

// ════════════════════════════════════════════════════════
// DASHBOARD & LAPORAN — KOORDINATOR
// ════════════════════════════════════════════════════════

function getDashboardKoordinator() {
  try {
    const baseRes = getDashboardAdmin();
    if (!baseRes.success) return baseRes;
    const guruRes = getAllDataGuru();
    const guruList = guruRes.success ? guruRes.data : [];
    const data = baseRes.data;
    data.totalGuru = guruList.filter(g => g.Status === 'Aktif').length;
    data.guruBreakdown = guruList.map(g => ({ nama: g.NamaLengkap, jumlahSantri: g.JumlahSantri, status: g.Status }));
    return createResponse(true, data, 'OK');
  } catch (error) {
    Logger.log('[getDashboardKoordinator] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

function getLaporanOrganisasi() {
  try {
    const guruRes = getAllDataGuru();
    const santriRes = getAllSantri();
    const hafalanRes = getAllJumlahHafalan();
    const absensiRes = getAllAbsensi();

    const guru = guruRes.success ? guruRes.data : [];
    const santri = santriRes.success ? santriRes.data : [];
    const hafalan = hafalanRes.success ? hafalanRes.data : [];
    const absensi = absensiRes.success ? absensiRes.data : [];

    const guruMap = {};
    guru.forEach(g => guruMap[g.ID] = g.NamaLengkap);

    const latestHafalanMap = {};
    hafalan.forEach(h => {
      if (!latestHafalanMap[h.IDSantri] || h.Bulan > latestHafalanMap[h.IDSantri].Bulan) latestHafalanMap[h.IDSantri] = h;
    });

    const rekapSantri = santri.map(s => {
      const absensiSantri = absensi.filter(a => a.IDSantri === s.ID);
      const latestH = latestHafalanMap[s.ID];
      return {
        nama: s.NamaLengkap,
        guru: s.IDGuru ? (guruMap[s.IDGuru] || '-') : '-',
        status: s.StatusAktif,
        totalJuz: latestH ? latestH.TotalJuz : 0,
        totalHalaman: latestH ? latestH.TotalHalaman : 0,
        bulanUpdateTerakhir: latestH ? formatBulanLabel(latestH.Bulan) : 'Belum pernah',
        persenHadir: hitungPersenKehadiran(absensiSantri)
      };
    });

    return createResponse(true, {
      totalGuru: guru.filter(g => g.Status === 'Aktif').length,
      totalSantri: santri.filter(s => s.StatusAktif === 'Aktif').length,
      rekapGuru: guru,
      rekapSantri: rekapSantri
    }, 'OK');
  } catch (error) {
    Logger.log('[getLaporanOrganisasi] ' + error.message);
    return createResponse(false, null, error.message);
  }
}

// ════════════════════════════════════════════════════════
// DATA MASTER TAMBAHAN — Sesi (Data Master kategori "Sesi")
// ════════════════════════════════════════════════════════

// Dipakai Input Setoran & Absensi. Jika belum dikustomisasi lewat Data Master,
// kembalikan 3 sesi default (Sesi 1/2/3) — SAMA PERSIS seperti perilaku sebelumnya.
function getSesiList() {
  try {
    const res = getDataMasterByKategori('Sesi');
    if (res.success && res.data.length > 0) {
      return createResponse(true, res.data.map(d => d.Nama), 'OK');
    }
    return createResponse(true, ['Sesi 1', 'Sesi 2', 'Sesi 3'], 'OK (data cadangan)');
  } catch (error) {
    Logger.log('[getSesiList] ' + error.message);
    return createResponse(true, ['Sesi 1', 'Sesi 2', 'Sesi 3'], 'OK (data cadangan setelah error)');
  }
}
