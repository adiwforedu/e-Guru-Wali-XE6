// ============================================================
//  SISTEM GURU WALI PRO — Code.gs
// ============================================================
//  SETUP AWAL (lakukan sekali):
//  1. Di editor: Pengaturan Proyek (⚙) → centang
//     "Tampilkan file manifes appsscript.json"
//  2. Ganti isi appsscript.json dengan file yang disertakan
//  3. Pilih fungsi requestDriveAuth → Run ▶ → setujui semua izin
//  4. Deploy ulang Web App
// ============================================================


// ─── ENTRY POINT ────────────────────────────────────────────
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('Sistem Guru Wali Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ─── OTORISASI DRIVE ────────────────────────────────────────


/**
 * Jalankan SATU KALI dari editor Apps Script untuk memicu dialog
 * izin OAuth. Setelah disetujui, scope Drive aktif permanen.
 */
function requestDriveAuth() {
  const root = DriveApp.getRootFolder(); // memicu dialog izin
  Logger.log('✓ Drive terotorisasi. Root: ' + root.getName());
  return 'OK: ' + root.getName();
}


/**
 * Dipanggil dari UI untuk cek status otorisasi Drive.
 *
 * PENTING: dengan executeAs=USER_DEPLOYING, semua kode berjalan
 * sebagai pemilik script. Session.getEffectiveUser() bisa gagal
 * pada mode ANYONE_ANONYMOUS — jadi kita cek DriveApp saja
 * sebagai satu-satunya indikator yang relevan.
 *
 * @returns {{ authorized: boolean, message: string }}
 */
function checkAuthStatus() {
  try {
    DriveApp.getRootFolder(); // jika ini berhasil, Drive sudah OK
    return {
      authorized: true,
      message:    'Google Drive sudah diotorisasi dengan benar.'
    };
  } catch (e) {
    return {
      authorized: false,
      // Sertakan error asli agar mudah dianalisis
      message: e.message
    };
  }
}


// ─── SETUP DATABASE ─────────────────────────────────────────
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tables = {
    'Settings': ['id','sekolah','kepsek','tahun','semester',
                 'provinsi','dinas','alamat','email','web',
                 'logoKiri','logoKanan','driveId'],
    'Users':    ['id','nama','username','password','role'],
    'Murid':    ['id','nis','nisn','nama','gender','agama',
                 'alamat','kelas','tahunMasuk','guruUsername',
                 'dataProfil','fotoUrl'],
    'Program':  ['id','pilar','bentuk','indikator','guruUsername'],
    'Jadwal':   ['id','waktu','kegiatan','keterangan','guruUsername'],
    'Jurnal':   ['id','siswaId','tanggal','topik','hasil','guruUsername'],
    'Laporan':  ['id','siswaId','aspek','deskripsi','rekomendasi','guruUsername'],
    'Rencana':  ['id','siswaId','fokus','tujuan','strategi','guruUsername']
  };
  for (const name in tables) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.appendRow(tables[name]);
      sh.getRange(1,1,1,tables[name].length).setFontWeight('bold').setBackground('#f1f5f9');
      if (name === 'Users')    sh.appendRow(["'1",'Administrator','admin','admin123','Admin']);
      if (name === 'Settings') sh.appendRow(["'1",'','','2024/2025','Ganjil','','','{}','','','','','']);
    } else {
      _migrateSheetColumns(sh, tables[name]);
    }
  }
  return 'Database berhasil disiapkan!';
}


function _migrateSheetColumns(sheet, expected) {
  const has = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h=>h.toString().trim());
  expected.forEach(h => {
    if (!has.includes(h)) {
      const c = sheet.getLastColumn() + 1;
      sheet.getRange(1,c).setValue(h).setFontWeight('bold').setBackground('#f1f5f9');
    }
  });
}


// ─── UPLOAD FOTO ────────────────────────────────────────────


/**
 * Upload foto murid ke Google Drive.
 * @param {string} muridId
 * @param {string} base64Data  tanpa prefix "data:..."
 * @param {string} mimeType    "image/jpeg" | "image/png" | "image/webp"
 * @param {string} namaFile    nama file + ekstensi
 * @returns {{ success, url?, fileId?, message }}
 */
function uploadFotoMurid(muridId, base64Data, mimeType, namaFile) {
  try {
    if (!muridId)    return _fail('ID murid kosong.');
    if (!base64Data) return _fail('Data foto tidak terkirim dari browser.');


    // Ambil folder ID dari Settings
    const folderId = _getDriveFolderId();
    if (!folderId) return _fail(
      'ID Folder Drive belum diisi.\n' +
      'Buka Identitas → isi ID Folder Drive → Simpan Pengaturan.'
    );


    // Akses folder utama
    let folder;
    try {
      folder = DriveApp.getFolderById(folderId);
      folder.getName();
    } catch (e) {
      return _fail(
        'Folder tidak dapat diakses. ID: ' + folderId + '\n' +
        'Error: ' + e.message
      );
    }


    // Dapatkan / buat sub-folder Foto_Murid
    let fotoFolder;
    try {
      const it = folder.getFoldersByName('Foto_Murid');
      fotoFolder = it.hasNext() ? it.next() : folder.createFolder('Foto_Murid');
    } catch (e) {
      return _fail('Gagal akses sub-folder Foto_Murid. Pastikan akses folder adalah Editor.\nError: ' + e.message);
    }


    // Hapus foto lama (non-fatal)
    try { _deleteOldPhoto(fotoFolder, muridId); } catch(_) {}


    // Decode base64 → Blob
    let blob;
    try {
      blob = Utilities.newBlob(
        Utilities.base64Decode(base64Data),
        mimeType || 'image/jpeg',
        namaFile || ('foto_' + muridId + '.jpg')
      );
    } catch (e) {
      return _fail('Data foto rusak. Error: ' + e.message);
    }


    // Upload
    let file;
    try {
      file = fotoFolder.createFile(blob);
      file.setName('murid_' + muridId + '_' + (namaFile || 'foto.jpg'));
    } catch (e) {
      return _fail('Gagal simpan ke Drive. Error: ' + e.message);
    }


    // Set izin publik (non-fatal)
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(_) {}


    // Simpan URL ke sheet
    const fileId   = file.getId();
    const driveUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
    try { _updateFotoUrl(muridId, driveUrl); } catch (e) {
      return { success:true, url:driveUrl, fileId:fileId,
               message:'Foto terunggah tapi gagal catat ke sheet. Refresh halaman.' };
    }


    Logger.log('uploadFotoMurid OK: ' + fileId);
    return { success:true, url:driveUrl, fileId:fileId, message:'Foto berhasil diunggah.' };


  } catch (err) {
    Logger.log('uploadFotoMurid ERR: ' + err.toString());
    return _fail('Error tidak terduga: ' + err.message);
  }
}


function deleteFotoMurid(muridId) {
  try {
    const id = _getDriveFolderId();
    if (!id) return _fail('ID Folder Drive belum dikonfigurasi.');
    const folder = DriveApp.getFolderById(id);
    const it     = folder.getFoldersByName('Foto_Murid');
    if (it.hasNext()) _deleteOldPhoto(it.next(), muridId);
    _updateFotoUrl(muridId, '');
    return { success:true, message:'Foto berhasil dihapus.' };
  } catch (e) {
    return _fail('Gagal hapus foto: ' + e.message);
  }
}


/**
 * Uji koneksi folder Drive dari tombol "Uji Koneksi" di UI.
 * Tidak memakai checkAuthStatus() lagi — langsung akses DriveApp
 * agar hasilnya akurat tanpa false-negative.
 * @returns {{ ok: boolean, message: string, folderName?: string }}
 */
function _testDriveFolderAccess(folderId) {
  if (!folderId) return { ok:false, message:'ID folder kosong.' };


  const id = folderId.toString().trim();
  try {
    const folder = DriveApp.getFolderById(id);
    const name   = folder.getName();


    // Uji izin tulis: coba akses/buat sub-folder
    try {
      const it = folder.getFoldersByName('Foto_Murid');
      if (!it.hasNext()) folder.createFolder('Foto_Murid');
    } catch (eW) {
      return {
        ok: false,
        message: 'Folder "' + name + '" ditemukan, tetapi izin BUKAN Editor ' +
                 '(tidak bisa tulis). Ubah akses berbagi ke "Editor".\nError: ' + eW.message
      };
    }


    return {
      ok: true,
      folderName: name,
      message: '✓ Folder "' + name + '" siap. Akses Editor terverifikasi.'
    };
  } catch (e) {
    // Bedakan error otorisasi scope vs folder tidak ditemukan
    const isAuthError = e.message && (
      e.message.includes('DriveApp') ||
      e.message.includes('authorization') ||
      e.message.includes('OAuth') ||
      e.message.includes('permission') ||
      e.message.includes('izin')
    );
    if (isAuthError) {
      return {
        ok: false,
        message:
          'Apps Script belum punya izin akses Drive.\n\n' +
          'Cara memperbaiki:\n' +
          '1. Buka editor Apps Script (Extensions → Apps Script)\n' +
          '2. Pilih fungsi "requestDriveAuth" di dropdown fungsi\n' +
          '3. Klik Run ▶ → setujui SEMUA izin di popup\n' +
          '4. Deploy ulang (Deploy → Manage deployments → versi baru)\n\n' +
          'Error asli: ' + e.message
      };
    }
    return {
      ok: false,
      message:
        'Folder ID "' + id + '" tidak ditemukan atau tidak bisa diakses.\n' +
        'Pastikan ID benar (salin dari URL setelah /folders/).\n' +
        'Error: ' + e.message
    };
  }
}


/** Jalankan dari editor untuk debug sheet Settings & Drive */
function debugDriveSettings() {
  Logger.log('=== checkAuthStatus ===');
  Logger.log(JSON.stringify(checkAuthStatus()));


  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sh) { Logger.log('Sheet Settings tidak ada!'); return; }
  const d = sh.getDataRange().getDisplayValues();
  Logger.log('Headers: ' + JSON.stringify(d[0]));
  for (let i = 1; i < d.length; i++) Logger.log('Row ' + (i+1) + ': ' + JSON.stringify(d[i]));


  const fid = _getDriveFolderId();
  Logger.log('driveId terbaca: "' + fid + '"');
  if (fid) Logger.log('Tes folder: ' + JSON.stringify(_testDriveFolderAccess(fid)));
}


// ─── HELPERS ────────────────────────────────────────────────
function _fail(msg) {
  Logger.log('FAIL: ' + msg);
  return { success:false, message:msg };
}


function _getDriveFolderId() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sh) return '';
  const data  = sh.getDataRange().getDisplayValues();
  if (data.length < 2) return '';
  const hdrs  = data[0].map(h => h.toString().trim().toLowerCase().replace(/\s+/g,''));
  const idCol = hdrs.indexOf('id');
  const drCol = hdrs.indexOf('driveid');
  if (drCol === -1) return '';
  let fallback = '';
  for (let i = 1; i < data.length; i++) {
    const val = (data[i][drCol] || '').toString().trim();
    if (!val) continue;
    if (idCol !== -1 && (data[i][idCol]||'').toString().replace(/^'/,'').trim() === '1') return val;
    if (!fallback) fallback = val;
  }
  return fallback;
}


function _deleteOldPhoto(fotoFolder, muridId) {
  const prefix = 'murid_' + muridId + '_';
  const files  = fotoFolder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getName().startsWith(prefix)) f.setTrashed(true);
  }
}


function _updateFotoUrl(muridId, url) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Murid');
  if (!sh) return;
  const data = sh.getDataRange().getDisplayValues();
  const hdrs = data[0].map(h => h.toString().trim());
  let col    = hdrs.indexOf('fotoUrl');
  if (col === -1) col = hdrs.map(h=>h.toLowerCase()).indexOf('fotourl');
  if (col === -1) {
    col = hdrs.length;
    sh.getRange(1, col+1).setValue('fotoUrl').setFontWeight('bold').setBackground('#f1f5f9');
  }
  const cid = muridId.toString().replace(/^'/,'');
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0]||'').toString().replace(/^'/,'') === cid) {
      sh.getRange(i+1, col+1).setValue(url);
      return;
    }
  }
}


// ─── CRUD STANDAR ───────────────────────────────────────────
function getTableData(sheetName) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) return [];
  const data = sh.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  const hdrs = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    hdrs.forEach((h, i) => {
      let v = (row[i]||'').toString();
      if (v.charAt(0) === "'") v = v.substring(1);
      const c = h.toString().trim();
      obj[c] = v;
      const l = c.toLowerCase().replace(/\s+/g,'');
      if (l !== c) obj[l] = v;
    });
    return obj;
  });
}


function saveRecord(sheetName, recordData) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) return false;
  const disp = sh.getDataRange().getDisplayValues();
  const hdrs = disp[0];
  const norm = {};
  for (const k in recordData) norm[k.toLowerCase().replace(/\s+/g,'')] = recordData[k];


  const cid = recordData.id ? recordData.id.toString().replace(/^'/,'') : null;
  let row   = -1;
  if (cid) {
    for (let i = 1; i < disp.length; i++) {
      if ((disp[i][0]||'').toString().replace(/^'/,'') === cid) { row = i+1; break; }
    }
  }
  if (row === -1) {
    const nid = sheetName==='Settings' ? '1'
      : (recordData.nisn||recordData.nis||recordData.username||Date.now().toString());
    norm['id'] = nid; recordData.id = nid;
    for (let i = 1; i < disp.length; i++) {
      if ((disp[i][0]||'').toString().replace(/^'/,'').trim() === nid.toString().trim()) { row = i+1; break; }
    }
  }
  const ID_COLS = ['id','siswaid','nis','nisn'];
  const vals = hdrs.map(h => {
    const nh = h.toString().toLowerCase().replace(/\s+/g,'');
    let v    = norm[nh];
    if (v !== null && typeof v === 'object') v = JSON.stringify(v);
    if (ID_COLS.includes(nh) && v !== undefined && v !== '') return "'" + v.toString().replace(/^'/,'');
    return v !== undefined ? v : '';
  });
  if (row > -1) sh.getRange(row, 1, 1, hdrs.length).setValues([vals]);
  else          sh.appendRow(vals);
  return true;
}


function deleteRecord(sheetName, id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) return false;
  const t  = id.toString().replace(/^'/,'');
  const d  = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < d.length; i++) {
    if ((d[i][0]||'').toString().replace(/^'/,'') === t) { sh.deleteRow(i+1); return true; }
  }
  return false;
}


function importMuridBatch(muridArray) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Murid');
  if (!sh) return false;
  const hdrs = sh.getDataRange().getValues()[0];
  muridArray.forEach(m => {
    m.id = m.nisn || m.nis || (Date.now().toString() + Math.floor(Math.random()*10000));
    const nis  = (m.nis||'').toString().replace(/^'/,'');
    const rows = sh.getDataRange().getDisplayValues();
    if (nis && rows.slice(1).some(r => (r[1]||'').toString().replace(/^'/,'') === nis)) return;
    const norm = {};
    for (const k in m) norm[k.toLowerCase().replace(/\s+/g,'')] = m[k];
    const rv = hdrs.map(h => {
      const nh = h.toString().toLowerCase().replace(/\s+/g,'');
      let v    = norm[nh];
      if (['id','nis','nisn'].includes(nh) && v !== undefined && v !== '')
        return "'" + v.toString().replace(/^'/,'');
      return v !== undefined ? v : '';
    });
    sh.appendRow(rv);
  });
  return true;
}

function doPost(e) {
  try {
    var contents = {};
    if (e && e.postData && e.postData.contents) {
      contents = JSON.parse(e.postData.contents);
    }
    var funcName = contents.funcName || (e.parameter ? e.parameter.funcName : '');
    var args = contents.args || (e.parameter && e.parameter.args ? JSON.parse(e.parameter.args) : []);
    
    var result = handleApiRequest(funcName, args);
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', result: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleApiRequest(funcName, args) {
  args = args || [];
  if (funcName === 'getTableData') {
    return getTableData(args[0]);
  } else if (funcName === 'saveRecord') {
    return saveRecord(args[0], args[1]);
  } else if (funcName === 'deleteRecord') {
    return deleteRecord(args[0], args[1]);
  } else if (funcName === 'importMuridBatch') {
    return importMuridBatch(args[0]);
  } else if (typeof this[funcName] === 'function') {
    return this[funcName].apply(this, args);
  }
  throw new Error('Fungsi ' + funcName + ' tidak ditemukan.');
}
