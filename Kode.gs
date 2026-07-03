// Code.gs (Updated with Monitoring & Account Lock)
const SPREADSHEET_ID = '1__EzshGr1VrjGfTcCxrnpmV-1UoMa8-Reu2CkNydpbU';

function doGet(e){
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('EXAM PORTAL')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

/* ---------- Utility ---------- */
function _open() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function _getSheetByName(name){
  const ss = _open();
  const sheet = ss.getSheetByName(name);
  if(!sheet) throw new Error('Sheet "'+name+'" tidak ditemukan.');
  return sheet;
}

/** * FITUR BARU 1: Update Aktivitas & Deteksi Akun
 * Fungsi ini mencatat status (Mengerjakan/Curang/Selesai) ke Kolom G
 * Dan mengunci akun ke Kolom F jika terjadi kecurangan / selesai
 */
function updateStudentActivity(nisn, status) {
  try {
    const sheet = _getSheetByName('DATA');
    const last = sheet.getLastRow();
    if (last < 2) return; // Antisipasi jika database siswa masih kosong
    
    const data = sheet.getRange(2, 1, last - 1, 8).getValues(); // Ambil data sampai kolom H
    const now = new Date();
    const jamMenit = Utilities.formatDate(now, "GMT+9", "HH:mm"); // Format waktu sesuai GMT Anda
    
    for (let i = 0; i < data.length; i++) {
      // Validasi NISN (Kolom B / indeks 1)
      if (String(data[i][1]).trim() === String(nisn).trim()) {
        const row = i + 2;
        
        // 1. Bersihkan status dan buat pengecekan huruf kecil (case-insensitive)
        let statusCek = String(status).trim().toLowerCase();
        
        // 2. SELALU TULIS status aktivitas apa pun (Mengerjakan/Curang/Selesai) ke Kolom G (Kolom 7)
        sheet.getRange(row, 7).setValue(status);
        
        // 3. Kunci Akun ke Kolom F (Kolom 6) menjadi 'Tidak Aktif' jika terdeteksi curang atau selesai
        // Logika ini mendeteksi kata kunci 'curang', 'selesai', atau jika teks status mengandung kata 'curang'
        if (statusCek === 'curang' || statusCek === 'selesai' || statusCek === 'tes tidak selesai' || statusCek.includes('curang')) {
          sheet.getRange(row, 6).setValue('Tidak Aktif');
        }
        
        // 4. Update Waktu hanya jika siswa baru pertama kali masuk ujian
        if (status === "Sedang Mengerjakan") {
          sheet.getRange(row, 8).setValue(jamMenit + " | " + now.getTime());
        }
        
        // Memaksa Google Sheets langsung menyimpan data detik itu juga
        SpreadsheetApp.flush();
        
        break; // Hentikan perulangan jika siswa sudah ditemukan dan diproses
      }
    }
  } catch (e) {
    Logger.log('Gagal memperbarui aktivitas siswa: ' + e.toString());
  }
}


/**
 * FITUR BARU 2: Cek Status Sebelum Masuk
 * Mencegah ganti akun atau buka tab baru jika sudah 'Curang'
 */
function checkPreFlightStatus(nisn) {
  const sheet = _getSheetByName('DATA');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(nisn).trim()) {
      return {
        statusSiswa: data[i][5], // Kolom F (Aktif/Tidak)
        detailAktivitas: data[i][6] || "" // Kolom G
      };
    }
  }
  return null;
}

/* ---------- Existing helpers (Modifikasi Sedikit) ---------- */

function authenticateStudent(nisn, tingkat, kelas, mataPelajaran){
  nisn = String(nisn || '').trim();
  const userEmail = Session.getActiveUser().getEmail(); // Proteksi Akun
  const sheet = _getSheetByName('DATA');
  const last = Math.max(sheet.getLastRow(),1);
  if(last < 2) return { success:false, message:'Tidak ada data siswa.' };
  
  const rows = sheet.getRange(2,1,last-1,7).getValues(); // Ambil sampai kolom G
  for(let i=0;i<rows.length;i++){
    const r = rows[i];
    const nisnCell = String(r[1] || '').trim();
    if(nisnCell === nisn){
      const status = String(r[5] || '');
      const detail = String(r[6] || '');

      // VALIDASI: Jika sudah pernah curang/selesai, blokir akses
      if(status !== 'Aktif' || detail.includes('Curang') || detail.includes('Selesai')) {
        return { success:false, message:'Akses ditolak. Status Anda sudah Tidak Aktif atau terdeteksi pengerjaan sebelumnya.' };
      }

      const tingkatCell = String(r[3] || '');
      const kelasCell = String(r[4] || '');
      if(tingkatCell !== String(tingkat)) return { success:false, message:'Tingkat tidak cocok.' };
      if(kelasCell !== String(kelas)) return { success:false, message:'Kelas tidak cocok.' };
      
      const meta = getSoalMeta(tingkat, mataPelajaran);
      if(!meta) return { success:false, message:'Soal tidak ditemukan.' };
      
      // Tandai sedang mengerjakan saat login berhasil
      updateStudentActivity(nisn, "Sedang Mengerjakan");

      return {
        success:true,
        data: {
          nama: r[2],
          nisn: nisnCell,
          tingkat: tingkatCell,
          kelas: kelasCell,
          mataPelajaran: mataPelajaran,
          waktuMenit: meta.durationMinutes,
          linkSoal: meta.link,
          userEmail: userEmail
        }
      };
    }
  }
  return { success:false, message:'NISN tidak ditemukan.' };
}

/* --- Sisanya adalah fungsi admin Anda (Tidak Berubah) --- */

function getLogoAndTitles(){
  const sheet = _getSheetByName('LOGO');
  const a2 = sheet.getRange('A2').getValue() || '';
  const b2 = sheet.getRange('B2').getValue() || '';
  return { logoUrl: a2, title: 'EXAM PORTAL', subtitle: b2 };
}

function getLevels(){
  const sheet = _getSheetByName('DATA');
  const last = Math.max( sheet.getLastRow(), 1 );
  if(last < 2) return [];
  const values = sheet.getRange(2,4,last-1,1).getValues().flat();
  const uniq = [...new Set(values.filter(v => v!==''))];
  return uniq;
}

function getClassesForLevel(tingkat){
  const sheet = _getSheetByName('DATA');
  const last = Math.max(sheet.getLastRow(), 1);
  if(last < 2) return [];
  const cols = sheet.getRange(2,4,last-1,2).getValues(); 
  const classes = [];
  cols.forEach(r => {
    if(String(r[0]) === String(tingkat) && r[1] && r[1] !== '') classes.push(r[1]);
  });
  return [...new Set(classes)];
}

function getSubjectsForLevel(tingkat){
  const sheet = _getSheetByName('SOAL');
  const last = Math.max(sheet.getLastRow(),1);
  if(last < 2) return [];
  const rows = sheet.getRange(2,1,last-1,sheet.getLastColumn()).getValues();
  const subjects = [];
  rows.forEach(r => {
    if(String(r[1]) === String(tingkat) && r[2] !== '') subjects.push(r[2]);
  });
  return [...new Set(subjects)];
}

function getSoalMeta(tingkat, mata){
  const sheet = _getSheetByName('SOAL');
  const last = Math.max(sheet.getLastRow(),1);
  if(last < 2) return null;
  const rows = sheet.getRange(2,1,last-1,sheet.getLastColumn()).getValues();
  for(let i=0;i<rows.length;i++){
    const r = rows[i];
    if(String(r[1]) === String(tingkat) && String(r[2]) === String(mata)){
      return { link: r[3] || '', token: r[4] || '', durationMinutes: Number(r[5]) || 0 };
    }
  }
  return null;
}

function validateToken(tingkat, mataPelajaran, token){
  const meta = getSoalMeta(tingkat, mataPelajaran);
  if(!meta) return { success:false, message:'Soal tidak ditemukan.' };
  if(String(meta.token) === String(token)) return { success:true, durationMinutes: meta.durationMinutes, link: meta.link };
  return { success:false, message:'Token salah.' };
}

function setStatusInactive(nisn){
  return updateStudentActivity(nisn, "Selesai");
}

/* ---------- Admin: autentikasi dari sheet 'AL' ---------- */
function authenticateAdmin(username, password){
  username = String(username || '').trim();
  password = String(password || '').trim();
  const sheetName = 'AL';
  const sheet = _getSheetByName(sheetName);
  const last = Math.max(sheet.getLastRow(),1);
  if(last < 2) return { success:false, message:'Data admin tidak ditemukan.' };
  const rows = sheet.getRange(2,1,last-1,3).getValues(); // A,B,C
  for(let i=0;i<rows.length;i++){
    const a = String(rows[i][0] || '').trim();
    const b = String(rows[i][1] || '').trim();
    const c = String(rows[i][2] || '').trim();
    if(a === username && b === password){
      return { success:true, name: c || username };
    }
  }
  return { success:false, message:'Username / password salah.' };
}


function getAdminCounts(){
  const sheet = _getSheetByName('DATA');
  const last = Math.max(sheet.getLastRow(),1);
  if(last < 2) return { total:0, aktif:0, tidakAktif:0 };
  const allF = sheet.getRange(2,6,last-1,1).getValues().flat(); 
  return { 
    total: allF.length, 
    aktif: allF.filter(v => v === 'Aktif').length, 
    tidakAktif: allF.filter(v => v === 'Tidak Aktif').length 
  };
}

function getStudentsAdmin() {
  try {
    const sheet = _getSheetByName('DATA');
    const last = Math.max(sheet.getLastRow(), 1);
    if (last < 2) return [];
    
    // Ambil data dari baris 2, kolom 1, sebanyak (last-1) baris dan 8 kolom (A-H)
    const rows = sheet.getRange(2, 1, last - 1, 8).getValues(); 
    const now = new Date().getTime();
    const out = [];

    for (let i = 0; i < rows.length; i++) {
      let r = rows[i];
      let rowNum = i + 2;
      let nisn = String(r[1]);
      let aktivitas = r[6] || 'Belum Mulai';
      
      // LOGIKA WAKTU
      let isiKolomH = String(r[7] || '');
      let waktuTampil = isiKolomH; 
      let waktuMulaiMurni = null;

      if (isiKolomH.includes("|")) {
        let parts = isiKolomH.split("|");
        waktuTampil = parts[0].trim(); 
        waktuMulaiMurni = Number(parts[1].trim()); 
      }

      // Durasi 60 Menit (Sesuaikan jika perlu)
      let durasiMiliDetik = 60 * 60 * 1000;

      if (aktivitas.includes("Mengerjakan") && waktuMulaiMurni) {
        if (now > (waktuMulaiMurni + durasiMiliDetik)) {
          aktivitas = "Tes tidak selesai";
          sheet.getRange(rowNum, 7).setValue(aktivitas);
          sheet.getRange(rowNum, 6).setValue("Tidak Aktif");
          
          let jamHabis = Utilities.formatDate(new Date(waktuMulaiMurni + durasiMiliDetik), "GMT+9", "HH:mm");
          waktuTampil = waktuTampil + " - " + jamHabis;
          sheet.getRange(rowNum, 8).setValue(waktuTampil);
        }
      }

      // Memasukkan data ke array out
      out.push({
        actualRow: rowNum,
        colA: r[0],
        nisn: nisn,
        nama: r[2],
        tingkat: r[3],
        kelas: r[4],
        status: r[5],
        aktivitas: aktivitas,
        waktuUjian: waktuTampil // Pastikan ini ada agar index.html tidak error
      });
    }
    return out; // Mengembalikan data ke frontend
    
  } catch (e) {
    // Jika ada error, catat di log agar Anda bisa melihatnya
    Logger.log(e.toString());
    return []; 
  }
}

/**
 * Fungsi untuk menyimpan perubahan data siswa dari modal edit Admin
 * @param {number} actualRow - Baris di Google Sheets
 * @param {object} data - Objek berisi data baru siswa
 */
function updateStudentAdmin(actualRow, data) {
  try {
    const sheet = _getSheetByName('DATA');
    actualRow = Number(actualRow);

    // Pastikan baris yang diakses valid
    if (actualRow < 2) throw new Error("Baris data tidak valid.");

    // Kita update kolom A sampai F
    // Kolom A: No/ID, B: NISN, C: Nama, D: Tingkat, E: Kelas, F: Status
    sheet.getRange(actualRow, 1, 1, 6).setValues([
      [
        data.colA, 
        "'" + data.nisn, // Tambahkan kutip agar NISN tidak berubah jadi format angka scientific
        data.nama, 
        data.tingkat, 
        data.kelas, 
        data.status
      ]
    ]);

    // Opsional: Jika status diubah jadi 'Aktif' kembali, 
    // kita reset kolom Aktivitas (G) agar siswa bisa ujian lagi
    if (data.status === 'Aktif') {
      sheet.getRange(actualRow, 7).setValue('Belum Mulai');
      sheet.getRange(actualRow, 8).setValue(''); // Kosongkan juga jam ujian lama
    }

    return { success: true };
  } catch (e) {
    Logger.log("Error updateStudentAdmin: " + e.toString());
    return { success: false, message: "Gagal menyimpan: " + e.toString() };
  }
}

function deleteStudentAdmin(actualRow){
  _getSheetByName('DATA').deleteRow(Number(actualRow));
  return { success:true };
}

function getSoalAdmin(){
  const sheet = _getSheetByName('SOAL');
  const last = Math.max(sheet.getLastRow(),1);
  if(last < 2) return [];
  return sheet.getRange(2,1,last-1,6).getValues().map((r, i) => ({
    actualRow: i+2, colA: r[0], tingkat: r[1], mata: r[2], link: r[3], token: r[4], waktu: r[5]
  }));
}

function updateSoalAdmin(actualRow, data){
  _getSheetByName('SOAL').getRange(Number(actualRow),1,1,6).setValues([[data.colA, data.tingkat, data.mata, data.link, data.token, data.waktu]]);
  return { success:true };
}

function deleteSoalAdmin(actualRow){
  _getSheetByName('SOAL').deleteRow(Number(actualRow));
  return { success:true };
}

function addStudentAdmin(data) {
  _getSheetByName('DATA').appendRow([data.colA, "'" + data.nisn, data.nama, data.tingkat, data.kelas, data.status || 'Aktif', '']);
  return { success: true };
}

function addSoalAdmin(data) {
  _getSheetByName('SOAL').appendRow([data.colA, data.tingkat, data.mata, data.link, data.token, data.waktu]);
  return { success: true };
}

function fungsiCariDiBackend(namaSiswa) {
  try {
    const sheet = _getSheetByName('DATA');
    const lastRow = Math.max(sheet.getLastRow(), 1);
    
    // Jika tidak ada data sama sekali selain header
    if (lastRow < 2) return []; 

    // Ambil semua data siswa dari baris 2 sampai baris terakhir
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    
    // Filter data berdasarkan nama yang dicari (tidak sensitif huruf besar/kecil)
    const hasilCari = data.map((r, i) => ({
      actualRow: i + 2, // Menyimpan nomor baris asli di spreadsheet
      colA: r[0],
      nisn: r[1],
      nama: r[2],
      tingkat: r[3],
      kelas: r[4],
      status: r[5]
    })).filter(siswa => siswa.nama.toLowerCase().includes(namaSiswa.toLowerCase()));

    return hasilCari; // Mengembalikan array berisi siswa yang cocok saja
  } catch (e) {
    Logger.log("Error fungsiCariDiBackend: " + e.toString());
    return [];
  }
}
function dapatkanUrlApp() {
  return ScriptApp.getService().getUrl();
}
