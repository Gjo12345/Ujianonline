// ==================== KONFIGURASI ====================
const SPREADSHEET_ID = '16ocKmxFu6OtwdaKXVQZ7KsDe01SyRmsr8fVxZ8uqgfc'; // Ganti dengan ID spreadsheet Anda
// const WAKTU_UJIAN = 30; // Waktu ujian dalam menit

// ==================== FUNGSI UTAMA (ROUTING) ====================

/**
 * Fungsi utama yang menangani permintaan GET dan routing halaman.
 */
function doGet(e) {
  const page = e.parameter.page || 'Index';
  const sessionId = e.parameter.sessionId; // Ambil sessionId dari URL
  let template;
  let title = 'Sistem Ujian Online';

  // Halaman yang memerlukan sesi yang valid
  if (page === 'Ujian' || page === 'Hasil' || page === 'Admin') {
    const session = getUserSession(sessionId);
    let hasAccess = session.success;

    // Pemeriksaan peran (role) tambahan
    // --- UNTUK DEBUGGING ---
    /*
    if (hasAccess) {
      if (page === 'Admin' && session.data.role !== 'admin') {
        hasAccess = false; // Hanya admin yang bisa akses halaman Admin
      }
      if ((page === 'Ujian' || page === 'Hasil') && session.data.role !== 'peserta') {
        hasAccess = false; // Hanya peserta yang bisa akses Ujian/Hasil
      }
    }
    */

    if (!hasAccess) {
      // Jika sesi tidak valid atau peran tidak sesuai, tampilkan pesan akses ditolak langsung
      const loginUrl = getScriptUrl();
      const htmlOutput = `
        <!DOCTYPE html>
        <html>
          <head>
            <base target="_top">
            <title>Akses Ditolak</title>
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
            <style>
              body, html {
                height: 100%;
                display: flex;
                align-items: center;
                justify-content: center;
                background-color: #f8f9fa;
              }
              .card {
                width: 100%;
                max-width: 400px;
                text-align: center;
                padding: 2rem;
                border-radius: 1rem;
                box-shadow: 0 4px 8px rgba(0,0,0,0.1);
              }
            </style>
          </head>
          <body>
            <div class="card">
              <h3 class="card-title mb-4">Akses Ditolak</h3>
              <p class="card-text">Anda tidak memiliki izin untuk mengakses halaman ini. Silakan login terlebih dahulu.</p>
              <a href="${loginUrl}" class="btn btn-primary mt-3">Kembali ke Halaman Login</a>
            </div>
          </body>
        </html>
      `;
      return HtmlService.createHtmlOutput(htmlOutput)
        .setTitle('Akses Ditolak')
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .setFaviconUrl('https://w7.pngwing.com/pngs/184/833/png-transparent-exam-test-checklist-online-learning-education-online-document-online-learning-icon.png');
    }
  }


  switch (page) {
    case 'Index':
      template = HtmlService.createTemplateFromFile('Index');
      title = 'Sistem Ujian Online';
      template.errorMessage = e.parameter.error || null;
      
      // Fetch config data and pass it to the template
      const configResponse = getConfiguration();
      template.config = configResponse.success ? configResponse.data : {};
      
      break;
    case 'Ujian':
      template = HtmlService.createTemplateFromFile('Ujian');
      title = 'Ujian Berlangsung';
      break;
    case 'Hasil':
      template = HtmlService.createTemplateFromFile('Hasil');
      title = 'Hasil Ujian';
      break;
    // --- PENAMBAHAN RUTE ADMIN ---
    case 'Admin':
      template = HtmlService.createTemplateFromFile('Admin');
      title = 'Dashboard Admin';
      break;
    // ---------------------------------
    default:
      template = HtmlService.createTemplateFromFile('Index');
      title = 'Sistem Ujian Online';
      template.errorMessage = null;
      break;
  }

  return template.evaluate()
    .setTitle(title)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://w7.pngwing.com/pngs/184/833/png-transparent-exam-test-checklist-online-learning-education-online-document-online-learning-icon.png');
}

/**
 * Menyertakan konten file HTML lain.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Mendapatkan URL web app.
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// ==================== FUNGSI MANAJEMEN SESI ====================

function createSession(data) {
  const sessionId = Utilities.getUuid();
  const cache = CacheService.getUserCache();
  cache.put(sessionId, JSON.stringify(data), 7200); // Sesi 2 jam
  return sessionId;
}

function getUserSession(sessionId) {
  try {
    const cache = CacheService.getUserCache();
    const sessionData = cache.get(sessionId);
    if (sessionData) {
      return { success: true, data: JSON.parse(sessionData) };
    }
    return { success: false, message: 'Sesi tidak ditemukan atau telah berakhir.' };
  } catch(e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

function clearSession(sessionId) {
  try {
    const cache = CacheService.getUserCache();
    cache.remove(sessionId);
    return { success: true };
  } catch(e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ==================== FUNGSI LOGIN & VALIDASI (MODIFIKASI) ====================

/**
 * Fungsi untuk mengambil PIN Sesi dari sheet 'setting' (sel A2).
 */
function getSessionPin() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const settingSheet = ss.getSheetByName('setting');
    if (!settingSheet) {
      return null;
    }
    // Mengambil PIN dari sel A2
    const pin = settingSheet.getRange('A2').getValue().toString().trim();
    return pin;
  } catch (e) {
    Logger.log('Gagal mengambil PIN Sesi: ' + e.message);
    return null; // Gagal mengambil PIN
  }
}

/**
 * Fungsi login baru yang menangani semua tipe login.
 * Dipanggil dari Index.html.
 */
function handleLogin(loginData) {
  if (loginData.loginType === 'peserta') {
    const result = loginPeserta(loginData.noPeserta, loginData.password, loginData.pinSesi);
    // Tambahkan redirectPage jika sukses (untuk peserta)
    if (result.success) {
      result.redirectPage = 'Ujian'; 
    }
    return result;
    
  } else if (loginData.loginType === 'admin') {
     // loginAdmin SEKARANG MENGEMBALIKAN redirectPage
    return loginAdmin(loginData.username, loginData.password);
    
  } else if (loginData.loginType === 'lihat_nilai') {
    return loginLihatNilai(loginData.noPeserta);
  } else {
    return { success: false, message: 'Tipe login tidak valid.' };
  }
}

function loginLihatNilai(noPeserta) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const dataPesertaSheet = ss.getSheetByName('data_peserta');
    if (!dataPesertaSheet) {
        return { success: false, message: 'Sheet "data_peserta" tidak ditemukan.' };
    }
    const data = dataPesertaSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == noPeserta) {
        const userData = {
          noPeserta: data[i][0],
          nama: data[i][1],
          role: 'peserta' // Anggap sebagai peserta untuk melihat nilai
        };
        const sessionId = createSession(userData);

        return {
          success: true,
          sessionId: sessionId,
          data: userData,
          redirectPage: 'Hasil'
        };
      }
    }

    return {
      success: false,
      message: 'Nomor peserta tidak ditemukan!'
    };
  } catch (error) {
    Logger.log('Error in loginLihatNilai: ' + error.message);
    return {
      success: false,
      message: 'Terjadi kesalahan: ' + error.message
    };
  }
}

function loginPeserta(noPeserta, password, pinSesi) {
  try {
    const correctPin = getSessionPin();
    if (!correctPin || pinSesi !== correctPin) {
      return {
        success: false,
        message: 'PIN Sesi yang Anda masukkan salah!'
      };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const dataPesertaSheet = ss.getSheetByName('data_peserta');
    if (!dataPesertaSheet) {
        return { success: false, message: 'Sheet "data_peserta" tidak ditemukan. Jalankan inisialisasi.' };
    }
    const data = dataPesertaSheet.getDataRange().getValues();
    const headers = data[0];
    const statusColIndex = headers.indexOf('Status');
    const subjectSheets = getSubjectSheets();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == noPeserta && data[i][2] == password) {
        if (statusColIndex > -1 && data[i][statusColIndex].toLowerCase() === 'tidak aktif') {
          return {
            success: false,
            message: 'Akun Anda tidak aktif. Silakan hubungi administrator.'
          };
        }
        
        const nilaiSheet = ss.getSheetByName('peserta');
        let needsReset = false; // Flag for retake

        if (nilaiSheet) {
          const nilaiData = nilaiSheet.getDataRange().getValues();
          const nilaiHeaders = nilaiData[0];
          const statusUjianCol = nilaiHeaders.indexOf('Status Ujian');
          
          let nilaiRow = null;
          let nilaiRowIndex = -1;
          for (let j = 1; j < nilaiData.length; j++) {
              if (nilaiData[j][1] == noPeserta) {
                  nilaiRow = nilaiData[j];
                  nilaiRowIndex = j;
                  break;
              }
          }

          if (nilaiRow) {
              // Check if it's a retake
              if (statusUjianCol > -1 && nilaiRow[statusUjianCol] === 'Ujian Ulang') {
                  needsReset = true;
                  // Reset the status back to blank after acknowledging the retake
                  nilaiSheet.getRange(nilaiRowIndex + 1, statusUjianCol + 1).setValue('');
              } else {
                  // Original logic to check if exam is already fully completed
                  let subjectsDoneCount = 0;
                  subjectSheets.forEach(sheetName => {
                      const subjectDisplayName = `Nilai ${sheetName.replace('soal_', '').replace(/_/g, ' ')}`;
                      const nilaiColIndex = nilaiHeaders.indexOf(subjectDisplayName);
                      if (nilaiColIndex > -1 && nilaiRow[nilaiColIndex] !== '') {
                          subjectsDoneCount++;
                      }
                  });

                  if (subjectsDoneCount >= subjectSheets.length) {
                      return {
                          success: false,
                          message: 'Anda sudah menyelesaikan semua mata pelajaran dalam ujian ini.'
                      };
                  }
              }
          }
        }

        const userData = {
          noPeserta: data[i][0],
          nama: data[i][1],
          rowIndex: i + 1,
          role: 'peserta'
        };
        const sessionId = createSession(userData);

        return {
          success: true,
          sessionId: sessionId,
          data: userData,
          needsReset: needsReset // Send the flag to the frontend
        };
      }
    }

    return {
      success: false,
      message: 'Nomor peserta atau password salah!'
    };
  } catch (error) {
    Logger.log('Error in loginPeserta: ' + error.message);
    return {
      success: false,
      message: 'Terjadi kesalahan: ' + error.message
    };
  }
}

function loginAdmin(username, password) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const settingSheet = ss.getSheetByName('setting');
    if (!settingSheet) {
      return { success: false, message: 'Sheet "setting" tidak ditemukan.' };
    }
    
    // Ambil kredensial admin dari B2 (username) dan C2 (password)
    const adminUser = settingSheet.getRange('B2').getValue().toString().trim();
    const adminPass = settingSheet.getRange('C2').getValue().toString().trim();

    if (username === adminUser && password === adminPass) {
      const adminData = {
        username: adminUser,
        role: 'admin' // Tambahkan role
      };
      const sessionId = createSession(adminData);
      
      return {
        success: true,
        sessionId: sessionId,
        data: adminData,
        redirectPage: 'Admin' // <-- INI DIA PERBAIKANNYA
      };
    } else {
      return {
        success: false,
        message: 'Username atau password Admin salah!'
      };
    }
  } catch (error) {
    return {
      success: false,
      message: 'Terjadi kesalahan admin: ' + error.message
    };
  }
}

// ==================== FUNGSI BARU UNTUK MENDAPATKAN SEMUA SHEET SOAL ====================

function getSubjectSheets() {

  try {

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const allSheets = ss.getSheets();

    const subjectSheets = allSheets

      .map(sheet => sheet.getName())

      .filter(name => name.toLowerCase().includes('soal') && name.toLowerCase() !== 'soal_esai');

    return subjectSheets;

  } catch (e) {

    Logger.log('Gagal mengambil daftar mata pelajaran: ' + e.message);

    return [];

  }

}

function getExamSubjects() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSheets = ss.getSheets();
    const allSheetNames = allSheets.map(sheet => sheet.getName());

    const subjectSheets = allSheetNames
      .filter(name => name.toLowerCase().startsWith('soal_') && name.toLowerCase() !== 'soal_esai');

    const hasEssay = allSheetNames.includes('soal_esai');

    return { success: true, subjects: subjectSheets, hasEssay: hasEssay };

  } catch (e) {
    Logger.log('Gagal mengambil daftar mata pelajaran ujian: ' + e.message);
    return { success: false, message: 'Gagal mengambil daftar mata pelajaran: ' + e.message };
  }
}



// ==================== FUNGSI LOAD SOAL (DIMODIFIKASI) ====================

function loadSoal(sheetName) {

  try {

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const sheetSoal = ss.getSheetByName(sheetName);

    if (!sheetSoal) {

      return {

        error: true,

        message: 'Sheet dengan nama "' + sheetName + '" tidak ditemukan.'

      };

    }



    const data = sheetSoal.getDataRange().getValues();

    

    const soalArray = [];

    // Read all questions into a single array first

    for (let i = 1; i < data.length; i++) {

      if (data[i].join('').trim() === '') continue;

      soalArray.push({

        no: data[i][0],

        tipe: data[i][1],

        pertanyaan: data[i][2],

        opsiA: data[i][3],

        opsiB: data[i][4],

        opsiC: data[i][5],

        opsiD: data[i][6],

        opsiE: data[i][7],

        jawaban: data[i][8]

      });

    }

    

    // Separate questions by type

    const pilihanGandaArray = soalArray.filter(soal => soal.tipe === 'Pilihan Ganda');

    const esaiArray = soalArray.filter(soal => soal.tipe === 'Esai');



    // Shuffle within each type

    const shuffledPG = shuffleArray(pilihanGandaArray);

    const shuffledEsai = shuffleArray(esaiArray);



    // Combine them: multiple choice first, then essays

    const finalSoalOrder = shuffledPG.concat(shuffledEsai);

    return finalSoalOrder;



  } catch (error) {

    return {

      error: true,

      message: 'Gagal memuat soal dari sheet "' + sheetName + '": ' + error.message

    };

  }

}



// ==================== FUNGSI SUBMIT JAWABAN (DIMODIFIKASI) ====================

function submitJawaban(noPeserta, jawabanPeserta, alasan, sheetName) {
  try {
    jawabanPeserta = jawabanPeserta || {};
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetSoal = ss.getSheetByName(sheetName);
    if (!sheetSoal) {
      return { success: false, message: 'Sheet soal tidak ditemukan: ' + sheetName };
    }

    const nilaiSheet = ss.getSheetByName('peserta');
    const rekapJawabanSheet = ss.getSheetByName('rekap');
    const dataSoal = sheetSoal.getDataRange().getValues();
    
    const semuaSoal = {};
    let totalSoalPG = 0;
    let maxSoalNum = 0;

    for (let i = 1; i < dataSoal.length; i++) {
      if (dataSoal[i].join('').trim() === '') continue;
      const noSoal = parseInt(dataSoal[i][0], 10);
      if (isNaN(noSoal)) continue;
      
      maxSoalNum = Math.max(maxSoalNum, noSoal);
      const tipeSoal = dataSoal[i][1];
      semuaSoal[noSoal] = {
        tipe: tipeSoal,
        pertanyaan: dataSoal[i][2],
        jawabanBenar: String(dataSoal[i][8])
      };
      if (tipeSoal === 'Pilihan Ganda') {
        totalSoalPG++;
      }
    }

    let benarPG = 0;
    const timestamp = new Date();
    
    const dataPesertaSheet = ss.getSheetByName('data_peserta');
    const dataPeserta = dataPesertaSheet.getDataRange().getValues();
    let namaPeserta = '';
    for (let i = 1; i < dataPeserta.length; i++) {
      if (dataPeserta[i][0] == noPeserta) {
        namaPeserta = dataPeserta[i][1];
        break;
      }
    }

    if (namaPeserta === '') {
      return { success: false, message: 'Data peserta tidak ditemukan di sheet data_peserta.' };
    }

    // --- LOGIKA REKAP JAWABAN DIPINDAHKAN KE submitAllAnswersAndFinalize ---

    // Hitung skor PG
    for (const noSoal in jawabanPeserta) {
      const soalInfo = semuaSoal[noSoal];
      const jawabanSiswa = jawabanPeserta[noSoal];
      if (!soalInfo) continue;
      if (soalInfo.tipe === 'Pilihan Ganda') {
        if (jawabanSiswa === soalInfo.jawabanBenar) {
          benarPG++;
        }
      }
    }

    const nilai = totalSoalPG > 0 ? (benarPG / totalSoalPG) * 100 : 0;
    const salahPG = totalSoalPG - benarPG;

    if (nilaiSheet) {
      const dataNilai = nilaiSheet.getDataRange().getValues();
      const headersNilai = dataNilai[0];
      let nilaiRowIndex = -1;

      for (let i = 1; i < dataNilai.length; i++) {
        if (dataNilai[i][1] == noPeserta) {
          nilaiRowIndex = i + 1;
          break;
        }
      }

      const subjectDisplayName = sheetName.replace('soal_', '').replace(/_/g, ' ');
      const nilaiColName = `Nilai ${subjectDisplayName}`;
      const nilaiColIndex = headersNilai.indexOf(nilaiColName);
      const statusColIndex = headersNilai.indexOf('Status Ujian'); // Dapatkan indeks kolom status

      if (nilaiRowIndex !== -1) {
        if (nilaiColIndex > -1) {
          nilaiSheet.getRange(nilaiRowIndex, nilaiColIndex + 1).setValue(nilai.toFixed(2));
        }
        // Perbarui status jika submit normal
        if (statusColIndex > -1 && alasan === 'Ujian Selesai') {
          nilaiSheet.getRange(nilaiRowIndex, statusColIndex + 1).setValue('Selesai');
        }
        nilaiSheet.getRange(nilaiRowIndex, 1).setValue(timestamp);
      } else {
        const newRowData = Array(headersNilai.length).fill('');
        newRowData[0] = timestamp;
        newRowData[1] = noPeserta;
        newRowData[2] = namaPeserta;
        if (nilaiColIndex > -1) {
          newRowData[nilaiColIndex] = nilai.toFixed(2);
        }
        // Set status untuk baris baru
        if (statusColIndex > -1 && alasan === 'Ujian Selesai') {
          newRowData[statusColIndex] = 'Selesai';
        }
        nilaiSheet.appendRow(newRowData);
        nilaiRowIndex = nilaiSheet.getLastRow();
      }

      const allSubjectSheets = getSubjectSheets();
      let totalNilai = 0;
      let subjectsDone = 0;
      const updatedRowData = nilaiSheet.getRange(nilaiRowIndex, 1, 1, headersNilai.length).getValues()[0];
      
      allSubjectSheets.forEach(subj => {
        const simpleName = subj.replace('soal_', '').replace(/_/g, ' ');
        const colIdx = headersNilai.indexOf(`Nilai ${simpleName}`);
        if (colIdx > -1 && updatedRowData[colIdx] !== '') {
          totalNilai += parseFloat(updatedRowData[colIdx]);
          subjectsDone++;
        }
      });
      
      const nilaiEsaiColIdx = headersNilai.indexOf('Nilai Esai');
      if (nilaiEsaiColIdx > -1 && updatedRowData[nilaiEsaiColIdx] !== '') {
          const nilaiEsai = parseFloat(updatedRowData[nilaiEsaiColIdx]);
          if(!isNaN(nilaiEsai)){
            totalNilai += nilaiEsai;
            subjectsDone++;
          }
      }

      const rataRata = subjectsDone > 0 ? totalNilai / subjectsDone : 0;
      const totalNilaiColIndex = headersNilai.indexOf('Total Nilai');
      const rataRataColIndex = headersNilai.indexOf('Rata-rata Nilai');

      if (totalNilaiColIndex > -1) nilaiSheet.getRange(nilaiRowIndex, totalNilaiColIndex + 1).setValue(totalNilai.toFixed(2));
      if (rataRataColIndex > -1) nilaiSheet.getRange(nilaiRowIndex, rataRataColIndex + 1).setValue(rataRata.toFixed(2));
    }

    return {
      success: true,
      nilai: nilai.toFixed(2),
      benar: benarPG,
      salah: salahPG,
      totalSoal: totalSoalPG,
      subject: sheetName
    };
  } catch (error) {
    Logger.log('Error in submitJawaban: ' + error.message + ' Stack: ' + error.stack);
    return {
      success: false,
      message: 'Gagal menyimpan jawaban untuk ' + sheetName + ': ' + error.message
    };
  }
}

// ==================== FUNGSI HELPER ====================

function getWaktuUjian() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const settingSheet = ss.getSheetByName('setting');
    if (!settingSheet) {
      return 30; // Default value if sheet not found
    }
    const waktu = settingSheet.getRange('G2').getValue();
    return !isNaN(waktu) && waktu > 0 ? waktu : 30; // Return saved time or default 30
  } catch (e) {
    Logger.log('Gagal mengambil waktu ujian dari sheet: ' + e.message);
    return 30; // Default value on error
  }
}

function shuffleArray(array) {

  const newArray = [...array];

  for (let i = newArray.length - 1; i > 0; i--) {

    const j = Math.floor(Math.random() * (i + 1));

    [newArray[i], newArray[j]] = [newArray[j], newArray[i]];

  }

  return newArray;

}

// ==================== FUNGSI INISIALISASI SHEET (MODIFIKASI) ====================

function onOpen() {

  SpreadsheetApp.getUi()

      .createMenu('Admin')

      .addItem('Setup & Inisialisasi Sheet', 'initializeSheets')

      .addItem('Perbarui Header Sheet Peserta', 'updatePesertaHeaders') 
      .addToUi();

}

function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetNames = ss.getSheets().map(s => s.getName());
  const requiredSubjects = ['soal_matematika', 'soal_bahasa_indonesia', 'soal_bahasa_inggris', 'soal_esai'];

  // 1. Buat sheet 'data_peserta' untuk login
  if (!sheetNames.includes('data_peserta')) {
    const dataPesertaSheet = ss.insertSheet('data_peserta');
    dataPesertaSheet.appendRow(['No Peserta', 'Nama Lengkap', 'Password', 'Status']);
    dataPesertaSheet.appendRow(['PESERTA01', 'Nama Peserta Satu', 'pass01', 'aktif']);
  }

  // 2. Buat sheet 'setting'
  if (!sheetNames.includes('setting')) {
    const settingSheet = ss.insertSheet('setting');
    settingSheet.getRange('A1:G1').setValues([['PIN Sesi', 'Admin Username', 'Admin Password', 'Judul Login', 'URL Logo', 'Sub Judul Login', 'Waktu Ujian (menit)']]).setFontWeight('bold');
    settingSheet.getRange('A2:G2').setValues([['123456', 'admin', 'admin123', 'UJIAN PSIKOTEST', 'https://upload.wikimedia.org/wikipedia/commons/9/98/Kota_Bengkulu.png', 'Sistem Ujian Online Kota Bengkulu', 30]]);
  }

  // 3. Buat sheet 'peserta' untuk REKAP NILAI
  if (!sheetNames.includes('peserta')) {
    const pesertaSheet = ss.insertSheet('peserta');
    const nilaiHeaders = ['Timestamp', 'Nomor Peserta', 'Nama Peserta', 'Status Ujian', 'Durasi Pengerjaan (menit)'];
    const allSubjectSheets = getSubjectSheets();
    allSubjectSheets.forEach(subject => {
      const subjectName = subject.replace('soal_', '').replace(/_/g, ' ');
      nilaiHeaders.push(`Nilai ${subjectName}`);
    });
    nilaiHeaders.push('Nilai Esai', 'Total Nilai', 'Rata-rata Nilai');
    pesertaSheet.appendRow(nilaiHeaders);
  }

  // 4a. Buat sheet 'rekap_pilihanganda'
  if (!sheetNames.includes('rekap_pilihanganda')) {
    const rekapPGSheet = ss.insertSheet('rekap_pilihanganda');
    const headersPG = ['Timestamp', 'Nomor Peserta', 'Nama Peserta', 'Mata Pelajaran'];
    for (let i = 1; i <= 50; i++) {
      headersPG.push('Soal ' + i);
    }
    rekapPGSheet.appendRow(headersPG);
  }

  // 4b. Buat sheet 'rekap_esai'
  if (!sheetNames.includes('rekap_esai')) {
    const rekapEsaiSheet = ss.insertSheet('rekap_esai');
    const headersEsai = ['Timestamp', 'Nomor Peserta', 'Nama Peserta', 'Mata Pelajaran'];
    for (let i = 1; i <= 20; i++) { // Assume max 20 essay questions
      headersEsai.push('Esai ' + i);
    }
    rekapEsaiSheet.appendRow(headersEsai);
  }

  // 4c. Hapus sheet 'rekap' lama jika ada
  const oldRekapSheet = ss.getSheetByName('rekap');
  if (oldRekapSheet) {
    ss.deleteSheet(oldRekapSheet);
  }

  // 5. Hapus sheet 'Jawaban Esai' jika ada
  const essaySheet = ss.getSheetByName('Jawaban Esai');
  if (essaySheet) {
    ss.deleteSheet(essaySheet);
  }

  // 6. Buat sheet-sheet soal
  requiredSubjects.forEach(subjectName => {
    if (!sheetNames.includes(subjectName)) {
      const soalSheet = ss.insertSheet(subjectName);
      soalSheet.appendRow(['No', 'Tipe', 'Pertanyaan', 'OpsiA', 'OpsiB', 'OpsiC', 'OpsiD', 'OpsiE', 'Jawaban']);
      soalSheet.appendRow([1, 'Pilihan Ganda', 'Contoh soal PG ' + subjectName, 'A', 'B', 'C', 'D', 'E', 'A']);
      soalSheet.appendRow([2, 'Esai', 'Contoh soal Esai ' + subjectName, '', '', '', '', '', '']);
    }
  });
  
  SpreadsheetApp.getUi().alert('Inisialisasi sheet dengan struktur baru telah selesai. Sheet "data_peserta" digunakan untuk login, "peserta" untuk nilai, dan "rekap" untuk jawaban.');
}

function submitAllAnswersAndFinalize(noPeserta, allAnswers, totalWaktuUjian, sisaDetik, status) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const nilaiSheet = ss.getSheetByName('peserta');
    const rekapPGSheet = ss.getSheetByName('rekap_pilihanganda');
    const rekapEsaiSheet = ss.getSheetByName('rekap_esai');
    const dataPesertaSheet = ss.getSheetByName('data_peserta');

    if (!nilaiSheet || !dataPesertaSheet) {
      return { success: false, message: 'Sheet penting (peserta atau data_peserta) tidak ditemukan.' };
    }

    // 1. Batch Read all necessary data
    const dataNilai = nilaiSheet.getDataRange().getValues();
    const headersNilai = dataNilai[0];
    const dataPeserta = dataPesertaSheet.getDataRange().getValues();
    const allSubjectSheetNames = getSubjectSheets().concat(['soal_esai']);
    
    const allSoalData = {};
    allSubjectSheetNames.forEach(sheetName => {
      const sheetSoal = ss.getSheetByName(sheetName);
      if (sheetSoal) {
        const dataSoal = sheetSoal.getDataRange().getValues();
        const soalMap = {};
        let pgCount = 0;
        for (let i = 1; i < dataSoal.length; i++) {
          if (dataSoal[i].join('').trim() === '') continue;
          const noSoal = parseInt(dataSoal[i][0], 10);
          if (isNaN(noSoal)) continue;
          const tipeSoal = dataSoal[i][1];
          soalMap[noSoal] = {
            tipe: tipeSoal,
            opsiA: dataSoal[i][3],
            opsiB: dataSoal[i][4],
            opsiC: dataSoal[i][5],
            opsiD: dataSoal[i][6],
            opsiE: dataSoal[i][7],
            jawabanBenar: String(dataSoal[i][8])
          };
          if (tipeSoal === 'Pilihan Ganda') {
            pgCount++;
          }
        }
        allSoalData[sheetName] = { soal: soalMap, totalPG: pgCount };
      }
    });

    // 2. Find participant info
    let namaPeserta = '';
    for (let i = 1; i < dataPeserta.length; i++) {
      if (dataPeserta[i][0] == noPeserta) {
        namaPeserta = dataPeserta[i][1];
        break;
      }
    }
    if (namaPeserta === '') return { success: false, message: 'Data peserta tidak ditemukan.' };

    let nilaiRowIndex = -1;
    let isNewNilaiRow = false;
    for (let i = 1; i < dataNilai.length; i++) {
      if (dataNilai[i][1] == noPeserta) {
        nilaiRowIndex = i;
        break;
      }
    }

    if (nilaiRowIndex === -1) {
      isNewNilaiRow = true;
      const newRow = Array(headersNilai.length).fill('');
      newRow[0] = new Date();
      newRow[1] = noPeserta;
      newRow[2] = namaPeserta;
      dataNilai.push(newRow);
      nilaiRowIndex = dataNilai.length - 1;
    }

    const timestamp = new Date();
    dataNilai[nilaiRowIndex][0] = timestamp;

    // --- New Rekap Logic ---
    const pgAnswers = {};
    const esaiAnswers = {};

    // 3. Process and Separate all answers by type
    for (const subjectSheet in allAnswers) {
      if (Object.hasOwnProperty.call(allAnswers, subjectSheet)) {
        const answersForSubject = allAnswers[subjectSheet];
        const soalInfoForSubject = allSoalData[subjectSheet];
        if (!soalInfoForSubject) continue;

        const subjectDisplayName = subjectSheet.replace('soal_', '').replace(/_/g, ' ');
        pgAnswers[subjectDisplayName] = {};
        esaiAnswers[subjectDisplayName] = {};

        let benarPG = 0;
        const totalSoalPG = soalInfoForSubject.totalPG;
        
        for (const noSoal in answersForSubject) {
          const soalDetail = soalInfoForSubject.soal[noSoal];
          if (soalDetail) {
            if (soalDetail.tipe === 'Pilihan Ganda') {
              pgAnswers[subjectDisplayName][noSoal] = answersForSubject[noSoal];
              if (answersForSubject[noSoal] === soalDetail.jawabanBenar) {
                benarPG++;
              }
            } else if (soalDetail.tipe === 'Esai') {
              esaiAnswers[subjectDisplayName][noSoal] = answersForSubject[noSoal];
            }
          }
        }

        if (totalSoalPG > 0) {
            const nilai = (benarPG / totalSoalPG) * 100;
            const nilaiColName = `Nilai ${subjectDisplayName}`;
            const nilaiColIndex = headersNilai.indexOf(nilaiColName);
            if (nilaiColIndex > -1) {
              dataNilai[nilaiRowIndex][nilaiColIndex] = nilai.toFixed(2);
            }
        }
      }
    }

    const processRekap = (sheet, answersBySubject, soalDetails) => {
      if (!sheet) return;
      const rekapData = sheet.getDataRange().getValues();
      if (!rekapData || rekapData.length === 0) return;

      const headers = rekapData[0];
      const noPesertaCol = headers.indexOf('Nomor Peserta');
      const mapelCol = headers.indexOf('Mata Pelajaran');
      
      const userExistingRows = new Map();
      for (let i = 1; i < rekapData.length; i++) {
        if (rekapData[i][noPesertaCol] == noPeserta) {
          userExistingRows.set(rekapData[i][mapelCol], i);
        }
      }

      const rowsToAppend = [];

      for (const subjectName in answersBySubject) {
        const answers = answersBySubject[subjectName];
        if (Object.keys(answers).length === 0) continue;

        const answerValues = [];
        const maxSoalNum = Object.keys(answers).length > 0 ? Math.max(...Object.keys(answers).map(k => parseInt(k, 10))) : 0;
        
        // --- MODIFIKASI DIMULAI DI SINI ---
        const isPilihanGanda = sheet.getName() === 'rekap_pilihanganda';
        const soalSheetName = 'soal_' + subjectName.toLowerCase().replace(/ /g, '_');
        const currentSoalDetails = isPilihanGanda && soalDetails ? soalDetails[soalSheetName] : null;

        for (let i = 1; i <= maxSoalNum; i++) {
          const jawabanMentah = answers[i] || '';
          let jawabanDisplay = jawabanMentah;

          if (isPilihanGanda && jawabanMentah && currentSoalDetails && currentSoalDetails.soal[i]) {
            const soalInfo = currentSoalDetails.soal[i];
            if (soalInfo.tipe === 'Pilihan Ganda') {
               // Buat peta terbalik untuk mencari kunci (A/B/C) berdasarkan nilai (teks jawaban)
               const reverseOsiMap = {};
               if (soalInfo.opsiA) reverseOsiMap[soalInfo.opsiA] = 'A';
               if (soalInfo.opsiB) reverseOsiMap[soalInfo.opsiB] = 'B';
               if (soalInfo.opsiC) reverseOsiMap[soalInfo.opsiC] = 'C';
               if (soalInfo.opsiD) reverseOsiMap[soalInfo.opsiD] = 'D';
               if (soalInfo.opsiE) reverseOsiMap[soalInfo.opsiE] = 'E';

               const kunciOpsi = reverseOsiMap[jawabanMentah];
               if (kunciOpsi) {
                jawabanDisplay = `${kunciOpsi}. ${jawabanMentah}`;
               }
            }
          }
          answerValues.push(jawabanDisplay);
        }
        // --- MODIFIKASI SELESAI ---
        
        const newRowData = [timestamp, noPeserta, namaPeserta, subjectName, ...answerValues];
        
        if (userExistingRows.has(subjectName)) {
          const rowIndex = userExistingRows.get(subjectName);
          rekapData[rowIndex] = newRowData;
        } else {
          rowsToAppend.push(newRowData);
        }
      }

      const allRows = rekapData.concat(rowsToAppend);
      const maxWidth = Math.max(...allRows.map(r => r.length));

      const soalPrefix = sheet.getName().includes('esai') ? 'Esai ' : 'Soal ';
      const baseHeaderCount = 4;
      if (maxWidth > headers.length) {
        for (let i = headers.length - baseHeaderCount + 1; i <= maxWidth - baseHeaderCount; i++) {
          headers.push(soalPrefix + i);
        }
      }
      
      const rectangularData = allRows.map(r => r.concat(Array(maxWidth - r.length).fill('')));

      sheet.clearContents();
      sheet.getRange(1, 1, rectangularData.length, maxWidth).setValues(rectangularData);
    };

    processRekap(rekapPGSheet, pgAnswers, allSoalData);
    processRekap(rekapEsaiSheet, esaiAnswers, null);

    // 4. Finalize: fill empty scores, calculate total/avg, set status
    let totalNilai = 0;
    let subjectsDone = 0;
    
    headersNilai.forEach((header, colIndex) => {
        if (header.startsWith('Nilai ')) {
          if (dataNilai[nilaiRowIndex][colIndex] === '') {
            dataNilai[nilaiRowIndex][colIndex] = 0;
          }
          const score = parseFloat(dataNilai[nilaiRowIndex][colIndex]);
          if (!isNaN(score)) {
            totalNilai += score;
            subjectsDone++;
          }
        }
    });

    const rataRata = subjectsDone > 0 ? totalNilai / subjectsDone : 0;
    const totalNilaiColIndex = headersNilai.indexOf('Total Nilai');
    const rataRataColIndex = headersNilai.indexOf('Rata-rata Nilai');
    if (totalNilaiColIndex > -1) dataNilai[nilaiRowIndex][totalNilaiColIndex] = totalNilai.toFixed(2);
    if (rataRataColIndex > -1) dataNilai[nilaiRowIndex][rataRataColIndex] = rataRata.toFixed(2);

    const statusColIndex = headersNilai.indexOf('Status Ujian');
    if (statusColIndex > -1) {
      dataNilai[nilaiRowIndex][statusColIndex] = status || 'Selesai';
    }
    
    const durasiColIndex = headersNilai.indexOf('Durasi Pengerjaan (menit)');
    if (durasiColIndex > -1 && totalWaktuUjian != null && sisaDetik != null) {
      const durasiDetik = totalWaktuUjian - sisaDetik;
      const durasiMenit = Math.round(durasiDetik / 60);
      dataNilai[nilaiRowIndex][durasiColIndex] = durasiMenit;
    }

    // 5. Write back to sheet
    if (isNewNilaiRow) {
      // If it's a brand new participant, append their row.
      nilaiSheet.appendRow(dataNilai[nilaiRowIndex]);
    } else {
      // If it's an existing participant, just update their specific row to avoid rewriting the whole sheet.
      nilaiSheet.getRange(nilaiRowIndex + 1, 1, 1, headersNilai.length).setValues([dataNilai[nilaiRowIndex]]);
    }

    return { success: true };

  } catch (e) {
    Logger.log('Error in submitAllAnswersAndFinalize: ' + e.message + ' Stack: ' + e.stack);
    return { success: false, message: 'Gagal memfinalisasi semua jawaban: ' + e.message };
  }
}


function saveEssayScores(sessionId, essayScores) {
  try {
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    // Target sheet 'peserta' untuk menyimpan nilai
    const nilaiSheet = ss.getSheetByName('peserta');
    if (!nilaiSheet) {
      return { success: false, message: 'Sheet "peserta" tidak ditemukan.' };
    }

    const dataRange = nilaiSheet.getDataRange();
    const dataNilai = dataRange.getValues();
    const headers = dataNilai[0];
    const noPesertaColIndex = 1; // Kolom 'Nomor Peserta'
    const nilaiEsaiColIndex = headers.indexOf('Nilai Esai');

    if (nilaiEsaiColIndex === -1) {
      return { success: false, message: 'Kolom "Nilai Esai" tidak ditemukan di sheet peserta.' };
    }

    const scoreMap = new Map(essayScores.map(s => [s.noPeserta, s.skor]));
    let changesMade = false;

    // Iterasi melalui data untuk memperbarui skor
    for (let i = 1; i < dataNilai.length; i++) {
      const noPeserta = dataNilai[i][noPesertaColIndex];

      if (scoreMap.has(noPeserta)) {
        const newScore = scoreMap.get(noPeserta);
        dataNilai[i][nilaiEsaiColIndex] = newScore; // Perbarui array data
        changesMade = true;

        // Hitung ulang Total dan Rata-rata untuk baris ini
        let totalNilai = 0;
        let subjectsDone = 0;
        
        headers.forEach((header, colIndex) => {
          if (header.startsWith('Nilai ')) {
            const score = parseFloat(dataNilai[i][colIndex]);
            if (!isNaN(score)) {
              totalNilai += score;
              subjectsDone++;
            }
          }
        });
        
        const rataRata = subjectsDone > 0 ? totalNilai / subjectsDone : 0;

        const totalNilaiColIndex = headers.indexOf('Total Nilai');
        const rataRataColIndex = headers.indexOf('Rata-rata Nilai');

        if (totalNilaiColIndex > -1) dataNilai[i][totalNilaiColIndex] = totalNilai.toFixed(2);
        if (rataRataColIndex > -1) dataNilai[i][rataRataColIndex] = rataRata.toFixed(2);
      }
    }
    
    if (changesMade) {
      dataRange.setValues(dataNilai); // Tulis semua data yang diperbarui kembali ke sheet
    }

    return { success: true, message: 'Nilai esai berhasil disimpan.' };

  } catch (e) {
    Logger.log('Error in saveEssayScores: ' + e.message);
    return { success: false, message: 'Terjadi kesalahan: ' + e.message };
  }
}

function getSemuaHasilUjian(noPeserta) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSubjectSheets = getSubjectSheets();
    if (allSubjectSheets.error) return allSubjectSheets;

    // Membaca dari sheet 'peserta' (sheet nilai)
    const nilaiSheet = ss.getSheetByName('peserta');
    let nilaiRow = null;
    let nilaiHeaders = [];

    if (nilaiSheet) {
      const nilaiData = nilaiSheet.getDataRange().getValues();
      nilaiHeaders = nilaiData[0];
      for (let i = 1; i < nilaiData.length; i++) {
        if (nilaiData[i][1] == noPeserta) { // Kolom B adalah 'Nomor Peserta'
          nilaiRow = nilaiData[i];
          break;
        }
      }
    }

    // Membuat objek hasil dari data baris nilai
    const finalResults = {
        subjects: [],
        total: '-',
        average: '-'
    };

    if (nilaiRow) {
        allSubjectSheets.forEach(sheetName => {
            const subjectDisplayName = sheetName.replace('soal_', '').replace(/_/g, ' ');
            const nilaiColName = `Nilai ${subjectDisplayName}`;
            const nilaiColIndex = nilaiHeaders.indexOf(nilaiColName);
            let nilai = '-';

            if (nilaiColIndex > -1 && nilaiRow[nilaiColIndex] !== '') {
                nilai = nilaiRow[nilaiColIndex];
            }
            
            finalResults.subjects.push({
                subject: subjectDisplayName,
                nilai: nilai,
                benar: '-', // Detail ini tidak lagi disimpan di sheet nilai
                salah: '-',
                totalSoal: '-'
            });
        });
        
        // Ambil juga nilai esai, total, dan rata-rata
        const nilaiEsaiCol = nilaiHeaders.indexOf('Nilai Esai');
        if(nilaiEsaiCol > -1 && nilaiRow[nilaiEsaiCol] !== '') {
            finalResults.subjects.push({
                subject: 'Esai',
                nilai: nilaiRow[nilaiEsaiCol],
                benar: '-',
                salah: '-',
                totalSoal: '-'
            });
        }

        const totalCol = nilaiHeaders.indexOf('Total Nilai');
        const avgCol = nilaiHeaders.indexOf('Rata-rata Nilai');
        if(totalCol > -1) finalResults.total = nilaiRow[totalCol];
        if(avgCol > -1) finalResults.average = nilaiRow[avgCol];
    }


    return { success: true, results: finalResults };
  } catch (e) {
    Logger.log('Error in getSemuaHasilUjian: ' + e.message);
    return { error: true, message: e.message };
  }
}

function updatePesertaHeaders() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const pesertaSheet = ss.getSheetByName('peserta');
  if (!pesertaSheet) {
    SpreadsheetApp.getUi().alert('Sheet "peserta" tidak ditemukan. Jalankan "Setup & Inisialisasi Sheet" dahulu.');
    return;
  }

  const idealHeaders = ['Timestamp', 'Nomor Peserta', 'Nama Peserta', 'Status Ujian'];
  const allSubjectSheets = getSubjectSheets();
  allSubjectSheets.forEach(subject => {
    const subjectName = subject.replace('soal_', '').replace(/_/g, ' ');
    idealHeaders.push(`Nilai ${subjectName}`);
  });
  idealHeaders.push('Nilai Esai', 'Total Nilai', 'Rata-rata Nilai');

  const oldData = pesertaSheet.getDataRange().getValues();
  if (oldData.length === 0) {
    pesertaSheet.appendRow(idealHeaders);
    SpreadsheetApp.getUi().alert('Sheet "peserta" kosong. Header baru telah ditambahkan.');
    return;
  }
  const oldHeaders = oldData.shift();

  if (JSON.stringify(oldHeaders) === JSON.stringify(idealHeaders)) {
    SpreadsheetApp.getUi().alert('Header sheet "peserta" sudah sesuai, tidak ada perubahan.');
    return;
  }

  const oldHeaderMap = {};
  oldHeaders.forEach((header, index) => {
    oldHeaderMap[header] = index;
  });

  const newData = [idealHeaders];

  oldData.forEach(oldRow => {
    const newRow = Array(idealHeaders.length).fill('');
    idealHeaders.forEach((newHeader, newIndex) => {
      if (oldHeaderMap.hasOwnProperty(newHeader)) {
        const oldIndex = oldHeaderMap[newHeader];
        newRow[newIndex] = oldRow[oldIndex];
      }
    });
    newData.push(newRow);
  });

  pesertaSheet.clear();
  pesertaSheet.getRange(1, 1, newData.length, idealHeaders.length).setValues(newData);

  SpreadsheetApp.getUi().alert('Sheet "peserta" telah berhasil diperbarui dengan header dan struktur data yang baru.');
}

// ==================== FUNGSI UNTUK ADMIN DASHBOARD ====================

function getDashboardStats(sessionId) {
  try {
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak. Sesi admin tidak valid.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. Get totalPeserta
    const dataPesertaSheet = ss.getSheetByName('data_peserta');
    let totalPeserta = 0;
    if (dataPesertaSheet) {
      totalPeserta = dataPesertaSheet.getLastRow() - 1; // Subtract header row
      if (totalPeserta < 0) totalPeserta = 0;
    }

    // 2. Get totalSoal
    const allSubjectSheets = getSubjectSheets();
    let totalSoal = 0;
    allSubjectSheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        totalSoal += (sheet.getLastRow() - 1); // Subtract header row
      }
    });
    if (totalSoal < 0) totalSoal = 0;

    // 3. Get ujianSelesai (based on all subjects completed)
    const pesertaSheet = ss.getSheetByName('peserta');
    let ujianSelesai = 0;
    if (pesertaSheet) {
      const data = pesertaSheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0];
        const allSubjectSheets = getSubjectSheets(); // Get all subject sheets dynamically

        for (let i = 1; i < data.length; i++) {
          let allSubjectsCompleted = true;
          // Check if all subject score columns are filled for this participant
          allSubjectSheets.forEach(sheetName => {
            const subjectDisplayName = `Nilai ${sheetName.replace('soal_', '').replace(/_/g, ' ')}`;
            const nilaiColIndex = headers.indexOf(subjectDisplayName);
            if (nilaiColIndex === -1 || data[i][nilaiColIndex] === '') {
              allSubjectsCompleted = false;
            }
          });

          // Also check for 'Nilai Esai' if it exists
          const nilaiEsaiColIndex = headers.indexOf('Nilai Esai');
          if (nilaiEsaiColIndex !== -1 && data[i][nilaiEsaiColIndex] === '') {
            allSubjectsCompleted = false;
          }

          if (allSubjectsCompleted) {
            ujianSelesai++;
          }
        }
      }
    }
    
    // 4. Get totalMapel
    const totalMapel = allSubjectSheets.length;

    return { success: true, data: { totalPeserta, totalSoal, ujianSelesai, totalMapel } };

  } catch (error) {
    Logger.log('Error in getDashboardStats: ' + error.message);
    return { success: false, message: 'Terjadi kesalahan: ' + error.message };
  }
}

function getPesertaList(sessionId) {
  try {
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak. Sesi admin tidak valid.' };
    }

    // Membaca dari sheet 'data_peserta'
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('data_peserta');
    if (!sheet) {
      return { success: false, message: 'Sheet "data_peserta" tidak ditemukan.' };
    }

    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { success: true, data: [] };
    }

    const headers = data.shift(); 
    const statusColIndex = headers.indexOf('Status'); // Get Status column index

    const peserta = data.map(row => {
      return {
        noPeserta: row[0],
        nama: row[1],
        password: row[2],
        status: statusColIndex > -1 ? row[statusColIndex] : 'aktif', // Include status, default to 'aktif' if column not found
      };
    });

    return { success: true, data: peserta };

  } catch (error) {
    Logger.log('Error in getPesertaList: ' + error.message);
    return { success: false, message: 'Terjadi kesalahan: ' + error.message };
  }
}

function getSoalList(sessionId, sheetName) {
  try {
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, message: `Sheet "${sheetName}" tidak ditemukan.` };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, data: [] };
    }

    const headers = data.shift();
    const soal = data.map(row => {
      // Assuming the structure is: No, Tipe, Pertanyaan, OpsiA-E, Jawaban
      return {
        no: row[0],
        tipe: row[1],
        pertanyaan: row[2],
        opsiA: row[3],
        opsiB: row[4],
        opsiC: row[5],
        opsiD: row[6],
        opsiE: row[7],
        jawaban: row[8] 
      };
    });

    return { success: true, data: soal };

  } catch (error) {
    Logger.log(`Error in getSoalList for sheet ${sheetName}: ` + error.message);
    return { success: false, message: 'Terjadi kesalahan: ' + error.message };
  }
}

function getAllSubjectSheetsForAdmin(sessionId) {
  try {
    // Admin session validation
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSheets = ss.getSheets();
    const subjectSheets = allSheets
      .map(sheet => sheet.getName())
      .filter(name => name.toLowerCase().startsWith('soal_')) // A more robust filter
      .map(sheetName => {
        // Create a user-friendly display name
        const displayName = sheetName.replace('soal_', '').replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
        return { value: sheetName, text: displayName };
      });
      
    return { success: true, data: subjectSheets };
  } catch (e) {
    Logger.log('Gagal mengambil semua mata pelajaran untuk admin: ' + e.message);
    return { success: false, message: 'Gagal mengambil daftar mata pelajaran: ' + e.message };
  }
}

function deleteSoal(sessionId, sheetName, noSoal) {
  try {
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, message: `Sheet "${sheetName}" tidak ditemukan.` };
    }

    const data = sheet.getDataRange().getValues();
    let rowIndexToDelete = -1;

    // Find the row index for the question number
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == noSoal) { // Column 0 is 'No'
        rowIndexToDelete = i + 1;
        break;
      }
    }

    if (rowIndexToDelete !== -1) {
      sheet.deleteRow(rowIndexToDelete);
      return { success: true, message: `Soal no. ${noSoal} berhasil dihapus.` };
    } else {
      return { success: false, message: `Soal no. ${noSoal} tidak ditemukan.` };
    }

  } catch (error) {
    Logger.log(`Error in deleteSoal: ${error.message}`);
    return { success: false, message: 'Gagal menghapus soal: ' + error.message };
  }
}

function getSoalDetail(sessionId, sheetName, noSoal) {
  try {
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, message: `Sheet "${sheetName}" tidak ditemukan.` };
    }

    const data = sheet.getDataRange().getValues();
    let soalData = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == noSoal) {
        soalData = {
          no: data[i][0],
          tipe: data[i][1],
          pertanyaan: data[i][2],
          opsiA: data[i][3],
          opsiB: data[i][4],
          opsiC: data[i][5],
          opsiD: data[i][6],
          opsiE: data[i][7],
          jawaban: data[i][8]
        };
        break;
      }
    }

    if (soalData) {
      return { success: true, data: soalData };
    } else {
      return { success: false, message: `Soal no. ${noSoal} tidak ditemukan.` };
    }
  } catch (error) {
    return { success: false, message: 'Gagal mengambil detail soal: ' + error.message };
  }
}

function updateSoal(sessionId, sheetName, soalData) {
  try {
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, message: `Sheet "${sheetName}" tidak ditemukan.` };
    }

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == soalData.originalNo) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: `Soal dengan nomor asli ${soalData.originalNo} tidak ditemukan.` };
    }

    // Create a new row array in the correct order
    const newRowData = [
      soalData.no,
      soalData.tipe,
      soalData.pertanyaan,
      soalData.opsiA,
      soalData.opsiB,
      soalData.opsiC,
      soalData.opsiD,
      soalData.opsiE,
      soalData.jawaban
    ];

    sheet.getRange(rowIndex, 1, 1, newRowData.length).setValues([newRowData]);

    return { success: true, message: 'Soal berhasil diperbarui.' };

  } catch (error) {
    return { success: false, message: 'Gagal memperbarui soal: ' + error.message };
  }
}

function addPeserta(sessionId, pesertaData) {
  try {
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak. Sesi admin tidak valid.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('data_peserta');
    if (!sheet) {
      return { success: false, message: 'Sheet "data_peserta" tidak ditemukan.' };
    }

    // Check for duplicates
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == pesertaData.noPeserta) {
        return { success: false, message: `Peserta dengan nomor ${pesertaData.noPeserta} sudah ada.` };
      }
    }

    // Add new participant with default 'aktif' status
    sheet.appendRow([pesertaData.noPeserta, pesertaData.nama, pesertaData.password, 'aktif']);

    return { success: true, message: 'Peserta berhasil ditambahkan.' };

  } catch (error) {
    Logger.log(`Error in addPeserta: ${error.message}`);
    return { success: false, message: 'Gagal menambahkan peserta: ' + error.message };
  }
}

function addSoal(sessionId, sheetName, soalData) {
  try {
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, message: `Sheet "${sheetName}" tidak ditemukan.` };
    }

    // Create a new row array in the correct order
    const newRowData = [
      soalData.no,
      soalData.tipe,
      soalData.pertanyaan,
      soalData.opsiA,
      soalData.opsiB,
      soalData.opsiC,
      soalData.opsiD,
      soalData.opsiE,
      soalData.jawaban
    ];

    sheet.appendRow(newRowData);

    return { success: true, message: 'Soal berhasil ditambahkan.' };

  } catch (error) {
    return { success: false, message: 'Gagal menambahkan soal: ' + error.message };
  }
}

function updatePeserta(sessionId, pesertaData) {
  try {
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak. Sesi admin tidak valid.' };
    }

    // Menulis ke sheet 'data_peserta'
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('data_peserta');
    if (!sheet) {
      return { success: false, message: 'Sheet "data_peserta" tidak ditemukan.' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0]; // Get headers
    const statusColIndex = headers.indexOf('Status'); // Find 'Status' column index
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == pesertaData.originalNoPeserta) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'Peserta dengan nomor ' + pesertaData.originalNoPeserta + ' tidak ditemukan.' };
    }

    sheet.getRange(rowIndex, 1).setValue(pesertaData.noPeserta);
    sheet.getRange(rowIndex, 2).setValue(pesertaData.nama);
    sheet.getRange(rowIndex, 3).setValue(pesertaData.password);
    
    if (statusColIndex > -1) {
      sheet.getRange(rowIndex, statusColIndex + 1).setValue(pesertaData.status);
    }

    return { success: true, message: 'Data peserta berhasil diperbarui.' };

  } catch (error) {
    return { success: false, message: 'Terjadi kesalahan: ' + error.message };
  }
}

function deletePeserta(sessionId, noPeserta) {
  try {
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak. Sesi admin tidak valid.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('data_peserta');
    if (!sheet) {
      return { success: false, message: 'Sheet "data_peserta" tidak ditemukan.' };
    }

    const data = sheet.getDataRange().getValues();
    let rowIndexToDelete = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == noPeserta) { // Column 0 is 'No Peserta'
        rowIndexToDelete = i + 1;
        break;
      }
    }

    if (rowIndexToDelete !== -1) {
      sheet.deleteRow(rowIndexToDelete);
      return { success: true, message: `Peserta dengan nomor ${noPeserta} berhasil dihapus.` };
    } else {
      return { success: false, message: `Peserta dengan nomor ${noPeserta} tidak ditemukan.` };
    }

  } catch (error) {
    Logger.log(`Error in deletePeserta: ${error.message}`);
    return { success: false, message: 'Gagal menghapus peserta: ' + error.message };
  }
}

function getHasilUjian(sessionId) {
  try {
    // 1. Validate Session
    const session = getUserSession(sessionId);
    if (!session.success || !session.data) {
      return { success: false, message: 'Akses ditolak. Sesi tidak valid.' };
    }

    // 2. Get Data from Spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('peserta');
    if (!sheet) {
      return { success: true, headers: [], data: [] }; // Return empty if sheet doesn't exist
    }

    const allData = sheet.getDataRange().getValues();
    
    if (allData.length <= 1) {
      return { success: true, headers: allData[0] || [], data: [] }; // No data besides header
    }

    const headers = allData.shift(); // Get and remove header row
    
    // Format the date in each row before sending
    const timezone = ss.getSpreadsheetTimeZone();
    allData.forEach(row => {
        if (row[0] instanceof Date) {
            row[0] = Utilities.formatDate(row[0], timezone, 'dd/MM/yyyy HH:mm:ss');
        }
    });

    return { success: true, headers: headers, data: allData };

  } catch (error) {
    Logger.log('Error in getHasilUjian: ' + error.message);
    return { success: false, message: 'Terjadi kesalahan: ' + error.message };
  }
}

function getRekapData(sessionId) {
  try {
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const rekapSheets = ['rekap_pilihanganda', 'rekap_esai'];
    const pesertaUnik = {};

    rekapSheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        if (data.length > 1) {
          const headers = data[0];
          const noPesertaCol = headers.indexOf('Nomor Peserta');
          const namaPesertaCol = headers.indexOf('Nama Peserta');

          for (let i = 1; i < data.length; i++) {
            const noPeserta = data[i][noPesertaCol];
            if (noPeserta && !pesertaUnik[noPeserta]) {
              pesertaUnik[noPeserta] = {
                noPeserta: noPeserta,
                nama: data[i][namaPesertaCol]
              };
            }
          }
        }
      }
    });

    const finalData = Object.values(pesertaUnik).map(p => {
      const pdfUrlAction = `google.script.run.withSuccessHandler(handlePdfUrl).createPdf('${p.noPeserta}')`;
      return [p.noPeserta, p.nama, pdfUrlAction];
    });

    const finalHeaders = ['Nomor Peserta', 'Nama Peserta', 'PDF'];

    return { success: true, headers: finalHeaders, data: finalData };

  } catch (error) {
    Logger.log('Error in getRekapData (aggregated): ' + error.message);
    return { success: false, message: 'Terjadi kesalahan saat mengambil data rekap agregat: ' + error.message };
  }
}

function createPdf(noPeserta) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const rekapSheets = {
      'Pilihan Ganda': ss.getSheetByName('rekap_pilihanganda'),
      'Esai': ss.getSheetByName('rekap_esai')
    };

    let pesertaInfo = {};
    let allAnswers = {};
    let soalDetails = {}; // Untuk menyimpan detail soal

    // Ambil detail soal dari sheet soal untuk mendapatkan opsi jawaban
    const allSubjectSheets = getSubjectSheets().concat(['soal_esai']);
    allSubjectSheets.forEach(sheetName => {
      const sheetSoal = ss.getSheetByName(sheetName);
      if (sheetSoal) {
        const dataSoal = sheetSoal.getDataRange().getValues();
        soalDetails[sheetName] = {};
        for (let i = 1; i < dataSoal.length; i++) {
          if (dataSoal[i].join('').trim() === '') continue;
          const noSoal = parseInt(dataSoal[i][0], 10);
          if (isNaN(noSoal)) continue;
          soalDetails[sheetName][noSoal] = {
            tipe: dataSoal[i][1],
            opsiA: dataSoal[i][3],
            opsiB: dataSoal[i][4],
            opsiC: dataSoal[i][5],
            opsiD: dataSoal[i][6],
            opsiE: dataSoal[i][7]
          };
        }
      }
    });

    // Ambil data dari rekap sheets
    for (const type in rekapSheets) {
      const sheet = rekapSheets[type];
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const noPesertaCol = headers.indexOf('Nomor Peserta');
        const namaPesertaCol = headers.indexOf('Nama Peserta');
        const mapelCol = headers.indexOf('Mata Pelajaran');
        const timestampCol = headers.indexOf('Timestamp');

        for (let i = 1; i < data.length; i++) {
          if (data[i][noPesertaCol] == noPeserta) {
            const row = data[i];
            const mapel = row[mapelCol];

            if (!pesertaInfo.nama) {
              pesertaInfo.nama = row[namaPesertaCol];
              pesertaInfo.noPeserta = row[noPesertaCol];
            }
            if (!pesertaInfo.timestamp) {
               pesertaInfo.timestamp = row[timestampCol] instanceof Date ? 
                 Utilities.formatDate(row[timestampCol], ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm:ss') : 
                 row[timestampCol];
            }

            if (!allAnswers[mapel]) {
              allAnswers[mapel] = {};
            }
            if (!allAnswers[mapel][type]) {
              allAnswers[mapel][type] = [];
            }

            // Tentukan nama sheet soal
            const soalSheetName = 'soal_' + mapel.toLowerCase().replace(/ /g, '_');

            for (let j = 4; j < headers.length; j++) {
              if (row[j]) {
                const soalNumMatch = headers[j].match(/\d+/);
                const soalNum = soalNumMatch ? parseInt(soalNumMatch[0], 10) : null;
                
                let jawabanDisplay = row[j];
                
                // Untuk Pilihan Ganda, format jawaban dengan opsi lengkap
                if (type === 'Pilihan Ganda' && soalNum && soalDetails[soalSheetName] && soalDetails[soalSheetName][soalNum]) {
                  const soalInfo = soalDetails[soalSheetName][soalNum];
                  const jawaban = row[j].toString().trim();
                  
                  // Ambil teks opsi sesuai jawaban
                  const opsiMap = {
                    'A': soalInfo.opsiA,
                    'B': soalInfo.opsiB,
                    'C': soalInfo.opsiC,
                    'D': soalInfo.opsiD,
                    'E': soalInfo.opsiE
                  };
                  
                  if (opsiMap[jawaban]) {
                    jawabanDisplay = `${jawaban}. ${opsiMap[jawaban]}`;
                  }
                }
                
                allAnswers[mapel][type].push({
                  soal: headers[j],
                  jawaban: jawabanDisplay
                });
              }
            }
          }
        }
      }
    }

    if (Object.keys(pesertaInfo).length === 0) {
      return { success: false, message: 'Data untuk peserta tidak ditemukan.' };
    }

    // Ambil nilai dari sheet peserta
    const nilaiSheet = ss.getSheetByName('peserta');
    let nilaiData = {};
    if (nilaiSheet) {
      const data = nilaiSheet.getDataRange().getValues();
      const headers = data[0];
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] == noPeserta) {
          const totalNilaiCol = headers.indexOf('Total Nilai');
          const rataRataCol = headers.indexOf('Rata-rata Nilai');
          const durasiCol = headers.indexOf('Durasi Pengerjaan (menit)');
          
          if (totalNilaiCol > -1) nilaiData.total = data[i][totalNilaiCol] || '-';
          if (rataRataCol > -1) nilaiData.rataRata = data[i][rataRataCol] || '-';
          if (durasiCol > -1) nilaiData.durasi = data[i][durasiCol] || '-';
          
          // Ambil nilai per mata pelajaran
          nilaiData.mapel = {};
          headers.forEach((header, idx) => {
            if (header.startsWith('Nilai ')) {
              const mapelName = header.replace('Nilai ', '');
              nilaiData.mapel[mapelName] = data[i][idx] || '-';
            }
          });
          break;
        }
      }
    }

    // Buat HTML dengan desain profesional
    let html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Rekap Jawaban ${pesertaInfo.nama}</title>
  <style>
    @page {
      margin: 20mm;
    }
    
    body {
      font-family: 'Arial', 'Helvetica', sans-serif;
      line-height: 1.6;
      color: #333;
      margin: 0;
      padding: 0;
    }
    
    .header {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 30px;
      border-radius: 10px;
      margin-bottom: 30px;
      text-align: center;
    }
    
    .header h1 {
      margin: 0 0 10px 0;
      font-size: 28px;
      font-weight: bold;
    }
    
    .header p {
      margin: 5px 0;
      font-size: 14px;
      opacity: 0.95;
    }
    
    .info-box {
      background: #f8f9fa;
      border-left: 4px solid #667eea;
      padding: 20px;
      margin-bottom: 30px;
      border-radius: 5px;
    }
    
    .info-row {
      display: flex;
      margin-bottom: 10px;
      align-items: center;
    }
    
    .info-label {
      font-weight: bold;
      color: #667eea;
      min-width: 150px;
      font-size: 14px;
    }
    
    .info-value {
      color: #333;
      font-size: 14px;
    }
    
    .score-summary {
      background: white;
      border: 2px solid #e9ecef;
      border-radius: 10px;
      padding: 20px;
      margin-bottom: 30px;
    }
    
    .score-summary h3 {
      color: #667eea;
      margin-top: 0;
      margin-bottom: 15px;
      font-size: 18px;
      border-bottom: 2px solid #667eea;
      padding-bottom: 10px;
    }
    
    .score-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 15px;
    }
    
    .score-item {
      background: #f8f9fa;
      padding: 15px;
      border-radius: 8px;
      border-left: 3px solid #667eea;
    }
    
    .score-item-label {
      font-size: 12px;
      color: #6c757d;
      margin-bottom: 5px;
    }
    
    .score-item-value {
      font-size: 24px;
      font-weight: bold;
      color: #667eea;
    }
    
    .mapel-section {
      page-break-inside: avoid;
      margin-bottom: 40px;
      border: 1px solid #e9ecef;
      border-radius: 10px;
      overflow: hidden;
    }
    
    .mapel-header {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 15px 20px;
      font-size: 18px;
      font-weight: bold;
    }
    
    .mapel-score {
      background: #f8f9fa;
      padding: 15px 20px;
      border-bottom: 1px solid #e9ecef;
      font-weight: bold;
      color: #667eea;
    }
    
    .type-section {
      padding: 20px;
    }
    
    .type-header {
      background: #f8f9fa;
      padding: 12px 15px;
      margin-bottom: 15px;
      border-radius: 5px;
      border-left: 4px solid #764ba2;
      font-weight: bold;
      color: #764ba2;
      font-size: 16px;
    }
    
    .answer-table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 20px;
      font-size: 13px;
    }
    
    .answer-table th {
      background: #667eea;
      color: white;
      padding: 10px;
      text-align: left;
      font-weight: bold;
      border: 1px solid #5568d3;
    }
    
    .answer-table td {
      padding: 8px 10px;
      border: 1px solid #dee2e6;
      vertical-align: top;
    }
    
    .answer-table tr:nth-child(even) {
      background-color: #f8f9fa;
    }
    
    .answer-table tr:hover {
      background-color: #e9ecef;
    }
    
    .soal-col {
      width: 80px;
      text-align: center;
      font-weight: bold;
      color: #667eea;
    }
    
    .jawaban-col {
      word-wrap: break-word;
      word-break: break-word;
    }
    
    .pg-answer {
      display: inline-block;
      background: #28a745;
      color: white;
      padding: 5px 15px;
      border-radius: 20px;
      font-weight: bold;
      font-size: 14px;
    }
    
    .pg-answer-text {
      color: #333;
      font-weight: normal;
    }
    
    .footer {
      margin-top: 40px;
      padding-top: 20px;
      border-top: 2px solid #e9ecef;
      text-align: center;
      color: #6c757d;
      font-size: 12px;
    }
    
    @media print {
      .mapel-section {
        page-break-inside: avoid;
      }
      
      .header {
        background: #667eea !important;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>📋 REKAP JAWABAN UJIAN</h1>
    <p>Sistem Ujian Online</p>
  </div>
  
  <div class="info-box">
    <div class="info-row">
      <span class="info-label">👤 Nama Peserta:</span>
      <span class="info-value">${pesertaInfo.nama}</span>
    </div>
    <div class="info-row">
      <span class="info-label">🔢 Nomor Peserta:</span>
      <span class="info-value">${pesertaInfo.noPeserta}</span>
    </div>
    <div class="info-row">
      <span class="info-label">📅 Waktu Submit:</span>
      <span class="info-value">${pesertaInfo.timestamp}</span>
    </div>
    ${nilaiData.durasi ? `
    <div class="info-row">
      <span class="info-label">⏱️ Durasi Pengerjaan:</span>
      <span class="info-value">${nilaiData.durasi} menit</span>
    </div>
    ` : ''}
  </div>`;

    // Tampilkan ringkasan nilai jika ada
    if (Object.keys(nilaiData).length > 0) {
      html += `
  <div class="score-summary">
    <h3>📊 Ringkasan Nilai</h3>
    <div class="score-grid">`;
      
      if (nilaiData.total && nilaiData.total !== '-') {
        html += `
      <div class="score-item">
        <div class="score-item-label">Total Nilai</div>
        <div class="score-item-value">${nilaiData.total}</div>
      </div>`;
      }
      
      if (nilaiData.rataRata && nilaiData.rataRata !== '-') {
        html += `
      <div class="score-item">
        <div class="score-item-label">Rata-rata Nilai</div>
        <div class="score-item-value">${nilaiData.rataRata}</div>
      </div>`;
      }
      
      html += `
    </div>
  </div>`;
    }

    // Loop untuk setiap mata pelajaran
    for (const mapel in allAnswers) {
      html += `
  <div class="mapel-section">
    <div class="mapel-header">📚 ${mapel.toUpperCase()}</div>`;
      
      // Tampilkan nilai mata pelajaran jika ada
      if (nilaiData.mapel && nilaiData.mapel[mapel] && nilaiData.mapel[mapel] !== '-') {
        html += `
    <div class="mapel-score">Nilai: ${nilaiData.mapel[mapel]}</div>`;
      }
      
      html += `
    <div class="type-section">`;
      
      // Loop untuk setiap tipe (Pilihan Ganda / Esai)
      for (const type in allAnswers[mapel]) {
        html += `
      <div class="type-header">${type === 'Pilihan Ganda' ? '✓' : '✍️'} ${type}</div>
      <table class="answer-table">
        <thead>
          <tr>
            <th class="soal-col">No. Soal</th>
            <th class="jawaban-col">Jawaban</th>
          </tr>
        </thead>
        <tbody>`;
        
        // Loop untuk setiap jawaban
        allAnswers[mapel][type].forEach(item => {
          html += `
          <tr>
            <td class="soal-col">${item.soal.replace('Soal ', '').replace('Esai ', '')}</td>
            <td class="jawaban-col">`;
          
          if (type === 'Pilihan Ganda') {
            // Cek apakah jawaban sudah dalam format lengkap (ada titik)
            if (item.jawaban.includes('.')) {
              const parts = item.jawaban.split('.');
              html += `<span class="pg-answer">${parts[0]}</span><span class="pg-answer-text">. ${parts.slice(1).join('.')}</span>`;
            } else {
              html += `<span class="pg-answer">${item.jawaban}</span>`;
            }
          } else {
            html += item.jawaban;
          }
          
          html += `</td>
          </tr>`;
        });
        
        html += `
        </tbody>
      </table>`;
      }
      
      html += `
    </div>
  </div>`;
    }

    html += `
  <div class="footer">
    <p>Dokumen ini dibuat secara otomatis oleh Sistem Ujian Online</p>
    <p>© ${new Date().getFullYear()} - Dicetak pada ${Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm:ss')}</p>
  </div>
</body>
</html>`;

    // Buat PDF
    const blob = Utilities.newBlob(html, MimeType.HTML, `Rekap_${pesertaInfo.noPeserta}.html`);
    const pdf = blob.getAs(MimeType.PDF);
    
    const FOLDER_ID = 'https://drive.google.com/file/d/1-3dngGu46KOnJE-4YOe6qyVT0DkkhRi2/view?usp=drive_link';
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const fileName = `Rekap_Lengkap_${pesertaInfo.noPeserta}.pdf`;
    
    // Hapus file lama jika ada
    const oldFiles = folder.getFilesByName(fileName);
    while(oldFiles.hasNext()){
      oldFiles.next().setTrashed(true);
    }

    // Upload file baru
    const file = folder.createFile(pdf).setName(fileName);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, url: file.getUrl() };

  } catch (e) {
    Logger.log('Error in createPdf: ' + e.message + ' Stack: ' + e.stack);
    return { success: false, message: 'Gagal membuat PDF: ' + e.message };
  }
}

function getConfiguration() {
   try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('setting');
    if (!sheet) {
      return { success: false, message: 'Sheet "setting" tidak ditemukan.' };
    }
    const data = sheet.getRange('A2:G2').getValues()[0];
    return {
      success: true,
      data: {
        pinSesi: data[0],
        adminUsername: data[1],
        loginTitle: data[3],
        logoUrl: data[4],
        loginSubtitle: data[5],
        waktuUjian: data[6]
      }
    };
  } catch (e) {
    Logger.log('Error in getConfiguration: ' + e.message);
    return { success: false, message: 'Gagal mengambil konfigurasi: ' + e.message };
  }
}

function saveGeneralSettings(sessionId, settings) {
  try {
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const settingSheet = ss.getSheetByName('setting');
    if (!settingSheet) {
      return { success: false, message: 'Sheet "setting" tidak ditemukan.' };
    }

    if (settings.pinSesi) {
      settingSheet.getRange('A2').setValue(settings.pinSesi);
    }
    if (settings.adminUsername) {
      settingSheet.getRange('B2').setValue(settings.adminUsername);
    }
    if (settings.adminPassword) {
      settingSheet.getRange('C2').setValue(settings.adminPassword);
    }
     if (settings.waktuUjian) {
      settingSheet.getRange('G2').setValue(settings.waktuUjian);
    }

    return { success: true, message: 'Pengaturan umum berhasil disimpan.' };

  } catch (e) {
    Logger.log('Error in saveGeneralSettings: ' + e.message);
    return { success: false, message: 'Gagal menyimpan pengaturan umum: ' + e.message };
  }
}

function saveLoginSettings(sessionId, settings) {
  try {
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const settingSheet = ss.getSheetByName('setting');
    if (!settingSheet) {
      return { success: false, message: 'Sheet "setting" tidak ditemukan.' };
    }

    let logoUrl = settingSheet.getRange('E2').getValue();

    if (settings.logoData) {
      const FOLDER_ID = 'wajib di ubah';
      const folder = DriveApp.getFolderById(FOLDER_ID);

      const [mimeType, data] = settings.logoData.split(',');
      const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType.replace('data:','').replace(';base64',''), 'logo_ujian_online');
      
      const oldFiles = folder.getFilesByName('logo_ujian_online');
      while(oldFiles.hasNext()){
        oldFiles.next().setTrashed(true);
      }

      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const fileId = file.getId();
      logoUrl = 'https://lh3.googleusercontent.com/d/' + fileId; // New URL format
    }

    settingSheet.getRange('D2').setValue(settings.loginTitle);
    settingSheet.getRange('E2').setValue(logoUrl);
    settingSheet.getRange('F2').setValue(settings.loginSubtitle);

    return { success: true, message: 'Pengaturan halaman login berhasil disimpan.' };

  } catch (e) {
    Logger.log('Error in saveLoginSettings: ' + e.message);
    return { success: false, message: 'Gagal menyimpan pengaturan login: ' + e.message };
  }
}

function retakeExam(sessionId, noPeserta) {
  try {
    // 1. Validate Admin Session
    const session = getUserSession(sessionId);
    if (!session.success || !session.data || session.data.role !== 'admin') {
      return { success: false, message: 'Akses ditolak. Sesi admin tidak valid.' };
    }

    if (!noPeserta) {
      return { success: false, message: 'Nomor Peserta tidak boleh kosong.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const pesertaSheet = ss.getSheetByName('peserta');
    const rekapPGSheet = ss.getSheetByName('rekap_pilihanganda');
    const rekapEsaiSheet = ss.getSheetByName('rekap_esai');

    // 2. Clear data in 'peserta' sheet
    if (pesertaSheet) {
      const data = pesertaSheet.getDataRange().getValues();
      const headers = data[0];
      const noPesertaCol = headers.indexOf('Nomor Peserta');
      
      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][noPesertaCol] == noPeserta) {
          rowIndex = i;
          break;
        }
      }

      if (rowIndex !== -1) {
        // Iterate through headers to clear scores, status, duration, and timestamp
        headers.forEach((header, index) => {
          const cell = pesertaSheet.getRange(rowIndex + 1, index + 1);
          if (header.startsWith('Nilai') || header === 'Total Nilai' || header === 'Rata-rata Nilai' || header === 'Durasi Pengerjaan (menit)' || header === 'Timestamp') {
            cell.setValue('');
          }
          if (header === 'Status Ujian') {
            cell.setValue('Ujian Ulang');
          }
        });
      } else {
        return { success: false, message: `Peserta dengan nomor ${noPeserta} tidak ditemukan di sheet peserta.` };
      }
    } else {
       return { success: false, message: 'Sheet "peserta" tidak ditemukan.' };
    }

    // 3. Delete rows from rekap sheets
    const clearRekapSheet = (sheet) => {
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return; // No data to clear
        const noPesertaCol = data[0].indexOf('Nomor Peserta');
        
        if (noPesertaCol > -1) {
          for (let i = data.length - 1; i > 0; i--) {
            if (data[i][noPesertaCol] == noPeserta) {
              sheet.deleteRow(i + 1);
            }
          }
        }
      }
    };

    clearRekapSheet(rekapPGSheet);
    clearRekapSheet(rekapEsaiSheet);

    return { success: true, message: `Ujian ulang untuk peserta ${noPeserta} telah berhasil diaktifkan.` };

  } catch (error) {
    Logger.log('Error in retakeExam: ' + error.message + ' Stack: ' + error.stack);
    return { success: false, message: 'Terjadi kesalahan server: ' + error.message };
  }
}


function getTotalQuestionCount() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSubjectSheets = getSubjectSheets().concat(['soal_esai']);
    let totalSoal = 0;
    allSubjectSheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        // Subtract header row and filter out empty rows
        const data = sheet.getDataRange().getValues();
        let questionCount = 0;
        for (let i = 1; i < data.length; i++) {
          if (data[i].join('').trim() !== '') {
            questionCount++;
          }
        }
        totalSoal += questionCount;
      }
    });
    return { success: true, total: totalSoal };
  } catch (e) {
    Logger.log('Error in getTotalQuestionCount: ' + e.message);
    return { success: false, message: e.message };
  }
}