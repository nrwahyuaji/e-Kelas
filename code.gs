/**
 * Fungsi utama untuk menangani semua request dari website
 */
function doGet(e) {
  const action = e.parameter.action;
  
  try {
    switch(action) {
      case 'login':
        return handleLogin(e.parameter.nisn, e.parameter.password);
      case 'getKehadiran':
        return handleGetKehadiran(e.parameter.nisn, e.parameter.tahun, e.parameter.bulan);
      case 'getCatatan':
        return handleGetCatatan(e.parameter.nisn, e.parameter.tahun, e.parameter.bulan);
      case 'getInfoSiswa':
        return handleGetInfoSiswa(e.parameter.nisn);
      default:
        return createResponse(false, 'Action tidak valid');
    }
  } catch (error) {
    console.error('Error in doGet:', error);
    return createResponse(false, 'Terjadi kesalahan server: ' + error.message);
  }
}

/**
 * Fungsi untuk menangani POST request
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    switch(action) {
      case 'login':
        return handleLogin(data.nisn, data.password);
      case 'getKehadiran':
        return handleGetKehadiran(data.nisn, data.tahun, data.bulan);
      case 'getCatatan':
        return handleGetCatatan(data.nisn, data.tahun, data.bulan);
      case 'getInfoSiswa':
        return handleGetInfoSiswa(data.nisn);
      default:
        return createResponse(false, 'Action tidak valid');
    }
  } catch (error) {
    console.error('Error in doPost:', error);
    return createResponse(false, 'Terjadi kesalahan server: ' + error.message);
  }
}

/**
 * Fungsi untuk membuat response JSON yang konsisten
 */
function createResponse(success, message, data = null) {
  const response = {
    success: success,
    message: message,
    data: data,
    timestamp: new Date().toISOString()
  };
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Fungsi untuk mendapatkan spreadsheet aktif
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (error) {
    console.error('Error getting spreadsheet:', error);
    throw new Error('Tidak dapat mengakses spreadsheet');
  }
}

/**
 * Fungsi untuk mendapatkan data dari sheet dengan caching
 */
function getSheetData(sheetName) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" tidak ditemukan`);
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) {
      return [];
    }
    
    const headers = data[0];
    const rows = data.slice(1);
    
    return rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] || '';
      });
      return obj;
    });
  } catch (error) {
    console.error(`Error getting data from sheet ${sheetName}:`, error);
    throw new Error(`Gagal mengambil data dari sheet ${sheetName}: ${error.message}`);
  }
}

/**
 * Fungsi untuk menangani login
 */
function handleLogin(nisn, password) {
  try {
    if (!nisn || !password) {
      return createResponse(false, 'NISN dan password harus diisi');
    }
    
    const dataSheet = getSheetData('Data');
    const user = dataSheet.find(row => 
      row.nisn && row.nisn.toString() === nisn.toString() && 
      row.paswd && row.paswd.toString() === password.toString()
    );
    
    if (user) {
      // Hapus password dari response untuk keamanan
      const safeUser = { ...user };
      delete safeUser.paswd;
      
      return createResponse(true, 'Login berhasil', safeUser);
    } else {
      return createResponse(false, 'NISN atau password salah');
    }
  } catch (error) {
    console.error('Error in handleLogin:', error);
    return createResponse(false, 'Gagal melakukan login: ' + error.message);
  }
}

/**
 * Fungsi untuk mendapatkan data kehadiran
 */
function handleGetKehadiran(nisn, tahun = null, bulan = null) {
  try {
    if (!nisn) {
      return createResponse(false, 'NISN harus diisi');
    }
    
    const kehadiranSheet = getSheetData('Kehadiran');
    let filteredData = kehadiranSheet.filter(row => 
      row.nisn && row.nisn.toString() === nisn.toString()
    );
    
    // Filter berdasarkan tahun jika ada
    if (tahun) {
      filteredData = filteredData.filter(row => 
        row.tahun && row.tahun.toString() === tahun.toString()
      );
    }
    
    // Filter berdasarkan bulan jika ada
    if (bulan) {
      filteredData = filteredData.filter(row => 
        row.bulan && row.bulan.toString().toLowerCase() === bulan.toString().toLowerCase()
      );
    }
    
    // Hitung statistik kehadiran
    const stats = calculateAttendanceStats(filteredData);
    
    return createResponse(true, 'Data kehadiran berhasil diambil', {
      kehadiran: filteredData,
      statistik: stats
    });
  } catch (error) {
    console.error('Error in handleGetKehadiran:', error);
    return createResponse(false, 'Gagal mengambil data kehadiran: ' + error.message);
  }
}

/**
 * Fungsi untuk menghitung statistik kehadiran
 */
function calculateAttendanceStats(kehadiranData) {
  let totalHadir = 0;
  let totalIzin = 0;
  let totalSakit = 0;
  let totalAlpha = 0;
  
  kehadiranData.forEach(row => {
    totalHadir += parseInt(row['j.khdrn'] || 0);
    totalIzin += parseInt(row.izin || 0);
    totalSakit += parseInt(row.sakit || 0);
    totalAlpha += parseInt(row.alpha || 0);
  });
  
  const totalHari = totalHadir + totalIzin + totalSakit + totalAlpha;
  const persentaseKehadiran = totalHari > 0 ? Math.round((totalHadir / totalHari) * 100) : 0;
  
  return {
    totalHadir,
    totalIzin,
    totalSakit,
    totalAlpha,
    totalHari,
    persentaseKehadiran
  };
}

/**
 * Fungsi untuk mendapatkan catatan (prestasi, pelanggaran, dll)
 */
function handleGetCatatan(nisn, tahun = null, bulan = null) {
  try {
    if (!nisn) {
      return createResponse(false, 'NISN harus diisi');
    }
    
    const kehadiranSheet = getSheetData('Kehadiran');
    let filteredData = kehadiranSheet.filter(row => 
      row.nisn && row.nisn.toString() === nisn.toString()
    );
    
    // Filter berdasarkan tahun jika ada
    if (tahun) {
      filteredData = filteredData.filter(row => 
        row.tahun && row.tahun.toString() === tahun.toString()
      );
    }
    
    // Filter berdasarkan bulan jika ada
    if (bulan) {
      filteredData = filteredData.filter(row => 
        row.bulan && row.bulan.toString().toLowerCase() === bulan.toString().toLowerCase()
      );
    }
    
    // Kumpulkan semua catatan
    const catatan = {
      prestasi: [],
      pelanggaran: [],
      waliCatatan: [],
      bkCatatan: []
    };
    
    filteredData.forEach(row => {
      if (row['wali.prestasi'] && row['wali.prestasi'].trim()) {
        catatan.prestasi.push({
          tahun: row.tahun,
          bulan: row.bulan,
          catatan: row['wali.prestasi'].trim()
        });
      }
      
      if (row['wali.pelanggaran'] && row['wali.pelanggaran'].trim()) {
        catatan.pelanggaran.push({
          tahun: row.tahun,
          bulan: row.bulan,
          catatan: row['wali.pelanggaran'].trim()
        });
      }
      
      if (row['wali.catatan'] && row['wali.catatan'].trim()) {
        catatan.waliCatatan.push({
          tahun: row.tahun,
          bulan: row.bulan,
          catatan: row['wali.catatan'].trim()
        });
      }
      
      if (row['bk.catatan'] && row['bk.catatan'].trim()) {
        catatan.bkCatatan.push({
          tahun: row.tahun,
          bulan: row.bulan,
          catatan: row['bk.catatan'].trim()
        });
      }
    });
    
    return createResponse(true, 'Data catatan berhasil diambil', catatan);
  } catch (error) {
    console.error('Error in handleGetCatatan:', error);
    return createResponse(false, 'Gagal mengambil data catatan: ' + error.message);
  }
}

/**
 * Fungsi untuk mendapatkan informasi siswa
 */
function handleGetInfoSiswa(nisn) {
  try {
    if (!nisn) {
      return createResponse(false, 'NISN harus diisi');
    }
    
    const dataSheet = getSheetData('Data');
    const siswa = dataSheet.find(row => 
      row.nisn && row.nisn.toString() === nisn.toString()
    );
    
    if (siswa) {
      const info = {
        nama: siswa.nama || '',
        nisn: siswa.nisn || '',
        ortu: siswa.ortu || '',
        kontakOrtu: siswa['kontak.ortu'] || ''
      };
      
      return createResponse(true, 'Informasi siswa berhasil diambil', info);
    } else {
      return createResponse(false, 'Data siswa tidak ditemukan');
    }
  } catch (error) {
    console.error('Error in handleGetInfoSiswa:', error);
    return createResponse(false, 'Gagal mengambil informasi siswa: ' + error.message);
  }
}

/**
 * Fungsi untuk mendapatkan semua data siswa (untuk testing)
 */
function getAllStudents() {
  try {
    const dataSheet = getSheetData('Data');
    return createResponse(true, 'Data semua siswa berhasil diambil', dataSheet);
  } catch (error) {
    console.error('Error in getAllStudents:', error);
    return createResponse(false, 'Gagal mengambil data siswa: ' + error.message);
  }
}

/**
 * Fungsi untuk mendapatkan semua data kehadiran (untuk testing)
 */
function getAllKehadiran() {
  try {
    const kehadiranSheet = getSheetData('Kehadiran');
    return createResponse(true, 'Data semua kehadiran berhasil diambil', kehadiranSheet);
  } catch (error) {
    console.error('Error in getAllKehadiran:', error);
    return createResponse(false, 'Gagal mengambil data kehadiran: ' + error.message);
  }
}

/**
 * Fungsi untuk testing koneksi
 */
function testConnection() {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets().map(sheet => sheet.getName());
    
    return createResponse(true, 'Koneksi berhasil', {
      spreadsheetName: ss.getName(),
      sheets: sheets,
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    console.error('Error in testConnection:', error);
    return createResponse(false, 'Gagal melakukan koneksi: ' + error.message);
  }
}

/**
 * Fungsi untuk membuat data sample (untuk testing)
 */
function createSampleData() {
  try {
    const ss = getSpreadsheet();
    
    // Buat atau update Sheet Data
    let dataSheet = ss.getSheetByName('Data');
    if (!dataSheet) {
      dataSheet = ss.insertSheet('Data');
    }
    
    const dataHeaders = ['nama', 'nisn', 'paswd', 'ortu', 'kontak.ortu'];
    const dataSample = [
      ['Ahmad Rizki Pratama', '1234567890', 'password123', 'Budi Santoso', '081234567891'],
      ['Siti Nurhaliza', '0987654321', 'siti2025', 'Hasan Basri', '081234567892'],
      ['Muhammad Fadli', '1122334455', 'fadli123', 'Ahmad Yani', '081234567893']
    ];
    
    dataSheet.clear();
    dataSheet.getRange(1, 1, 1, dataHeaders.length).setValues([dataHeaders]);
    dataSheet.getRange(2, 1, dataSample.length, dataHeaders[0].length).setValues(dataSample);
    
    // Buat atau update Sheet Kehadiran
    let kehadiranSheet = ss.getSheetByName('Kehadiran');
    if (!kehadiranSheet) {
      kehadiranSheet = ss.insertSheet('Kehadiran');
    }
    
    const kehadiranHeaders = ['tahun', 'bulan', 'nisn', 'j.khdrn', 'izin', 'sakit', 'alpha', 'wali.pelanggaran', 'wali.catatan', 'wali.prestasi', 'bk.catatan'];
    const kehadiranSample = [
      ['2025', 'Juli', '1234567890', '20', '2', '1', '0', '', 'Siswa aktif dan rajin', 'Juara 1 Olimpiade Matematika', ''],
      ['2025', 'Agustus', '1234567890', '22', '1', '0', '0', '', 'Prestasi meningkat', '', ''],
      ['2025', 'Juli', '0987654321', '18', '3', '2', '0', '', 'Perlu peningkatan kedisiplinan', '', 'Konseling motivasi belajar'],
      ['2025', 'Agustus', '0987654321', '21', '2', '0', '0', '', 'Sudah ada peningkatan', 'Juara 2 Lomba Puisi', '']
    ];
    
    kehadiranSheet.clear();
    kehadiranSheet.getRange(1, 1, 1, kehadiranHeaders.length).setValues([kehadiranHeaders]);
    kehadiranSheet.getRange(2, 1, kehadiranSample.length, kehadiranHeaders.length).setValues(kehadiranSample);
    
    return createResponse(true, 'Data sample berhasil dibuat', {
      dataRows: dataSample.length,
      kehadiranRows: kehadiranSample.length
    });
  } catch (error) {
    console.error('Error in createSampleData:', error);
    return createResponse(false, 'Gagal membuat data sample: ' + error.message);
  }
}

/**
 * Fungsi untuk validasi struktur sheet
 */
function validateSheetStructure() {
  try {
    const ss = getSpreadsheet();
    const results = {};
    
    // Validasi Sheet Data
    const dataSheet = ss.getSheetByName('Data');
    if (dataSheet) {
      const dataHeaders = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
      const requiredDataHeaders = ['nama', 'nisn', 'paswd', 'ortu', 'kontak.ortu'];
      const missingDataHeaders = requiredDataHeaders.filter(header => !dataHeaders.includes(header));
      
      results.dataSheet = {
        exists: true,
        headers: dataHeaders,
        missingHeaders: missingDataHeaders,
        rowCount: dataSheet.getLastRow() - 1
      };
    } else {
      results.dataSheet = { exists: false };
    }
    
    // Validasi Sheet Kehadiran
    const kehadiranSheet = ss.getSheetByName('Kehadiran');
    if (kehadiranSheet) {
      const kehadiranHeaders = kehadiranSheet.getRange(1, 1, 1, kehadiranSheet.getLastColumn()).getValues()[0];
      const requiredKehadiranHeaders = ['tahun', 'bulan', 'nisn', 'j.khdrn', 'izin', 'sakit', 'alpha', 'wali.pelanggaran', 'wali.catatan', 'wali.prestasi', 'bk.catatan'];
      const missingKehadiranHeaders = requiredKehadiranHeaders.filter(header => !kehadiranHeaders.includes(header));
      
      results.kehadiranSheet = {
        exists: true,
        headers: kehadiranHeaders,
        missingHeaders: missingKehadiranHeaders,
        rowCount: kehadiranSheet.getLastRow() - 1
      };
    } else {
      results.kehadiranSheet = { exists: false };
    }
    
    return createResponse(true, 'Validasi struktur sheet selesai', results);
  } catch (error) {
    console.error('Error in validateSheetStructure:', error);
    return createResponse(false, 'Gagal melakukan validasi: ' + error.message);
  }
}

/**
 * Fungsi untuk logging aktivitas (opsional)
 */
function logActivity(action, nisn, details = '') {
  try {
    const ss = getSpreadsheet();
    let logSheet = ss.getSheetByName('ActivityLog');
    
    if (!logSheet) {
      logSheet = ss.insertSheet('ActivityLog');
      logSheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Action', 'NISN', 'Details']]);
    }
    
    const timestamp = new Date().toISOString();
    const lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow + 1, 1, 1, 4).setValues([[timestamp, action, nisn, details]]);
    
  } catch (error) {
    console.error('Error in logActivity:', error);
    // Jangan throw error untuk logging, karena ini tidak kritis
  }
}