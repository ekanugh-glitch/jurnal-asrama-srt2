const SPREADSHEET_ID = '1gc88wC_Pm_YUD9-n3Wemls0AwxNUqjXFJ0zVFS19SUQ';
const SHEET_NAME = 'DATA_SISWA';
const HISTORY_SHEET_NAME = 'RIWAYAT_REKAP';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Jurnal Asrama SRT 2 Kota Surabaya')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getStudentData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];
  
  const headers = data[0].map(h => String(h).trim().toUpperCase());
  
  const idxId = headers.indexOf('ID SISWA');
  const idxNama = headers.indexOf('NAMA');
  const idxKelas = headers.indexOf('KELAS');
  const idxJk = headers.indexOf('JENIS KELAMIN');
  const idxKamar = headers.indexOf('KAMAR');
  const idxStatus = headers.indexOf('STATUS');
  const idxKetStatus = headers.indexOf('KETERANGAN STATUS');
  const idxKetSakit = headers.indexOf('KETERANGAN SAKIT');
  const idxCatatan = headers.indexOf('CATATAN KHUSUS');

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[idxNama]) continue;
    
    result.push({
      rowIndex: i + 1,
      id: row[idxId] || '',
      nama: String(row[idxNama] || '').trim(),
      kelas: String(row[idxKelas] || '').trim(),
      jk: String(row[idxJk] || '').trim(),
      kamar: String(row[idxKamar] || '').trim(),
      status: String(row[idxStatus] || 'Ada').trim(),
      ketStatus: String(row[idxKetStatus] || '').trim(),
      ketSakit: String(row[idxKetSakit] || '').trim(),
      catatan: String(row[idxCatatan] || '').trim()
    });
  }
  
  return result;
}

function updateStudentStatus(payload) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const rowIndex = parseInt(payload.rowIndex);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toUpperCase());
    
    const colStatus = headers.indexOf('STATUS') + 1;
    const colKetStatus = headers.indexOf('KETERANGAN STATUS') + 1;
    const colKetSakit = headers.indexOf('KETERANGAN SAKIT') + 1;
    const colCatatan = headers.indexOf('CATATAN KHUSUS') + 1;

    // Lakukan update di sheet utama
    if (colStatus > 0) sheet.getRange(rowIndex, colStatus).setValue(payload.status);
    if (colKetStatus > 0) sheet.getRange(rowIndex, colKetStatus).setValue(payload.ketStatus);
    if (colKetSakit > 0) sheet.getRange(rowIndex, colKetSakit).setValue(payload.ketSakit);
    if (colCatatan > 0) sheet.getRange(rowIndex, colCatatan).setValue(payload.catatan);

    // Kelola pencatatan ke sheet RIWAYAT_REKAP
    manageHistoryLog(ss, {
      id: payload.id || sheet.getRange(rowIndex, headers.indexOf('ID SISWA') + 1).getValue(),
      nama: payload.nama || sheet.getRange(rowIndex, headers.indexOf('NAMA') + 1).getValue(),
      kelas: payload.kelas || sheet.getRange(rowIndex, headers.indexOf('KELAS') + 1).getValue(),
      status: payload.status,
      ketStatus: payload.ketStatus,
      ketSakit: payload.ketSakit,
      catatan: payload.catatan
    });

    return { success: true, message: 'Data berhasil diperbarui & dicatat ke riwayat!' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Fungsi kelola riwayat: hitung jumlah hari dan bersihkan jika kembali "Ada" / sembuh / catatan kosong
function manageHistoryLog(ss, data) {
  let histSheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  
  // Jika sheet riwayat belum ada, buat otomatis
  if (!histSheet) {
    histSheet = ss.insertSheet(HISTORY_SHEET_NAME);
    histSheet.appendRow([
      'TIMESTAMP', 
      'ID SISWA', 
      'NAMA', 
      'KELAS', 
      'STATUS', 
      'KETERANGAN STATUS', 
      'KETERANGAN SAKIT', 
      'CATATAN PEMBINAAN', 
      'JUMLAH HARI TIDAK ADA/SAKIT/CATATAN'
    ]);
    histSheet.getRange(1, 1, 1, 9).setFontWeight('bold');
  }

  const now = new Date();
  const isSelesai = (data.status.toLowerCase() === 'ada') && (!data.ketSakit || data.ketSakit.trim() === '') && (!data.catatan || data.catatan.trim() === '');

  const lastRow = histSheet.getLastRow();
  let existingRowIndex = -1;
  let startDate = null;

  // Cek apakah siswa ini sudah memiliki catatan aktif di riwayat sebelumnya
  if (lastRow > 1) {
    const histData = histSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    for (let i = histData.length - 1; i >= 0; i--) {
      if (String(histData[i][2]).trim().toLowerCase() === String(data.nama).trim().toLowerCase()) {
        existingRowIndex = i + 2; // Baris riwayat di sheet (1-indexed + header)
        startDate = new Date(histData[i][0]);
        break;
      }
    }
  }

  if (isSelesai) {
    // Jika status kembali normal (Ada, sakit kosong, catatan kosong)
    if (existingRowIndex !== -1) {
      // Hitung total hari dari awal mulai sampai tanggal hari ini
      const diffTime = Math.abs(now - startDate);
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      const totalHari = diffDays > 0 ? diffDays : 1;

      // Update baris terakhir siswa tersebut menjadi status selesai/normal dengan total hari akhir
      histSheet.getRange(existingRowIndex, 1).setValue(Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
      histSheet.getRange(existingRowIndex, 5).setValue('Selesai / Kembali Ada');
      histSheet.getRange(existingRowIndex, 6).setValue('-');
      histSheet.getRange(existingRowIndex, 7).setValue('-');
      histSheet.getRange(existingRowIndex, 8).setValue('-');
      histSheet.getRange(existingRowIndex, 9).setValue(totalHari + ' Hari (Selesai)');
    }
  } else {
    // Jika masih dalam status Tidak Ada, Sakit, atau ada Catatan Pembinaan
    let jumlahHari = 1;
    if (existingRowIndex !== -1) {
      const diffTime = Math.abs(now - startDate);
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      jumlahHari = diffDays > 0 ? diffDays : 1;
    }

    // Tambahkan baris baru riwayat update terkini beserta jumlah hari berjalan
    histSheet.appendRow([
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      data.id,
      data.nama,
      data.kelas,
      data.status,
      data.ketStatus || '-',
      data.ketSakit || '-',
      data.catatan || '-',
      jumlahHari + ' Hari'
    ]);
  }
}