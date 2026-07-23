/**
 * ================================================================
 *  GOOGLE APPS SCRIPT – Absensi KKN Desa Muara Tembulih (v4)
 * ================================================================
 */

const SHEET_NAME = 'Absensi';
const MEMBERS_SHEET_NAME = 'Anggota';

const DEFAULT_MEMBERS = [
  { name: 'Wahyu Zainal Mutakin', npm: '230201061', jurusan: 'Teknik Informatika' },
  { name: 'Muhamad Iqbal', npm: '240101222', jurusan: 'S1 Keperawatan' },
  { name: 'Revani Sukma Dewi', npm: '240105011', jurusan: 'D3 Kebidanan' },
  { name: 'Diesta Putri Marantika', npm: '240105005', jurusan: 'D3 Kebidanan' },
  { name: 'Zenisa Auvin Naimah', npm: '230104029', jurusan: 'S1 Gizi' },
  { name: 'Astri Berlian Putri Atriska', npm: '230104020', jurusan: 'S1 Gizi' },
  { name: 'Amelia Al-Qoyyima Saputri', npm: '220104013', jurusan: 'S1 Gizi' },
  { name: 'Adila Nailah Savina', npm: '230104046', jurusan: 'S1 Gizi' },
  { name: 'Anisa Juspayana', npm: '230107009', jurusan: 'S1 Kebidanan' },
  { name: 'Seliya Wati', npm: '230107001', jurusan: 'S1 Kebidanan' },
  { name: 'Novita Widianti', npm: '230107095', jurusan: 'S1 Kebidanan' },
  { name: 'Islati Arum Sheila', npm: '230107090', jurusan: 'S1 Kebidanan' },
  { name: 'Vina Desita Sari', npm: '230107083', jurusan: 'S1 Kebidanan' },
  { name: 'Keyzia Alferita', npm: '240101167', jurusan: 'S1 Keperawatan' },
  { name: 'Helya Sari', npm: '240101251', jurusan: 'S1 Keperawatan' },
  { name: 'Rahmah Ulya Azizah', npm: '240101189', jurusan: 'S1 Keperawatan' },
  { name: 'Suci Inesa Putri', npm: '240101242', jurusan: 'S1 Keperawatan' },
  { name: 'Sonya Nanda Aulia', npm: '240101072', jurusan: 'S1 Keperawatan' },
  { name: 'Rahma Annisa', npm: '240101102', jurusan: 'S1 Keperawatan' },
  { name: 'Evi Nur Halimah', npm: '240101078', jurusan: 'S1 Keperawatan' }
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'ping') {
      return responseJSON({ status: 'pong', message: 'Koneksi berhasil!' });
    }

    if (data.action === 'addAbsensi') {
      return responseJSON(tambahAbsensi(data));
    }

    if (data.action === 'clearAllRecords') {
      return responseJSON(clearAllRecordsSheets());
    }

    // Central Member Actions
    if (data.action === 'addMember') {
      return responseJSON(addMemberSheets(data));
    }

    if (data.action === 'removeMember') {
      return responseJSON(removeMemberSheets(data));
    }

    if (data.action === 'clearAllMembers') {
      return responseJSON(clearAllMembersSheets());
    }

    if (data.action === 'loadDefaultMembers') {
      return responseJSON(loadDefaultMembersSheets());
    }

    return responseJSON({ status: 'error', message: 'Action tidak dikenal' });

  } catch (err) {
    return responseJSON({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  if (e && e.parameter) {
    if (e.parameter.action === 'getRecords') {
      return responseJSON(getRecordsData());
    }
    if (e.parameter.action === 'getMembers') {
      return responseJSON(getMembersData());
    }
  }
  return responseJSON({ status: 'pong', message: 'Script aktif dan berjalan!' });
}

function responseJSON(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// === CENTRAL MEMBERS SHEET CODE ===
function getMembersData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MEMBERS_SHEET_NAME);
    initMembersSheet(sheet);
  }
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return [];
  }
  const members = [];
  for (let i = 1; i < data.length; i++) {
    members.push({
      name: String(data[i][0]).trim(),
      npm: String(data[i][1]).trim(),
      jurusan: String(data[i][2]).trim()
    });
  }
  return members;
}

function initMembersSheet(sheet) {
  sheet.appendRow(['Nama Mahasiswa', 'NPM', 'Prodi/Jurusan']);
  const range = sheet.getRange(1, 1, 1, 3);
  range.setBackground('#9c27b0'); // Violet color for members header
  range.setFontColor('#ffffff');
  range.setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 200);

  // Populate default members
  DEFAULT_MEMBERS.forEach(m => {
    sheet.appendRow([m.name, m.npm, m.jurusan]);
  });
}

function addMemberSheets(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MEMBERS_SHEET_NAME);
    initMembersSheet(sheet);
  }
  const values = sheet.getDataRange().getValues();
  const name = data.name.trim();
  const exists = values.some(row => row[0].toString().toLowerCase() === name.toLowerCase());
  if (exists) {
    return { status: 'error', message: 'Nama sudah ada di Google Sheets!' };
  }
  sheet.appendRow([name, data.npm || '-', data.jurusan || '-']);
  return { status: 'success', message: 'Anggota berhasil ditambahkan ke Google Sheets' };
}

function removeMemberSheets(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (!sheet) return { status: 'error', message: 'Sheet anggota tidak ditemukan' };
  
  const values = sheet.getDataRange().getValues();
  const name = data.name.trim();
  let rowToDelete = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0].toString().toLowerCase() === name.toLowerCase()) {
      rowToDelete = i + 1;
      break;
    }
  }
  if (rowToDelete !== -1) {
    sheet.deleteRow(rowToDelete);
    return { status: 'success', message: 'Anggota berhasil dihapus dari Google Sheets' };
  }
  return { status: 'error', message: 'Anggota tidak ditemukan di Google Sheets' };
}

function clearAllMembersSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (sheet) {
    sheet.clear();
    sheet.appendRow(['Nama Mahasiswa', 'NPM', 'Prodi/Jurusan']);
    const range = sheet.getRange(1, 1, 1, 3);
    range.setBackground('#9c27b0');
    range.setFontColor('#ffffff');
    range.setFontWeight('bold');
  }
  return { status: 'success', message: 'Semua anggota berhasil dihapus' };
}

function loadDefaultMembersSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MEMBERS_SHEET_NAME);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(MEMBERS_SHEET_NAME);
  }
  initMembersSheet(sheet);
  return { status: 'success', message: 'Data default berhasil dimuat ulang' };
}

// ================================================================
//  FUNGSI: Tambah data absensi ke sheet
// ================================================================
function tambahAbsensi(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    buatHeader(sheet);
  }

  if (sheet.getLastRow() === 0) {
    buatHeader(sheet);
  }

  sheet.appendRow([
    data.tanggal  || '',
    data.waktu    || '',
    data.nama     || '',
    data.npm      || '-',
    data.jurusan  || '-',
    data.status   || '',
    data.kegiatan || '',
    new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' }),
  ]);

  const lastRow = sheet.getLastRow();
  formatBaris(sheet, lastRow, data.status);

  return { status: 'success', message: `Data ${data.nama} berhasil disimpan.`, row: lastRow };
}

// ================================================================
//  FUNGSI: Buat baris header
// ================================================================
function buatHeader(sheet) {
  const headers = ['Tanggal', 'Waktu', 'Nama Mahasiswa', 'NPM', 'Prodi/Jurusan', 'Status', 'Kegiatan', 'Timestamp Server'];
  sheet.appendRow(headers);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a56db');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');

  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2,  80);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 300);
  sheet.setColumnWidth(8, 180);
}

// ================================================================
//  FUNGSI: Warnai baris berdasarkan status
// ================================================================
function formatBaris(sheet, rowNum, status) {
  const range = sheet.getRange(rowNum, 1, 1, 6);

  let bgColor;
  switch (status) {
    case 'Hadir':  bgColor = '#e6f4ea'; break;
    case 'Izin':   bgColor = '#fef9e7'; break;
    case 'Sakit':  bgColor = '#e8f4f8'; break;
    case 'Alpha':  bgColor = '#fce8e6'; break;
    default:       bgColor = '#ffffff';
  }

  range.setBackground(bgColor);
  range.setFontSize(10);
  range.setBorder(null, null, true, null, null, null, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(rowNum, 3).setFontWeight('bold');

  sheet.getRange(rowNum, 1).setHorizontalAlignment('center');
  sheet.getRange(rowNum, 2).setHorizontalAlignment('center');
  sheet.getRange(rowNum, 4).setHorizontalAlignment('center');
}

// ================================================================
//  FUNGSI BONUS: Buat sheet Rekap per anggota (jalankan manual)
// ================================================================
function buatRekapAnggota() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAbsensi = ss.getSheetByName(SHEET_NAME);
  if (!sheetAbsensi) { SpreadsheetApp.getUi().alert('Sheet Absensi tidak ditemukan!'); return; }

  const data = sheetAbsensi.getDataRange().getValues();
  if (data.length <= 1) { SpreadsheetApp.getUi().alert('Tidak ada data absensi.'); return; }

  const namaSet = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2]) namaSet.add(data[i][2]);
  }

  const rekapData = [];
  namaSet.forEach(nama => {
    const rows   = data.filter(r => r[2] === nama);
    const hadir  = rows.filter(r => r[5] === 'Hadir').length;
    const izin   = rows.filter(r => r[5] === 'Izin').length;
    const sakit  = rows.filter(r => r[5] === 'Sakit').length;
    const alpha  = rows.filter(r => r[5] === 'Alpha').length;
    const total  = rows.length;
    const persen = total > 0 ? Math.round((hadir / total) * 100) : 0;
    rekapData.push([nama, hadir, izin, sakit, alpha, total, persen + '%']);
  });

  let rekapSheet = ss.getSheetByName('Rekap Per Anggota');
  if (rekapSheet) ss.deleteSheet(rekapSheet);
  rekapSheet = ss.insertSheet('Rekap Per Anggota');

  rekapSheet.appendRow(['Nama Mahasiswa', 'Hadir', 'Izin', 'Sakit', 'Alpha', 'Total', '% Kehadiran']);
  rekapData.forEach(row => rekapSheet.appendRow(row));

  const hRange = rekapSheet.getRange(1, 1, 1, 7);
  hRange.setBackground('#1a56db');
  hRange.setFontColor('#ffffff');
  hRange.setFontWeight('bold');
  hRange.setHorizontalAlignment('center');

  rekapSheet.setColumnWidth(1, 200);
  for (let i = 2; i <= 7; i++) rekapSheet.setColumnWidth(i, 90);
  rekapSheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('✅ Sheet "Rekap Per Anggota" berhasil dibuat!');
}

function getRecordsData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const records = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let tglStr = '';
      if (row[0] instanceof Date) {
        const d = row[0];
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        tglStr = `${y}-${m}-${day}`;
      } else {
        tglStr = row[0] ? String(row[0]).trim() : '';
      }

      let waktuStr = '';
      if (row[1] instanceof Date) {
        const d = row[1];
        const hh = String(d.getHours()).padStart(2, '0');
        const mm = String(d.getMinutes()).padStart(2, '0');
        waktuStr = `${hh}:${mm}`;
      } else {
        waktuStr = row[1] ? String(row[1]).trim() : '';
      }

      records.push({
        id:        Date.now() + i,
        tanggal:   tglStr,
        waktu:     waktuStr,
        nama:      row[2] ? String(row[2]).trim() : '',
        npm:       row[3] ? String(row[3]).trim() : '',
        jurusan:   row[4] ? String(row[4]).trim() : '',
        status:    row[5] ? String(row[5]).trim() : '',
        kegiatan:  row[6] ? String(row[6]).trim() : '',
      });
    }
    return records;
  } catch (err) {
    return [];
  }
}

function clearAllRecordsSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    return { status: 'error', message: 'Sheet Absensi tidak ditemukan.' };
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    // Delete all data rows, keeping only the header (row 1)
    sheet.deleteRows(2, lastRow - 1);
  }
  
  return { status: 'success', message: 'Semua data absensi berhasil dihapus.' };
}
