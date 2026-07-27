function doGet() {
  let template;
  try {
    template = HtmlService.createTemplateFromFile('Index');
  } catch (e) {
    try {
      template = HtmlService.createTemplateFromFile('index');
    } catch (e2) {
      return HtmlService.createHtmlOutput('Error: File utama tidak ditemukan. Pastikan Anda memiliki file bernama "Index.html" atau "Index" di proyek Google Apps Script Anda.');
    }
  }

  try {
    return template
      .evaluate()
      .setTitle('My Absen - Manajemen Kelas')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (e) {
    return HtmlService.createHtmlOutput('Terjadi kesalahan saat memuat aplikasi. Kemungkinan file "JavaScript.html" belum dibuat atau namanya tidak sesuai. Detail Error: ' + e.message);
  }
}

/**
 * Helper untuk menggabungkan file HTML terpisah (mis. JavaScript.html) ke dalam Index.html
 * lewat scriptlet <?!= include("JavaScript"); ?>. Ini praktik baku Apps Script untuk
 * memisahkan JavaScript besar dari file HTML utama, supaya lebih stabil saat dirender.
 */
function include(filename) {
  const possibleNames = [filename, filename.toLowerCase(), filename.toLowerCase() + '.html', filename + '.html'];
  for (let name of possibleNames) {
    try {
      return HtmlService.createHtmlOutputFromFile(name).getContent();
    } catch (e) {
      // Coba kemungkinan nama berikutnya
    }
  }
  return '<script>alert("Gagal memuat file ' + filename + '. Pastikan Anda telah membuat file tersebut di Google Apps Script dengan nama yang benar.");</script>';
}

/**
 * Jalankan fungsi ini pertama kali untuk membuat struktur database (Sheets).
 * Aman dijalankan ulang kapan saja: hanya membuat sheet yang belum ada,
 * dan memastikan seluruh kolom data diformat sebagai Teks (lihat catatan di writeSheet_).
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [
    {name: 'Config', headers: ['Key', 'Value']},
    {name: 'Classes', headers: ['ClassID', 'ClassName']},
    {name: 'Students', headers: ['StudentID', 'ClassID', 'StudentName', 'ParentPhone']},
    {name: 'Attendance', headers: ['RecordID', 'ClassID', 'MonthYear', 'DatesData', 'AttendanceData']},
    {name: 'Grades', headers: ['RecordID', 'ClassID', 'Subject', 'Type', 'MonthYear', 'DatesData', 'Cakupan', 'GradesData']},
    {name: 'Koku', headers: ['RecordID', 'ClassID', 'Tanggal', 'Jenis', 'Projek', 'Deskripsi', 'Dimensi', 'KokuData']},
    {name: 'Journal', headers: ['RecordID', 'Date', 'ClassID', 'Subject', 'Content', 'AttachmentUrl']},
    {name: 'Schedule', headers: ['RecordID', 'Day', 'TimeStart', 'TimeEnd', 'Subject', 'ClassID']},
    {name: 'Subjects', headers: ['SubjectName']},
    {name: 'Poin', headers: ['RecordID', 'StudentID', 'ClassID', 'Tanggal', 'PoinReward', 'PoinPelanggaran', 'Keterangan']}
  ];

  sheets.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);
      sheet.getRange(1, 1, 1, s.headers.length).setFontWeight('bold').setBackground('#f4f7f6');
      sheet.setFrozenRows(1);
    }
    // PENTING: paksa seluruh kolom data (baris di bawah header) sebagai format Teks.
    // Tanpa ini, Google Sheets otomatis mengubah tipe data saat ditulis lewat setValues(),
    // misalnya string "07" (bulan Juli) menjadi angka 7, atau "2026-07-11" menjadi objek Date.
    // Akibatnya pencocokan data di aplikasi (yang membandingkan string persis, mis. filter bulan)
    // gagal menemukan data tersebut walau datanya benar-benar ada di spreadsheet.
    const maxRows = sheet.getMaxRows();
    if (maxRows > 1) {
      sheet.getRange(2, 1, maxRows - 1, s.headers.length).setNumberFormat('@');
    }
  });
}

/**
 * Helper function untuk membaca Sheet menjadi array of object
 */
function cleanData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Kosong atau hanya header

  const headers = data[0];
  const result = [];

  for (let i = 1; i < data.length; i++) {
    let obj = {};
    for (let j = 0; j < headers.length; j++) {
      let val = data[i][j];
      // Jaga-jaga jika ada sel lama yang masih tersimpan sebagai Date/Number akibat bug lama
      // (sebelum kolom diformat Teks) supaya tetap terbaca benar oleh frontend.
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  return result;
}

/**
 * Mengambil semua data dari spreadsheet ke frontend
 */
function getAllData() {
  const configData = cleanData('Config');
  let configObj = {};
  configData.forEach(row => { configObj[row.Key] = row.Value; });

  const subjectsData = cleanData('Subjects').map(row => row.SubjectName);

  return {
    config: configObj,
    classes: cleanData('Classes'),
    students: cleanData('Students'),
    attendance: cleanData('Attendance'),
    grades: cleanData('Grades'),
    koku: cleanData('Koku'),
    journal: cleanData('Journal'),
    schedule: cleanData('Schedule'),
    subjects: subjectsData,
    poin: cleanData('Poin')
  };
}

function writeSheet_(ss, sheetName, headers, dataArray) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  // Hapus data lama, pertahankan header
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }

  if (dataArray && dataArray.length > 0) {
    const rows = dataArray.map(obj => headers.map(h => {
      let val = obj[h];
      if (val === undefined || val === null) return '';
      // Jika nilai adalah array/objek (seperti data nilai/absen), ubah ke string
      if (typeof val === 'object') return JSON.stringify(val);
      return val;
    }));
    const range = sheet.getRange(2, 1, rows.length, headers.length);
    // PENTING: format sebagai Teks SEBELUM menulis nilai. Ini mencegah Google Sheets
    // otomatis mengonversi string seperti "07" menjadi angka 7, atau "2026-07-11"
    // menjadi objek Date — akar penyebab data "hilang" dari aplikasi meski ada di spreadsheet.
    range.setNumberFormat('@');
    range.setValues(rows);
  }
}

/**
 * Menyimpan seluruh data JSON dari frontend ke masing-masing Sheet.
 * Dibungkus LockService agar dua proses simpan tidak berjalan bersamaan dan saling
 * menimpa/merusak data (mis. dua tab/perangkat menyimpan pada waktu yang hampir sama).
 */
function saveAllData(dbString, masterId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // tunggu maks 30 detik jika ada proses simpan lain yang sedang berjalan
  } catch (e) {
    throw new Error('Server sedang sibuk menyimpan data lain, silakan coba lagi beberapa saat.');
  }

  try {
    const db = JSON.parse(dbString);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // PENTING untuk keamanan multi-perangkat (HP/Tablet/Laptop dengan akun sama):
    // nomor revisi (_rev) TIDAK dipercaya dari client, melainkan SELALU dibaca ulang dan
    // dinaikkan di sini, di server, sehingga menjadi satu penghitung tunggal yang sama-sama
    // dipakai oleh semua perangkat — tidak peduli perangkat mana yang terakhir menyimpan.
    const currentConfig = {};
    cleanData('Config').forEach(row => { currentConfig[row.Key] = row.Value; });
    const currentRev = Number(currentConfig._rev) || 0;
    const newRev = currentRev + 1;
    db.config._rev = newRev;

    const configArr = [];
    for (let key in db.config) {
      configArr.push({Key: key, Value: db.config[key]});
    }

    writeSheet_(ss, 'Config', ['Key', 'Value'], configArr);
    writeSheet_(ss, 'Classes', ['ClassID', 'ClassName'], db.classes);
    writeSheet_(ss, 'Students', ['StudentID', 'ClassID', 'StudentName', 'ParentPhone'], db.students);
    writeSheet_(ss, 'Attendance', ['RecordID', 'ClassID', 'MonthYear', 'DatesData', 'AttendanceData'], db.attendance);
    writeSheet_(ss, 'Grades', ['RecordID', 'ClassID', 'Subject', 'Type', 'MonthYear', 'DatesData', 'Cakupan', 'GradesData'], db.grades);
    writeSheet_(ss, 'Koku', ['RecordID', 'ClassID', 'Tanggal', 'Jenis', 'Projek', 'Deskripsi', 'Dimensi', 'KokuData'], db.koku);
    writeSheet_(ss, 'Journal', ['RecordID', 'Date', 'ClassID', 'Subject', 'Content', 'AttachmentUrl'], db.journal);
    writeSheet_(ss, 'Schedule', ['RecordID', 'Day', 'TimeStart', 'TimeEnd', 'Subject', 'ClassID'], db.schedule);
    writeSheet_(ss, 'Poin', ['RecordID', 'StudentID', 'ClassID', 'Tanggal', 'PoinReward', 'PoinPelanggaran', 'Keterangan'], db.poin);

    const subjArr = (db.subjects || []).map(s => ({SubjectName: s}));
    writeSheet_(ss, 'Subjects', ['SubjectName'], subjArr);

    return { ok: true, rev: newRev };
  } finally {
    lock.releaseLock();
  }
}

/**
 * JALANKAN SEKALI SAJA secara manual dari editor Apps Script (pilih fungsi ini lalu klik Run)
 * untuk memperbaiki data yang SUDAH TERLANJUR rusak akibat bug lama:
 *  - Kolom MonthYear di Attendance/Grades yang berubah dari "07" menjadi angka 7 (kehilangan
 *    angka nol di depan), sehingga data absensi/nilai bulan Jan-Sep tidak muncul di aplikasi.
 *  - Kolom Date di Journal yang berubah dari teks "2026-07-11" menjadi objek Date.
 * Aman dijalankan berkali-kali (idempotent) — tidak akan merusak data yang sudah benar.
 *
 * SETELAH menjalankan ini, buka aplikasi lalu klik "Segarkan Data dari Spreadsheet"
 * (di menu profil) di SETIAP perangkat/browser yang pernah dipakai, supaya cache lama di
 * perangkat tersebut tidak menimpa balik data yang baru saja diperbaiki ini.
 */
function repairLegacyDataTypes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function fixMonthYearColumn(sheetName, colIndex) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;
    const range = sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1);
    const values = range.getValues();
    const fixed = values.map(row => {
      let v = row[0];
      if (v === '' || v === null || v === undefined) return [v];
      let num = Number(v);
      if (!isNaN(num) && String(v).trim() !== '') return [String(num).padStart(2, '0')];
      return [v];
    });
    range.setNumberFormat('@');
    range.setValues(fixed);
  }

  function fixDateColumn(sheetName, colIndex) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;
    const range = sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1);
    const values = range.getValues();
    const fixed = values.map(row => {
      let v = row[0];
      if (v instanceof Date) return [Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd')];
      return [v];
    });
    range.setNumberFormat('@');
    range.setValues(fixed);
  }

  fixMonthYearColumn('Attendance', 3); // kolom C = MonthYear
  fixMonthYearColumn('Grades', 5);     // kolom E = MonthYear
  fixDateColumn('Journal', 2);         // kolom B = Date

  SpreadsheetApp.getUi().alert('Perbaikan data selesai. Sekarang buka aplikasi lalu klik "Segarkan Data dari Spreadsheet" di menu profil pada setiap perangkat yang pernah dipakai.');
}

/**
 * JALANKAN SEKALI SAJA secara manual dari editor Apps Script untuk memperbaiki fitur
 * Nilai Kokurikuler: sheet 'Koku' versi lama hanya punya kolom [Siswa, Kategori, Narasi]
 * tanpa ClassID sama sekali, sehingga aplikasi tidak pernah bisa menampilkan ulang nilai
 * yang sudah diinput per kelas. Fungsi ini mengubah skema sheet 'Koku' menjadi
 * [RecordID, ClassID, Tanggal, Jenis, Projek, Deskripsi, Dimensi, KokuData], memetakan
 * data lama ke kelasnya masing-masing (berdasarkan nama siswa yang cocok dengan sheet
 * Students saat ini), dan mencadangkan data lama apa adanya ke sheet 'Koku_Backup_Lama'
 * sebelum ditimpa. Aman dijalankan berkali-kali (idempotent) — kalau skema sudah baru,
 * fungsi ini tidak melakukan apa-apa.
 *
 * SETELAH menjalankan ini, buka aplikasi lalu klik "Segarkan Data dari Spreadsheet"
 * (di menu profil) di setiap perangkat yang pernah dipakai.
 */
function migrateKokuSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Koku');
  if (!sheet) { setupDatabase(); return; }

  const newHeaders = ['RecordID', 'ClassID', 'Tanggal', 'Jenis', 'Projek', 'Deskripsi', 'Dimensi', 'KokuData'];
  const currentHeaders = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];

  if (currentHeaders.length === newHeaders.length && currentHeaders.every((h, i) => h === newHeaders[i])) {
    SpreadsheetApp.getUi().alert('Sheet Koku sudah memakai skema baru. Tidak ada yang perlu dimigrasikan.');
    return;
  }

  // Baca data lama (format lama: Siswa, Kategori, Narasi) SEBELUM header ditimpa
  const oldRows = cleanData('Koku');

  // Cadangkan data lama apa adanya, untuk jaga-jaga
  if (oldRows.length > 0) {
    let backupSheet = ss.getSheetByName('Koku_Backup_Lama');
    if (!backupSheet) backupSheet = ss.insertSheet('Koku_Backup_Lama');
    backupSheet.clearContents();
    const oldHeaders = Object.keys(oldRows[0]);
    backupSheet.appendRow(oldHeaders);
    backupSheet.getRange(2, 1, oldRows.length, oldHeaders.length)
      .setValues(oldRows.map(r => oldHeaders.map(h => r[h])));
  }

  // Petakan nama siswa -> ClassID memakai data Siswa saat ini, lalu kelompokkan nilai lama per kelas
  const students = cleanData('Students');
  const nameToStudent = {};
  students.forEach(s => { nameToStudent[s.StudentName] = s; });

  const grouped = {}; // ClassID -> { StudentID: {Kategori, Narasi} }
  let skipped = 0;
  oldRows.forEach(row => {
    const student = nameToStudent[row.Siswa];
    if (!student) { skipped++; return; }
    if (!grouped[student.ClassID]) grouped[student.ClassID] = {};
    grouped[student.ClassID][student.StudentID] = { Kategori: row.Kategori, Narasi: row.Narasi };
  });

  const newRows = Object.keys(grouped).map((cid, i) => ({
    RecordID: 'K' + Date.now() + i,
    ClassID: cid, Tanggal: '', Jenis: 'LintasMapel',
    Projek: 'Data Lama (Migrasi Otomatis)', Deskripsi: '', Dimensi: '[]',
    KokuData: JSON.stringify(grouped[cid])
  }));

  sheet.clear();
  sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight('bold').setBackground('#f4f7f6');
  sheet.setFrozenRows(1);
  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) sheet.getRange(2, 1, maxRows - 1, newHeaders.length).setNumberFormat('@');

  if (newRows.length > 0) {
    const rows = newRows.map(r => newHeaders.map(h => r[h]));
    sheet.getRange(2, 1, rows.length, newHeaders.length).setValues(rows);
  }

  SpreadsheetApp.getUi().alert(
    'Migrasi sheet Koku selesai.\n' +
    newRows.length + ' kelas berhasil dimigrasikan ke format baru.\n' +
    (skipped > 0 ? skipped + ' data lama dilewati karena nama siswanya tidak ditemukan di data Siswa saat ini (sudah dicadangkan di sheet "Koku_Backup_Lama").' : 'Semua data berhasil dimigrasikan.')
  );
}

/**
 * --- FITUR BARU: UPLOAD KE GOOGLE DRIVE ---
 */
function uploadFileToDrive(base64Data, filename, mimeType) {
  try {
    var folder;
    var folders = DriveApp.getFoldersByName("AgendaGuru_Bukti");
    if (folders.hasNext()) { folder = folders.next(); }
    else { folder = DriveApp.createFolder("AgendaGuru_Bukti"); folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }

    var base64 = base64Data.split(',')[1];
    var blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, filename);
    var file = folder.createFile(blob);
    return file.getUrl();
  } catch (e) {
    return "";
  }
}

/**
 * --- FITUR BARU: SINKRONISASI KE GOOGLE CALENDAR ---
 */
function syncToCalendar(type, data) {
  try {
    var cal = CalendarApp.getDefaultCalendar();

    if(type === 'jurnal') {
      // Buat event 1 hari penuh untuk Jurnal Harian
      var title = "Tugas Mengajar: " + data.subject + " (" + data.className + ")";
      cal.createAllDayEvent(title, new Date(data.date), {description: data.content});
      return true;

    } else if (type === 'schedule') {
      // Buat event rutinan mingguan (recurring)
      var daysMap = { 'Minggu': CalendarApp.Weekday.SUNDAY, 'Senin': CalendarApp.Weekday.MONDAY, 'Selasa': CalendarApp.Weekday.TUESDAY, 'Rabu': CalendarApp.Weekday.WEDNESDAY, 'Kamis': CalendarApp.Weekday.THURSDAY, 'Jumat': CalendarApp.Weekday.FRIDAY, 'Sabtu': CalendarApp.Weekday.SATURDAY };
      var targetDayObj = {'Minggu': 0, 'Senin': 1, 'Selasa': 2, 'Rabu': 3, 'Kamis': 4, 'Jumat': 5, 'Sabtu': 6};
      var targetDay = targetDayObj[data.day];

      // Cari tanggal terdekat dari hari yang dipilih
      var d = new Date();
      d.setDate(d.getDate() + (targetDay + 7 - d.getDay()) % 7);

      var startT = data.start.split(':');
      var endT = data.end.split(':');
      var startDateTime = new Date(d.getFullYear(), d.getMonth(), d.getDate(), startT[0], startT[1]);
      var endDateTime = new Date(d.getFullYear(), d.getMonth(), d.getDate(), endT[0], endT[1]);

      var recurrence = CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekday(daysMap[data.day]);
      var titleSch = "Jadwal: " + data.subject + " (" + data.className + ")";
      cal.createEventSeries(titleSch, startDateTime, endDateTime, recurrence);
      return true;
    }
  } catch (e) {
    return false;
  }
}
