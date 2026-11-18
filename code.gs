
const SHEET_NAME = 'data_kelas';
const EVENT_DATE = '2026-11-21';
const REFRESH_SECONDS = 5;

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Informasi Kelas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Membuat kolom last_update jika belum tersedia.
 * Dipanggil saat mengambil data dashboard.
 */
function prepareSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error('Tab "' + SHEET_NAME + '" tidak ditemukan.');
  }

  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(h => h.trim().toLowerCase());

  const required = [
    'kategori',
    'judul_kelas',
    'pembicara',
    'jenis_kelas',
    'sesi',
    'ruang',
    'waktu',
    'kuota',
    'jumlah_hadir',
    'sisa_kuota',
    'status'
  ];

  required.forEach(name => {
    if (!headers.includes(name)) {
      throw new Error('Kolom "' + name + '" tidak ditemukan.');
    }
  });

  if (!headers.includes('last_update')) {
    sheet.getRange(1, lastColumn + 1).setValue('last_update');
  }

  return sheet;
}

/**
 * Mencatat timestamp hanya pada baris kelas yang diedit.
 * Berlaku untuk perubahan jumlah_hadir, sisa_kuota, status.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  if (sheet.getName() !== SHEET_NAME) return;
  if (e.range.getRow() < 2) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map(h => h.trim().toLowerCase());

  const lastUpdateCol = headers.indexOf('last_update') + 1;
  const watchedColumns = [
    'jumlah_hadir',
    'sisa_kuota',
    'status'
  ].map(name => headers.indexOf(name) + 1);

  if (lastUpdateCol < 1) return;

  const editedFirstCol = e.range.getColumn();
  const editedLastCol = editedFirstCol + e.range.getNumColumns() - 1;

  const relevantEdit = watchedColumns.some(col =>
    col >= editedFirstCol && col <= editedLastCol
  );

  if (!relevantEdit) return;

  const firstRow = Math.max(2, e.range.getRow());
  const lastRow = e.range.getLastRow();
  const now = new Date();

  const timestamps = Array.from(
    { length: lastRow - firstRow + 1 },
    () => [now]
  );

  sheet.getRange(
    firstRow,
    lastUpdateCol,
    timestamps.length,
    1
  ).setValues(timestamps);
}

/**
 * Mengambil data untuk website.
 */
function getDashboardData() {
  const sheet = prepareSheet();
  const values = sheet.getDataRange().getValues();
  const displayValues = sheet.getDataRange().getDisplayValues();

  if (values.length < 2) {
    return { classes: [], lastUpdated: Date.now() };
  }

  const headers = displayValues[0].map(h =>
    h.trim().toLowerCase()
  );

  const col = {};
  headers.forEach((header, index) => {
    col[header] = index;
  });

  const timezone = sheet.getParent().getSpreadsheetTimeZone();

  const now = new Date();
  const today = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');
  const currentTime = Utilities.formatDate(now, timezone, 'HH:mm');

  const [hour, minute] = currentTime.split(':').map(Number);
  const currentMinutes = hour * 60 + minute;

  // Sebelum acara: semua kelas tetap tampil.
  // Tanggal acara: kelas hilang setelah jam selesai.
  // Setelah acara: semua kelas disembunyikan.
  if (today > EVENT_DATE) {
    return { classes: [], lastUpdated: now.getTime() };
  }

  const classes = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const displayRow = displayValues[i];

    if (!String(displayRow[col.judul_kelas] || '').trim()) {
      continue;
    }

    const waktu = String(displayRow[col.waktu] || '').trim();

    let endMinutes = null;

    const timeMatch = waktu.match(
      /(\d{1,2})[.:](\d{2})\s*[-–—]\s*(\d{1,2})[.:](\d{2})/
    );

    if (timeMatch) {
      endMinutes =
        Number(timeMatch[3]) * 60 +
        Number(timeMatch[4]);
    }

    // Hanya menyaring jam kelas pada tanggal acara.
    if (
      today === EVENT_DATE &&
      endMinutes !== null &&
      endMinutes <= currentMinutes
    ) {
      continue;
    }

    const timestamp = row[col.last_update];

    classes.push({
      kategori: displayRow[col.kategori],
      judul_kelas: displayRow[col.judul_kelas],
      pembicara: displayRow[col.pembicara],
      jenis_kelas: displayRow[col.jenis_kelas],
      sesi: displayRow[col.sesi],
      ruang: displayRow[col.ruang],
      waktu: waktu,
      sisa_kuota: displayRow[col.sisa_kuota],
      status: displayRow[col.status],
      last_update: timestamp instanceof Date
        ? timestamp.getTime()
        : null
    });
  }

  return {
    classes: classes,
    lastUpdated: now.getTime()
  };
}
