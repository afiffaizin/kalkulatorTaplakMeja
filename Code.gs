/**
 * ============================================================
 * KALKULATOR KKDA & DA - Google Apps Script
 * ============================================================
 * 
 * Script ini menghitung KKDA dan DA secara otomatis
 * berdasarkan input Panjang, Lebar, dan Tinggi.
 * 
 * Layout Spreadsheet (Sheet: "Kalkulator"):
 *   Baris 1 : Header
 *   Kolom A : Panjang (cm)
 *   Kolom B : Lebar (cm)
 *   Kolom C : Tinggi (cm)
 *   Kolom D : KKDA Ukuran (e.g., "350 × 200")
 *   Kolom E : KKDA (m²)
 *   Kolom F : DA Ukuran (e.g., "200 × 200")
 *   Kolom G : DA (m²)
 * 
 * Admin cukup mengisi kolom A, B, C.
 * Kolom D–G diisi otomatis oleh script.
 * ============================================================
 */

// ============================================================
// KONFIGURASI
// ============================================================

const CONFIG = {
  SHEET_NAME: "Kalkulator",
  INPUT_COLS: { PANJANG: 1, LEBAR: 2, TINGGI: 3 },  // A, B, C
  OUTPUT_COLS: {
    KKDA_LUAS: 4,    // D
    DA_LUAS: 5       // E
  },
  TITLE_ROW: 1,
  HEADER_ROW: 2,
  DATA_START_ROW: 3,
  ROUNDING_MULTIPLE: 50,
  CM2_TO_M2_DIVISOR: 10000
};

// ============================================================
// TRIGGER: onEdit
// ============================================================

/**
 * Trigger otomatis saat admin mengedit sel.
 * Jika sel yang diedit berada di kolom Panjang, Lebar, atau Tinggi,
 * maka KKDA dan DA dihitung ulang secara otomatis.
 * 
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - Event object
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  // Hanya proses sheet "Kalkulator"
  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  // Hanya proses baris data (bukan header)
  if (row < CONFIG.DATA_START_ROW) return;

  // Hanya proses jika kolom yang diedit adalah Panjang, Lebar, atau Tinggi
  const inputCols = Object.values(CONFIG.INPUT_COLS);
  if (!inputCols.includes(col)) return;

  // Proses semua baris yang diedit (mendukung paste multi-baris)
  const numRows = range.getNumRows();
  for (let i = 0; i < numRows; i++) {
    processRow(sheet, row + i);
  }
}

// ============================================================
// FUNGSI UTAMA: processRow
// ============================================================

/**
 * Memproses satu baris data: membaca input, menghitung KKDA & DA,
 * dan menulis hasil ke kolom output.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet aktif
 * @param {number} row - Nomor baris yang diproses
 */
function processRow(sheet, row) {
  const panjang = sheet.getRange(row, CONFIG.INPUT_COLS.PANJANG).getValue();
  const lebar = sheet.getRange(row, CONFIG.INPUT_COLS.LEBAR).getValue();
  const tinggi = sheet.getRange(row, CONFIG.INPUT_COLS.TINGGI).getValue();

  // Validasi: semua input harus angka positif
  if (!isValidInput(panjang) || !isValidInput(lebar) || !isValidInput(tinggi)) {
    clearOutputRow(sheet, row);
    return;
  }

  // Hitung KKDA
  const kkda = calculateKKDA(panjang, lebar, tinggi);

  // Hitung DA
  const da = calculateDA(panjang, lebar, tinggi);

  // Tulis hasil ke spreadsheet
  writeResults(sheet, row, kkda, da);
}

// ============================================================
// FUNGSI PERHITUNGAN
// ============================================================

/**
 * Membulatkan nilai ke atas ke kelipatan tertentu (default: 50).
 * 
 * Contoh:
 *   roundUp50(105) → 150
 *   roundUp50(180) → 200
 *   roundUp50(200) → 200
 *   roundUp50(340) → 350
 * 
 * @param {number} value - Nilai yang akan dibulatkan
 * @returns {number} Nilai yang sudah dibulatkan
 */
function roundUp50(value) {
  return Math.ceil(value / CONFIG.ROUNDING_MULTIPLE) * CONFIG.ROUNDING_MULTIPLE;
}

/**
 * Menghitung luas dari panjang × lebar (dalam cm²).
 * 
 * @param {number} panjang - Panjang dalam cm
 * @param {number} lebar - Lebar dalam cm
 * @returns {number} Luas dalam cm²
 */
function calculateArea(panjang, lebar) {
  return panjang * lebar;
}

/**
 * Mengkonversi luas dari cm² ke m².
 * 
 * @param {number} areaCm2 - Luas dalam cm²
 * @returns {number} Luas dalam m²
 */
function cm2ToM2(areaCm2) {
  return areaCm2 / CONFIG.CM2_TO_M2_DIVISOR;
}

/**
 * Menghitung KKDA.
 * 
 * Rumus:
 *   Panjang KKDA = roundUp50(Panjang + 2 × Tinggi)
 *   Lebar KKDA   = roundUp50(Lebar + Tinggi)
 *   Luas KKDA    = Panjang KKDA × Lebar KKDA (cm²) → m²
 * 
 * @param {number} panjang - Panjang meja (cm)
 * @param {number} lebar - Lebar meja (cm)
 * @param {number} tinggi - Tinggi meja (cm)
 * @returns {Object} { panjang, lebar, ukuran, luasCm2, luasM2 }
 */
function calculateKKDA(panjang, lebar, tinggi) {
  const panjangKKDA = roundUp50(panjang + (2 * tinggi));
  const lebarKKDA = roundUp50(lebar + tinggi);
  const luasCm2 = calculateArea(panjangKKDA, lebarKKDA);
  const luasM2 = cm2ToM2(luasCm2);

  return {
    panjang: panjangKKDA,
    lebar: lebarKKDA,
    ukuran: panjangKKDA + " × " + lebarKKDA,
    luasCm2: luasCm2,
    luasM2: luasM2
  };
}

/**
 * Menghitung DA (Depan Atas).
 * 
 * Rumus:
 *   Panjang DA = roundUp50(Panjang)
 *   Lebar DA   = roundUp50(Lebar + Tinggi)
 *   Luas DA    = Panjang DA × Lebar DA (cm²) → m²
 * 
 * @param {number} panjang - Panjang meja (cm)
 * @param {number} lebar - Lebar meja (cm)
 * @param {number} tinggi - Tinggi meja (cm)
 * @returns {Object} { panjang, lebar, ukuran, luasCm2, luasM2 }
 */
function calculateDA(panjang, lebar, tinggi) {
  const panjangDA = roundUp50(panjang);
  const lebarDA = roundUp50(lebar + tinggi);
  const luasCm2 = calculateArea(panjangDA, lebarDA);
  const luasM2 = cm2ToM2(luasCm2);

  return {
    panjang: panjangDA,
    lebar: lebarDA,
    ukuran: panjangDA + " × " + lebarDA,
    luasCm2: luasCm2,
    luasM2: luasM2
  };
}

// ============================================================
// FUNGSI OUTPUT
// ============================================================

/**
 * Menulis hasil perhitungan KKDA dan DA ke spreadsheet.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet aktif
 * @param {number} row - Nomor baris
 * @param {Object} kkda - Hasil perhitungan KKDA
 * @param {Object} da - Hasil perhitungan DA
 */
function writeResults(sheet, row, kkda, da) {
  const outputValues = [
    [kkda.luasM2, da.luasM2]
  ];

  sheet.getRange(row, CONFIG.OUTPUT_COLS.KKDA_LUAS, 1, 2).setValues(outputValues);

  // Format angka: 2 desimal untuk kolom luas (m²)
  sheet.getRange(row, CONFIG.OUTPUT_COLS.KKDA_LUAS).setNumberFormat("0.00");
  sheet.getRange(row, CONFIG.OUTPUT_COLS.DA_LUAS).setNumberFormat("0.00");
}

/**
 * Menghapus kolom output jika input tidak valid.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet aktif
 * @param {number} row - Nomor baris
 */
function clearOutputRow(sheet, row) {
  sheet.getRange(row, CONFIG.OUTPUT_COLS.KKDA_LUAS, 1, 2).clearContent();
}

// ============================================================
// FUNGSI VALIDASI
// ============================================================

/**
 * Mengecek apakah input valid (angka positif).
 * 
 * @param {*} value - Nilai yang dicek
 * @returns {boolean} true jika valid
 */
function isValidInput(value) {
  return typeof value === "number" && !isNaN(value) && value > 0;
}

// ============================================================
// FUNGSI SETUP (Jalankan sekali)
// ============================================================

/**
 * Fungsi untuk setup awal spreadsheet.
 * Jalankan sekali untuk membuat header dan format sheet.
 * 
 * Cara menjalankan:
 *   1. Buka Apps Script Editor
 *   2. Pilih fungsi "setupSheet" dari dropdown
 *   3. Klik tombol Run (▶)
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // Buat sheet baru jika belum ada
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  // Set judul
  sheet.getRange(CONFIG.TITLE_ROW, 1).setValue("KALKULATOR TAPLAK MEJA");

  // Merge judul across all columns
  sheet.getRange(CONFIG.TITLE_ROW, 1, 1, 5).merge();

  // Format judul
  const titleRange = sheet.getRange(CONFIG.TITLE_ROW, 1);
  titleRange.setFontSize(18);
  titleRange.setFontWeight("bold");
  titleRange.setHorizontalAlignment("center");
  titleRange.setVerticalAlignment("middle");
  titleRange.setBackground("#1a3c6e");
  titleRange.setFontColor("#ffffff");
  sheet.setRowHeight(CONFIG.TITLE_ROW, 50);

  // Set header
  const headers = [
    ["Panjang (cm)", "Lebar (cm)", "Tinggi (cm)", "KKDA (m²)", "DA (m²)"]
  ];
  sheet.getRange(CONFIG.HEADER_ROW, 1, 1, 5).setValues(headers);

  // Format header
  const headerRange = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, 5);
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  headerRange.setBackground("#4a86c8");
  headerRange.setFontColor("#ffffff");

  // Set lebar kolom
  sheet.setColumnWidth(1, 120);  // Panjang
  sheet.setColumnWidth(2, 100);  // Lebar
  sheet.setColumnWidth(3, 100);  // Tinggi
  sheet.setColumnWidth(4, 100);  // KKDA (m²)
  sheet.setColumnWidth(5, 100);  // DA (m²)

  // Format kolom output agar tidak bisa diedit secara visual (warna background berbeda)
  if (sheet.getLastRow() > CONFIG.HEADER_ROW) {
    const dataRows = sheet.getLastRow() - CONFIG.HEADER_ROW;
    sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.OUTPUT_COLS.KKDA_LUAS, dataRows, 2).setBackground("#f0f4f8");
  }

  // Freeze judul + header
  sheet.setFrozenRows(CONFIG.HEADER_ROW);

  SpreadsheetApp.getUi().alert(
    "✅ Setup selesai!\n\n" +
    "Sheet '" + CONFIG.SHEET_NAME + "' sudah siap digunakan.\n\n" +
    "Silakan isi kolom Panjang, Lebar, dan Tinggi.\n" +
    "KKDA dan DA akan dihitung otomatis."
  );
}

/**
 * Fungsi untuk menghitung ulang semua baris data yang sudah ada.
 * Berguna jika ada perubahan rumus atau ingin rekalkulasi massal.
 * 
 * Cara menjalankan:
 *   1. Buka Apps Script Editor
 *   2. Pilih fungsi "recalculateAll" dari dropdown
 *   3. Klik tombol Run (▶)
 */
function recalculateAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ Sheet '" + CONFIG.SHEET_NAME + "' tidak ditemukan.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("ℹ️ Tidak ada data untuk dihitung.");
    return;
  }

  let count = 0;
  for (let row = CONFIG.DATA_START_ROW; row <= lastRow; row++) {
    processRow(sheet, row);
    count++;
  }

  SpreadsheetApp.getUi().alert("✅ Selesai! " + count + " baris berhasil dihitung ulang.");
}
