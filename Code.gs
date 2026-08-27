/**
 * ============================================================
 * KALKULATOR KKDA & DA - Google Apps Script
 * ============================================================
 * 
 * Script ini menghitung KKDA dan DA secara otomatis
 * berdasarkan input Panjang, Lebar, dan Tinggi.
 * 
 * Layout Spreadsheet (Sheet: "Kalkulator"):
 *   A1:C2 : Judul "KALKULATOR TAPLAK MEJA" (merged)
 *   A3    : Header "PANJANG"
 *   B3    : Header "LEBAR"
 *   C3    : Header "TINGGI"
 *   F3    : Label "JENIS TAPLAK"
 *   G3    : Label "BANYAK CHECKOUT"
 *   A4    : Input Panjang
 *   B4    : Input Lebar
 *   C4    : Input Tinggi
 *   F4    : Label "KKDA"        | G4 : Hasil KKDA (m²)
 *   F5    : Label "DA"          | G5 : Hasil DA (m²)
 * 
 * Admin cukup mengisi A4, B4, C4.
 * G4 dan G5 diisi otomatis oleh script.
 * ============================================================
 */

// ============================================================
// KONFIGURASI
// ============================================================

const CONFIG = {
  SHEET_NAME: "Kalkulator",
  // Input: baris 4, kolom A/B/C
  INPUT_ROW: 4,
  INPUT_COLS: { PANJANG: 1, LEBAR: 2, TINGGI: 3 },  // A, B, C
  // Output: kolom G
  OUTPUT_CELLS: {
    KKDA_ROW: 4, KKDA_COL: 7,  // G4
    DA_ROW: 5,   DA_COL: 7     // G5
  },
  // Layout rows
  TITLE_ROWS: [1, 2],    // A1:C2 merged
  HEADER_ROW: 3,         // PANJANG, LEBAR, TINGGI di baris 3
  DATA_START_ROW: 4,     // Input data di baris 4
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

  // Hanya proses jika baris input (baris 4)
  if (row !== CONFIG.INPUT_ROW) return;

  // Hanya proses jika kolom yang diedit adalah Panjang, Lebar, atau Tinggi
  const inputCols = Object.values(CONFIG.INPUT_COLS);
  if (!inputCols.includes(col)) return;

  // Proses perhitungan
  processRow(sheet);
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
function processRow(sheet) {
  const row = CONFIG.INPUT_ROW;
  const panjang = sheet.getRange(row, CONFIG.INPUT_COLS.PANJANG).getValue();
  const lebar = sheet.getRange(row, CONFIG.INPUT_COLS.LEBAR).getValue();
  const tinggi = sheet.getRange(row, CONFIG.INPUT_COLS.TINGGI).getValue();

  // Validasi: semua input harus angka positif
  if (!isValidInput(panjang) || !isValidInput(lebar) || !isValidInput(tinggi)) {
    clearOutputCells(sheet);
    return;
  }

  // Hitung KKDA
  const kkda = calculateKKDA(panjang, lebar, tinggi);

  // Hitung DA
  const da = calculateDA(panjang, lebar, tinggi);

  // Tulis hasil ke spreadsheet
  writeResults(sheet, kkda, da);
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
function writeResults(sheet, kkda, da) {
  const oc = CONFIG.OUTPUT_CELLS;

  // Tulis hasil KKDA ke G2
  const kkdaCell = sheet.getRange(oc.KKDA_ROW, oc.KKDA_COL);
  kkdaCell.setValue(kkda.luasM2);
  kkdaCell.setNumberFormat("0.00");

  // Tulis hasil DA ke G5
  const daCell = sheet.getRange(oc.DA_ROW, oc.DA_COL);
  daCell.setValue(da.luasM2);
  daCell.setNumberFormat("0.00");

  // Tampilkan notifikasi hasil
  const msg = 
    "KKDA  :  " + kkda.luasM2.toFixed(2) + " m²\n" +
    "DA      :  " + da.luasM2.toFixed(2) + " m²";
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Perhitungan berhasil! Hasil KKDA dan DA telah diperbarui.", 5);
}

/**
 * Menghapus output jika input tidak valid.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet aktif
 */
function clearOutputCells(sheet) {
  const oc = CONFIG.OUTPUT_CELLS;
  sheet.getRange(oc.KKDA_ROW, oc.KKDA_COL).clearContent();
  sheet.getRange(oc.DA_ROW, oc.DA_COL).clearContent();
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

  // ── Reset: hapus merge yang sudah ada ──
  sheet.getRange("A1:G5").breakApart();

  // ── JUDUL: A1:C2 (merged) ──
  sheet.getRange("A1:C2").merge();
  const titleCell = sheet.getRange("A1");
  titleCell.setValue("KALKULATOR TAPLAK MEJA");
  titleCell.setFontSize(17);
  titleCell.setFontWeight("bold");
  titleCell.setHorizontalAlignment("center");
  titleCell.setVerticalAlignment("middle");
  titleCell.setBackground("#FFFF00");

  // ── Baris 3: Input headers + Jenis Taplak header ──
  sheet.getRange("A3").setValue("PANJANG").setHorizontalAlignment("center").setBackground("#FFD966");
  sheet.getRange("B3").setValue("LEBAR").setHorizontalAlignment("center").setBackground("#FFD966");
  sheet.getRange("C3").setValue("TINGGI").setHorizontalAlignment("center").setBackground("#FFD966");
  sheet.getRange("F3").setValue("JENIS TAPLAK").setHorizontalAlignment("center").setBackground("#FFD966");
  sheet.getRange("G3").setValue("BANYAK CHECKOUT").setHorizontalAlignment("center").setBackground("#FFD966");

  // ── Baris 4: KKDA label & result ──
  sheet.getRange("F4").setValue("KKDA").setHorizontalAlignment("center").setBackground("#F6B26B");
  sheet.getRange("G4").setBackground("#FFFF00").setNumberFormat("0.00");

  // ── Baris 5: DA label & result ──
  sheet.getRange("F5").setValue("DA").setHorizontalAlignment("center").setBackground("#F6B26B");
  sheet.getRange("G5").setBackground("#FFFF00").setNumberFormat("0.00");

  // ── Set lebar kolom ──
  sheet.setColumnWidth(1, 140);  // A - Panjang
  sheet.setColumnWidth(2, 80);   // B - Lebar
  sheet.setColumnWidth(3, 80);   // C - Tinggi
  sheet.setColumnWidth(4, 80);   // D
  sheet.setColumnWidth(5, 80);   // E
  sheet.setColumnWidth(6, 120);  // F - Jenis Taplak / Label
  sheet.setColumnWidth(7, 160);  // G - Banyak Checkout / Hasil

  SpreadsheetApp.getUi().alert(
    "✅ Setup selesai!\n\n" +
    "Sheet '" + CONFIG.SHEET_NAME + "' sudah siap digunakan.\n\n" +
    "Silakan isi Panjang (A4), Lebar (B4), dan Tinggi (C4).\n" +
    "Hasil KKDA (G4) dan DA (G5) akan dihitung otomatis."
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

  processRow(sheet);

  SpreadsheetApp.getUi().alert("✅ Selesai! Perhitungan KKDA dan DA berhasil diperbarui.");
}
