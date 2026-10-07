// Google Apps Script - Buku Tamu / Ucapan Gita & Ridwan
// Spreadsheet: https://docs.google.com/spreadsheets/d/1PhmbrxFlX9HxOqt3m9sLmglY67jw-FroH_dkNcD5QDE/edit?gid=0#gid=0
// Nama sheet: Sheet1
// Kolom wajib (baris 1 header):
// A: timestamp | B: nama tamu | C: ucapan | D: konfirmasi kehadiran | E: jumlah tamu
//
// CARA PAKAI (sekali saja):
// 1. Buka https://script.google.com > New project > paste semua file ini > Save.
// 2. Ganti SHEET_ID di bawah jika beda (sudah diisi sesuai link kamu).
// 3. Run > setupSheet (sekali, untuk buat header). Izinkan akses saat diminta.
// 4. Deploy > New deployment > type: Web app:
//    - Execute as: Me
//    - Who has access: Anyone
//    - Copy Web app URL (https://script.google.com/macros/s/XXXX/exec)
// 5. Paste URL itu ke index.html variabel GOOGLE_APPS_SCRIPT_URL.
// 6. Test: buka Web app URL + ?action=list di browser, harus JSON.
//    Contoh: https://script.google.com/macros/s/XXXX/exec?action=list

const SHEET_ID = '1PhmbrxFlX9HxOqt3m9sLmglY67jw-FroH_dkNcD5QDE';
const SHEET_NAME = 'Sheet1';
const HEADERS = ['timestamp', 'nama tamu', 'ucapan', 'konfirmasi kehadiran', 'jumlah tamu'];

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  return sh;
}

function setupSheet() {
  const sh = getSheet_();
  const h1 = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const empty = h1.every(function (v) { return !v; });
  if (empty) sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
}

function doGet(e) {
  try {
    setupSheet();
    const sh = getSheet_();
    const vals = sh.getDataRange().getValues();
    const data = [];
    for (let i = 1; i < vals.length; i++) {
      const r = vals[i];
      if (!r[1] && !r[2]) continue;
      data.push({
        timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0] || ''),
        nama: String(r[1] || ''),
        ucapan: String(r[2] || ''),
        kehadiran: String(r[3] || ''),
        jumlah: String(r[4] || '')
      });
    }
    data.reverse();
    const out = JSON.stringify({ result: 'ok', data: data });
    const cb = e && e.parameter && e.parameter.callback;
    if (cb && /^[A-Za-z_$][\w$]*$/.test(cb)) {
      return ContentService.createTextOutput(cb + '(' + out + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(out)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    const out = JSON.stringify({ result: 'error', message: String(err) });
    const cb = e && e.parameter && e.parameter.callback;
    if (cb && /^[A-Za-z_$][\w$]*$/.test(cb)) {
      return ContentService.createTextOutput(cb + '(' + out + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(out)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    setupSheet();
    let nama = '', ucapan = '', kehadiran = '', jumlah = '';
    if (e && e.postData && e.postData.contents) {
      try {
        const b = JSON.parse(e.postData.contents);
        nama = b.nama || b.name || '';
        ucapan = b.ucapan || b.comment || b.message || '';
        kehadiran = b.kehadiran || b.attendance || b.confirmation || '';
        jumlah = b.jumlah || b.guest_count || b.count || '';
      } catch (_ignore) {}
    }
    if (e && e.parameter) {
      nama = nama || e.parameter.nama || e.parameter.name || '';
      ucapan = ucapan || e.parameter.ucapan || e.parameter.comment || '';
      kehadiran = kehadiran || e.parameter.kehadiran || e.parameter.attendance || '';
      jumlah = jumlah || e.parameter.jumlah || e.parameter.guest_count || '';
    }
    nama = String(nama).trim().slice(0, 100);
    ucapan = String(ucapan).trim().slice(0, 1000);
    kehadiran = String(kehadiran).trim().slice(0, 50);
    jumlah = String(jumlah).trim().slice(0, 10);
    if (!nama || !ucapan) {
      return ContentService
        .createTextOutput(JSON.stringify({ result: 'error', message: 'nama dan ucapan wajib diisi' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    getSheet_().appendRow([new Date(), nama, ucapan, kehadiran, jumlah]);
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
