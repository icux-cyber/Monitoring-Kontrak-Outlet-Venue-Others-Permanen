const BACKUP_SHEET_NAME = "Monitoring Kontrak Outlet";
const BACKUP_COLUMNS = [
  "ID Outlet", "Nama Customer", "Alamat", "Zona", "Sektor", "Pos Budget",
  "Kategori Outlet", "Kategori Pekerjaan", "Brand", "Nominal Kontrak",
  "Periode Kontrak Awal", "Periode Kontrak Akhir", "Panel Toko", "Showcase",
  "Layar Toko", "Shop Sign / Neon Box", "Ukuran Panel Toko",
  "Ukuran Sticker Panel Toko",
];

function doPost(event) {
  try {
    const payload = JSON.parse(event.postData.contents);
    if (payload.action !== "backup" || payload.sheet_name !== BACKUP_SHEET_NAME) {
      throw new Error("Permintaan backup tidak valid.");
    }
    if (
      JSON.stringify(payload.columns) !== JSON.stringify(BACKUP_COLUMNS) ||
      !Array.isArray(payload.records)
    ) {
      throw new Error("Struktur data backup tidak valid.");
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(BACKUP_SHEET_NAME)
      || spreadsheet.insertSheet(BACKUP_SHEET_NAME);
    const rows = payload.records.map((record) => BACKUP_COLUMNS.map((column) => {
      const value = record[column];
      if (value === null || value === undefined) {
        return "";
      }
      if (
        (column === "Periode Kontrak Awal" || column === "Periode Kontrak Akhir") &&
        typeof value === "string" &&
        value
      ) {
        const parsedDate = new Date(value);
        return Number.isNaN(parsedDate.getTime()) ? value : parsedDate;
      }
      return value;
    }));

    sheet.clearContents();
    sheet.getRange(1, 1, 1, BACKUP_COLUMNS.length).setValues([BACKUP_COLUMNS]);
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, BACKUP_COLUMNS.length).setValues(rows);
      sheet.getRange(2, 11, rows.length, 2).setNumberFormat("dd/mm/yyyy");
    }
    sheet.setFrozenRows(1);
    return jsonResponse({
      ok: true,
      row_count: rows.length,
      backed_up_at: payload.backup_at,
    });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error) });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
