/**
 * Fungsi ini dipicu saat spreadsheet dibuka.
 * Ini membuat menu kustom di bilah menu Google Sheets.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Akses Cepat') // Nama menu utama
      .addItem('Buka Tautan Google', 'openGoogleLink') // Item menu dan fungsi yang dipanggil
      .addToUi();
}

/**
 * Fungsi ini membuka tautan (URL) di tab atau jendela baru browser pengguna.
 */
function openGoogleLink() {
  var html = HtmlService.createHtmlOutput('<script>window.open("https://www.google.com", "_blank");</script>')
      .setWidth(100)
      .setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, 'Membuka Tautan...');
}
