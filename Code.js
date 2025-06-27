function onOpen() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const menu = SpreadsheetApp.getUi().createMenu('Invoice Tools');

  if (sheet.getName().startsWith("Invoice")) {
    menu.addItem('Export to PDF', 'exportActiveSheetToPDF');
  }

  menu.addToUi();
}

function exportActiveSheetToPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (!sheet.getName().startsWith("Invoice")) {
    SpreadsheetApp.getUi().alert("❌ This is not an invoice sheet.");
    return;
  }

  const folder = DriveApp.getFileById(ss.getId()).getParents().next();
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` +
    `format=pdf&size=a4&portrait=true&fitw=true&sheetnames=false&printtitle=false&` +
    `pagenumbers=false&gridlines=false&fzr=false&gid=${sheet.getSheetId()}`;

  const pdfBlob = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }
  }).getBlob().setName(`${sheet.getName()}.pdf`);

  folder.createFile(pdfBlob);

  SpreadsheetApp.getUi().alert(`✅ PDF saved to:\n${folder.getName()}/${sheet.getName()}.pdf`);
}