function onDriveItemsSelected(e) {
  const items = e.drive?.selectedItems;
  const ui = CardService.newCardBuilder();

  if (!items || items.length === 0) {
    ui.setHeader(CardService.newCardHeader().setTitle("No file selected"))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText("❗ Please select a file in Google Drive and refresh the add-on."))
        .addWidget(CardService.newTextButton()
          .setText("🔄 Refresh Selection")
          .setOnClickAction(CardService.newAction().setFunctionName("onDriveItemsSelected"))));
    return ui.build();
  }

  const file = items[0];
  const fileId = file.id;
  const fileName = file.title;

  // Store the file ID temporarily for the button callback
  PropertiesService.getUserProperties().setProperty("selectedFileId", fileId);

  ui.setHeader(CardService.newCardHeader().setTitle("Invoice PDF Generator v2"))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextButton()
        .setText("🔄 Refresh Selection")
        .setOnClickAction(CardService.newAction().setFunctionName("onDriveItemsSelected")))
      .addWidget(CardService.newTextParagraph().setText("Selected file: <b>" + fileName + "</b>"))
      .addWidget(CardService.newTextParagraph().setText("📄 File ready — click below to generate PDF."))
      .addWidget(CardService.newTextButton()
        .setText("Generate PDF")
        .setOnClickAction(CardService.newAction()
          .setFunctionName("generatePDFFromUI")))
    );

  return ui.build();
}

function generatePDFFromUI() {
  const fileId = PropertiesService.getUserProperties().getProperty("selectedFileId");
  if (!fileId) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("❌ No file ID found"))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText("Please select a file and refresh the add-on.")))
      .build();
  }

  try {
    const pdfUrl = generateInvoicePDF(fileId);
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("✅ PDF Created"))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText("✅ PDF generated and saved. You may now close this panel."))
        .addWidget(CardService.newTextButton()
          .setText("Open PDF")
          .setOpenLink(CardService.newOpenLink().setUrl(pdfUrl)))
        .addWidget(CardService.newTextButton()
          .setText("🔄 Refresh Selection")
          .setOnClickAction(CardService.newAction().setFunctionName("onDriveItemsSelected"))))
      .build();
  } catch (err) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle("❌ Error"))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText(err.message)))
      .build();
  }
}

function generateInvoicePDF(fileId) {
  const ss = SpreadsheetApp.openById(fileId);
  const sheet = ss.getSheetByName("Invoice") || ss.getSheets()[0];
  const pdfName = ss.getName() + ".pdf";

  const url = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?" +
    "format=pdf&size=a4&portrait=true&fitw=true&sheetnames=false&printtitle=false&" +
    "pagenumbers=false&gridlines=false&fzr=false&gid=" + sheet.getSheetId();

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + token
    }
  });

  const pdfBlob = response.getBlob().setName(pdfName);
  const parent = DriveApp.getFileById(fileId).getParents().next();
  const file = parent.createFile(pdfBlob);
  return file.getUrl();
}

function onSheetsHomepage() {
  const fileId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const fileName = SpreadsheetApp.getActiveSpreadsheet().getName();

  PropertiesService.getUserProperties().setProperty("selectedFileId", fileId);

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Invoice PDF Generator (Sheets)"))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText("📄 Current file: <b>" + fileName + "</b>"))
      .addWidget(CardService.newTextButton()
        .setText("Generate PDF")
        .setOnClickAction(CardService.newAction()
          .setFunctionName("generatePDFFromUI"))))
    .build();
}