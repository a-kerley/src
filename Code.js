function onDriveItemsSelected(e) {
  const items = e.drive?.selectedItems;
  const ui = CardService.newCardBuilder();

  if (!items || items.length === 0) {
    ui.setHeader(CardService.newCardHeader().setTitle("No file selected"))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextButton()
          .setText("🔄 Refresh Selection")
          .setOnClickAction(CardService.newAction().setFunctionName("onDriveItemsSelected")))
        .addWidget(CardService.newTextParagraph().setText("❗ Please select a file in Google Drive and refresh the add-on.")));
    return ui.build();
  }

  const file = items[0];
  const fileId = file.id;
  const fileName = file.title;
  const mimeType = file.mimeType;

  PropertiesService.getUserProperties().setProperty("selectedFileId", fileId);

  const section = CardService.newCardSection()
    .addWidget(CardService.newTextButton()
      .setText("🔄 Refresh Selection")
      .setOnClickAction(CardService.newAction().setFunctionName("onDriveItemsSelected")))
    .addWidget(CardService.newTextParagraph().setText("Selected file: <b>" + fileName + "</b>"));

  if (mimeType === "application/vnd.google-apps.spreadsheet") {
    section.addWidget(CardService.newTextParagraph().setText("📄 File ready — choose an action below."))
      .addWidget(CardService.newTextButton()
        .setText("Generate PDF")
        .setOnClickAction(CardService.newAction().setFunctionName("generatePDFFromUI")))
      .addWidget(CardService.newTextButton()
        .setText("Generate PDF & Attach to Email Draft")
        .setOnClickAction(CardService.newAction().setFunctionName("generatePDFAndAttachToDraft")));
  } else {
    section.addWidget(CardService.newTextParagraph().setText("⚠️ Unsupported file type."));
  }

  ui.setHeader(CardService.newCardHeader().setTitle("Invoice PDF Generator v2"))
    .addSection(section);

  return ui.build();
}

// Make sure the following OAuth scopes are included in your appsscript.json:
// "oauthScopes": [
//   "https://www.googleapis.com/auth/script.external_request",
//   "https://www.googleapis.com/auth/drive.readonly",
//   "https://www.googleapis.com/auth/gmail.compose",
//   "https://www.googleapis.com/auth/gmail.modify"
// ]

function attachPDFToDraft() {
  const fileId = PropertiesService.getUserProperties().getProperty("selectedFileId");
  if (!fileId) {
    throw new Error("No file selected.");
  }

  const file = DriveApp.getFileById(fileId);
  const pdfBlob = file.getBlob();

  const invoiceNumberMatch = file.getName().match(/Invoice\s+(BA\d+)/i);
  const subject = invoiceNumberMatch ? `Invoice ${invoiceNumberMatch[1]}` : "Invoice";
  const body = "<p>Please find the attached invoice PDF.</p>";
  const draft = GmailApp.createDraft(
    "",
    subject,
    "",
    {
      htmlBody: body,
      attachments: [pdfBlob]
    }
  );

  const draftId = draft.getId();
  const draftUrl = `https://mail.google.com/mail/u/0/#drafts`;

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("✅ Email Draft Created"))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextButton()
        .setText("🔄 Refresh Selection")
        .setOnClickAction(CardService.newAction().setFunctionName("onDriveItemsSelected")))
      .addWidget(CardService.newTextParagraph()
        .setText("A Gmail draft has been created with the PDF attached. Please check your Drafts folder."))
      .addWidget(CardService.newTextButton()
        .setText("📬 Open Gmail Drafts")
        .setOpenLink(CardService.newOpenLink().setUrl(draftUrl))))
    .build();

  return card;
}
function generatePDFAndAttachToDraft() {
  const fileId = PropertiesService.getUserProperties().getProperty("selectedFileId");
  if (!fileId) {
    throw new Error("No file selected.");
  }

  const file = DriveApp.getFileById(fileId);
  const url = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=pdf&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  const pdfBlob = response.getBlob().setName(file.getName() + ".pdf");

  // Save the PDF to the same folder as the sheet, overwriting if a file with the same name exists
  const parentFolder = file.getParents().next();
  const existingFiles = parentFolder.getFilesByName(pdfBlob.getName());
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true); // Move existing file to trash
  }
  parentFolder.createFile(pdfBlob);

  const invoiceNumberMatch = file.getName().match(/Invoice\s+(BA\d+)/i);
  const subject = invoiceNumberMatch ? `Invoice ${invoiceNumberMatch[1]}` : "Invoice";

  const sheet = SpreadsheetApp.openById(fileId).getSheets()[0];
  const amount = sheet.getRange("G21").getDisplayValue();

  const issueDate = sheet.getRange("F3:H3").getDisplayValue();
  const paymentTerm = sheet.getRange("C13:D13").getDisplayValue();

  // Extract client name from cell range "F5:H5"
  const clientName = sheet.getRange("F5:H5").getDisplayValue();

  const body = `
    <p>Dear ${clientName},</p>
    <p>Please find attached invoice <b>${subject}</b>.</p>
    <p><b>Invoice Summary:</b><br>
    • Invoice Number: <b>${subject}</b><br>
    • Date Issued: ${issueDate}<br>
    • Due Date: ${paymentTerm}<br>
    • Amount Due: ${amount}</p>
    <p>All relevant details are included in the attached PDF.</p>
    <p>Please don’t hesitate to reach out if you have any questions or require further information.</p>
    <p>Many thanks.</p>
  `;

  const draft = GmailApp.createDraft(
    "",
    subject,
    "",
    {
      htmlBody: body,
      attachments: [pdfBlob]
    }
  );

  const draftId = draft.getId();
  const draftUrl = `https://mail.google.com/mail/u/0/#drafts`;

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("✅ PDF Created & Draft Email Prepared"))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextButton()
        .setText("🔄 Refresh Selection")
        .setOnClickAction(CardService.newAction().setFunctionName("onDriveItemsSelected")))
      .addWidget(CardService.newTextParagraph()
        .setText("The PDF has been generated and attached to a new Gmail draft. Please check your Drafts folder."))
      .addWidget(CardService.newTextButton()
        .setText("📬 Open Gmail Drafts")
        .setOpenLink(CardService.newOpenLink().setUrl(draftUrl))))
    .build();

  return card;
}