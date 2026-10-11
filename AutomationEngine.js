/**
 * =========================================================================
 * OPERATIONS & INVOICING AUTOMATION ENGINE
 * Author: Shubham S. | Upwork Automation Showcase
 * Features: Dynamic PDF Export, Drive Archiving, Gmail API Dispatch & UI Menu
 * =========================================================================
 */

// 1. Google Sheet open hone par Custom Menu create karna
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ Automation Suite')
    .addItem('📄 Generate PDF for Selected Row', 'runGeneratePDFOnly')
    .addItem('📧 Generate & Email Invoice to Client', 'runGenerateAndEmail')
    .addSeparator()
    .addItem('🔄 Refresh Dashboard Metrics', 'refreshDashboard')
    .addToUi();
}

function runGeneratePDFOnly() {
  processInvoiceAutomation(false);
}

function runGenerateAndEmail() {
  processInvoiceAutomation(true);
}

function processInvoiceAutomation(sendEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const tasksSheet = ss.getSheetByName('📋 Orders_Tasks');
  const templateSheet = ss.getSheetByName('📄 Invoice_Template');
  const settingsSheet = ss.getSheetByName('⚙️ Settings');

  const activeCell = tasksSheet.getActiveCell();
  const rowIndex = activeCell.getRow();

  if (rowIndex < 2) {
    ui.alert('⚠️ Selection Error', 'Kripya "📋 Orders_Tasks" sheet me kisi order row (Row 2 se neeche) ko select karein.', ui.ButtonSet.OK);
    return;
  }

  const rowData = tasksSheet.getRange(rowIndex, 1, 1, 10).getValues()[0];
  const invoiceId = rowData[0];
  const clientName = rowData[1];
  const clientEmail = rowData[2];
  const serviceDesc = rowData[3];
  const hoursQty = rowData[4];
  const rate = rowData[5];
  const totalAmount = rowData[6];

  if (!invoiceId || !clientName) {
    ui.alert('⚠️ Missing Data', 'Selected row me Invoice ID ya Client Name missing hai.', ui.ButtonSet.OK);
    return;
  }

  const settingsData = settingsSheet.getRange("B4:C10").getValues();
  const config = {};
  settingsData.forEach(row => { config[row[0]] = row[1]; });

  const folderName = config['Google Drive Output Folder'] || 'Automated_Invoices_Archive';
  const companyName = config['Company Name'] || 'Nexus Automation Labs Inc.';

  ss.toast('Invoice data template me inject ho raha hai...', 'Processing', 3);

  const todayFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy');
  templateSheet.getRange('E3').setValue(`Invoice No: ${invoiceId}\nDate: ${todayFormatted}\nDue: Net 15 Days`);
  templateSheet.getRange('B7').setValue(`${clientName}\nAttn: Accounts Payable\n${clientEmail}`);
  templateSheet.getRange('B11').setValue(serviceDesc);
  templateSheet.getRange('C11').setValue(hoursQty);
  templateSheet.getRange('D11').setValue(typeof rate === 'number' ? `$${rate.toFixed(2)}` : rate);
  templateSheet.getRange('E11').setValue(typeof totalAmount === 'number' ? `$${totalAmount.toFixed(2)}` : totalAmount);
  templateSheet.getRange('E16').setValue(typeof totalAmount === 'number' ? `$${totalAmount.toFixed(2)}` : totalAmount);
  templateSheet.getRange('E18').setValue(typeof totalAmount === 'number' ? `$${totalAmount.toFixed(2)}` : totalAmount);

  SpreadsheetApp.flush();
  Utilities.sleep(1000);

  const folder = getOrCreateFolder(folderName);
  const pdfBlob = createSheetPDFBlob(ss, templateSheet, `${invoiceId}_${clientName.replace(/\s+/g, '_')}.pdf`);
  const createdFile = folder.createFile(pdfBlob);

  let emailDispatched = false;
  if (sendEmail) {
    if (clientEmail && clientEmail.includes('@')) {
      const subject = `Invoice ${invoiceId} - ${companyName}`;
      const body = `Hi ${clientName},\n\nPlease find attached the official invoice (${invoiceId}) for recent automation and technical deliverables.\n\nSummary:\n- Invoice ID: ${invoiceId}\n- Deliverable: ${serviceDesc}\n- Total Due: $${totalAmount}\n\nThank you for your business.\n\nBest regards,\nAccounts & Automation Team\n${companyName}`;

      GmailApp.sendEmail(clientEmail, subject, body, {
        attachments: [pdfBlob],
        name: companyName
      });
      emailDispatched = true;
    }
  }

  const nowStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
  tasksSheet.getRange(rowIndex, 9).setValue(emailDispatched ? 'Dispatched' : 'Generated');
  tasksSheet.getRange(rowIndex, 10).setValue(nowStamp);

  ui.alert(
    '✅ Workflow Complete!',
    `Invoice: ${invoiceId}\nClient: ${clientName}\nStatus: ${emailDispatched ? 'PDF Generated & Emailed' : 'PDF Saved in Drive'}\nFolder: ${folderName}\nFile Name: ${createdFile.getName()}`,
    ui.ButtonSet.OK
  );
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function createSheetPDFBlob(spreadsheet, sheet, fileName) {
  const spreadsheetId = spreadsheet.getId();
  const sheetId = sheet.getSheetId();
  
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
    `format=pdf&size=letter&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&fzr=false&gid=${sheetId}`;

  const params = {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, params);
  return response.getBlob().setName(fileName);
}

function refreshDashboard() {
  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard metrics synched successfully!', 'Refreshed', 2);
}

function onEdit(e) {
  if (!e) return;
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== '📋 Orders_Tasks') return;

  const range = e.range;
  const col = range.getColumn();
  const row = range.getRow();

  if (col === 8 && row > 1) {
    const timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    sheet.getRange(row, 10).setValue(timeStamp);
  }
}
