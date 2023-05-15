/*-- Clear form --*/
function clearForm() {
  let formRange = sheetForm.getRange('A1:N50');
  formRange.clear();
  formRange.setDataValidation(null);
  formRange.setNote(null);
}

/*-- Check form name --*/
function checkFormName(formName) {
  if (submittedFormName != formName) {
    msg = `Incorrect form, Please fill up ${formName} form`;
    return msg;
  }
  return false;
}

/*-- Format text --*/
function formatText() {
  let textStyle = SpreadsheetApp.newTextStyle()
                            .setFontSize(14)
                            .build();
  let cellName = sheetForm.getRange('A2:B2');
  cellName.setTextStyle(textStyle);

  styleName = SpreadsheetApp.newTextStyle()
                            .setFontSize(12)
                            .build();
  cellName = sheetForm.getRange('A3:A10');
  cellName.setTextStyle(styleName);
}

/*-- Format date cell --*/
function formatDateCell(cellRange) {
  note = 'Double click to insert a date';
  sheetForm.getRange(cellRange).setNote(note);
  cell = sheetForm.getRange(cellRange);
  rule = SpreadsheetApp.newDataValidation().requireDate().build();
  cell.setDataValidation(rule);
}

/*-- Create drop down list --*/
function createList(listName, cellRange) {
  let listItems;
  if (listName == 'banks') {
    listItems = sheetItems.getRange('C1:C');
  }
  else if (listName == 'employees') {
    listItems = sheetParticipants.getRange('A6:A');
  }
  else if (listName == 'months') {
    listItems = sheetItems.getRange('A1:A');
  }
  else if (listName == 'years') {
    listItems = sheetItems.getRange('B1:B');
  }
  
  let rule = SpreadsheetApp.newDataValidation().requireValueInRange(listItems).build();
  cells = sheetForm.getRange(cellRange);
  cells.setDataValidation(rule);
}

/*-- Employee PF statement --*/
function createStatementForm() {
  clearForm();
  sheetForm.appendRow([" "]);
  sheetForm.appendRow(['Employee PF statement']);
  sheetForm.appendRow(['Employee Name']);

  let cell = sheetForm.getRange('B3');
  let range = sheetParticipants.getRange('A6:A');
  let rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
  cell.setDataValidation(rule);

  formatText();
  sheetForm.getRange("A2:B2").setBackground("#c1e5b4");
}

/*-- Disburse fund --*/
function createFundDisburseForm() {
  clearForm();
  sheetForm.appendRow([' ']);
  sheetForm.appendRow(['Disburse Fund']);
  sheetForm.appendRow(['Date']);
  sheetForm.appendRow(['PF of Month']);
  sheetForm.appendRow(['PF of Year']);

  cell = sheetForm.getRange('B4');
  range = sheetItems.getRange('A1:A');
  rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
  cell.setDataValidation(rule);
  
  cell = sheetForm.getRange('B5');
  range = sheetItems.getRange('B1:B');
  rule = SpreadsheetApp.newDataValidation().requireValueInRange(range).build();
  cell.setDataValidation(rule);

  formatText();
  sheetForm.getRange('A2:B2').setBackground("#c0d5ff");
  formatDateCell('B3');
  createList('months', 'B4');
  createList('years', 'B5');
}

/*-- Bank deposit --*/
function createBankDepositForm() {
  clearForm();
  sheetForm.appendRow([" "]);
  sheetForm.appendRow(["Bank deposit"]);
  sheetForm.appendRow(['Date']);
  sheetForm.appendRow(["Amount"]);
  sheetForm.appendRow(["Bank Name"]);
  sheetForm.getRange('A2:B2').setBackground("#ffcccc");
  formatText();
  formatDateCell('B3');
  createList('banks', 'B5');
}

/*-- Bank fund relocation --*/
function createBankFundRelocationForm() {
  clearForm();
  sheetForm.appendRow([" "]);
  sheetForm.appendRow(["Bank fund relocation"]);
  sheetForm.appendRow(['Date']);
  sheetForm.appendRow(['Amount']);
  sheetForm.appendRow(['From Bank']);
  sheetForm.appendRow(['To Bank']);
  sheetForm.getRange('A2:B2').setBackground("#bf86a2");
  let textStyle = SpreadsheetApp.newTextStyle().setForegroundColor("#ffffff").build();
  sheetForm.getRange('A2:B2').setTextStyle(textStyle);
  formatText();
  formatDateCell('B3');
  createList('banks', 'B5:B6');
}

/*-- Loan from PF --*/
function createLoanForm() {
  clearForm();
  sheetForm.appendRow([' ']);
  sheetForm.appendRow(['Loan from PF']);
  sheetForm.appendRow(['Date']);
  sheetForm.appendRow(['Amount']);
  sheetForm.appendRow(['Loan receiver']);
  sheetForm.appendRow(['Disurse from Bank']);
  formatText();
  sheetForm.getRange('A2:B2').setBackground("#ffe599");
  formatDateCell('B3');
  createList('employees', 'B5');
  createList('banks', 'B6');
}

/*-- FDR issue --*/
function createFdrIssueForm() {
  clearForm();
  sheetForm.appendRow([' ']);
  sheetForm.appendRow(['FDR issue']);
  sheetForm.appendRow(['Date']);
  sheetForm.appendRow(['FDR amount']);
  sheetForm.appendRow(['From bank']);
  sheetForm.appendRow(['To bank']);
  sheetForm.appendRow(['FDR number']);

  formatDateCell('B3');
  let bankNames = sheetItems.getRange('C1:C');
  let rule = SpreadsheetApp.newDataValidation().requireValueInRange(bankNames).build();
  cellRange = sheetForm.getRange('B5:B6');
  cellRange.setDataValidation(rule);
  formatText();
  sheetForm.getRange('A2:B2').setBackground("#9ee3f2");
}

/*-- FDR encash --*/
function createFdrEncashForm() {
  clearForm();
  sheetForm.appendRow([' ']);
  sheetForm.appendRow(['FDR encash']);
  sheetForm.appendRow(['Date']);
  sheetForm.appendRow(['Profit amount']);
  sheetForm.appendRow(['From bank']);
  sheetForm.appendRow(['To bank']);
  sheetForm.appendRow(['FDR number']);

  formatDateCell('B3');
  let bankNames = sheetItems.getRange('C1:C');
  let rule = SpreadsheetApp.newDataValidation().requireValueInRange(bankNames).build();
  cellRange = sheetForm.getRange('B5:B6');
  cellRange.setDataValidation(rule);
  formatText();
  sheetForm.getRange('A2:B2').setBackground("#a3f082");
}

/*-- Profit distribution --*/
function createProfitDistributionForm() {
  clearForm();
  sheetForm.appendRow([' ']);
  sheetForm.appendRow(['Profit distribution']);
  sheetForm.appendRow(['From date']);
  sheetForm.appendRow(['To date']);

  formatDateCell('B3:B4');
  formatText();
  sheetForm.getRange('A2:B2').setBackground("#c6b4f4");
}

/*-- Employee PF closure --*/
function createFundCloseForm() {
  clearForm();
  sheetForm.appendRow([' ']);
  sheetForm.appendRow(['Employee PF closure']);
  sheetForm.appendRow(['Date']);
  sheetForm.appendRow(['Employee name']);
  sheetForm.appendRow(['Pay from bank']);

  sheetForm.getRange('A2:B2').setBackground("#f9cb9c");
  formatText();
  formatDateCell('B3');
  createList('employees', 'B4');
  createList('banks', 'B5');
}

/*-- New PF spreadsheet create --*/
function createNewSpreadsheetForm() {
  clearForm();
  sheetForm.appendRow([' ']);
  sheetForm.appendRow(['New PF spreadsheet create']);
  sheetForm.appendRow(['Account start date']);
  sheetForm.appendRow(['New spreadsheet name']);

  sheetForm.getRange('A2:B2').setBackground("#abe9de");
  formatText();
  formatDateCell('B3');
}

