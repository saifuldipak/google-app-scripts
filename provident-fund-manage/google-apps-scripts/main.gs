const ss = SpreadsheetApp.getActive();
const sheetEmployees = ss.getSheetByName('Participants');
const sheetReceivable = ss.getSheetByName('Receivable');
const sheetFund = ss.getSheetByName('Fund');
const sheetBank = ss.getSheetByName('Bank');
const sheetForms = ss.getSheetByName('Forms');
const sheetReport = ss.getSheetByName('Report');
const sheetFDR = ss.getSheetByName('FDR');
const sheetProfit = ss.getSheetByName('Profit');
const sheetForm = ss.getSheetByName('Form');
const sheetMonthlyContribution = ss.getSheetByName('MonthlyContribution');
const sheetLoan = ss.getSheetByName('Loan');
const sheetItems = ss.getSheetByName('Items');
const sheetParticipants = ss.getSheetByName('Participants');
const sheetLapses = ss.getSheetByName('Lapses');
const currentDateTime = new Date();
const ui = SpreadsheetApp.getUi();
const month = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const formValues = sheetForm.getRange('A1:B10').getValues();
const submittedFormName = formValues[1][0];
const date = formValues[2][1];

/*-- Format date and numbers columns --*/
function formatDateAndNumbers(spreadSheetObject, sheetName) {
  let sheet = spreadSheetObject.getSheetByName(sheetName);
  sheet.getRange('A5:A').setNumberFormat('d-mmm-yy');
  sheet.getRange('I5:I').setNumberFormat('0,000.00');
  sheet.getRange('H5:H').setNumberFormat('0,000.00');
}

/* Check balance */
function checkBalance(sheet, name) {
  let entries;
  if (sheet == 'Bank') {
    entries = sheetBank.getDataRange().getValues();
  }
  else if (sheet == 'Fund') {
    entries = sheetFund.getDataRange().getValues();
  }

  let drSum = 0;
  let crSum = 0;
  for (let i = 0; i < entries.length; i++) {
    if (entries[i][1] == name) {
      if (entries[i][5]) {
        drSum += entries[i][5];
      }
      if (entries[i][6]) {
        crSum += entries[i][6];
      }
    }
  }

  let balance;
  if (sheet == 'Bank') {
    balance = drSum - crSum;
  }
  else if (sheet == 'Fund') {
    balance = crSum - drSum;
  }

  return balance;
}

/*-- Get debit & credit balance --*/
function getBalance(sheet, name = null, type = null) {
  if (sheet == '' || name == '') {
    ui.alert('Internal error, please check console log');
    Logger.log('getDebitCredit(): Sheet and/or name not given');
    return;
  }
  let entries = ss.getSheetByName(sheet).getDataRange().getValues();
  let debit = 0;
  let credit = 0;
  for (let i = 4; i < entries.length; i++) {
    if (name == null) {
       if (entries[i][7] && typeof(entries[i][7]) == 'number') {
          debit += entries[i][7];
        }
      if (entries[i][8] && typeof(entries[i][8]) == 'number') {
        credit += entries[i][8];
      }
    }
    else {
      if (entries[i][1] == name) {
        if (type != null) {
          if (entries[i][7] && entries[i][2] == type) {
            debit += entries[i][7];
          }
          if (entries[i][8] && entries[i][2] == type) {
            credit += entries[i][8];
          }
        }  
        else {
          if (entries[i][7]) {
            debit += entries[i][7];
          }
          if (entries[i][8]) {
            credit += entries[i][8];
          }
        }
      }
    }
  }

  let balance;
  if (debit > credit) {
    balance = debit - credit;
  }
  else {
    balance = credit - debit;
  }

  return {debit, credit, balance};
}


/*-- Create new PF spreadsheet --*/
function createNewSpreadsheet() {
  let formName = 'New PF spreadsheet create';
  let date = formValues[2][1];
  let newSpreadsheetName = formValues[3][1];

  /*//check form inputs
  if (submittedFormName != formName) {
    msg = `Incorrect form. Please fill up "${formName}" form.`;
    ui.alert(msg);
    return;
  }
  if (date == '' || newSpreadsheetName == '') {
    msg = 'You must provide date and spreadsheet name';
    ui.alert(msg);
    return;
  }

  //check new spreadsheet name exists or not
  let openSpreadsheetId = ss.getId();
  let folders = DriveApp.getFileById(openSpreadsheetId).getParents();
  while (folders.hasNext()) {
    var folder= folders.next();
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      if (file == newSpreadsheetName) {
        msg = `File "${file}" exists, choose another name`;
        ui.alert(msg);
        return;
      }
    }
  }
  
  //create new spreadsheet and clear contents of specific sheets
  let newSheetId = DriveApp.getFileById(openSpreadsheetId).makeCopy(newSpreadsheetName).getId();
  let ssNew = SpreadsheetApp.openById(newSheetId);*/
  
  let ssNew = SpreadsheetApp.openById('1afQ4JCS6uOOxlccw5-5m1HuW2LaHFCdMtSesHIIHVqE');
  let ranges = ['A5:A', 'B5:B', 'C5:C', 'D5:D', 'E5:E', 'F5:F', 'G5:G', 'H5:H', 'I5:I'];
  let sheets = ['Participants', 'MonthlyContribution', 'Receivable', 'Fund', 'Bank', 'Loan', 'FDR', 'Profit', 'Lapses'];

  for (let i = 0; i < sheets.length; i++) {
    for (let j = 0; j < ranges.length; j++) {
      ssNew.getSheetByName(sheets[i]).getRange(ranges[j]).clear();
    }
  }

  //Inserting data in the new spreadsheet Participants sheet
  let sheetParticipantsNew = ssNew.getSheetByName('Participants');
  let employees = sheetParticipants.getRange('A5:C200').getValues();
  for (let i = 0; i < employees.length; i++) {
    if (employees[i][0] && !employees[i][2]) {
      sheetParticipantsNew.appendRow([employees[i][0], employees[i][1]]);
    }
  }
  sheetParticipantsNew.getRange('B5:B').setNumberFormat('d-mmm-yy');

  //Inserting data in the new spreadsheet "Receivable" sheet
  let receivable = getBalance('Receivable');
  if (receivable.balance > 0) {
    let sheetReceivableNew = ssNew.getSheetByName('Receivable');
    sheetReceivableNew.appendRow([date, 'Opening balance', '', '', '', '', '', receivable.balance, '']);
  }
  formatDateAndNumbers(ssNew, 'Receivable');

  //Inserting opening balance in the new spreadsheet "Fund" and "Loan" sheet
  let sheetFundNew = ssNew.getSheetByName('Fund');
  let sheetLoanNew = ssNew.getSheetByName('Loan');
  for (let i = 0; i < employees.length; i++) {
    if (employees[i][0] && !employees[i][2]) {
      let ownFund = getBalance('Fund', employees[i][0], 'Own');
      let companyFund = getBalance('Fund', employees[i][0], 'Company');
      let loan = getBalance('Loan', employees[i][0]);
      if (loan.balance > 0) {
        sheetLoanNew.appendRow([date, employees[i][0], '', '', '', '', '', loan.balance, '']);
      }
      sheetFundNew.appendRow([date, employees[i][0], 'Own', '', '', '', '', '', ownFund.credit]);
      sheetFundNew.appendRow([date, employees[i][0], 'Company', '', '', '', '', '', companyFund.credit]);
    }
  }
  formatDateAndNumbers(ssNew, 'Fund');
  formatDateAndNumbers(ssNew, 'Loan');

  //Inserting opening balance in the new spreadsheet "Bank" & "FDR" sheet
  let items = ssNew.getSheetByName('Items').getDataRange().getValues();
  for (let i = 0; i < items.length; i++) {
    let bankName = items[i][2];
    if (bankName) {
      let bank = getBalance('Bank', bankName);
      let fdr = getBalance('FDR', bankName);
      if (bank.balance > 0){
        ssNew.getSheetByName('Bank').appendRow([date, bankName, '', '', '', '', '', bank.balance, '']);
      }
      if (fdr.balance > 0) {
        ssNew.getSheetByName('FDR').appendRow([date, bankName, '', '', '', '', '', fdr.balance, '']);
      }
    }
  }

  formatDateAndNumbers(ssNew, 'Bank');
  formatDateAndNumbers(ssNew, 'FDR');

  //Inserting opening balance in the new "Profit" & "Lapses" sheet
  let profit = getBalance('Profit');
  let lapses = getBalance('Lapses');
  if (profit.balance > 0) {
    ssNew.getSheetByName('Profit').appendRow([date, '', '', '', '', '', '', '', profit.balance]);
  }
  if (lapses.balance > 0) {
    ssNew.getSheetByName('Lapses').appendRow([date, '', '', '', '', '', '', '', lapses.balance]);
  }
  formatDateAndNumbers(ssNew, 'Profit');
  formatDateAndNumbers(ssNew, 'Lapses');
}

