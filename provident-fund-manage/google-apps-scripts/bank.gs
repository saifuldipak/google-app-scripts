/*-- Bank deposits --*/
function depositBank() {
  let formName = 'Bank deposit';
  let date = formValues[2][1];
  let amount = formValues[3][1];
  let bankName = formValues[4][1];

  //check form inputs
  let msg = checkFormName(formName);
  if (msg) {
    ui.alert(msg)
    return;
  }

  if (date == '' || amount == "" || bankName == "") {
    ui.alert('You must provide date, amount and bank name');
    return
  }

  //check balance in receivable sheet
  let receivable = getBalance('Receivable');
  if (amount > receivable.balance) {
    ui.alert('Amount is greater than the balance in Receivable sheet');
    return;
  }

  let receivableCr = [date, bankName, '', '', '', '', '', '', amount];
  let bankDr = [date, bankName, 'Fund',  '', '', '', '', amount, ''];

  sheetReceivable.appendRow(receivableCr);
  sheetBank.appendRow(bankDr);

  //format date columns
  let dateColumn;
  dateColumn = sheetBank.getRange('A6:A');
  dateColumn.setNumberFormat('d-mmm-yy');
  dateColumn = sheetReceivable.getRange('A6:A');
  dateColumn.setNumberFormat('d-mmm-yy')
}

/*-- FDR issue --*/
function issueFdr() {
  let formName = 'FDR issue';
  let date = formValues[2][1];
  let amount = formValues[3][1];
  let fromBank = formValues[4][1];
  let toBank = formValues[5][1];
  let fdrNumber = formValues[6][1];
  let entriesFDR = sheetFDR.getDataRange().getValues();

  //checking form inputs
  if (submittedFormName != formName) {
    msg = `Incorrect form, Please fill up ${formName} form`;
    ui.alert(msg);
    return;
  }
  if (date == '' || amount == '' || fromBank == '' || toBank == '' || fdrNumber == '') {
    ui.alert('You must provide date, amount, from bank, to bank & fdr number');
    return;
  }

  //checking balance in bank
  let bank = getBalance('Bank', fromBank);
  if (amount > bank.balance) {
    msg = `Your balance at ${fromBank} is insufficient`;
    ui.alert(msg);
    return;
  }
  
  //checking fdr number 
  for (let i = 0; i < entriesFDR.length; i++){
    if (entriesFDR[i][1] === toBank && entriesFDR[i][2] === fdrNumber){
      ui.alert('FDR number already exists in FDR sheet');
      return;
    }
  }

  //sheet entries
  let entryBank = [date, fromBank, 'FDR', toBank, fdrNumber, '', '', '', amount];
  let entryFDR = [date, toBank, fdrNumber, fromBank, '', '', '', amount, ''];
  sheetBank.appendRow(entryBank);
  sheetFDR.appendRow(entryFDR);

  //format date columns
  let bankDateColumn = sheetBank.getRange('A6:A');
  let fdrDateColumn = sheetFDR.getRange('A6:A');
  bankDateColumn.setNumberFormat('d-mmm-yy');
  fdrDateColumn.setNumberFormat('d-mmm-yy');
}

/*-- FDR encash --*/
function encashFdr() {
  let formName = 'FDR encash';
  let profit = formValues[3][1];
  let fromBank = formValues[4][1];
  let toBank = formValues[5][1];
  let fdrNumber = formValues[6][1];
  let entriesFDR = sheetFDR.getDataRange().getValues();

  //checking form inputs
  if (submittedFormName != formName) {
    msg = `Incorrect form, Please fill up ${formName} form`;
    ui.alert(msg);
    return;
  }
  if (profit == '' || fromBank == '' || toBank == '' || fdrNumber == '') {
    ui.alert('You must provide date, profit amount, from bank, to bank & fdr number');
    return;
  }

  //checking fdr number
  for (let i = 0; i < entriesFDR.length; i++){
    if (entriesFDR[i][1] === fromBank && entriesFDR[i][2] === fdrNumber && entriesFDR[i][7]){
      var found = true;
      var fdrAmount = entriesFDR[i][7];
    }
    if (entriesFDR[i][1] === fromBank && entriesFDR[i][2] === fdrNumber && entriesFDR[i][8]){
      var encashed = true;
      var fdrAmount = entriesFDR[i][7];
    }

  }
  if (!found){
    ui.alert('FDR not found');
    return;
  }
  if (encashed){
    ui.alert('FDR already encashed');
    return;
  }
  
  //sheet entries
  let entryFDR = [date, fromBank, fdrNumber, '', '', '', '', '', fdrAmount];
  let entryProfit = [date, fromBank, fdrNumber, '', '','', '', '', profit];
  let entryBankFDR = [date, toBank, 'FDR', fromBank, fdrNumber, '', '', fdrAmount];
  let entryBankProfit = [date, toBank, 'Profit', fromBank, fdrNumber,'', '', profit];
  sheetFDR.appendRow(entryFDR);
  sheetProfit.appendRow(entryProfit);
  sheetBank.appendRow(entryBankFDR);
  sheetBank.appendRow(entryBankProfit);

  //format date columns
  let bankDateColumn = sheetBank.getRange('A6:A');
  let fdrDateColumn = sheetFDR.getRange('A6:A');
  let profitDateColumn = sheetProfit.getRange('A6:A');
  bankDateColumn.setNumberFormat('d-mmm-yy');
  fdrDateColumn.setNumberFormat('d-mmm-yy');
  profitDateColumn.setNumberFormat('d-mmm-yy');
}

/*-- Bank fund relocation --*/
function relocateBankFund(){
  let formName = 'Bank fund relocation';
  let date = formValues[2][1];
  let amount = formValues[3][1];
  let fromBank = formValues[4][1];
  let toBank = formValues[5][1];

  //checking form inputs
  let msg = checkFormName(formName) 
  if (msg) {
    ui.alert(msg);
    return;    
  }
  if (date == '' || amount == '' || fromBank == '' || toBank == '') {
    ui.alert('You must provide date, amount, from bank and to bank');
    return;
  }

  //checking balance in fromBank
  let bank = getBalance('Bank', fromBank);
  if (amount > bank.balance) {
    msg = `Insufficient fund in ${fromBank}`;
    ui.alert(msg);
    return;
  }

  //update Bank records
  let entryFromBank = [date, fromBank, 'Relocate', toBank, '', '', '', '',  amount];
  sheetBank.appendRow(entryFromBank);
  let entryToBank = [date, toBank, 'Relocate', fromBank, '', '', '', amount, ''];
  sheetBank.appendRow(entryToBank);

  //format date columns
  let bankDateColumn = sheetBank.getRange('A6:A');
  bankDateColumn.setNumberFormat('d-mmm-yy');
}
