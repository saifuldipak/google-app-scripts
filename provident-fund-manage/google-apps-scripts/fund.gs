/*-- Genarate statement --*/
function generateStatement() {
  let formName = 'Employee PF statement';
  let fund = sheetFund.getDataRange().getValues();
  let loan = sheetLoan.getDataRange().getValues();
  let employeeName = sheetForm.getRange("B3").getValue();

  //checking form inputs
  let msg = checkFormName(formName);
  if (msg) {
    ui.alert(smg);
    return;
  }
  if (submittedFormName != formName) {
    msg = `Incorrect form. Please fill up "${formName}" form.`;
    ui.alert(msg);
    return;
  }
  if(employeeName == ''){
    ui.alert('You must provide employee name');
    return;
  }

  //get PF joining date
  let participants = sheetParticipants.getDataRange().getValues();
  for (i = 0; i < participants.length; i++) {
    if (participants[i][0] == employeeName) {
      var pfJoinDate = participants[i][1];
      var pfLeaveDate = '';
      if (participants[i][2]) {
        pfLeaveDate = participants[i][2];
      }
    }
  }
  
  //clear report sheet
  let rangeReport = sheetReport.getRange('A1:N200');
  rangeReport.clear()

  //creating headline and table headers
  sheetReport.appendRow(['PF Statement']);
  sheetReport.appendRow([employeeName]);
  sheetReport.appendRow(['PF join date']);
  sheetReport.appendRow(['PF leave date']);
  sheetReport.appendRow([' ']);   
  sheetReport.appendRow(['Balance']);
  sheetReport.appendRow(['Own Contribution']);
  sheetReport.appendRow(['Company Contribution']);
  sheetReport.appendRow(['Profit']);
  sheetReport.appendRow(['Loan']);
  
  sheetReport.getRange('B3').setValue(pfJoinDate);
  sheetReport.getRange('B4').setValue(pfLeaveDate);
  sheetReport.getRange('B6').setFormula('=SUM(E11:E) - SUM(D11:D)');
  sheetReport.getRange('B7').setFormula('=SUMIF(B11:B, "Own", E11:E)');
  sheetReport.getRange('B8').setFormula('=SUMIF(B11:B, "Company", E11:E)');
  sheetReport.getRange('B9').setFormula('=ROUND(SUMIF(B11:B, "Profit", E11:E))');
  sheetReport.getRange('B10').setFormula('=SUMIF(B11:B, "Loan", D11:D)-SUMIF(B11:B, "Repay", E11:E)');
  sheetReport.appendRow(['Date', 'Month', 'Year', 'Dr', 'Cr']);

  //Inserting own and company contribution data
  for(i = 0; i < fund.length; i++) {
    if(fund[i][1] == employeeName) {
      if (fund[i][2] == 'Paid') {
        var datePaid = fund[i][0];
        var paid = fund[i][7];
        continue;
      }
      if (fund[i][2] == 'Lapses') {
        var dateLapses = fund[i][0];
        var lapses = fund[i][7];
        continue;
      }

      let monthName = '';
      if (fund[i][3]) {
        monthName = month[fund[i][3].getMonth()];
      }
      
      entry = [fund[i][0], fund[i][2], monthName, fund[i][7], fund[i][8]];
      if(fund[i][3] == 'Profit'){
        entry[1] = ' ';
      }
      sheetReport.appendRow(entry);
    }
  }

  //Inserting loan data 
  for(i = 0; i < loan.length; i++) {
    if(loan[i][1] == employeeName) {
      entry = [loan[i][0], loan[i][2], '', loan[i][7], loan[i][8]];
      sheetReport.appendRow(entry);
    }
  }

  if (paid) {
    sheetReport.appendRow([datePaid, 'Paid', '', paid]);
  }
  if (lapses) {
    sheetReport.appendRow([dateLapses, 'Lapses', '', lapses]);
  }

  //format date column
  sheetReport.getRange('A10:A').setNumberFormat('d-mmm-yy');
  sheetReport.getRange('B3').setNumberFormat('d-mmm-yy');
  sheetReport.getRange('B4').setNumberFormat('d-mmm-yy');

  //formatting name
  let styleName = SpreadsheetApp.newTextStyle()
                            .setFontSize(14)
                            .build();
  let cellName = sheetReport.getRange('A2');
  cellName.setTextStyle(styleName);
  cellName.setBorder(false, false, false, false, false, false);

  //formatting table header
  let styleTableHeader = SpreadsheetApp.newTextStyle()
                            .setFontSize(11)
                            .setBold(true)
                            .build();
  let rangeTableHeader = sheetReport.getRange('A11:E11');
  rangeTableHeader.setTextStyle(styleTableHeader);
  rangeTableHeader.setBackground('#eeeeee');
  rangeTableHeader.setBorder(true, false, true, false, false, false)
}

/*-- Monthly fund disburse --*/
function disburseFund() {
  let date = formValues[2][1];
  let month = formValues[3][1];
  let year = formValues[4][1];
  let employees = sheetEmployees.getDataRange().getValues();

  //check form inputs
  if (date == '' || month == "" || year == "") {
    msg = "Must give date, month and year";
    ui.alert(msg);
    return;
  }
  
  //check whether submitted month & year exists 
  let fund = sheetFund.getDataRange().getValues();
  for (let i = 0; i < fund.length; i++){
    if (typeof(fund[i][3]) === "object") {
      if (fund[i][3].getUTCMonth() + 1 == month && fund[i][3].getUTCFullYear() == year){
        msg = `Fund already disbursed for ${month}, ${year}`
        ui.alert(msg);
        return;
      }
    }
  }

  //check whether employees in "MonthlyContribution" sheet exists in "Participants" sheet
  let monthlyContribution = sheetMonthlyContribution.getDataRange().getValues();
  let found;
  for (let i = 0; i < monthlyContribution.length; i++) {
    if (monthlyContribution[i][0] && monthlyContribution[i][1] > 0) {
      found = false;
      for (let j = 0; j < employees.length; j++) {
        if (monthlyContribution[i][0] == employees[j][0] && employees[j][1] && !employees[j][2]) {
          found = true;
          break;
        }
      }
      if (found == false) {
        msg = `Employee ${monthlyContribution[i][0]} not found in Participants sheet`;
        ui.alert(msg);
        return;
      }
    }
  }

  //update Fund, Loan & Receivable sheet
  let monthlyReceivable = 0;
  let fundContributionDate = new Date(year, month - 1, 1, 12);
  for (let j = 3; j < monthlyContribution.length; j++) {
    let ownContributionCr = [date, monthlyContribution[j][0], 'Own', fundContributionDate, '', '', '', '', monthlyContribution[j][1]];
    let companyContributionCr = [date, monthlyContribution[j][0], 'Company', fundContributionDate, '', '', '', '', monthlyContribution[j][1]];
    let loanRepayCr = [date, monthlyContribution[j][0], 'Repay', '', '', '', '', '', monthlyContribution[j][2]];
    
    sheetFund.appendRow(ownContributionCr);
    sheetFund.appendRow(companyContributionCr);

    let loanRepay = 0;
    if (monthlyContribution[j][2]){
      loanRepay = monthlyContribution[j][2];
      sheetLoan.appendRow(loanRepayCr);
    }

    monthlyReceivable += monthlyContribution[j][1] * 2;
    monthlyReceivable += loanRepay;    
  }

  let receivableDr = [date, 'Fund', fundContributionDate, '', '', '', '', monthlyReceivable];
  sheetReceivable.appendRow(receivableDr);

  //formate date columns
  let dateColumn;
  dateColumn = sheetReceivable.getRange('A5:A');
  dateColumn.setNumberFormat('d-mmm-yy');
  dateColumn = sheetReceivable.getRange('C5:C');
  dateColumn.setNumberFormat('d-mmm-yy')
  dateColumn = sheetFund.getRange('A5:A');
  dateColumn.setNumberFormat('d-mmm-yy');
  dateColumn = sheetFund.getRange('D5:D');
  dateColumn.setNumberFormat('d-mmm-yy');
}

/*-- Loan disbursement --*/
function disburseLoan() {
  let formName = 'Loan from PF';
  let date = formValues[2][1];
  let amount = formValues[3][1];
  let loanReceiver = formValues[4][1];
  let fromBank = formValues[5][1];

  //check form inputs
  let msg = checkFormName(formName);
  if (msg) {
    ui.alert(msg);
    return;
  }

  if (date == '' || amount == '' || loanReceiver == '' || fromBank == '') {
    ui.alert('You must provide date, amount, loan receiver & bank name');
    return;
  }

  //checking employee loan status
  let loan = getBalance('Loan', loanReceiver);
  if (loan.balance > 0) {
    msg = `${loanReceiver} has a loan`;
    ui.alert(msg);
    return;
  }
  
  //checking employee fund amount
  let ownFund = getBalance('Fund', loanReceiver, 'Own');
  if(amount > (ownFund.credit * 0.8)){
    ui.alert('Loan amount is higher than 80% of own contribution');
    return;
  }

  //checking balance in bank
  let bank = getBalance('Bank', fromBank);
  if (amount > bank.balance) {
    msg = `Your balance at ${fromBank} is insufficient`;
    ui.alert(msg);
    return;
  }
  
  let entryBank = [date, fromBank, 'Loan', loanReceiver, '', '', '', '', amount];
  let entryLoan = [date, loanReceiver, 'Loan', '', '', '', '', amount, ''];
  sheetBank.appendRow(entryBank);
  sheetLoan.appendRow(entryLoan);
  
  //format date columns
  let dateColumn = sheetBank.getRange('A6:A');
  dateColumn.setNumberFormat('d-mmm-yy')
  dateColumn = sheetLoan.getRange('A6:A');
  dateColumn.setNumberFormat('d-mmm-yy')
}

/*-- Profit distribution --*/
function distributeProfit(){
  let fromDate = formValues[2][1];
  let toDate = formValues[3][1];
  let employees = sheetEmployees.getDataRange().getValues();
  let funds = sheetFund.getDataRange().getValues();
  let profits = sheetProfit.getDataRange().getValues();
  let totalFund = sheetFund.getRange('I4').getValue();

  //checking form inputs
  if (fromDate == '' || toDate == '') {
    ui.alert('You must provide from date and to date')
    return;
  }

  //calculate fund of each employee of specified period
  let employeeFunds = [];
  let dr;
  let cr;
  for(let i = 0; i < employees.length; i++) {
    if (employees[i][2] == 'A'){
      dr = 0;
      cr = 0;
      for (let j = 0; j < funds.length; j++) {
        if (funds[j][1] == employees[i][0]) {
          if (funds[j][0] >= fromDate && funds[j][0] <= toDate) {
            if (funds[j][8]) {
              cr += funds[j][8];
            }
            if (funds[j][7]) {
              dr += funds[j][7];
            }
          }
        }
      }
    
      let fund = cr - dr;
      let profitShare = 0;
      let employeeFund = [employees[i][0], fund, profitShare];
      employeeFunds.push(employeeFund);
    }
  }

  //calculate profit of specified period and update profit sheet
  let totalProfit = 0;
  let profitDisbursed;
  for(let i = 0; i < profits.length; i++){
    if(profits[i][0] >= fromDate && profits[i][0] <= toDate){ //get values within specified dates
      if(profits[i][8]){
        profitDisbursed = false;
        for(let j = 0; j < profits.length; j++){ //checking profit disbursed or not
          if(profits[j][1] == profits[i][1] && profits[j][2] == profits[i][2] && profits[j][7] == profits[i][8]){
            profitDisbursed = true;
          }
        }
      
        if(profitDisbursed == false){
          totalProfit += profits[i][8];
          let entryProfit = [currentDateTime, profits[i][1], profits[i][2], '', '', '', '', profits[i][8], ''];
          sheetProfit.appendRow(entryProfit);     
        }
      }
    }
  }

  if(totalProfit == 0){
    ui.alert('No profit found for the specified period');
    return;
  }

  //distribute profit to each employee
  for(let i=0; i < employeeFunds.length; i++){
    employeeFunds[i][2] = (employeeFunds[i][1] / totalFund) * totalProfit;
    let entryFund = [currentDateTime, employeeFunds[i][0], currentDateTime, 'Profit', '', '', '', '', employeeFunds[i][2]];
    sheetFund.appendRow(entryFund);
  }

  //format date columns
  let dateColumn;
  dateColumn = sheetFund.getRange('A6:A');
  dateColumn.setNumberFormat('d-mmm-yy');
  dateColumn = sheetProfit.getRange('A6:A');
  dateColumn.setNumberFormat('d-mmm-yy');
}

/*-- Close employee PF --*/
function closeFund() {
  let formName = 'Employee PF closure';
  let date = formValues[2][1];
  let employeeName = formValues[3][1];
  let fromBank = formValues[4][1];

  //checking form inputs
  if (submittedFormName != formName) {
    msg = `Incorrect form. Please fill up "${formName}" form`;
    ui.alert(msg);
    return;
  }
  if (date == '' || employeeName == '' || fromBank == '') {
    ui.alert('You must provide date, employee name & from bank');
    return;
  }

  //checking employee exists
  let participants = sheetParticipants.getDataRange().getValues();
  found = false;
  for (let i = 0; i < participants.length; i++) {
    if (participants[i][0] == employeeName && participants[i][2] == '') {
      var fundDurationDays = (currentDateTime.getTime() - participants[i][1].getTime()) / (1000*3600*24);
      found = true;
      foundRow = i;
      break;
    }
  }

  if (!found) {
    msg = `Employee "${employeeName}" not found or account already closed`;
    ui.alert(msg);
    return;
  }

  //getting balance of Fund and Bank
  let bank = getBalance('Bank', fromBank);
  let fund = getBalance('Fund', employeeName);
  let ownFund = getBalance('Fund', employeeName, 'Own');
  let companyFund = getBalance('Fund', employeeName, 'Company');
  let profit = getBalance('Fund', employeeName, 'Profit');
  let loan = getBalance('Loan', employeeName);

  let fundBalance = fund.balance - loan.balance;
  if (fundBalance == 0) {
    msg = `"${employeeName}" PF account already closed`;
    ui.alert(msg);
    return;
  }

  //caculating employee fund based on the fund joining duration
  let companyContribution = companyFund.credit;
  let companyProfit = profit.credit / 2;
  let ownProfit = profit.credit / 2;
  
  if (fundDurationDays < (365 * 3)) {
    companyContribution = 0;
    companyProfit = 0;
  }
  else if (fundDurationDays > (365 * 3) && fundDurationDays < (365 * 5)) {
    companyContribution = companyContribution / 2;
    companyProfit = companyProfit / 2;
  }
  
  let fundPayable = (ownFund.balance - loan.balance) + companyContribution + ownProfit  + companyProfit;
  if (fundPayable > bankBalance) {
    msg = `Not sufficient fund in ${fromBank}`;
    ui.alert(msg);
    return;
  }

  //update bank, fund, lapses & participants table
  let entryBank = [date, fromBank, 'Paid', employeeName, '', '', '', '', fundPayable];
  sheetBank.appendRow(entryBank);
  if (Math.round(fundPayable) == Math.round(fundBalance)) {
    let entryFund = [date, employeeName, 'Paid', '', '', '', '', fundBalance, ''];
    sheetFund.appendRow(entryFund);
  }
  else {
    let lapsesBalance = fundBalance - fundPayable;
    let fundEntry1 = [date, employeeName, 'Paid', '', '', '', '', fundPayable, ''];
    sheetFund.appendRow(fundEntry1);
    let fundEntry2 = [date, employeeName, 'Lapses', '', '', '', '', lapsesBalance, ''];
    sheetFund.appendRow(fundEntry2);
    let lapsesEntry = [date, employeeName, '', '', '', '', '', '', lapsesBalance];
    sheetLapses.appendRow(lapsesEntry);
  }
  
  sheetParticipants.getRange(foundRow + 1, 3).setValue(date);
  formatDateAndNumbers(ss, 'Fund');
  formatDateAndNumbers(ss, 'Bank');
}

