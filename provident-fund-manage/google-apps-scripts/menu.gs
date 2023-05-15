/*-- Create menu --*/
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Manage PF')
      .addSubMenu(ui.createMenu('Forms')
        .addItem('Employee PF statement', 'createStatementForm')
        .addItem('Monthly PF disburse', 'createFundDisburseForm')
        .addItem('Bank deposit', 'createBankDepositForm')
        .addItem('Bank fund relocation', 'createBankFundRelocationForm')
        .addItem('Loan from PF', 'createLoanForm')
        .addItem('FDR issue', 'createFdrIssueForm')
        .addItem('FDR encash', 'createFdrEncashForm')
        .addItem('Profit distribution', 'createProfitDistributionForm')
        .addItem('Employee PF closure', 'createFundCloseForm')
        .addItem('New PF spreadsheet create', 'createNewSpreadsheetForm'))
      .addSubMenu(ui.createMenu('Actions')
        .addItem('Genarate PF statement', 'generateStatement')
        .addItem('Disburse monthly PF', 'disburseFund')
        .addItem('Deposit bank', 'depositBank')
        .addItem('Relocate bank fund', 'relocateBankFund')
        .addItem('Disburse loan', 'disburseLoan')
        .addItem('Issue FDR', 'issueFdr')
        .addItem('Encash FDR', 'encashFdr')
        .addItem('Distribute profit', 'distributeProfit')
        .addItem('Close employee PF', 'closeFund')
        .addItem('Create new PF spreadsheet', 'createNewSpreadsheet'))
      .addToUi();
}
