const ss = SpreadsheetApp.getActive();
const forms = ss.getSheetByName('Forms');
const acs = ss.getSheetByName('ACs');
const requests = ss.getSheetByName('Requests');
const jobs = ss.getSheetByName('Jobs');
const parts = ss.getSheetByName('Parts');
const items = ss.getSheetByName('Items');
const emails = ss.getSheetByName('Emails');
const ui = SpreadsheetApp.getUi();

//-- Submit maintenance request --//
function submitRequest() {
  let currentDateTime = new Date();
  let formValues = forms.getRange('B3:B5').getValues();
  let popName = formValues[0][0];
  let acName = formValues[1][0];
  let problemType = formValues[2][0];
  
  //check form values
  if (popName == '' || acName == '' || problemType == '') {
    ui.alert('You must give PoP Name, AC Name, Problem Type, Ac Brand & Ac Capacity');
    return;
  }

  let lastRowNumber = requests.getLastRow();

  let job = [lastRowNumber + 1, popName, acName, problemType, currentDateTime];
  requests.appendRow(job);
}