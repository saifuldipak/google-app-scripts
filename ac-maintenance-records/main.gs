const ss = SpreadsheetApp.getActive();
const forms = ss.getSheetByName('Forms');
const acs = ss.getSheetByName('ACs');
const jobs = ss.getSheetByName('Jobs');
const parts = ss.getSheetByName('Parts');
const items = ss.getSheetByName('Items');
const emails = ss.getSheetByName('Emails');
const ui = SpreadsheetApp.getUi();

function submitJob() {
  let formValues = forms.getRange('B3:B5').getValue();
  let popName = formValues[0];
  let acName = formValues[1];
  let problemType = formValues[2];
  let acBrand = formValues[3];
  let acCapacity = formValues[4];
  
  //check form values
  if (popName == '' || acName == '' || problemType == '' || acBrand == '' || acCapacity == '') {
    ui.alert('You must give PoP Name, AC Name, Problem Type, Ac Brand & Ac Capacity');
    return;
  }



}