const ss = SpreadsheetApp.getActive();
const forms = ss.getSheetByName('Forms');
const acs = ss.getSheetByName('ACs');
const requests = ss.getSheetByName('Requests');
const jobs = ss.getSheetByName('Jobs');
const parts = ss.getSheetByName('Parts');
const items = ss.getSheetByName('Items');
const emails = ss.getSheetByName('Emails');
const ui = SpreadsheetApp.getUi();

//------- Send custom mail ----------//
function sendEmail(actionType, popName, acName, problemType, jobId) {
  let recipients = emails.getDataRange().getValues();

  if (actionType != 'request') {
    Logger.log('sendEmail(): Unknown actionType');
    return;
  }
  
  let recipientCcList = [];
  if (actionType == 'request') {
    for (i = 0; i < recipients.length; i++) {
      if (recipients[i][1] == 'Admin' && recipients[i][2] == 'Manager') {
        var recipientTo = recipients[i][3];
      }
      if (recipients[i][0] == 'Technical' && recipients[i][2] == 'Head') {
        recipientCcList.push(recipients[i][3]);
      }
    }
  }

  const options = {cc:recipientCcList.join(',')};
  const subject = `AC maintenance ${actionType}: ${popName}-${acName}-${jobId}`;
  const body = `Job Id: ${jobId} \nPoP Name: ${popName} \nAC name: ${acName} \nProblem type: ${problemType} \n\n${ss.getUrl()}`;

  MailApp.sendEmail(recipientTo, subject, body, options);
}

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
  let jobId = lastRowNumber + 1;
  let job = [jobId, popName, acName, problemType, currentDateTime];
  requests.appendRow(job);

  sendEmail('request', popName, acName, problemType, jobId);
}
