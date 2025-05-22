function test_onEdit() {
  let sheet = SpreadsheetApp.openById(config.get('testSheetId'));
  let range = sheet.getRange('Projects In-Progress!E2'); // test with change in cell E2
  let size = sheet.getLastRow();
  sheet.setActiveRange(range);

  onTriggeredEdit({
    user : Session.getActiveUser().getEmail(),
    source : sheet,
    range : range,
    value : range.getValue(),
    authMode : "LIMITED"
  });
}

function test_createProjectFolder() {
  let projectInfo = {
    projectId: '123456',
    description: 'Test Job',
    customer: 'New Customer',
  };
  createProjectFolder(projectInfo);
}