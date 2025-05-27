function test_onEdit() {
  let sheet = SpreadsheetApp.openById(config.get('testSheetId'));
  let range = sheet.getRange('Projects In-Progress!E2'); // test with change in cell E2
  let size = sheet.getLastRow();
  sheet.setActiveRange(range);
  
  let event = {
    user : Session.getActiveUser().getEmail(),
    source : sheet,
    range : range,
    value : range.getValue(),
    authMode : "LIMITED"
  };

  const scriptProperties = PropertiesService.getScriptProperties();
  let config = {
    projectsFolderId: scriptProperties.getProperty('projectsFolderId'),
    templateFolderId: scriptProperties.getProperty('templateFolderId'),
  }

  onTriggeredEdit(event, config);
}

function test_createProjectFolder() {
  let projectInfo = {
    projectId: '123456',
    description: 'Test Job',
    customer: 'New Customer',
  };
  createProjectFolder(projectInfo);
}

function testProperty(name) {
    const scriptProperties = PropertiesService.getScriptProperties();
    let propery = scriptProperties.getProperty(name);
    console.log(`${name}: `, propery);
}
