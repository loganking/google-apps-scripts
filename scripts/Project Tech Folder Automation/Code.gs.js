/**
 * @OnlyCurrentDoc
 */

// need to add manual trigger for Spreadsheet onEdit event
function onTriggeredEdit(e, config) {
  checkForProjectFolder(e, config);
}

function checkForProjectFolder(e, config) {
  let sheet = e.source.getActiveSheet();

  // Ensure we are working on the correct sheet
  if (sheet.getName() !== "Projects In-Progress") {
    console.log('Ending. Triggered edit not in Projects sheet.');
    return;
  }

  // Only create project folders when Status is updated to "Not Started Yet"
  if (!isUpdateInColumn(e, 'Status') || e.range.getValues()[0][0] != 'Not Started Yet') {
    console.log('Ending. Triggered edit not changing Status to "Not Started Yet"');
    return;
  }

  let editedRow = e.range.getRow();
  let projectId = getValueByColumnName(sheet, editedRow, 'QB Project');
  let description = getValueByColumnName(sheet, editedRow, 'Description');
  let customer = getValueByColumnName(sheet, editedRow, 'Customer Name');
  let folder = getValueByColumnName(sheet, editedRow, 'Project Folder (Technician)');

  let projectInfo = {
    'projectsFolderId': config.projectsFolderId,
    'templateFolderId': config.templateFolderId,
    'projectId': projectId,
    'description': description,
    'customer': customer,
    'projectFolder': folder,
  }

  if (projectId == "" || description == "" || folder != "") {
    console.log("Not creating folder with context", projectInfo);
    return;
  }

  let folderUrl = createProjectFolder(projectInfo);
  updateValueByColumnName(sheet, editedRow, 'Project Folder (Technician)', folderUrl);
}

function getColIdByName(sheet, name) {
  let data = sheet.getDataRange().getValues();
  let header = data[0].map((name)=>{
    return name.trim().toLowerCase();
  });
  let index = header.indexOf(name.toLowerCase());
  if (index == '-1') {
    console.log(`Error. Unable to find index for ${name}`);
  }
  return index;
}

function isUpdateInColumn(e, column){
  let colId = getColIdByName(e.source.getActiveSheet(), column)
  return (e.range.getColumn() == colId+1);
}

function getValueByColumnName(sheet, row, name) {
  let colId = getColIdByName(sheet, name);
  return sheet.getRange(row, colId + 1).getValues()[0][0];
}

function updateValueByColumnName(sheet, row, name, value) {
  let colId = getColIdByName(sheet, name);
  return sheet.getRange(row, colId + 1).setValue(value);
}

function createProjectFolder(projectInfo) {
  if (projectInfo.customer.startsWith('Larimer County')) {
    projectInfo.customer = 'Larimer County';
  }

  console.log('attempting to create folder for', projectInfo);

  try {
    var projectsFolder = DriveApp.getFolderById(projectInfo.projectsFolderId);
  } catch (e) {
    console.log('Error. Unable to find or open projects folder.');
    throw e;
  }

  let customerFolder = findOrCreateFolder(projectsFolder, projectInfo.customer);

  let jobFolderName = `${projectInfo.projectId} - ${projectInfo.description}`
  let jobFolder = findOrCreateFolder(customerFolder, jobFolderName);

  copyTemplateToProjectFolder(jobFolder, projectInfo);
  updateJobFilesInFolder(jobFolder, projectInfo)

  return jobFolder.getUrl();
}

function copyTemplateToProjectFolder(projectFolder, projectInfo) {
  try {
    var templateFolder = DriveApp.getFolderById(projectInfo.templateFolderId);
  } catch (e) {
    console.log('Error. Unable to find or open template folder.');
    throw e;
  }
  let count = recursivelyCopyFilesAndFolders(templateFolder, projectFolder);
  console.log(`Copied ${count} template files to new job folder.`)
}

function recursivelyCopyFilesAndFolders(orig, dest, count = 0) {
  let files = orig.getFiles();
  while (files.hasNext()) {
    let file = files.next();
    file.makeCopy(file.getName(), dest);
    count++;
  }
  
  let folders = orig.getFolders();
  while (folders.hasNext()) {
    let folder = folders.next();
    let newFolder = dest.createFolder(folder.getName());
    count ++;
    count = count + recursivelyCopyFilesAndFolders(folder, newFolder);
  }

  return count;
}

function findOrCreateFolder(parentFolder, name) {
  let childFolders = parentFolder.getFoldersByName(name);
  while(childFolders.hasNext()) {
    return childFolders.next()
  }

  return parentFolder.createFolder(name);
}

function updateJobFilesInFolder(jobFolder, projectInfo) {
  let templateFiles = jobFolder.getFiles();
  while (templateFiles.hasNext()) {
    let file = templateFiles.next();
    let name = file.getName();
    if (name == "BOM tracking") {
      file.setName(`${projectInfo.projectId} - BOM Tracking`);
      SpreadsheetApp.openById(file.getId())
        .getRange('A1:D1').setValue(`${projectInfo.projectId} - ${projectInfo.description}`);
    }
    if (name == "Punch-List") {
      file.setName(`${projectInfo.projectId} - Punch-List`);
      SpreadsheetApp.openById(file.getId())
        .getRange('B2').setValue(`${projectInfo.projectId} - ${projectInfo.description}`);
    }
  }
}
