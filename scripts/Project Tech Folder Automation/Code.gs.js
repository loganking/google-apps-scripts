/**
 * @OnlyCurrentDoc
 */

// example trigger function to be used in the source spreadsheet script
function triggerForProjectFolders(e) {
  const scriptProperties = PropertiesService.getScriptProperties();
  let config = {
    adminFolderId: scriptProperties.getProperty('adminFolderId'),
    projectsFolderId: scriptProperties.getProperty('projectsFolderId'),
    templateFolderId: scriptProperties.getProperty('templateFolderId'),
  }

  ProjectTechFolderAutomation.onTriggeredEdit(e, config);
}

// need to add manual trigger for Spreadsheet onEdit event
function onTriggeredEdit(e, config) {
  let sheet = e.source.getActiveSheet();

  // Ensure we are working on the correct sheet
  if (sheet.getName() !== "Projects In-Progress") {
    console.log('Ending. Triggered edit not in Projects sheet.');
    return;
  }


  if (isUpdateInColumn(e, 'Status') && e.range.getValues()[0][0] === 'Not Started Yet') {
    checkForProjectFolders(e, config);
  }
  if (isUpdateInColumn(e, 'Status') && e.range.getValues()[0][0] === 'Completed/Invoiced') {
    archiveProjectFolders(e, config);
  }
}

function checkForProjectFolders(e, config) {
  let projectInfo = setProjectInfo(e, config);
  let sheet = e.source.getActiveSheet();

  if (projectInfo.projectId == "" || projectInfo.description == "" || projectInfo.customer == "") {
    console.log("Not creating folders. Context: ", projectInfo);
    return;
  }

  if (projectInfo.projectFolder == "") {
    let folder = createJobFolder(projectInfo, 'project');
    updateValueByColumnName(sheet, projectInfo.editedRow, 'Project Folder (Technician)', folder.getUrl());
    console.log("Created project folder ", folder.getUrl());
  }

  if (projectInfo.adminFolder == "") {
    let folder = createJobFolder(projectInfo, 'admin');
    updateValueByColumnName(sheet, projectInfo.editedRow, 'Administration Project Folder', folder.getUrl());
    console.log("Created admin folder ", folder.getUrl());
  }
}

function archiveProjectFolders(e, config) {
  const projectInfo = setProjectInfo(e, config);

  if (projectInfo.projectId == "" || projectInfo.customer == "") {
    return;
  }

  console.log(`Attempting to archive project folders for`, projectInfo);
  const customerFolderName = getCustomerFolderName(projectInfo);
  if (projectInfo.projectFolder != "") {
    moveToArchiveFolder(projectInfo.projectsFolderId, customerFolderName, projectInfo.projectFolder);
  }
  if (projectInfo.adminFolder != "") {
    moveToArchiveFolder(projectInfo.adminFolderId, customerFolderName, projectInfo.adminFolder);
  }
}

function moveToArchiveFolder(rootId, customerFolderName, folderName) {
  const rootFolder = DriveApp.getFolderById(rootId);
  const customerFolder = findFolder(rootFolder, customerFolderName);
  const archiveFolder = findOrCreateFolder(customerFolder, 'Completed');
  const folder = findFolder(customerFolder, folderName);
  folder.moveTo(archiveFolder);
  console.log(`Moved folder ${folderName} to archive.`);
}

function setProjectInfo(e, config) {
  let sheet = e.source.getActiveSheet();
  let editedRow = e.range.getRow();

  let projectId = getValueByColumnName(sheet, editedRow, 'QB Project');
  let description = getValueByColumnName(sheet, editedRow, 'Description');
  let customer = getValueByColumnName(sheet, editedRow, 'Customer Name');
  let projectFolder = getValueByColumnName(sheet, editedRow, 'Project Folder (Technician)');
  let adminFolder = getValueByColumnName(sheet, editedRow, 'Administration Project Folder');

  let projectInfo = {
    'editedRow': editedRow,
    'adminFolderId': config.adminFolderId,
    'projectsFolderId': config.projectsFolderId,
    'templateFolderId': config.templateFolderId,
    'projectId': projectId,
    'description': description,
    'customer': customer,
    'projectFolder': projectFolder,
    'adminFolder': adminFolder,
  }
  return projectInfo;
}

function getColIdByName(sheet, name) {
  let data = sheet.getDataRange().getValues();
  let header = data[0].map((name)=>{
    return name.trim().toLowerCase();
  });
  let index = header.indexOf(name.toLowerCase());
  if (index == '-1') {
    console.log(`Headers: `,data[0]);
    throw new Error(`Unable to find column ${name} in header row.`);
  }
  return index;
}

function isUpdateInColumn(e, column){
  let colId = getColIdByName(e.source.getActiveSheet(), column)
  return (e.range.getColumn() == colId+1);
}

function getValueByColumnName(sheet, row, name) {
  let colId = getColIdByName(sheet, name);
  let value = sheet.getRange(row, colId + 1).getValues()[0][0];
  if (typeof value === 'string') {
    return value.trim();
  }
  return value;
}

function updateValueByColumnName(sheet, row, name, value) {
  let colId = getColIdByName(sheet, name);
  return sheet.getRange(row, colId + 1).setValue(value);
}

function createJobFolder(projectInfo, type) {
  console.log(`attempting to create ${type} folder for`, projectInfo);

  let rootFolder;
  try {
    let folderId = (type === 'admin') ? projectInfo.adminFolderId : projectInfo.projectsFolderId;
    rootFolder = DriveApp.getFolderById(folderId);
  } catch (e) {
    console.log(`Error. Unable to find or open ${type} folder.`);
    throw e;
  }

  let customerFolderName = getCustomerFolderName(projectInfo);
  let customerFolder = findOrCreateFolder(rootFolder, customerFolderName);

  let jobFolderName = `${projectInfo.projectId} - ${(type === 'admin')? '(Admin) ': ''}${projectInfo.description}`
  let jobFolder = findOrCreateFolder(customerFolder, jobFolderName);

  if (type !== 'admin') {
    copyTemplateToFolder(projectInfo.templateFolderId, jobFolder);
    updateJobFilesInFolder(jobFolder, projectInfo)
  }

  return jobFolder;
}

function getCustomerFolderName(projectInfo) {
  if (projectInfo.customer.startsWith('Larimer County')) {
    return 'Larimer County';
  }
  return projectInfo.customer;
}

function copyTemplateToFolder(templateFolderId, folder) {
  try {
    var templateFolder = DriveApp.getFolderById(templateFolderId);
  } catch (e) {
    console.log('Error. Unable to find or open template folder.');
    throw e;
  }
  let count = recursivelyCopyFilesAndFolders(templateFolder, folder);
  console.log(`Copied ${count} template files to folder.`)
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

function findFolder(parentFolder, name) {
  let childFolders = parentFolder.getFoldersByName(name);
  if (childFolders.hasNext()) {
    return childFolders.next()
  }
  throw new Error(`Error. Unable to find folder ${name} in ${parentFolder.getName()}`);
}

function findOrCreateFolder(parentFolder, name) {
  let childFolders = parentFolder.getFoldersByName(name);
  if (childFolders.hasNext()) {
    return childFolders.next()
  }
  return parentFolder.createFolder(name);
}

function findOrCreateFolderPath(root, path) {
  let parts = path.split('/');
  if (parts[parts.length - 1].contains('.')) {
    parts.pop(); // Remove file separator at the end if it exists
  }

  let currentFolder = root;
  for (let part of parts) {
    if (part.trim() === '') continue; // Skip empty parts
    let childFolders = currentFolder.getFoldersByName(part);
    if (childFolders.hasNext()) {
      currentFolder = childFolders.next();
    } else {
      currentFolder = currentFolder.createFolder(part);
    }
  }
  return currentFolder;
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
