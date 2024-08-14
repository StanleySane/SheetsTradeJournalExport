
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Export')
      .addItem('Export trades to Portfolio Performance', 'showSelectSheetToExport')
      .addToUi();
}

function getFileFolder(file) {
  var folders = file.getParents();
  while (folders.hasNext()) {
    var folder = folders.next();
    return folder;
  }
  return null;
}

function getFilePath(file) {
  var parents = [];
  var folders = file.getParents();
  while (folders.hasNext()) {
    var folder = folders.next();
    parents.push(folder.getName());
    folders = folder.getParents();
  }
  parents = parents.reverse();  
  return parents;
}

function getThisTradeSheetsConfig() {
  return getTradeSheetsConfig(exportConfigSheet(), instrumentsConfigSheet());
}

function exportConfigSheet() {
  return SpreadsheetApp.getActive().getSheetByName(exportConfigSheetName);
}

function instrumentsConfigSheet() {
  return SpreadsheetApp.getActive().getSheetByName(instrumentsConfigSheetName);
}

function showSelectSheetToExport() {
  const initialSheetName = SpreadsheetApp.getActiveSheet().getName();
  showSelectSheetToExport(initialSheetName);
}

function showSelectSheetToExport(initialSheetName) {
  var html = HtmlService.createTemplateFromFile('SelectSheet');
  html.initialSheetName = initialSheetName;
  html = html.evaluate();

  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function processSelectSheetForm(formObject) {
  const sheetName = formObject.sheet;
  const sheetToExport = SpreadsheetApp.getActive().getSheetByName(sheetName);

  if (sheetToExport === null) {
    SpreadsheetApp.getUi().alert(`Sheet "${sheetName}" not found`);
    return;
  }

  var availableRows = getDatesFromSheet(sheetToExport);

  var html = HtmlService.createTemplateFromFile('SelectInterval');
  html.availableRows = availableRows;
  html.sheetName = sheetName;
  html = html.evaluate();

  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function processExportForm(formObject) {
  const sheetName = formObject.sheetName;
  const firstRow = parseInt(formObject.firstRow);
  const lastRow = parseInt(formObject.lastRow);
  const stateId = formObject.stateId;
  const ui = SpreadsheetApp.getUi();

  const sheetToExport = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (sheetToExport === null) {
    ui.alert(`Sheet "${sheetName}" not found`);
    return {startExport: false};
  }
  if (isNaN(firstRow)) {
    ui.alert(`Expression "${formObject.firstRow}" is not valid as first row number`);
    return {startExport: false};
  }
  if (isNaN(lastRow)) {
    ui.alert(`Expression "${formObject.lastRow}" is not valid as last row number`);
    return {startExport: false};
  }
  if (firstRow > lastRow) {
    ui.alert(`Last row number must be greater or equal to first row number`);
    return {startExport: false};
  }

  var result = ui.alert(
    'Export trades to CSV',
    `Are you sure you want to export data from sheet "${sheetName}" (rows ${firstRow} to ${lastRow})?`,
    ui.ButtonSet.YES_NO);

  if (result == ui.Button.YES) {
    return {startExport: true, sheetName: sheetName, firstRow: firstRow, lastRow: lastRow, stateId: stateId};
  }

  return {startExport: false};
}

function getCurrentState(stateId) {
  const cache = CacheService.getScriptCache();

  const stateString = cache.get(stateId);
  if (stateString === null) {
    return null;
  }

  const state = JSON.parse(stateString);
  return state;
}

function propagateState(state, cache) {
  const stateId = state.stateId;
  cache.put(stateId, JSON.stringify(state), 30 * 60); // 30 minutes
}

function startExport(sheetName, firstRow, lastRow, stateId) {
  const cache = CacheService.getScriptCache();
  const currentState = {stateId: stateId, msg: "Initialize"};
  propagateState(currentState, cache);

  const sheetToExport = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const spreadsheet = sheetToExport.getParent();
  const spreadsheetId = spreadsheet.getId();
  const spredsheetInDrive = DriveApp.getFileById(spreadsheetId);
  const spreadsheetFolder = getFileFolder(spredsheetInDrive);

  const folderNameToExport = "Export";
  var foldersToExportIteator = spreadsheetFolder.getFoldersByName(folderNameToExport);

  const nowTicks = Date.now(); // for uniqueness
  const now = new Date(nowTicks);
  const localeNow = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
  const dateString = localeNow.toISOString().slice(0, 10);
  const fileName = `${dateString}.${nowTicks}.${sheetName}.${firstRow}-${lastRow}.csv`;

  var folderToExport = undefined;
  var exportedFile = undefined;

  try {
    if (foldersToExportIteator.hasNext()) {
      folderToExport = foldersToExportIteator.next();
    }
    else {
      Logger.log(`Folder '${folderNameToExport}' not found`);
      folderToExport = spreadsheetFolder.createFolder(folderNameToExport);
      Logger.log(`New folder '${folderNameToExport}' created`);
    }

    const config = getThisTradeSheetsConfig();
    const csv = getCsvFromTradeSheet(sheetToExport, firstRow, lastRow, config, ";", currentState);

    currentState.msg = "Create Drive file";
    propagateState(currentState, cache);

    exportedFile = folderToExport.createFile(fileName, csv, MimeType.CSV);
  }
  catch(err) {
    currentState.msg = err;
    Browser.msgBox(err);
    throw err;
  }

  const fileLink = exportedFile.getUrl();
  const folderLink = folderToExport.getUrl();

  currentState.msg = "Export finished successfully";
  currentState.finished = true;
  propagateState(currentState, cache);
  
  showSuccess(fileName, fileLink, folderLink);
}

function showSuccess(fileName, fileLink, folderLink) {
  var html = HtmlService.createTemplateFromFile('SuccessExport');
  html.fileName = fileName;
  html.fileLink = fileLink;
  html.folderLink = folderLink;
  html = html.evaluate().setWidth(1000);

  SpreadsheetApp.getUi()
    .showModalDialog(html, 'Export trades to Portfolio Performance CSV');
}

function activateSheet(sheetName) {
  const sheetToActivate = SpreadsheetApp.getActive().getSheetByName(sheetName);

  if (sheetToActivate === null) {
    SpreadsheetApp.getUi().alert(`Sheet "${sheetName}" not found`);
    return;
  }

  sheetToActivate.activate();
}

function activateRow(sheetName, rowNum) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);

  if (sheet === null) {
    SpreadsheetApp.getUi().alert(`Sheet "${sheetName}" not found`);
    return;
  }

  var rowToActivate = sheet.getRange(`${rowNum}:${rowNum}`);
  rowToActivate.activate();
}
