//
// constant values
//
const panelSheetName = "パネル";
const workSheetName = "ワークシート";
let srcFolderID;
let dstFolderID;
let startRow;
let processingUnit

let spreadsheet;
let sheet;

const colSrcFolderID = 2;
const colDstFolderID = 3;
const colStatus = 4;

//
// パネルシートの設定情報を読み込む
//
function readPanel() {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    sheet = spreadsheet.getSheetByName(panelSheetName);

    srcFolderID = sheet.getRange("コピー元").getValue();
    dstFolderID = sheet.getRange("コピー先").getValue();
    startRow = sheet.getRange("開始行").getValue();
    processingUnit = sheet.getRange("処理単位").getValue();
}

function showMessage(msg) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(msg);
}

function dispLog(msg) {
  Logger.log(msg);
}


//
// 全フォルダのフォルダ名とIDのリストを作成
//
let sheetRow = 1;
function CreateFolderList() {
  readPanel();
  var srcFolder = DriveApp.getFolderById(srcFolderId);
  var srcFolders = srcFolder.getFolders();//フォルダ内フォルダをゲット
  
  while(srcFolders.hasNext()) {
    var nextSrcFolder = srcFolders.next();
    setFolderList(srcFolder.getName(), nextSrcFolder, sheet); //再帰処理
  }
  showMessage('完了しました');
}



//
// Google Driveのフォルダをサブフォルダも含めて全部コピー（再起処理） 
//
function CopyWholeFolder() {
  readPanel(); // 設定情報読み込み
  // ui.alert(srcFolderId + 'から' + dstFolderId + 'へコピーします');

  var srcFolder = DriveApp.getFolderById(srcFolderId);
  var dstFolder = DriveApp.getFolderById(dstFolderId);
  
  var dstFolderName = srcFolder.getName();
  
  var newFolder = dstFolder.createFolder(dstFolderName);
  copyFolderRecursive(srcFolder, newFolder); // フォルダの中身を再起コピー
  ui.alert('完了しました');
}

//
// ワークシートの指定行の範囲にあるフォルダ間でファイルだけをコピー（再起処理しない）
//
function CopyAllFiles() {
  readPanel(); // 設定情報読み込み

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();

  for (i = startRow; i < startRow + processingUnit; i++) {
    var src = sheet.getRange(i, colSrcFolderID).getValue();
    var dst = sheet.getRange(i, colDstFolderID).getValue();
    // var log = "From " + src + " To " + dst;
    copyFilesOnly(src, dst);
    sheet.getRange(i, colStatus).setValue("済");
  }
  showMessage('完了しました');
}

//
//  ワークシートに記載のフォルダ構成をターゲットフォルダ内に作成
//
function CreateNewTree() {
  var sheetRow = 1;
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var lists = sheet.getRange('A1:A710').getValues();



  for (list of lists) {
    dispLog(list);
    if (list == null) break;
    target = createTree(dstFolderID, list.toString().split("/"));
    sheet.getRange(sheetRow, 3).setValue(target.getId());
    sheetRow++;
  }
  showMessage('完了しました');
}

//
// Utility functions
// 再帰関数
// Google Driveのフォルダをサブフォルダも含めて全部コピー（再起処理） 
//
function copyFolderRecursive(srcFolder, newFolder){
  var srcFiles = srcFolder.getFiles();//フォルダ内ファイルをゲット
  while(srcFiles.hasNext()) {
    var srcFile = srcFiles.next();
    dispLog(srcFile.getName());
    srcFile.makeCopy(srcFile.getName(), newFolder);
  }
  var srcFolders = srcFolder.getFolders();//フォルダ内フォルダをゲット
  while(srcFolders.hasNext()) {
    var nextSrcFolder = srcFolders.next();
    dispLog(nextSrcFolder.getName());
    var nextNewFolder = newFolder.createFolder(nextSrcFolder.getName());
    copy(nextSrcFolder, nextNewFolder); //再帰処理
  }
}

//
// Utility functions
// フォルダ間でファイルだけをコピー
//
function copyFilesOnly(srcFolderID, dstFolderID) {
  var srcFolder = DriveApp.getFolderById(srcFolderID);
  var dstFolder = DriveApp.getFolderById(dstFolderID);
  dispLog(srcFolder);

  var srcFiles = srcFolder.getFiles();//フォルダ内ファイルをゲット
  while(srcFiles.hasNext()) {
    var srcFile = srcFiles.next();
    dispLog(srcFile.getName());
    srcFile.makeCopy(srcFile.getName(), dstFolder);
  }
}


//
//
//
function createTree(targetFolder, folderList) {
  dispLog(folderList);
  target = DriveApp.getFolderById(targetFolder)
  for (const folder of folderList) {
    dispLog(folder);
    target = createNewFolder(target, folder);
  }
  return target;
}

//
// 指定フォルダ内に新しいフォルダを作成。
// 既に同名のフォルダがあればそれを使う。
//
function createNewFolder(target, folderName){
  var folders = target.getFolders();
  while(folders.hasNext()) {
    var folder = folders.next();
    if (folder.getName() == folderName ) {
      dispLog(folderName +"作成をスキップ")
      return folder;
    }
  }
  folder = target.createFolder(folderName);
  return folder;
}


//
//
//
function setFolderList(parentFolderName, srcFolder, sheet){
  var srcFolders = srcFolder.getFolders();//フォルダ内フォルダをゲット
  var row = 1;
  while(srcFolders.hasNext()) {
    var nextSrcFolder = srcFolders.next();
    let folderName = parentFolderName + '/' + srcFolder + '/' + nextSrcFolder.getName();
    let folderID = nextSrcFolder.getId();
    dispLog(folderName);
    dispLog(folderID);
    sheet.getRange(sheetRow,1).setValue(folderName);
    sheet.getRange(sheetRow,2).setValue(folderID);
    sheetRow++;
    setFolderList(folderName, nextSrcFolder, sheet); //再帰処理
  }
}

