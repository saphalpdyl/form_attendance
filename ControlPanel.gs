/**
 * Copyright © 2019 , 2020 , 2023  saphalpdyl
 * This file is part of Form Attendance which is released under MIT license.
 * See file LICENSE for full license details.
 * 
 */

let runC4 = () => runStart("C4" , 3);
let runC5 = () => runStart("C5" , 4);
let runC6 = () => runStart("C6" , 5);
let runC7 = () => runStart("C7" , 6);


function runStart(CName , rowNum) {

  //Start Timer
  let date1 = new Date();
  let time1 = date1.getTime();

  //Primary References
  let sprdst = SpreadsheetApp.getActiveSpreadsheet();
  let panel = sprdst.getSheetByName("Panel");
  let drv = DriveApp;
  let ui = SpreadsheetApp.getUi();

  //Data 
  let className = CName;
  let sampleFileName = panel.getRange(rowNum , 2).getValue() + " " + className;
  let destFolderName = panel.getRange(rowNum , 3).getValue();
  let destSpreadsheetName = panel.getRange(rowNum , 4).getValue() + " " + className;
  let copyName;
  let namePrompt =  ui.prompt("Write assignment Name" , ui.ButtonSet.OK_CANCEL);
  
  if(namePrompt.getSelectedButton() == ui.Button.CANCEL) return;

  if(namePrompt.getResponseText() == "" || namePrompt.getResponseText() == null){
    ui.alert("Please enter a valid name");
    return;
  }

  //Writing to copyName
  copyName = namePrompt.getResponseText();

  //Folder and file References 
  let locFolder;
  let destSpreadsheet;
  let toBeCopiedFile;
  let destFolder
  
  //For class Folder
  let searchLocFolders = drv.getFoldersByName(className);
  if(searchLocFolders.hasNext()) locFolder = searchLocFolders.next();
  else return;

  //For Form  Attendance reference
  let searchDestSS = locFolder.getFilesByName(destSpreadsheetName);
  if(searchDestSS.hasNext()) destSpreadsheet = searchDestSS.next();
  else return;

  //For Sample Form
  let searchSampleForm = locFolder.getFilesByName(sampleFileName);
  if(searchSampleForm.hasNext()) toBeCopiedFile = searchSampleForm.next();
  else return;

  //Destination Folder
  let searchDestFolder = locFolder.getFoldersByName(destFolderName);
  if(searchDestFolder.hasNext()) destFolder = searchDestFolder.next();
  else return;

  let copiedFile = toBeCopiedFile.makeCopy(copyName , destFolder);
  let copiedFileId = copiedFile.getId();
  
  let finalForm = FormApp.openById(copiedFileId);
  finalForm.setDestination(FormApp.DestinationType.SPREADSHEET , destSpreadsheet.getId());

  //Referencing to Form Attendance
  let attendance = SpreadsheetApp.openById(destSpreadsheet.getId());

  //Reference to sheets array
  let sheets = attendance.getSheets();

  //The first sheet , the added one
  let firstSheet = sheets[0];

  //The total number of days extracted from B1 cell of Main sheet
  let numDayValue = parseInt(sheets[1].getRange(1 , 2).getValue());
  firstSheet.setName(numDayValue + 1);

  //Tweaking the name of the copied Assignment by add *Day number* - *assingment* - .....
  let newNameforCopiedFile = (numDayValue + 1) + " - " + copyName  + ' | ' + className;
  copiedFile.setName(newNameforCopiedFile);

  //Set as active sheet
  firstSheet.activate();
  attendance.moveActiveSheet(sheets.length);

  // Referencing 'Main' Sheet
  let mainSheet = attendance.getSheetByName("Main");

  //Reference to Settings sheet
  let settings = sprdst.getSheetByName("Settings");

  //Values from Settings
  let columnBias = settings.getRange(1 , 2).getValue();
  let lastRowBias = settings.getRange(2 , 2).getValue();

  //Inserts empty column
  mainSheet.insertColumnsAfter(numDayValue + columnBias, 1);
  //Copying value of previous column
  let cpyValue = mainSheet.getRange(1 , numDayValue + columnBias  , mainSheet.getLastRow() - lastRowBias , 1).getValues();
  //Pasting the copied value
  mainSheet.getRange(1 , numDayValue + columnBias + 1 , mainSheet.getLastRow() - lastRowBias , 1).setValues(cpyValue);
  //Converting to checkboxes
  mainSheet.getRange(3 , numDayValue + columnBias + 1 , mainSheet.getLastRow() - lastRowBias , 1).insertCheckboxes();
  //Changing the day number
  mainSheet.getRange(1 ,  numDayValue + columnBias + 1 ).setValue("D-" + (numDayValue + 1));

  //Finish Timer
  let date2 = new Date();
  let time2 = date2.getTime();

  //For Logs
  let logs = sprdst.getSheetByName('Logs');
  let tmpRows = logs.getLastRow() + 1
  logs.getRange(tmpRows , 1).setValue(date1);
  logs.getRange(tmpRows , 2).setValue(className);
  logs.getRange(tmpRows , 3).setValue(newNameforCopiedFile);
  logs.getRange(tmpRows , 4).setValue( (time2 - time1) + " ms");
}