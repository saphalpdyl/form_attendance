/**
 * Copyright © 2019 , 2020 , 2023  saphalpdyl
 * This file is part of Form Attendance which is released under MIT license.
 * See file LICENSE for full license details.
 * 
 */

function checkStart(e){
  let sheet = SpreadsheetApp.openById(e.source.getId());
  sheet.toast('Made by Saphal' , 'Recalculating');
  let allStudents = [] , inputStudents = [] , rows = 3 , days = sheet.getRange('B1').getValue() , main = sheet.getSheetByName('Main');
  while(main.getRange(rows, 1).isBlank() != true){
      allStudents.push(main.getRange(rows, 2).getValue());
      rows++;
    }
  for(let cDays = 1 ; cDays <= days ; cDays++){sheet.getSheetByName(cDays).getRange('C1').setFormula('=COUNTA(C2:C1000)');}
  for(let checkDaysC = 1 ; checkDaysC <= days ; checkDaysC++ ){
    let currentSheet = sheet.getSheetByName(checkDaysC);
    let inStudentsCount = currentSheet.getRange('C1').getValue();
    inputStudents = [];
    for(let tmpStudentCount = 1 ; tmpStudentCount <= inStudentsCount ; tmpStudentCount++){inputStudents.push(currentSheet.getRange('C' + (tmpStudentCount + 1)).getValue());} //FIlling students from form responses
    for(let checkStudentsCounter = 0 ; checkStudentsCounter < allStudents.length ; checkStudentsCounter++)
    {
      let hasDone = false;
      for(let tmpCounter2 = 0 ; tmpCounter2 < inputStudents.length ; tmpCounter2++){
        if(allStudents[checkStudentsCounter].toString().trim() == inputStudents[tmpCounter2].toString().trim()){
          hasDone = true;
          break;
        }
      }
      if(hasDone == true){main.getRange(checkStudentsCounter + 3, checkDaysC + 4).setValue('TRUE');} // Side bias and up bias
      else{main.getRange(checkStudentsCounter + 3, checkDaysC + 4).setValue('FALSE');}
    }
  }
}