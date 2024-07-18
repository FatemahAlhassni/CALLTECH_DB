

/**
 * Constants
 */ 
const spreadsheet = SpreadsheetApp.openById("1U0F1Y3uyqEDK3lAPqdIkqyFAKPJN4w365s0a0OxHXBk");
const schoolsSheet = spreadsheet.getSheetByName("schools");
const contactsSheet = spreadsheet.getSheetByName("Copy of contacts");

function getCombinedData() { // دمج اسم المدرسة مع رمز المدرسة بالصيغة المطلوبة
  const destlastRow1 = schoolsSheet.getLastRow();
  var dataRange = schoolsSheet.getRange(`A2:B${destlastRow1}`); // لتحديد أخر صف يحتوي على بيانات
  var values = dataRange.getValues();
 
  var combinedData = [];
  for (var i = 0; i < values.length; i++) {
    combinedData.push([values[i][0] + " | " + values[i][1]]);
  }
  return combinedData;
}

function createNamedRangesFromDifferentSheets() {
  var sheetNames = ["admins&offices", "admins&offices", "clients" ,"schools","schools","clients"];
  var ranges =  ["B2:B31",// Adminstration
                "C2:C11",// Office
                "A2:A18",//Contractor
                "AI2:AI4",//Project_Name
                "AF2:AF7",//Entity_name
                "A2:A18"]; //Subcontractor

  var names =  ["Adminstration", "Office", "Contractor", "Project_Name", "Entity_name", "Subcontractor"];

  // Loop through the ranges and set the values
  for (var i = 0; i < sheetNames.length; i++) {
    var sheet = spreadsheet.getSheetByName(sheetNames[i]);
    var range = sheet.getRange(ranges[i]);
    spreadsheet.setNamedRange(names[i], range);
  }
   
}

function removeDuplicates() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var newData = new Array();
  for (var i in data) {
    var row = data[i];
    var duplicate = false;
    for (var j in newData) {
      if (row.join() == newData[j].join()) {
        duplicate = true;
      }
    }
    if (!duplicate) {
      newData.push(row);
    }
  }
  //sheet.clearContents(); // لمسح المحتوى فقط في النطاق الحالي
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

function getDropdownOptions() {
  var namedRanges = ["Adminstration", "Office", "Contractor", "Project_Name", "Entity_name", "Subcontractor"];
  var dropdownOptions = [];  
 
  namedRanges.forEach(function(namedRange) {
    var range = spreadsheet.getRangeByName(namedRange);
    if (range) {
      var values = range.getValues();
      // تجميع جميع القيم في مصفوفة واحدة
      dropdownOptions = dropdownOptions.concat(values.flat());
    }
  });

  var dropdownOptions = removeEmptyValues(dropdownOptions)
  dropdownOptions = dropdownOptions.filter((value, index, self) => self.indexOf(value) === index);
  return dropdownOptions;
}

function addDataToSheet(selectedValue,column) {   // تعدلت
  var lastRow = getLastRowInColumnA(1,contactsSheet)
  var range = contactsSheet.getRange(lastRow, column); // الصف الأخير + 1 في العمود
  range.setValue(selectedValue);
}

function getLastRowInColumnA(column, gool_sheet) {

  var data = gool_sheet.getDataRange().getValues();

  for (var row = data.length - 1; row >= 0; row--) {
    if (data[row][column - 1]) {
      break;
    }
  }
  var lastRow = row + 1;
  return lastRow + 1;
}

function getLastRowInColumn(column,gool_sheet) {  // اضافة جديد// column, gool_sheet

  var sheet = spreadsheet.getSheetByName(gool_sheet)// اسم صفحة البيانات
  var data = sheet.getDataRange().getValues();

  for (var row = data.length - 1; row >= 0; row--) {
    if (data[row][column - 1]) {
      break;
    }
  }
  var lastRow = row + 1;
  return lastRow+1;
}

function dropdownOptions_ScooleCode() {
  var namedRange = ["schoolcodeToSalse"]
  var dropdownOptions = []

  namedRange.forEach(function(namedRange) {
  var range = schoolsSheet.getRange(namedRange);
 
    if (range) {
      var values = range.getValues();
      // تجميع جميع القيم في مصفوفة واحدة
      dropdownOptions = dropdownOptions.concat(values.flat());
    }
    })
    var filteredOptions = removeEmptyValues(dropdownOptions)
    return filteredOptions
}

function dropdownOptions_ScooleName() {
  var namedRange = ["schoolnameToSalse"]
  var dropdownOptions = []

  namedRange.forEach(function(namedRange) {
  var range = schoolsSheet.getRange(namedRange);
 
    if (range) {
      var values = range.getValues();
      // تجميع جميع القيم في مصفوفة واحدة
      dropdownOptions = dropdownOptions.concat(values.flat());
    }
    })
    var filteredOptions = removeEmptyValues(dropdownOptions);
    return filteredOptions
}

function removeEmptyValues(array) {
  return array.filter(value => value !== "");
}

function insertCodebyNameToContacts(sc_name) {
  var lastRow = getLastRowInColumnA(1,contactsSheet)
  var sc_code
  var sc_nameR = schoolsSheet.getRange("B2:B");
  for (var i = 0; i < sc_nameR.getNumRows(); i++) {
    if (sc_nameR.getCell(i + 1, 1).getValue() === sc_name) {


    var rowNum = (i+2)
    var cellInRowA = schoolsSheet.getRange("A" + rowNum);
    sc_code = cellInRowA.getValue();
    break;
  }
  }
   contactsSheet.getRange(lastRow, 1).setValue(sc_code); // تعيين قيمة sc_code في العمود A
   contactsSheet.getRange(lastRow, 2).setValue(sc_name); // تعيين قيمة sc_name في العمود B
}

function insertNamebyCodeToContacts(sc_code) {
  var lastRow = getLastRowInColumnA(1,contactsSheet)
  var sc_name
  var sc_nameR = schoolsSheet.getRange("A2:A");
  for (var i = 0; i < sc_nameR.getNumRows(); i++) {
     if (sc_nameR.getCell(i + 1, 1).getValue() === sc_code) {
      var rowNum = (i+2)
      var cellInRowB = schoolsSheet.getRange("B" + rowNum);
      sc_name = cellInRowB.getValue();
      break;
  }
  }
  contactsSheet.getRange(lastRow, 1).setValue(sc_code); // تعيين قيمة sc_code في العمود A
  contactsSheet.getRange(lastRow, 2).setValue(sc_name); // تعيين قيمة sc_name في العمود B
}

function getCodebyNameToContacts(sc_name) {
 
  var sc_code
  var sc_nameR = schoolsSheet.getRange("B2:B");
  for (var i = 0; i < sc_nameR.getNumRows(); i++) {
    if (sc_nameR.getCell(i + 1, 1).getValue() === sc_name) {


    var rowNum = (i+2)
    var cellInRowA = schoolsSheet.getRange("A" + rowNum);
    sc_code = cellInRowA.getValue();
    break;
  }
  }
   return sc_code;
}

function getNamebyCodeToContacts(sc_code) {
 
  var sc_name
  var sc_nameR = schoolsSheet.getRange("A2:A");
  for (var i = 0; i < sc_nameR.getNumRows(); i++) {
     if (sc_nameR.getCell(i + 1, 1).getValue() === sc_code) {
      var rowNum = (i+2)
      var cellInRowB = schoolsSheet.getRange("B" + rowNum);
      sc_name = cellInRowB.getValue();
      break;
  }
  }
  return sc_name;
}

function updateContacts() {
    const lastRow = contactsSheet.getLastRow();
    var rowNumber = getLastRowInColumn(1,"Copy of contacts")

    rowNumber = rowNumber -1 ;

  for (let col = 1; col <= contactsSheet.getLastColumn(); col++) {
    const cellValue = contactsSheet.getRange(rowNumber, col).getValue();
    if (cellValue === "") {
      contactsSheet.getRange(rowNumber, col).setValue("");
    }
  }
    Logger.log (rowNumber+1)// الصف التالي الفارغ
}

function shiftRowsDown() {// عمل إزاحة للصفوف إلى الأسفل وانشاء صف جديد فارغ
  var rowToShift = 2; // رقم الصف الذي ترغب في نقله
  var numColumns = contactsSheet.getLastColumn();

  // احصل على البيانات الموجودة في الصف الثاني
  var existingData = contactsSheet.getRange(rowToShift, 1, 1, contactsSheet.getLastColumn()).getValues();

  // أضف البيانات المحفوظة من الصف الثاني السابق إلى الصف الثالث
  contactsSheet.insertRowAfter(rowToShift);
  contactsSheet.getRange(rowToShift + 1, 1, 1, existingData[0].length).setValues(existingData);

    var emptyRow = [];// تعبئة الصف الجديد الثاني بقيم فارغة
    for (var i = 0; i < numColumns; i++) {
    emptyRow.push("");
  }
  contactsSheet.getRange(2,1,1,numColumns).setValue([emptyRow]);
}

function getValueFromSheet(searchValue) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("schools");
  var range = sheet.getRange("A:A");
  var lastRow = getLastRowInColumnA(1,contactsSheet)
  var value
  
  try {
    var matchRow = range.createTextFinder(searchValue).findNext().getRow();
    value = sheet.getRange("AC" + matchRow).getValue();
    //return value;
  } catch (error) {
    Logger.log("Error: " + error.message);
    value = null; // إذا حدث خطأ، يتم استرجاع قيمة فارغة
  }
  contactsSheet.getRange(lastRow, 5).setValue(value);
}

/**
 * Creates a custom menu titled "My Menu" in the spreadsheet's UI. The menu includes
 * an item "Open Form" that, when clicked, triggers the openForm function.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("My Menu")
    .addItem("Open Form", "openForm").addToUi(); 
}
/**
 * Opens the form  as a modal dialog titled "Contact Details".
 */
function openForm() {
  updateContacts();
  var form = HtmlService.createTemplateFromFile('Index').evaluate();
  form.setWidth(700).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(form, "Contact Details");
}

/**
 * Appends form data (first_name, last_name, etc.) as a new row in the active spreadsheet's data sheet.
 *
 * @param {Object} formObject - The submitted form data object.
 */

function processForm(formObject) {  // تعديل

   var lastRow = getLastRowInColumnA(1,contactsSheet)

    try {
      //contactsSheet.getRange(2, 3).setValue(formObject.jobtitle)
      contactsSheet.getRange(lastRow, 9).setValue(formObject.fName)
      contactsSheet.getRange(lastRow, 12).setValue(formObject.lName)
      contactsSheet.getRange(lastRow, 13).setValue(formObject.phoneno)
      contactsSheet.getRange(lastRow, 15).setValue(formObject.emailaddress1)
  } catch (error) {
    Logger.log('Error appending data: ' + error.message);
  }
  Logger.log(lastRow)
}

/**
 * Includes the content of an external HTML file.
 *
 * @param {string} fileName The name of the HTML file to include.
 * @returns {string} The HTML content of the file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
