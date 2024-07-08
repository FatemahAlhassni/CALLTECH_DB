
/**
 * Fatemah Alshareef - technical Support specialist - CALLTECH 
 * Constants
 */
const DATA_SHEET = "Sheet6";

function getCombinedData() { // دمج اسم المدرسة مع رمز المدرسة بالصيغة المطلوبة
  var destinationSS = SpreadsheetApp.openById("1wfV7hZ4ZBrwnweqGJBFF7o2T50NY7fRoNBMo61ir6dc");
  var sheet = destinationSS.getSheetByName("schools");
  const destlastRow1 = sheet.getLastRow();
  var dataRange = sheet.getRange(`A2:B${destlastRow1}`); // لتحديد أخر صف يحتوي على بيانات
  var values = dataRange.getValues();
 
  var combinedData = [];
  for (var i = 0; i < values.length; i++) {
    combinedData.push([values[i][0] + " | " + values[i][1]]);
  }
  return combinedData;
}

function createNamedRangesFromDifferentSheets() {
  var spreadsheet = SpreadsheetApp.openById("1wfV7hZ4ZBrwnweqGJBFF7o2T50NY7fRoNBMo61ir6dc");
    
  var sheetNames = ["admins&offices", "admins&offices", "clients" ,"schools","schools","clients"];
  var ranges =  ["B2:B12",// Adminstration
                "C2:C11",// Office
                "A2:A18",//Contractor
                "AI2:AI4",//Project_Name
                "AF2:AF7",//Entity_name
                "A2:A" +spreadsheet.getSheetByName(sheetNames[5]).getLastRow()]; //Subcontractor

  var names =  ["School_code_and_name", "Adminstration", "Office", "Contractor", "Project_Name", "Entity_name", "Subcontractor"];

  // Loop through the ranges and set the values
  for (var i = 1; i < sheetNames.length; i++) {
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var namedRanges = ["Adminstration", "Office", "Contractor", "Project_Name", "Entity_name", "Subcontractor"];
  var dropdownOptions = [];  
    
  var combinedData = getCombinedData();
  dropdownOptions = dropdownOptions.concat(combinedData.flat());
  
  namedRanges.forEach(function(namedRange) {
    var range = sheet.getRangeByName(namedRange);
    if (range) {
      var values = range.getValues();
      // تجميع جميع القيم في مصفوفة واحدة
      dropdownOptions = dropdownOptions.concat(values.flat());
    }
  });

  removeDuplicates();
  return dropdownOptions;
}


function addDataToSheet(value,column) {   // تعدلت
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = getLastRowInColumn(2) // استدعاء الدالة للحصول على رقم الصف الأخير في العمود A
  var range = sheet.getRange(lastRow, column); // الصف الأخير + 1 في العمود 
  range.setValue(value);
}


function getLastRowInColumn(column) {  // اضافة جديد
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
   // رقم العمود الذي ترغب في الوصول إلى الصف الأخير منه column
  var data = sheet.getDataRange().getValues();

  for (var row = data.length - 1; row >= 0; row--) {
    if (data[row][column - 1]) {
      break;
    }
  }

  var lastRow = row + 1;
  //Logger.log("مكان الصف الأخير في العمود: " + lastRow);
  return(lastRow+1)
}


/**
 * Creates a custom menu titled "My Menu" in the spreadsheet's UI. The menu includes 
 * an item "Open Form" that, when clicked, triggers the openForm function.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("My Menu")
    .addItem("Open Form", "openForm")
    .addItem('Salse Sheet', 'salseSheet')
    .addToUi();
}

/**
 * Opens the form  as a modal dialog titled "Contact Details". 
 */
function openForm() {
  var form = HtmlService.createTemplateFromFile('Index').evaluate();
  form.setWidth(700).setHeight(400); 
  SpreadsheetApp.getUi().showModalDialog(form, "Contact Details");
}

function salseSheet() {
  var salseHtml = HtmlService.createHtmlOutputFromFile('salse')
    .setWidth(700)
    .setHeight(400);
     SpreadsheetApp.getUi()
    .showModalDialog(salseHtml,'Salse Sheet');
}

function dropdownOptions_ScooleCode() {
  var spreadsheet = SpreadsheetApp.openById("1wfV7hZ4ZBrwnweqGJBFF7o2T50NY7fRoNBMo61ir6dc");
  var dataSheet = spreadsheet.getSheetByName("schools"); // اسم صفحة البيانات

  var namedRange = ["schoolcodeToSalse"]
  var dropdownOptions = []


  namedRange.forEach(function(namedRange) {
  var range = dataSheet.getRange(namedRange);
 
    if (range) {
      var values = range.getValues();
      // تجميع جميع القيم في مصفوفة واحدة
      dropdownOptions = dropdownOptions.concat(values.flat());
    }
    })
    return dropdownOptions
}/////////////////////////////////////////////////////////////

function getNamebyCode(sc_code) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadsheet.getSheetByName("schools"); // اسم صفحة البيانات
  var salseSheet = spreadsheet.getSheetByName("salse");
  
  var sc_name

  var sc_nameR = dataSheet.getRange("A2:A");
  for (var i = 0; i < sc_nameR.getNumRows(); i++) {
     if (sc_nameR.getCell(i + 1, 1).getValue() === sc_code) {
      var rowNum = (i+2)
      var cellInRowB = dataSheet.getRange("B" + rowNum);
      sc_name = cellInRowB.getValue();
      //salseSheet.appendRow([sc_code, sc_name]);
      break;
  }
  }
  salseSheet.appendRow([sc_code, sc_name]);
 // Logger.log (sc_name)
}

function dropdownOptions_ScooleName() {
  var spreadsheet = SpreadsheetApp.openById("1wfV7hZ4ZBrwnweqGJBFF7o2T50NY7fRoNBMo61ir6dc");
  var dataSheet = spreadsheet.getSheetByName("schools"); // اسم صفحة البيانات

  var namedRange = ["schoolnameToSalse"]
  var dropdownOptions = []

  namedRange.forEach(function(namedRange) {
  var range = dataSheet.getRange(namedRange);
 
    if (range) {
      var values = range.getValues();
      // تجميع جميع القيم في مصفوفة واحدة
      dropdownOptions = dropdownOptions.concat(values.flat());
    }
    })
    return dropdownOptions
}

function getCodebyName(sc_name) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = spreadsheet.getSheetByName("schools"); // اسم صفحة البيانات
  var salseSheet = spreadsheet.getSheetByName("salse");

  var sc_code

  var sc_nameR = dataSheet.getRange("B2:B");
  for (var i = 0; i < sc_nameR.getNumRows(); i++) {
    if (sc_nameR.getCell(i + 1, 1).getValue() === sc_name) {

    var rowNum = (i+2)
    var cellInRowA = dataSheet.getRange("A" + rowNum);
    sc_code = cellInRowA.getValue();
    break;
  }}
   salseSheet.appendRow([sc_code, sc_name]);
   //Logger.log([sc_code, sc_name]);
}
/**
 * Appends form data (first_name, last_name, etc.) as a new row in the active spreadsheet's data sheet.
 *
 * @param {Object} formObject - The submitted form data object.
 */

function processForm(formObject) {  // تعديل
  const dataSheet = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET);

  try {
    const selectedRow = (getLastRowInColumn(2));

    dataSheet.getRange(selectedRow , 2, 1, 5).//  تحديد العمود رقم 2، لإدخال صف 1 من البيانات، 4 هي عدد الأعمدة
    setValues([
      //new Date().toLocaleString(),
      [//formObject.schoolcode,
       formObject.fName,
       formObject.lName,
       formObject.jobtitle,
       formObject.phoneno, 
       formObject.emailaddress1,
     //  formObject.dropdown
       ],
      //Add your new field values here
    ]);
  } catch (error) {
    Logger.log('Error appending data: ' + error.message);
  }
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