//TODO: Create the logic for every 2 weeks. Create a function that determines how many paychecks are in each month into an array and then call that for each month you build.
//TODO: Also change previous balance link to account for lines being added.

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'New Month...', functionName: 'generateNewMonth_'}
    ,{name: 'Several New Months...', functionName: 'generateSeveralNewMonths_'}
  ];
  spreadsheet.addMenu('Budget Utilities', menuItems);
}

function generateNewMonth_(sheetName) {
  if (sheetName != undefined) {
    var name = sheetName;
  }
  else {
    var name = Browser.inputBox('Enter month name', Browser.Buttons.OK_CANCEL);
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  /* Before cloning the sheet, check for existing name */
  var old = ss.getSheetByName(name);
  if (old) {
    throw new Error("The sheet " + name + " already exists");
    return;
  }
  
  var sheet = ss.getSheetByName('Master').copyTo(ss);
  
  SpreadsheetApp.flush(); // Utilities.sleep(2000);
  sheet.setName(name);
  
  /* Make the new sheet active */
  ss.setActiveSheet(sheet);
  
  ss.moveActiveSheet(1);
  
  /* Call Format Sheet Function */
  formatMonthSheet_(name);
  
}

function generateSeveralNewMonths_() {
  /*
  for each month in months array 
  call generateNewMonth_()
  and call formatMonthSheet_()
  */
  
  var numOfMonths = Browser.inputBox('Enter the number of months you want to add', Browser.Buttons.OK_CANCEL);
  
  /* Get the last month created */
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var lastSheet = sheets[0];
  var lastSheetName = lastSheet.getSheetName();
  
  /* Create an array of the next n months to create */
  var monthNames = ["January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"];
  
  var d = new Date();
  //throw new Error("The next month is " + monthNames[d.getMonth()+1]);
  //throw new Error("The last sheet is " + monthNames[monthNames.indexOf(lastSheetName)]);
  
  var monthsToAdd = [];
  for (i=1; i <= numOfMonths; i++) {
    monthsToAdd.push(monthNames[monthNames.indexOf(lastSheetName)+i]);
    generateNewMonth_(monthsToAdd[monthsToAdd.length-1]);
  }
  
  //throw new Error("Months: " + monthsToAdd);
  
  
  
  
  
}

function formatMonthSheet_(sheetToFormat) {
  
  
  
  //var payDay = "" 
  var payDay = Browser.inputBox('Enter day number of your first pay day', Browser.Buttons.OK_CANCEL);
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheet = sheets[0];
  var sheetName = sheet.getSheetName();
  
  //TODO: Get this to work.
  //getPayPeriods(sheetName);

  /* Set the dates that are dynamic */
  var range;
  range = sheet.getRange("A5:A12");
  range.setValue(payDay);
  
  //payDay = "" 
  payDay = Browser.inputBox('Enter day number of your second pay day', Browser.Buttons.OK_CANCEL);
  range = sheet.getRange("A14:A21");
  range.setValue(payDay);
  
  range = sheet.getRange("A5:C26");
  range.sort(1);
  
  
  /* Link the beginning balance to the previous month's ending balance */
  if (sheetName == "January") return;
  var sheetNumber = sheet.getIndex();
  var nextSheet = sheets[1];
  var nextSheetName = nextSheet.getSheetName();
  range = sheet.getRange("C3");
  range.setValue("=" + nextSheetName + "!C27");
  
  SpreadsheetApp.flush();
}


/**
 * This function is used to get the pay period dates for the 
 * month specified. In this version a start date of 1/19/17 is assumed.
 */
function getPayPeriods(monthName) {
  var currentMonth = new Date(monthName + " 1, 2017");
  var currentMonthNumber = currentMonth.getMonth();
  var startDate = new Date('01/19/2017');
  var payDate = new Date(startDate);
  var payMonth = startDate.getMonth();
  var dayOfMonth = startDate.getDate();
  
  
  var periodsToAdd = [];
  while (payMonth <= currentMonthNumber) {
    payDate.setDate(payDate.getDate() + 14);
    if (payMonth = currentMonthNumber) {
      periodsToAdd.push(payDate.getDate());
    }
    payMonth += 1;
  }
  
  throw new Error("The pay months are " + currentMonth);
  
  //TODO: Look at these for a possible samples: 
  //  http://stackoverflow.com/questions/13146418/find-all-the-days-in-a-month-with-date-object
  //  https://answers.acrobatusers.com/auto-populating-pay-period-form-q285209.aspx
  
  return periodsToAdd;
}






















