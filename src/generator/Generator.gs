var daysCount = 0;

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Generate Template', functionName: 'generateTemplate'}
    ,
    {name: 'Generate Model', functionName: 'generateModel'},
    {name: 'Generate Chart', functionName: 'generateChart'}
  ];
  spreadsheet.addMenu('Scrum', menuItems);
  generateTemplate();
}

/**
 * Creates a new sheet 'Scope' containing sprint issues and work in progress.
 */
function generateTemplate() {
  var spreadsheet = SpreadsheetApp.getActive();
  var scopeSheet = spreadsheet.getActiveSheet();
  scopeSheet.setName('Scope');
   
  userInput();
}

/**
 * Creates a new sheet 'Model' containing sprint progress data for praphs.
 */
function generateModel() {
  var spreadsheet = SpreadsheetApp.getActive();
}

/**
 * Creates a new sheet 'Charts' containing burn down graphs.
 */
function generateChart() {
  var spreadsheet = SpreadsheetApp.getActive();
}

function userInput() {
  var sh = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setHeight('300').setWidth('400');
  var panel = app.createVerticalPanel();
  var startDateBox = app.createDateBox().setName("startDate");
  var endDateBox = app.createDateBox().setName("endDate");
  var button = app.createButton('submit');
  var handler = app.createServerHandler('getDate');
  
  var startDateLabel = app.createLabel("Please enter Sprint start date: ");
  var endDateLabel = app.createLabel("Please enter Sprint end date: ");
  var delimiter = app.createLabel("\n")
  
  handler.addCallbackElement(panel);
  button.addClickHandler(handler);
  panel.add(startDateLabel).add(startDateBox).add(delimiter).add(endDateLabel).add(endDateBox).add(button);
  app.add(panel);
  sh.show(app);
  app.close();
}

function getDate(e){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startDate = new Date(e.parameter.startDate);
  var endDate = new Date(e.parameter.endDate);
  sheet.getRange('A1').setValue(startDate);
  sheet.getRange('A2').setValue(endDate);
  
  drawHeader(sheet, startDate, endDate);
  
  
  var today = new Date(e.parameter.startDate);
  var nextDate = today.setDate(today.getDate() + 1);
  sheet.getRange('A3').setValue(nextDate);
  
  // day of week
//  sheet.getRange('A4').setValue(getDayOfWeek(startDate));
//  sheet.getRange('A5').setValue(getNextDay(startDate));
//  sheet.getRange('A6').setValue(getDaysCount(startDate, endDate));
  
}

function drawHeader(sheet, startDate, endDate) {
  var currentColumn = 1;
  setHeader(sheet, 1, currentColumn++, '#', 60);
  setHeader(sheet, 1, currentColumn++, 'Summary', 200);
  setHeader(sheet, 1, currentColumn++, 'Ext', 20);
  setHeader(sheet, 1, currentColumn++, 'Pilot', 60);
  setHeader(sheet, 1, currentColumn++, 'Copilot', 60);
  setHeader(sheet, 1, currentColumn++, 'Verified', 60);
  setHeader(sheet, 1, currentColumn++, 'Est.', 40);
  setHeader(sheet, 1, currentColumn++, ' ', 20);
  drawWorkingDays(sheet, startDate, endDate, 1, currentColumn);
}

function setHeader(sheet, row, column, value, width) {
  sheet.setColumnWidth(column, width);
  var range = sheet.getRange(row, column);
  range.setValue(value);
  range.setBackgroundRGB(100, 100, 100);
  range.setFontWeight('bold');
  var cell = range.getCell(1, 1);
  cell.setHorizontalAlignment('center');
}

function drawWorkingDays(sheet, startDate, endDate, startRow, startColumn) {
  var limit = 100;
  var date = startDate;
  daysCount = getDaysCount(startDate, endDate);
  var dayNumber = 0;
  while (dayNumber < daysCount && limit > 0) {
    var dayOfWeek = getDayOfWeek(date);
    if (dayOfWeek < 6) {
      sheet.getRange(startRow, startColumn++).setValue(date);
    } else if (dayOfWeek == 7) {
      // making border at end of week
      sheet.getRange(1, startColumn - 1, 100).setBorder(false, false, false, true, false, false);
    } else {
      sheet.getRange(1, startColumn - 1, 100).setBorder(false, false, false, false, false, false);
    }
    date = getNextDay(date);
    dayNumber++;
    limit--;
  }
}

/**
 * Week starts with Monday.
 */
function getDayOfWeek(date) {
  var dayOfWeek = date.getDay();
  if (dayOfWeek == 0) {
    return 7;
  } 
  return dayOfWeek;
}

/**
 * 
 */
function getNextDay(date) {
  var d = new Date(date);
  d.setDate(d.getDate()+1);
  return d;
}

/**
 * Returns days count for provided two dates.
 */
function getDaysCount(startDate, endDate) {
  // set hours, minutes, seconds and milliseconds to 0 if necessary and get number of days
  var startDay = startDate.setHours(0,0,0,0)/(24*3600000);
  var endDay = endDate.setHours(0,0,0,0)/(24*3600000);
  
  // get the difference in days (integer value )
  return parseInt(endDay - startDay) + 1;
}
