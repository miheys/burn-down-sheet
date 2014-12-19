var daysCount = 0;
var sheet;
var startDate;
var endDate;

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
  sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  startDate = new Date(e.parameter.startDate);
  endDate = new Date(e.parameter.endDate);
  
  drawHeader(startDate, endDate);
  
  var today = new Date(e.parameter.startDate);
  var nextDate = today.setDate(today.getDate() + 1);
}

function drawHeader(startDate, endDate) {
  var currentColumn = 1;
  setHeader(1, currentColumn++, '#', 60);
  setHeader(1, currentColumn++, 'Summary', 600);
  setHeader(1, currentColumn++, 'Ext', 20);
  setHeader(1, currentColumn++, 'Pilot', 60);
  setHeader(1, currentColumn++, 'Copilot', 60);
  setHeader(1, currentColumn++, 'Verified', 60);
  setHeader(1, currentColumn++, 'Est.', 40);
  setHeader(1, currentColumn++, ' ', 20);
  drawBorder(currentColumn - 1, false, true);
  drawWorkingDays(1, currentColumn);
}

function setHeader(row, column, value, width) {
  sheet.setColumnWidth(column, width);
  var range = sheet.getRange(row, column);
  range.setValue(value);
  range.setBackgroundRGB(200, 200, 200);
  range.setFontWeight('bold');
  var cell = range.getCell(1, 1);
  cell.setHorizontalAlignment('center');
}

function drawWorkingDays(startRow, startColumn) {
  var limit = 100;
  var date = startDate;
  
  daysCount = getDaysCount(startDate, endDate);
  var dayNumber = 0;
  while (dayNumber < daysCount && limit > 0) {
    var dayOfWeek = getDayOfWeek(date);
    
    if (dayOfWeek < 6) {
      setHeader(startRow, startColumn, date, 40);
      var range = sheet.getRange(startRow, startColumn);
      range.setValue(Utilities.formatDate(date, "GMT+10", "''dd.MM"));
      startColumn++;
    }
    if (dayOfWeek == 5) {
      // making border at end of week
      drawBorder(startColumn - 1, false, true);
    }
    date = getNextDay(date);
    dayNumber++;
    limit--;
  }
  drawBorder(startColumn, true, false);
}

function drawBorder(column, left, right) {
  var range = sheet.getRange(1, column, 100);
  range.getCell(1, 1).get
  sheet.getRange(1, column, 100).setBorder(false, left, false, right, false, false);
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
