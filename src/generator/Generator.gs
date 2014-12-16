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
  
  var limit = 100;
  var datesColumn = 5;
  var date = startDate;
  sheet.getRange(2, datesColumn++).setValue(date);
  while (date != endDate && limit > 0) {
    date = getNextDay(date);
    sheet.getRange(2, datesColumn++).setValue(date);
    limit--;
  }
  
  var today = new Date(e.parameter.startDate);
  var nextDate = today.setDate(today.getDate() + 1);
  sheet.getRange('A3').setValue(nextDate);
  
  // day of week
  sheet.getRange('A4').setValue(getDayOfWeek(startDate));
  sheet.getRange('A5').setValue(getNextDay(startDate));
  
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
