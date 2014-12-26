// constants
var MAX_ROWS = 100;
var STORY_MARKER = 's';
var COLUMNS_INITIAL_COUNT = 8;

// app objects
//var spreadsheet;
//var variablesSheet;
//var scopeSheet;

// local variables
var daysCount = 0;
var workingDaysCount = 0;
var startDate;
var endDate;
var columnsCount;

/***********************************************
 * Menu items functions                        *
 ***********************************************/

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Generate Template', functionName: 'generateTemplate'},
    {name: 'Process stories', functionName: 'processStories'},
    {name: 'Generate Model', functionName: 'generateModel'},
    {name: 'Generate Chart', functionName: 'generateChart'}
  ];
  spreadsheet.addMenu('Scrum', menuItems);
  
  createVariablesSheet();
  
//  generateTemplate();
  
  // TODO: MVO: remove
  processStories();
}

/***********************************************
 * Functions for initializing local variables  *
 ***********************************************/
function createVariablesSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var scopeSheet = spreadsheet.getSheetByName('Scope');
  if (scopeSheet == null) {
    scopeSheet = spreadsheet.getActiveSheet();
    scopeSheet.setName('Scope');
  }
  var variablesSheet = spreadsheet.getSheetByName('Variables');
  if (variablesSheet == null) {
    variablesSheet = spreadsheet.insertSheet('Variables');
  }
  variablesSheet.hideSheet();
}
function initVariables() {
//  spreadsheet = SpreadsheetApp.getActive();
//  scopeSheet = spreadsheet.getSheetByName('Scope');
//  variablesSheet = spreadsheet.getSheetByName('Variables');
  if (variablesSheet() == null) {
    SpreadsheetApp.getUi().alert('You should run Generate Template first');
    return;
  }
  daysCount = readDaysCount();
  workingDaysCount = readWorkingDaysCount();
  startDate = readStartDate();
  endDate = readEndDate();
  columnsCount = readColumnsCount();
}
function writeDaysCount(daysCount) {
  variablesSheet().getRange(1, 1).setValue('daysCount');
  variablesSheet().getRange(1, 2).setValue(daysCount);
}
function readDaysCount() {
  return variablesSheet().getRange(1, 2).getValue();
}
function writeWorkingDaysCount(workingDaysCount) {
  variablesSheet().getRange(2, 1).setValue('workingDaysCount');
  variablesSheet().getRange(2, 2).setValue(workingDaysCount);
}
function readWorkingDaysCount() {
  return variablesSheet().getRange(2, 2).getValue();
}
function writeStartDate(startDate) {
  variablesSheet().getRange(3, 1).setValue('startDate');
  variablesSheet().getRange(3, 2).setValue(startDate);
}
function readStartDate() {
  return variablesSheet().getRange(3, 2).getValue();
}
function writeEndDate(endDate) {
  variablesSheet().getRange(4, 1).setValue('endDate');
  variablesSheet().getRange(4, 2).setValue(endDate);
}
function readEndDate() {
  return variablesSheet().getRange(4, 2).getValue();
}
function writeColumnsCount(columnsCount) {
  variablesSheet().getRange(5, 1).setValue('columnsCount');
  variablesSheet().getRange(5, 2).setValue(columnsCount);
}
function readColumnsCount() {
  return variablesSheet().getRange(5, 2).getValue();
}
function variablesSheet() {
  return spreadsheet().getSheetByName('Variables');
}
function scopeSheet() {
  return spreadsheet().getSheetByName('Scope');
}
function spreadsheet() {
  return SpreadsheetApp.getActive();
}

/**
 * Iterates over input stories and subtasks. Row marked with 's' are considered to be a story.
 * Creates total cells and completes 'Scope' sheet.
 */
function processStories() {
  initVariables();
  var row = 2;
  var storyRow = 2;
  while (row < MAX_ROWS) {
    processStoryItem(row);
    row++;
  }
  alert('Processing stories completed');
}

function processStoryItem(row) {
  var isStory = scopeSheet().getRange(row, 3).getValue() == STORY_MARKER;
  if (isStory) {
    drawStoryHeader(row);
  }
}

/**
 * Creates a new sheet 'Scope' containing sprint issues and work in progress.
 */
function generateTemplate() {
  userInput();
}

/**
 * Creates a new sheet 'Model' containing sprint progress data for praphs.
 */
function generateModel() {
  initVariables();
}

/**
 * Creates a new sheet 'Charts' containing burn down graphs.
 */
function generateChart() {
  initVariables();
}

function userInput() {
  var sh = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setHeight('300').setWidth('400');
  var panel = app.createVerticalPanel();
  var startDateBox = app.createDateBox().setName("startDate");
  var endDateBox = app.createDateBox().setName("endDate");
  var button = app.createButton('submit');
  var handler = app.createServerHandler('drawTemplate');
  
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

function drawTemplate(e){
  startDate = new Date(e.parameter.startDate);
  if (startDate == null) {
    startDate = new Date("2014/12/03 14:25:58 +00");
  }
  writeStartDate(startDate);
  if (endDate == null) {
    endDate = new Date("2014/12/22 14:25:58 +00");
  }
  endDate = new Date(e.parameter.endDate);
  writeEndDate(endDate);
  
  drawHeader(startDate, endDate);
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
  scopeSheet().setFrozenRows(1);
  scopeSheet().setFrozenColumns(4);
  
  // extending sheet to MAX_ROWS size
  var lastRow = scopeSheet().getLastRow();
  while (lastRow < MAX_ROWS) {
    scopeSheet().appendRow(['0']);
    lastRow = scopeSheet().getLastRow();
  }
  
  drawBorder(currentColumn - 1, false, true);
  drawWorkingDays(1, currentColumn);
  deleteObsoleteColumns();
}

function setHeader(row, column, value, width) {
  scopeSheet().setColumnWidth(column, width);
  var range = scopeSheet().getRange(row, column);
  range.setValue(value);
  range.setBackgroundRGB(200, 200, 200);
  range.setFontWeight('bold');
  var cell = range.getCell(1, 1);
  cell.setHorizontalAlignment('center');
}

function drawStoryHeader(row) {
  var range = scopeSheet().getRange(row, 1, 1, columnsCount);
  range.setBackgroundRGB(220, 220, 220);
  range.setFontWeight('bold');
}

function drawWorkingDays(startRow, startColumn) {
  var limit = MAX_ROWS;
  var date = startDate;
  
  // Display a sidebar with custom HtmlService content.
 var htmlOutput = HtmlService
     .createHtmlOutput('<p>A change of speed, a change of style...</p>')
     .setSandboxMode(HtmlService.SandboxMode.IFRAME)
     .setTitle('TAF add-on');
 SpreadsheetApp.getUi().showSidebar(htmlOutput);
  
  // Display a sidebar with custom UiApp content.
 var uiInstance = UiApp.createApplication().setTitle('TAF add-on');
 uiInstance.add(uiInstance.createLabel('Please add your Stories with Subtasks in two columns. In third column mark story items with "s".'));
 SpreadsheetApp.getUi().showSidebar(uiInstance);
  daysCount = getDaysCount(startDate, endDate);
  writeDaysCount(daysCount);
  var dayNumber = 0;
  var workingDaysCount = 0;
  while (dayNumber < daysCount && limit > 0) {
    var dayOfWeek = getDayOfWeek(date);
    
    if (dayOfWeek < 6) {
      setHeader(startRow, startColumn, date, 40);
      var range = scopeSheet().getRange(startRow, startColumn);
      range.setValue(Utilities.formatDate(date, "GMT+10", "''dd.MM"));
      startColumn++;
      workingDaysCount++;
    }
    if (dayOfWeek == 5) {
      // making border at end of week
      drawBorder(startColumn - 1, false, true);
    }
    date = getNextDay(date);
    dayNumber++;
    limit--;
  }
  writeWorkingDaysCount(workingDaysCount);
  columnsCount = COLUMNS_INITIAL_COUNT + workingDaysCount;
  writeColumnsCount(columnsCount);
  drawBorder(startColumn, true, false);
}

function deleteObsoleteColumns() {
  var columnsCount = readColumnsCount();
  var counter = 100;
  var scopeSheet = scopeSheet();
  while (counter > 0) {
    try {
      scopeSheet.deleteColumns(columnsCount + 1, 1);
    } catch(e) {
      // ignore
    }
    counter--;
  }
}

function drawBorder(column, left, right) {
  scopeSheet().getRange(1, column, MAX_ROWS).setBorder(false, left, false, right, false, false);
}

function alert(message) {
  SpreadsheetApp.getUi().alert(message);
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
