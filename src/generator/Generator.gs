// constants
var MAX_ROWS = 100;
var STORY_MARKER = 's';
var EMPTY_ROW_MARKER = '0';
var COLUMNS_INITIAL_COUNT = 8;
var ESTIMATES_COLUMN = 7;

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

var storyRows = [];
var subtaskRows = [];

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
  
//generateTemplate();
  
  // TODO: MVO: remove
//  processStories();
  generateModel();
}

/**
 * Creates a new sheet 'Model' containing sprint progress data for praphs.
 */
function generateModel() {
  initVariables();
  createModelSheet();
  updateModelHeader();
  updateModel();
}
function createModelSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var modelSheet = spreadsheet.getSheetByName('Model');
  if (modelSheet == null) {
    modelSheet = spreadsheet.duplicateActiveSheet();
    modelSheet.setName('Model');
    modelSheet.deleteColumn(1);
    modelSheet.getRange(1, 2, readRowsCount(), 1).copyTo(modelSheet.getRange(1, 7, readRowsCount(), 1))
  }
}
function updateModelHeader() {
  modelSheet().getRange(1, 1, 1, 7).clearContent();
  modelSheet().getRange(1, 3).setValue('Init');
  modelSheet().getRange(1, 4).setValue('Real');
  modelSheet().getRange(1, 5).setValue('Done');
  modelSheet().getRange(1, 6).setValue('Δ');
  modelSheet().setFrozenColumns(7);
}
function updateModel() {
  readStoriesRows(7);
  var row = 2;
  // unified formatting for both stories and subtasks
  while (row <= readRowsCount() - 2) {
    var range = modelSheet().getRange(row, 1);
    range.setFormula('=Scope!' + range.getA1Notation());
    row++;
  }
  // processing stories
  for (index in storyRows) {
    var storyRow = storyRows[index];
    var estimatesLocation = modelSheet().getRange(storyRow, ESTIMATES_COLUMN).getA1Notation();
    var subtaskDoneRange = modelSheet().getRange(subtaskRows[index][0], 5, subtaskRows[index].length).getA1Notation();
    var storyCompletionFormula = '=IF((Scope!' + estimatesLocation + '=0);"";(COUNTIF(' + subtaskDoneRange + ';"=TRUE")/COUNTA(' + subtaskDoneRange + ')))';
    modelSheet().getRange(storyRow, 2).setFormula(storyCompletionFormula);
  }
  
  for (story in subtaskRows) {
    var subtasks = subtaskRows[story];
  }
}
function readStoriesRows(storyMarkerColumn) {
  storyRows = [];
  subtaskRows = [];
  var row = 2;
  var storyRow = 2;
  var storyIndex = 0;
  while (row <= readRowsCount()) {
    var isStory = modelSheet().getRange(row, storyMarkerColumn).getValue() == STORY_MARKER;
    if (isStory) {
      storyRows = storyRows.concat(row);
      var subtasksStarted = storyRow + 1;
      var subtasksStopped = row - 1;
      var storySubtasks = [];
      while (subtasksStopped - subtasksStarted >= 0) {
        storySubtasks = storySubtasks.concat(subtasksStarted);
        subtasksStarted++;
      }
      if (row > 2) {
        subtaskRows[storyIndex] = storySubtasks;
        storyIndex++;
      }
      storyRow = row;
    }
    row++;
  }
  var subtasksStarted = storyRow + 1;
  var subtasksStopped = row - 1;
  var storySubtasks = [];
  while (subtasksStopped - subtasksStarted >= 0) {
    storySubtasks = storySubtasks.concat(subtasksStarted);
    subtasksStarted++;
  }
  subtaskRows[storyIndex] = storySubtasks;
}

/**
 * Iterates over input stories and subtasks. Row marked with 's' are considered to be a story.
 * Creates total cells and completes 'Scope' sheet.
 */
function processStories() {
  initVariables();
  cleanFormulas();
  var row = 2;
  var storyRow = 2;
  while (row < MAX_ROWS) {
    var isStory = processStoryItem(row, storyRow);
    if (isStory[0]) {
      storyRow = row;
      storyRows = storyRows.concat(row);
    }
    var shouldBreak = isStory[1];
    if (shouldBreak) {
      trimRows(row);
      MAX_ROWS = row;
    }
    row++;
  }
  
  writeRowsCount(row - 2);
  appendTotalEstimate(row - 1);
  appendTotalDevelopers(row);
  addConditionalFormatting();
  
  alert('Processing stories completed');
}

function addConditionalFormatting() {
  ScriptApp.newTrigger('updateConditionalFormatting')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
}

function updateConditionalFormatting(e) {
  
  // passing single cell changes only
  var value = e.value;  
  if (value != null) {
    var range = e.range;
    
    // passing Scope sheet changes only
    if (range.getSheet().getName() == 'Scope') {
      
      // passing work in progress range changes only
      if (isCellInRange(range, 3, 100, COLUMNS_INITIAL_COUNT, readColumnsCount())) {
        if (value == 0) {
          range.setBackgroundRGB(0, 200, 0);
          scopeSheet().getRange(range.getRow(), 7).setBackgroundRGB(0, 200, 0);
        } else if(value > 0) {
          range.setBackgroundRGB(255, 255, 255);
          scopeSheet().getRange(range.getRow(), 7).setBackgroundRGB(255, 255, 255);
        }
      }
    }
  }
}

function isCellInRange(cell, fromRow, toRow, fromColumn, toColumn) {
  var row = cell.getRow();
  var column = cell.getColumn();
  return row >= fromRow && row <= toRow && column >= fromColumn && column <= toColumn;
}

function trimRows(row) {
  scopeSheet().deleteRows(row, scopeSheet().getLastRow() - row + 1);
  var i = 0;
  while (i < 100) {
    try {
      i++;
      scopeSheet().deleteRows(row, 2);
    } catch(e) {
      // ignore
    }
  }
}

function appendTotalEstimate(row) {
  scopeSheet().appendRow(['','','','Total estimate:','','','']);
  scopeSheet().getRange(row, 1, 2, columnsCount).setHorizontalAlignment('right');
  scopeSheet().getRange(row, 1, 2, columnsCount).setFontWeight('bold');
  scopeSheet().getRange(row, 1, 2, columnsCount).setBackgroundRGB(200, 200, 200);
  for (storyRow in storyRows) {
    var storyEstimateAddress = scopeSheet().getRange(storyRows[storyRow], ESTIMATES_COLUMN).getA1Notation();
    var totalEstimatesCell = scopeSheet().getRange(row, ESTIMATES_COLUMN);
    totalEstimatesCell.setFormula(totalEstimatesCell.getFormula() + ' + ' + storyEstimateAddress);
  }
  var column = COLUMNS_INITIAL_COUNT + 1;
  while (column <= readColumnsCount()) {
    var start = scopeSheet().getRange(2, column).getA1Notation();
    var stop = scopeSheet().getRange(row - 1, column).getA1Notation()
    scopeSheet().getRange(row, column).setFormula('=COUNTIF(' + start + ':' + stop + ';">0")');
    column++;
  }
}

function appendTotalDevelopers(row) {
  scopeSheet().appendRow(['','','','Developer days available:','','']);
  scopeSheet().getRange(row, 2, 1, 3).merge();
  var developersCount = readDevelopersCount();
  scopeSheet().getRange(row, COLUMNS_INITIAL_COUNT + 1, 1, readWorkingDaysCount()).setValue(developersCount);
  var start = scopeSheet().getRange(row, COLUMNS_INITIAL_COUNT + 1).getA1Notation();
  var stop = scopeSheet().getRange(row, readColumnsCount()).getA1Notation();
  scopeSheet().getRange(row, ESTIMATES_COLUMN).setFormula('=SUM(' + start + ':' + stop + ')');
}

function processStoryItem(row, storyRow) {
  var isStory = scopeSheet().getRange(row, 3).getValue() == STORY_MARKER;
  var isSubtask = !isStory && scopeSheet().getRange(row, 1).getValue() != EMPTY_ROW_MARKER && scopeSheet().getRange(row, 1).getValue() != '';
  if (isStory) {
    drawStoryHeader(row);
  } else if(isSubtask) {
    updateStoryFormula(row, storyRow);
  } else {
    return [false, true];
  }
  return [isStory, false];
}

function drawStoryHeader(row) {
  var range = scopeSheet().getRange(row, 1, 1, columnsCount);
  range.setBackgroundRGB(220, 220, 220);
  range.setFontWeight('bold');  
}

function updateStoryFormula(row, storyRow) {
  var currentFormula = scopeSheet().getRange(storyRow, ESTIMATES_COLUMN).getFormula();
  if (currentFormula == '') {
    currentFormula = '0';
  }
  var subtaskEstimateCell = scopeSheet().getRange(row, ESTIMATES_COLUMN).getA1Notation();
  scopeSheet().getRange(storyRow, ESTIMATES_COLUMN).setFormula(currentFormula + ' + ' + subtaskEstimateCell);
}

function cleanFormulas() {
  scopeSheet().getRange(2, ESTIMATES_COLUMN, 100).setFormula('');
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
  
  // TODO: remove
  variablesSheet.showSheet();
}
function initVariables() {
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
  writeKeyValue(1, 'daysCount', daysCount);
}
function readDaysCount() {
  return readKeyValue(1);
}
function writeWorkingDaysCount(workingDaysCount) {
  writeKeyValue(2, 'workingDaysCount', workingDaysCount);
}
function readWorkingDaysCount() {
  return readKeyValue(2);
}
function writeStartDate(startDate) {
  writeKeyValue(3, 'startDate', startDate);
}
function readStartDate() {
  return readKeyValue(3);
}
function writeEndDate(endDate) {
  writeKeyValue(4, 'endDate', endDate);
}
function readEndDate() {
  return readKeyValue(4);
}
function writeColumnsCount(columnsCount) {
  writeKeyValue(5, 'columnsCount', columnsCount);
}
function readColumnsCount() {
  return readKeyValue(5);
}
function writeDevelopersCount(developersCount) {
  writeKeyValue(6, 'developersCount', developersCount);
}
function readDevelopersCount() {
  return readKeyValue(6);
}
function writeRowsCount(rowsCount) {
  writeKeyValue(7, 'rowsCount', rowsCount);
}
function readRowsCount() {
  return readKeyValue(7);
}
function writeKeyValue(row, key, value) {
  variablesSheet().getRange(row, 1).setValue(key);
  variablesSheet().getRange(row, 2).setValue(value);
}
function readKeyValue(row) {
  return variablesSheet().getRange(row, 2).getValue();
}
function variablesSheet() {
  return spreadsheet().getSheetByName('Variables');
}
function scopeSheet() {
  return spreadsheet().getSheetByName('Scope');
}
function modelSheet() {
  return spreadsheet().getSheetByName('Model');
}
function spreadsheet() {
  return SpreadsheetApp.getActive();
}

/**
 * Creates a new sheet 'Scope' containing sprint issues and work in progress.
 */
function generateTemplate() {
  userInput();
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
  var developersCountBox = app.createTextBox().setName("developersCount");
  var button = app.createButton('submit');
  var handler = app.createServerHandler('drawTemplate');
  
  var startDateLabel = app.createLabel("Please enter Sprint start date: ");
  var endDateLabel = app.createLabel("Please enter Sprint end date: ");
  var developersCountLabel = app.createLabel("Please enter available developers count: ");
  var delimiter = app.createLabel("\n\n");
  
  handler.addCallbackElement(panel);
  button.addClickHandler(handler);
  panel.add(startDateLabel).add(startDateBox).add(delimiter);
  panel.add(endDateLabel).add(endDateBox).add(delimiter);
  panel.add(developersCountLabel).add(developersCountBox).add(delimiter);
  panel.add(button);
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
  
  var developersCount = +e.parameter.developersCount;
  writeDevelopersCount(developersCount);
  
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
//  alert('Drawing headers completed');
  
  scopeSheet().setFrozenRows(1);
  scopeSheet().setFrozenColumns(4);
//  alert('Freezing completed');
  
  // extending sheet columns
  var lastColumn = scopeSheet().getLastColumn();
  var maxLimit = 100;
  while (lastColumn < MAX_ROWS && maxLimit > 0) {
    maxLimit--;
    scopeSheet().insertColumnAfter(lastColumn);
    lastColumn = scopeSheet().getLastColumn();
  }
  
  // extending sheet rows
  var lastRow = scopeSheet().getLastRow();
  maxLimit = 100;
  while (lastRow < MAX_ROWS && maxLimit > 0) {
    maxLimit--;
    scopeSheet().appendRow([EMPTY_ROW_MARKER]);
    lastRow = scopeSheet().getLastRow();
  }
//  alert('Extending rows completed');
  
  drawBorder(currentColumn - 1, false, true);
//  alert('Drawing border completed');
  
  drawWorkingDays(1, currentColumn);
  alert('Drawing working days completed');
  
  deleteObsoleteColumns();
//  alert('Deletion obsolete columns completed');
  alert('Drawing template completed');
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
  while (counter > 0) {
    try {
      counter--;
      scopeSheet().deleteColumns(columnsCount + 1, 1);
    } catch(e) {
      // ignore
    }
  }
  alert('Max Column: ' + scopeSheet().getMaxColumns());
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
