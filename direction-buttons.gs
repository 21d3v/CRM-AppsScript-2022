/** @OnlyCurrentDoc 
*functions that takes you to one sheet to another
*/

function VaiADash() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Dashboard'), true);
};

function VaiA1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('1 - Lista Consulenti'), true);
};

function VaiA2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('2 - Lista Ufficio'), true);
};

function VaiSi() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DB SI'), true);
};

function VaiNo() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DB NO'), true);
};

function VaiForse() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DB FORSE'), true);
};

function openForm() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('leadForm'), true);
  spreadsheet.getRange('A1').activate();
  clean();
};

function VaiALeadAggiunti() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('leadAggiunti'), true);
};

function VaiAIstruzioni() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Istruzioni'), true);
};

/** Funzione filtra istruzioni in dashboard*/

function hideIstruz() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('22:36').activate();
  spreadsheet.getActiveSheet().hideRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  spreadsheet.getRange('B19').activate();
};

function showIstruz() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('20:36').activate();
  spreadsheet.getActiveSheet().showRows(20, 17);
  spreadsheet.getRange('B19').activate();
};

