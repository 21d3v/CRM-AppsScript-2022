/** Funzioni relative al Form di inserimento lead e relativa lista*/
function clean() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadForm = ss.getSheetByName('leadForm');
  var codice = ss.getSheetByName('CodiceNomi');
  let rowValues = codice.getRange("CodiceNomi!BB2:BL2").getValues();

  leadForm.getRange("leadForm!D5:N5").setValues(rowValues);
};

function hideForm() {

  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('leadForm'), true);
  spreadsheet.getActiveSheet().hideSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LeadAggiunti'), true);
  spreadsheet.getActiveSheet().hideSheet();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName('1 - Lista Consulenti'), true);
};

function annulla() {
  hideForm();
  clean();
};

function annulla2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('LeadAggiunti'), true);
  spreadsheet.getActiveSheet().hideSheet();
  clean();
};

function addLead() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getRange('A1').activate();
  var leadForm = ss.getSheetByName('leadForm');
  var addedLead = ss.getSheetByName('leadAggiunti');
  var addedLeadLastRow = addedLead.getLastRow();
  let rowValues = leadForm.getRange("leadForm!D5:N5").getValues();

  addedLead.getRange(addedLeadLastRow+1, 1, 1, 11).setValues(rowValues);
 
  hideForm();
  clean();
};

function addLead2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getRange('A1').activate();
  var leadForm = ss.getSheetByName('leadForm');

  leadForm.getRange("M5").setFormula('=TEXTJOIN(" // ";TRUE;H5:J5)');

  var addedLead = ss.getSheetByName('leadAggiunti');
  var addedLeadLastRow = addedLead.getLastRow();

  var addedLead2 = ss.getSheetByName('2 - Lista Ufficio');
  var addedLead2LastRow = addedLead2.getLastRow();

  let rowValues = leadForm.getRange("leadForm!D5:N5").getValues();
  let posizione = leadForm.getRange(5, 3).getValue();
  
  if (posizione == '2 - Lista Ufficio') {
    addedLead2.getRange(addedLead2LastRow+1, 1, 1, 11).setValues(rowValues);
    ss.setActiveSheet(ss.getSheetByName('2 - Lista Ufficio'), true);
    ss.getRange('A1').activate();
    } else {
      addedLead.getRange(addedLeadLastRow+1, 1, 1, 11).setValues(rowValues);
      ss.setActiveSheet(ss.getSheetByName('1 - Lista Consulenti'), true);
      ss.getRange('A1').activate();
    }
  
  hideForm();
  clean();
};
