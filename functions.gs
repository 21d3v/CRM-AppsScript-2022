/** funzione di rimozione duplicati, che viene richiamata durante la funzione 'Trasferisci' e di 
 * formattazione dei paragrafi
*/
function rimDupl() {
  var spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('2 - Lista Ufficio'), true);

  spreadsheet.getRange('B4:B').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('C4:C').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('D4:D').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('d"/"mm"/"yy');
  spreadsheet.getRange('B4:L').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('A4:V').activate();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  .setVerticalAlignment('middle');

  spreadsheet.getRange('A2:M').activate();
  // spreadsheet.setCurrentCell(spreadsheet.getRange('A3'));
  // spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).activate();
  spreadsheet.getActiveRange().removeDuplicates().activate();

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DB SI'), true);
  spreadsheet.getRange('A2:Q').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('A3'));
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).activate();
  spreadsheet.getActiveRange().removeDuplicates().activate();
  spreadsheet.getRange('B4:B').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('C4:C').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('D4:D').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('d"/"mm"/"yy');
  spreadsheet.getRange('H4:H').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('F4:F').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('A4:Q').activate();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  .setVerticalAlignment('middle');

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DB NO'), true);
  spreadsheet.getRange('A2:Q').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('A3'));
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).activate();
  spreadsheet.getActiveRange().removeDuplicates().activate();
  spreadsheet.getRange('B4:B').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('C4:C').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('D4:D').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('d"/"mm"/"yy');
  spreadsheet.getRange('H4:H').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('I4:I').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('A4:Q').activate();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  .setVerticalAlignment('middle');

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DB FORSE'), true);
  spreadsheet.getRange('A2:Q').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('A3'));
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).activate();
  spreadsheet.getActiveRange().removeDuplicates().activate();
  spreadsheet.getRange('B4:B').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('C4:C').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('D4:D').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('d"/"mm"/"yy');
  spreadsheet.getRange('H4:H').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('I4:I').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('A4:Q').activate();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  .setVerticalAlignment('middle');

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DB ERR'), true);
  spreadsheet.getRange('A2:Q').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('A3'));
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).activate();
  spreadsheet.getActiveRange().removeDuplicates().activate();
  spreadsheet.getRange('B4:B').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('C4:C').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('D4:D').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('d"/"mm"/"yy');
  spreadsheet.getRange('H4:H').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('I4:I').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getRange('A4:Q').activate();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  .setVerticalAlignment('middle');

  
};

/**Funzione filtraggio di celle con SI NO FORSE, che viene richiamata in 'Trasferisci' */
function filterRows() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("1 - Lista Consulenti");
  var data = sheet.getDataRange().getValues();

  for(var x = 1; x<data.length; x++) {
    var stato = data[x][11];
    if (stato==='si') {
      sheet.hideRows(x+1);
    } else if (stato==='no') {
      sheet.hideRows(x+1);
    } else if (stato==='forse') {
      sheet.hideRows(x+1);
    } else if (stato==='err') {
      sheet.hideRows(x+1);
    } else if (stato==='inizio') {
      sheet.hideRows(x+1);
    }
  }
};

function filterRows2() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("2 - Lista Ufficio");
  var data = sheet.getDataRange().getValues();

  for(var x = 1; x<data.length; x++) {
    var stato = data[x][13];
    if (stato==='no') {
      sheet.hideRows(x+1);
    // } else if (stato==='si') {
    //   sheet.hideRows(x+1);
    } else if (stato==='forse') {
      sheet.hideRows(x+1);
    } else if (stato==='inizio') {
      sheet.hideRows(x+1);
    }
  }
};

function showAllRows2() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("2 - Lista Ufficio");
  sheet.showRows(1, sheet.getMaxRows());
}

function showAllRows() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("1 - Lista Consulenti");
  sheet.showRows(1, sheet.getMaxRows());
}

function formatCell() {
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName("2 - Lista Ufficio");;
  spreadsheet.getRange('H:V').activate();
  spreadsheet.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  spreadsheet.setCurrentCell(spreadsheet.getRange('A2'));
};

/** Funzioni di trasferimento dati a lista avanzata, db si no e forse */
function Trasferisci1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lista1 = ss.getSheetByName('1 - Lista Consulenti');
  var lista1LastRow = lista1.getLastRow();
  let lista1Range = lista1.getSheetValues(3, 1, lista1LastRow, 13);

  var lista2 = ss.getSheetByName('2 - Lista Ufficio');
  var lista2LastRow = lista2.getLastRow();
  let lista2Counter = 1;

  var dbNo = ss.getSheetByName('DB NO');
  var dbNoLastRow = dbNo.getLastRow();
  let dbNoCounter = 1;
  
  var dbForse = ss.getSheetByName('DB FORSE');
  var dbForseLastRow = dbForse.getLastRow();
  let dbForseCounter = 1;

  var dbErr = ss.getSheetByName('DB ERR');
  var dbErrLastRow = dbErr.getLastRow();
  let dbErrCounter = 1;

  for (var i = 4; i <= lista1Range.length; i++) {
    let stato = lista1.getRange(i,12).getValue();
    let rowValues = lista1.getRange(i, 1, 1, 13).getDisplayValues();
    // Logger.log(rowValues);

      if (stato == 'si') {
        lista2.getRange(lista2LastRow+lista2Counter, 1, 1, 13).setValues(rowValues);
        lista2Counter++;

      } else if (stato == 'no') {
        dbNo.getRange(dbNoLastRow+dbNoCounter, 1, 1, 13).setValues(rowValues);
        dbNoCounter++;
        
      } else if (stato == 'forse') {
        dbForse.getRange(dbForseLastRow+dbForseCounter, 1, 1, 13).setValues(rowValues);
        dbForseCounter++;

      } else if (stato == 'err') {
        dbErr.getRange(dbErrLastRow+dbErrCounter, 1, 1, 13).setValues(rowValues);
        dbErrCounter++;
      }
  };
  SpreadsheetApp.flush();
  rimDupl();
  filterRows();
  formatCell();
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName("1 - Lista Consulenti");
  spreadsheet.getRange('A2').activate();
};

function Trasferisci2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lista2 = ss.getSheetByName('2 - Lista Ufficio');
  var lista2LastRow = lista2.getLastRow();
  let lista2Range = lista2.getSheetValues(3, 1, lista2LastRow, 17);

  // var dbSi = ss.getSheetByName('DB SI');
  // var dbSiLastRow = dbSi.getLastRow();
  // var dbSiCounter = 1;
  var dbNo = ss.getSheetByName('DB NO');
  var dbNoLastRow = dbNo.getLastRow();
  var dbNoCounter = 1;
  var dbForse = ss.getSheetByName('DB FORSE');
  var dbForseLastRow = dbForse.getLastRow();
  var dbForseCounter = 1;

  for (var i = 1; i <= lista2Range.length; i++) {
    let stato = lista2.getRange(i,14).getValue();
    let rowValues = lista2.getRange(i, 1, 1, 17).getValues();
      if (stato == 'no') {
        dbNo.getRange(dbNoLastRow+dbNoCounter, 1, 1, 17).setValues(rowValues);
        dbNoCounter++;
      // } else if (stato == 'si') {
      //   dbSi.getRange(dbSiLastRow+dbSiCounter, 1, 1, 17).setValues(rowValues);
      //   dbSiCounter++;
      } else if (stato == 'forse') {
        dbForse.getRange(dbForseLastRow+dbForseCounter, 1, 1, 17).setValues(rowValues);
        dbForseCounter++;
      }
  }
  rimDupl();
  filterRows2();
  var spreadsheet = SpreadsheetApp.getActive().getSheetByName('2 - Lista Ufficio');
  spreadsheet.getRange('A2').activate();
};
