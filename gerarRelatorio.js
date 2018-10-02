function gerarRelatorio() {
  var ss = SpreadsheetApp.getActive();
  
  cloneTemplate();
  
  ss.getSheetByName("Loading").showSheet().activate();
  
  var reportSheet = ss.getSheetByName('Relatório');
  var tasksSheet = ss.getSheetByName('Tarefas');
  
  reportSheet.hideSheet();  
  
  // tasksSheet.activate();
  var dataRange = tasksSheet.getDataRange();
  var rangeValues = dataRange.getValues();
  var numberOfRows = dataRange.getLastRow();
  
  for(i = 2; i < numberOfRows; i++){
    
    var progression = (i - 2) / (numberOfRows - 2);
    ss.getSheetByName("Loading").getRange('B1').setValue(progression);
    
    var situacao = rangeValues[i][3];
    
    switch(situacao){
      case "Metas":
        copyPasteTasks(i+1, 1);
        break;
        
      case "Concluídas":
        copyPasteTasks(i+1, 2);
        break;
        
      case "À espera de execução":
        copyPasteTasks(i+1, 3);
    }

  }
  
  reportSheet.deleteColumn(4);
  
  reportSheet.showSheet();
  SpreadsheetApp.getActive().setActiveSheet(reportSheet);
  SpreadsheetApp.getActive().getSheetByName("Loading").hideSheet();
  
}

function copyPasteTasks(rowIndex, taskCondition){
  var ss = SpreadsheetApp.getActive();
  var reportSheet = ss.getSheetByName('Relatório');
  
  
  var aux = reportSheet.getRange('D1');
  
  for (j = 0; j < taskCondition; j++){
    aux = aux.getNextDataCell(SpreadsheetApp.Direction.DOWN);
  }
  
  reportSheet.insertRowAfter(aux.getLastRow());
  aux = aux.offset(1,0);
  var lastRow = aux.getLastRow();
  
  aux.offset(-1, 0).moveTo(aux);
  aux = aux.getNextDataCell(SpreadsheetApp.Direction.PREVIOUS);
  ss.getRange('Tarefas!A' + rowIndex + ':C' + rowIndex).copyTo(aux, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
  reportSheet.setRowHeight(lastRow, 80);
  var taskRow = reportSheet.getRange('A' + lastRow + ':C' + lastRow);
  taskRow.setVerticalAlignment('middle')
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  SpreadsheetApp.flush();
  
}

function cloneTemplate(){
  var name = "Relatório";
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Template").copyTo(ss);
  
  /* Before cloning the sheet, delete any previous copy */
  var old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old); // or old.setName(new Name);
  
  SpreadsheetApp.flush(); // Utilities.sleep(2000);
  sheet.setName(name);
  
}









