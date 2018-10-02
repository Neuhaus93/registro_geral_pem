function gerarRelatorio() {
  var ss = SpreadsheetApp.getActive();
  
  cloneTemplate();
  
  var reportSheet = ss.getSheetByName('Relatório');
  var tasksSheet = ss.getSheetByName('Tarefas');
  
  
  var dataRange = tasksSheet.getDataRange();
  var rangeValues = dataRange.getValues();
  var numberOfRows = dataRange.getLastRow();
  
  for(i = 2; i < numberOfRows; i++){
    
    var situacao = rangeValues[i][3];
    
    switch(situacao){
      case "Metas":
        Logger.log(i + ": Metas");
        copyPasteTasks(i+1, 1);
        break;
        
      case "Concluídas":
        Logger.log(i + ": Concluídas");
        copyPasteTasks(i+1, 2);
        break;
        
      case "À espera de execução":
        Logger.log(i + ": À espera de execução");
        copyPasteTasks(i+1, 3);
    }

  }
  
  ss.getRange('D:D').activate();
  ss.getActiveSheet().deleteColumns(ss.getActiveRange().getColumn(), ss.getActiveRange().getNumColumns());
  
}

function copyPasteTasks(rowIndex, taskCondition){
  var ss = SpreadsheetApp.getActive();
  var reportSheet = ss.getSheetByName('Relatório');
  
  
  reportSheet.getRange('D1').activate();
  
  for (j = 0; j < taskCondition; j++){
    ss.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  }
  
  ss.getActiveSheet().insertRowsAfter(ss.getActiveRange().getLastRow(), 1);
  ss.getActiveRange().offset(1, 0).activate();
  var lastRow = ss.getActiveRange().getLastRow();
  Logger.log(lastRow);
  
  ss.getActiveRange().offset(-1, 0).moveTo(ss.getActiveRange());
  ss.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.PREVIOUS).activate();
  ss.getRange('Tarefas!A' + rowIndex + ':C' + rowIndex).copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
  ss.getActiveSheet().setRowHeight(lastRow, 80);
  ss.getRange('A' + lastRow + ':C' + lastRow).activate();
  ss.getActiveRangeList().setVerticalAlignment('middle')
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
}

function cloneTemplate(){
  var name = "Relatório";
  var ss = SpreadsheetApp.getActive();
  var template = ss.getSheetByName("Template").copyTo(ss);
  
  /* Before cloning the sheet, delete any previous copy */
  var old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old); // or old.setName(new Name);
  
  SpreadsheetApp.flush(); // Utilities.sleep(2000);
  template.setName(name);

  /* Make the new sheet active */
  ss.setActiveSheet(template);
  
}









