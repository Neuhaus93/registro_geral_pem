var lin1 = 13;
var lin2 = 18;
var lin3 = 23;


function gerarRelatorio() {
  var ss = SpreadsheetApp.getActive();
  
  cloneTemplate();
  
  ss.getSheetByName("Processando").showSheet().activate();
  
  var reportSheet = ss.getSheetByName('Relatório');
  var tasksSheet = ss.getSheetByName('Tarefas');
  
  reportSheet.hideSheet();  
  
  var dataRange = tasksSheet.getDataRange();
  var rangeValues = dataRange.getValues();
  var numberOfRows = dataRange.getLastRow();
  
  for(i = 2; i < numberOfRows; i++){
    
    var progression = (i - 2) / (numberOfRows - 2) * 100;
    progression = progression.toFixed(1);
    ss.getSheetByName("Processando").getRange('B1').setValue("Progresso: " + progression + "%");
    
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
  
  reportSheet.showSheet();
  SpreadsheetApp.getActive().setActiveSheet(reportSheet);
  SpreadsheetApp.getActive().getSheetByName("Processando").hideSheet();
  
}

function copyPasteTasks(rowIndex, taskCondition){
  var ss = SpreadsheetApp.getActive();
  var reportSheet = ss.getSheetByName('Relatório');
  
  var aux = reportSheet.getRange('A1');
  var lastRow = 1;
  
  switch(taskCondition){
    case 1:
      reportSheet.insertRowAfter(lin1);
      lin1++; lin2++; lin3++;
      lastRow = lin1;
      aux = reportSheet.getRange('A'+lin1);
      break;
      
    case 2:
      reportSheet.insertRowAfter(lin2);
      lin2++, lin3++;
      lastRow = lin2;
      aux = reportSheet.getRange('A'+lin2);
      break;
      
    case 3:
      reportSheet.insertRowAfter(lin3);
      lin3++;
      lastRow = lin3;
      aux = reportSheet.getRange('A'+lin3);
  }
  
  ss.getRange('Tarefas!A' + rowIndex + ':C' + rowIndex).copyTo(aux, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
  reportSheet.setRowHeight(lastRow, 80);
  var taskRow = reportSheet.getRange('A' + lastRow + ':C' + lastRow);
  Logger.log(taskRow.getA1Notation());
  taskRow.setVerticalAlignment('middle')
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  SpreadsheetApp.flush();
  
}

function cloneTemplate(){
  var name = "Relatório";
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Template").copyTo(ss);
  
  var old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  
  SpreadsheetApp.flush(); 
  sheet.setName(name);
  
}









