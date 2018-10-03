var START_LINE = 13; // Linha a partir da qual será inserido às tarefas (após esta linha)
var OFFSET = 5;      // Quantidae de colunas em branco + Títulos entre o fim de uma classe de tarefas e o início de outra


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
  
  var arrayA = new Array();
  var arrayB = new Array();
  var arrayC = new Array();

  for(i = 2; i < numberOfRows; i++){
    
    var temp = tasksSheet.getRange('A' + (i+1) + ':C' + (i+1)).getValues();

    var situacao = rangeValues[i][3];
    
    switch(situacao){
      case "Metas":
        arrayA = arrayA.concat(temp);
        break;
        
      case "Concluídas":
        arrayB = arrayB.concat(temp);
        break;
        
      case "À espera de execução":
        arrayC = arrayC.concat(temp);
    }
    
  }  

  // Colando os falore e formatando as células para cada situação de tarefa
  var totalOffset = START_LINE;
  
  if (arrayA.length != 0) {
      copyPasteTasks(totalOffset, arrayA.length, arrayA);
  }
  
  totalOffset += arrayA.length + OFFSET;
  
  if (arrayB.length != 0) {
      copyPasteTasks(totalOffset, arrayB.length, arrayB);
  }
  
  totalOffset += arrayB.length + OFFSET;
  
  if (arrayC.length != 0) {
      copyPasteTasks(totalOffset, arrayC.length, arrayC);
  }
  
  
  reportSheet.showSheet();
  SpreadsheetApp.getActive().setActiveSheet(reportSheet);
  SpreadsheetApp.getActive().getSheetByName("Processando").hideSheet();
  
}


function copyPasteTasks(afterPosition, howMany, values){
  var ss = SpreadsheetApp.getActive();
  var reportSheet = ss.getSheetByName('Relatório');
 
  reportSheet.insertRowsAfter(afterPosition, howMany);
  
  reportSheet.getRange(afterPosition+1, 1, howMany, 3)
  .setValues(values)
  .setVerticalAlignment('middle')
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  reportSheet.setRowHeights(afterPosition+1, howMany, 80);
  
  SpreadsheetApp.flush();
  
}


function cloneTemplate(){
  var name = "Relatório";
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Template").copyTo(ss);
  
  // Caso exista uma planilha de relatório antiga, será apagada
  var old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  
  SpreadsheetApp.flush(); 
  sheet.setName(name);
  
}