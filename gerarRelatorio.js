var LIN1 = 13;


function gerarRelatorioTeste() {
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
  
  if (arrayA.length != 0) {
      copyPasteTasksTeste(LIN1, arrayA.length, arrayA);
  }
  
  if (arrayB.length != 0) {
      copyPasteTasksTeste(LIN1+arrayA.length+5, arrayB.length, arrayB);
  }
  
  if (arrayC.length != 0) {
      copyPasteTasksTeste(LIN1+arrayA.length+arrayB.length+10, arrayC.length, arrayC);
  }
  
  reportSheet.showSheet();
  SpreadsheetApp.getActive().setActiveSheet(reportSheet);
  SpreadsheetApp.getActive().getSheetByName("Processando").hideSheet();
  
}


function copyPasteTasksTeste(afterPosition, howMany, values){
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


function cloneTemplateTeste(){
  var name = "Relatório";
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Template").copyTo(ss);
  
  var old = ss.getSheetByName(name);
  if (old) ss.deleteSheet(old);
  
  SpreadsheetApp.flush(); 
  sheet.setName(name);
  
}