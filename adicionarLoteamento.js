function addCliente() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var addSheet = ss.getSheetByName('Adicionar cliente');
  var lotSheet = ss.getSheetByName('Loteamentos');
  
  var nomeEmpresa = addSheet.getRange(3, 3);
  
  var row = 5;
  var emailsEmpresa = new Array();
  
  while(!addSheet.getRange(row, 3).isBlank()) {
    emailsEmpresa.push(addSheet.getRange(row,3).getValue());
    row ++;
  }
  Logger.log(emailsEmpresa);
  
  if((nomeEmpresa.getValue() == '') || (lotSheet.getRange(5, 3).getValue() == '')) {
    
    SpreadsheetApp.getUi().alert('Por favor termine de inserir os dados antes de adicionar');
    
    return;
  }
  
  var i = 0;
  var j = 1;
  
  while (i < 2) {
    if(lotSheet.getRange(j, 2).isBlank()) i++;
    j++;
  }
  lotSheet.insertRowAfter(j - 2);
  
  var codigo = lotSheet.getRange(j-2, 2).getValue();
  codigo = codigo + 1;
  
  lotSheet.getRange(j-1, 2).setValue(codigo);
  lotSheet.getRange(j-1, 3).setValue(nomeEmpresa.getValue());
  
  var emails = emailsEmpresa[0];
  for(i = 1; i < emailsEmpresa.length; i++) {
    emails += ',' + emailsEmpresa[i];
  }  
  lotSheet.getRange(j-1, 4).setValue(emails);
  
  SpreadsheetApp.setActiveSheet(lotSheet);
  
  nomeEmpresa.setValue('');
  addSheet.getRange(5, 3).setValue('');
  addSheet.getRange(6, 3).setValue('');
  
  
}




function addLoteamento() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var addSheet = ss.getSheetByName('Adicionar loteamento');
  var lotSheet = ss.getSheetByName('Loteamentos');
  
  var nomeEmpresa = addSheet.getRange(3, 3);
  var nomeLoteamento = addSheet.getRange(5, 3);
  var dataInicio = addSheet.getRange(7,3);
  
  if((nomeEmpresa.getValue() == '') || (nomeLoteamento.getValue() == '') || (dataInicio.getValue() == '')) {
    
    SpreadsheetApp.getUi().alert('Por favor termine de inserir os dados antes de adicionar');
    
    return;
  }
  
  var i = 0;
  var j = 1;
  
  while (i < 4) {
    if(lotSheet.getRange(j, 2).isBlank()) i++;
    j++;
  }
  lotSheet.insertRowAfter(j - 2);
  
  var codigo = lotSheet.getRange(j-2, 2).getValue();
  codigo = codigo + 1;
  
  Logger.log(j);
  
  lotSheet.getRange(j-1, 2).setValue(codigo);
  lotSheet.getRange(j-1, 3).setValue(nomeLoteamento.getValue());
  lotSheet.getRange(j-1, 4).setValue(nomeEmpresa.getValue());
  lotSheet.getRange(j-1, 5).setValue('Em andamento');
  lotSheet.getRange(j-1, 6).setValue(dataInicio.getValue());
  
  
  // Criar nova planilha relacionada ao loteamento
  var name = '#' + codigo + ' - ' + nomeLoteamento.getValue();
  var newSheet = ss.getSheetByName("Loteamento - Template").copyTo(ss);
  
  SpreadsheetApp.flush(); 
  newSheet.setName(name);
  
  newSheet.getRange(1,1).setValue(nomeLoteamento.getValue());
  newSheet.getRange(1,5).setValue(dataInicio.getValue());
  
  
  nomeEmpresa.setValue('');
  nomeLoteamento.setValue('');
  dataInicio.setValue('');
  
  SpreadsheetApp.setActiveSheet(newSheet);
}

