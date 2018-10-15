function mandarTarefas() {
 
  var ss = SpreadsheetApp.getActive();
  var FILE_NAME = 'Relatório de Tarefas - P&M';
  var TITULO = 'Registro Geral de Tarefas - P&M';
  var CORPO = 'Encontra-se em anexo o registro geral de tarefas programadas e concluídas.';
  
  // Pega o nome do cliente
  var clientTable = ss.getRangeByName('clientes').getValues();
  
  // Pega os emails, e remove possíves espaços
  var emails = clientTable[1][2].split(",");
  for(var i = 0; i < emails.length; i++) {
   emails[i] = emails[i].trim(); 
  }
  
  // Gera um relatório atualizado referente a planilha 'Tarefas'
  gerarRelatorio();
  
  // Deleta os relatórios antigos da pasta do Drive
  deleteFile(FILE_NAME + '.pdf');
  
  // Converte a planilha 'Relatório' criada em pdf
  var pdfFile = convertSpreadsheetToPdf(null, 'Relatório - Tarefas', FILE_NAME);
  
  // Envia o arquivo para os emails 
  sendEmail(emails, pdfFile, TITULO, CORPO);
  
}


function mandarTodosRelatorios() {
  
  var clientTable = SpreadsheetApp.getActive().getRangeByName('clientes').getValues();
  
  for(var i = 1; i < (clientTable.length - 1); i++) {
   mandarRelatorio(i); 
  }
             
}



function mandarRelatorio(num) {
  
  var ss = SpreadsheetApp.getActive();
  var LINHA_TITULO = 19;
  var TITULO = 'Relatório semanal - P&M';
  var CORPO = 'Encontra-se em anexo o relatório semanal com as informações sobre cada um dos loteamentos em andamento.';
  
  /** Pega o nome do cliente */
  var clientTable = ss.getRangeByName('clientes').getValues();
  var cliente = clientTable[num][1];
  
  var FILE_NAME = 'Relatório Semanal - ' + cliente;
  
  var emails = clientTable[num][2].split(",");
  for(var i = 0; i < emails.length; i++) {
   emails[i] = emails[i].trim(); 
  }
  
  // Adiciona o nome do cliente na capa
  ss.getSheetByName('Capa - Template').getRange(LINHA_TITULO,1).setValue(cliente); 
  
  // Mostrar todos os negócios referentes ao cliente específico
  var loteamentos = getLoteamentos(cliente);
  
  // Caso não exista loteamentos do cliente, não continua o código
  if(loteamentos.length == 0) return;
  
  // Adiciona a capa antes de mandar o relatório
  loteamentos.push('Capa - Template');
  
  // Mostrar todos as sheets referentes aos loteamentos do cliente
  showSheets(loteamentos);
  hideSheets(loteamentos);
  
  // Deleta os relatórios antigos da pasta do Drive
  deleteFile(FILE_NAME + '.pdf');
  
  // Cria o pdf na pasta e salva ele em na variável pdfFile
  var pdfFile = convertSpreadsheetToPdf(null, null, FILE_NAME);
  
  // Esconde todas as sheets referentes ao relatório;
  showSheetsDefault();
  hideSheetsDefault();
  
  // Apaga o nome do cliente na capa template
  ss.getSheetByName('Capa - Template').getRange(25,1).setValue('');
  
  ss.getSheetByName('Loteamentos').activate();
  
  // Manda os e-mails com o arquivo anexado
  if(emails.length == 0) {
    return;
  }   
  sendEmail(emails, pdfFile, TITULO, CORPO);
  
}




function hideSheetsDefault() {

  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var toShow = ['Tarefas', 'Loteamentos', 'Adicionar cliente', 'Adicionar loteamento', 'Relatório'];
  
  sheets.forEach(function(sheet) {
    if (toShow.indexOf(sheet.getName()) == -1) {
      sheet.hideSheet();
    }
  })
  
}

function hideSheets(lot) {
 var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  
  sheets.forEach(function(sheet) {
    if (lot.indexOf(sheet.getName()) == -1) {
      sheet.hideSheet();
    }
  })
}

function showSheetsDefault() {  
  
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var toShow = ['Tarefas', 'Loteamentos', 'Adicionar cliente', 'Adicionar loteamento', 'Relatório'];
  
  sheets.forEach(function(sheet) {
    if (toShow.indexOf(sheet.getName()) != -1) {
        sheet.showSheet();
    }
  })
};

function showSheets(lot) {  
  
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  
  sheets.forEach(function(sheet) {
    if (lot.indexOf(sheet.getName()) != -1) {
        sheet.showSheet();
    }
  })
};



function getLoteamentos(cliente) {
  var ss = SpreadsheetApp.getActive();
  var loteamentosSheet = ss.getSheetByName('Loteamentos');
  
  var lotTable = ss.getRangeByName('loteamentos').getValues();
  
  var loteamentos = new Array();
  
  
  for(var i = 1; i < lotTable.length; i++) {
   
    if (lotTable[i][2] == cliente) {
      
      var temp = '#' + lotTable[i][0] + ' - ' + lotTable[i][1];
      
      loteamentos.push(temp);
      
    }
    
  }
  
  return loteamentos;
}




function getEmails(num) {
  var ss = SpreadsheetApp.getActive();
  var loteamentosSheet = ss.getSheetByName('Loteamentos');
  
  var emails = new Array();
  var j = 3;
  
  while(!loteamentosSheet.getRange(j, num+1).isBlank()) {
    var email = loteamentosSheet.getRange(j,num+1).getValue();
    
    emails.push(email);
    j++;
  }
  return emails;
}







function deleteFile(myFileName) {
  var allFiles, idToDLET, myFolder, rtrnFromDLET, thisFile;

  myFolder = DriveApp.getFolderById('1W_X9yDGhu4epK98covYD7I7lQSh3aekQ');

  allFiles = myFolder.getFilesByName(myFileName);

  while (allFiles.hasNext()) {//If there is another element in the iterator
    thisFile = allFiles.next();
    idToDLET = thisFile.getId();
    //Logger.log('idToDLET: ' + idToDLET);

    rtrnFromDLET = Drive.Files.remove(idToDLET);
  };
}





function convertSpreadsheetToPdf(spreadsheetId, sheetName, pdfName) {
  
  var spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  spreadsheetId = spreadsheetId ? spreadsheetId : spreadsheet.getId()  
  var sheetId = sheetName ? spreadsheet.getSheetByName(sheetName).getSheetId() : null;  
  var pdfName = pdfName ? pdfName : spreadsheet.getName();
  var parents = DriveApp.getFileById(spreadsheetId).getParents();
  var folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var url_base = spreadsheet.getUrl().replace(/edit$/,'');

  var url_ext = 'export?exportFormat=pdf&format=pdf'   //export as pdf

      // Print either the entire Spreadsheet or the specified sheet if optSheetId is provided
      + (sheetId ? ('&gid=' + sheetId) : ('&id=' + spreadsheetId)) 
      // following parameters are optional...
      + '&size=a4'      // paper size
      + '&portrait=true'    // orientation, false for landscape
      + '&fitw=true'        // fit to width, false for actual size
      + '&sheetnames=false&printtitle=false&pagenumbers=false'  //hide optional headers and footers
      + '&gridlines=false'  // hide gridlines
      + '&fzr=false';       // do not repeat row headers (frozen rows) on each page

  var options = {
    headers: {
      'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken(),
    }
  }
  
  var response = UrlFetchApp.fetch(url_base + url_ext, options);
  var blob = response.getBlob().setName(pdfName + '.pdf');
  folder.createFile(blob);
  return blob;
} // convertSpreadsheetToPdf()




function sendEmail(emails, file, titulo, corpo) {
  
  var mailOptions = {
      attachments:file
  }
  
  emails.forEach(function(email) {
    MailApp.sendEmail(
      email, 
      titulo, 
      corpo, 
      mailOptions);
  })
  
}
