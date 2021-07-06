function archiving() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetCatalog = ss.getSheetByName('Абонементы');
  var dataCatalog = sheetCatalog.getDataRange().getValues();  
  var sheetArchive = ss.getSheetByName('Архив абонементов');
    
  // идем по всем абонементам с конца
  for(var i=dataCatalog.length-1;i>=0;i--){
    var rowData = dataCatalog[i]; // строка с данными
    // сравниваем кол-во использовано и всего в абоне
    if(rowData[6] == rowData[7] && rowData[6] != "") {
      // удаляем строку
      sheetCatalog.deleteRow(1+i);
      // переносим в архив
      sheetArchive.appendRow(rowData)
    }
  }
  
    return;
}



 