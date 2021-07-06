// Добавить абонемент
function addClient() {
  var res = getDataForModal();
  var htmlOutputTMP = HtmlService.createTemplateFromFile("addClientModal");
  htmlOutputTMP.res = res;
  var htmlOutputText = htmlOutputTMP.evaluate().getContent();
  var htmlOutput = HtmlService.createHtmlOutput(htmlOutputText)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(960)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(
    htmlOutput,
    "Добавление нового абонемента"
  );
}

function checkCountLesson() {
  var admin = getSpreadsheetById(ADMIN_PANEL_SPREADSHEET_ID);
  var teachers = admin.getSheetByName("педагоги");
  var teachersData = teachers.getDataRange().getValues();

  //собираю id и имя перподавателя из справочника
  var teachers = [];
  var test = [];
  for (var i = 1; i < teachersData.length; i++) {
    teachers.push({
      name: teachersData[i][3],
      id: teachersData[i][4],
      flag: true
    });
  }

  var allClientsSheet = getSpreadsheetById(SUBSCRIPTIONS_SPREADSHEET_ID);
  var clients = allClientsSheet.getSheetByName("Абонементы в работе");
  var clientsData = clients.getDataRange().getValues();

  // нахожу преподавателей, у которых нужно проверить кол-во занятий
  for (var i = 1; i < clientsData.length; i++) {
    // цикл по талице "Абонементы в работе"
    var wasLessons = clientsData[i][9];
    var name = clientsData[i][5];
    var numClient = clientsData[i][1];

    for (var j = 0; j < teachers.length; j++) {
      // по списку перподавателей, подбираем их id страниц
      // если имя совпадает с справочником, беру id страницы преподавателя

      if (name == teachers[j]["name"] && teachers[j]["flag"]) {
        var idTeacher = teachers[j]["id"];
        var teacherSheet = SpreadsheetApp.openById(idTeacher);
        var clientsOfTeacher = teacherSheet.getSheetByName("Абонементы");
        var clientsDataTeacher = clientsOfTeacher.getDataRange().getValues();
        for (var k = 1; k < clientsDataTeacher.length; k++) {
          // по странице конкретного преподавателя
          if (
            numClient == clientsDataTeacher[k][0] &&
            wasLessons !== clientsDataTeacher[k][7]
          ) {
            var cell = clients.getRange(i + 1, 10);
            var cellOstatok = clients.getRange(i + 1, 11);
            var cellAllValue = clients.getRange(i + 1, 8).getValue();

            var ostatokCount = cellAllValue - clientsDataTeacher[k][7];

            test.push(clientsDataTeacher[k][7]);
            cell.setValue(clientsDataTeacher[k][7]);
            cellOstatok.setValue(ostatokCount);
          }
        }
      }
    }
  }

  Logger.log(test);
}

// Архивация израсходованных абонементов, у которых занятий столько же сколько использованых
function clientArchive() {
  var clients = getSpreadsheetById(SUBSCRIPTIONS_SPREADSHEET_ID);
  var clientsSheet = clients.getSheetByName("Абонементы в работе");
  var clientsSheetData = clients
    .getSheetByName("Абонементы в работе")
    .getDataRange()
    .getValues();
  var archive = clients.getSheetByName("Архив абонементов");

  var usedLessons = [];
  var delRows = []; // массив с номерами рядов, которые позже удалим...
  for (var i = 1; i < clientsSheetData.length; i++) {
    if (clientsSheetData[i][7] <= clientsSheetData[i][9]) {
      usedLessons.push(clientsSheetData[i]);
      delRows.push(i + 1);
    }
  }

  // копируем необходимые ряды в вкладку "архив"...
  for (var j = 0; j < usedLessons.length; j++) {
    archive.appendRow(usedLessons[j]);
  }
  // удаляем из вкладки "абонементы в работе" выбраные ряды по их номеру...
  for (k = delRows.length - 1; k >= 0; k--) {
    clientsSheet.deleteRows(delRows[k], 1);
  }
}
