//добавление новых месяцов в таблицу
function addNewSheetForNextMonth() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
    sheetTemplate = spreadSheet.getSheetByName("шаблон"),
    tempDay = new Date(),
    today = new Date(
      tempDay.getFullYear(),
      tempDay.getMonth() + 3,
      tempDay.getDate()
    );
  var tomorrow = new Date(
      today.getFullYear(),
      today.getMonth(),
      today.getDate() + 1
    ),
    lastDayInMonth = new Date(
      tomorrow.getFullYear(),
      tomorrow.getMonth(),
      0
    ).getDate(); //последний день месяца

  //если завтра другой месяц
  if (today.getMonth() != tomorrow.getMonth()) {
    var sheetName = MONTHS[tomorrow.getMonth()] + " " + tomorrow.getFullYear(); //название листа
    spreadSheet.insertSheet(sheetName, { template: sheetTemplate }); //вставляю в таблицу новый лист по шаблону
    var sheet = spreadSheet.getSheetByName(sheetName),
      data = sheet.getDataRange().getValues(),
      bgrds = sheet.getDataRange().getBackgrounds();
    var indexFirstDayInMonth = getWeekDay(tomorrow); //cмотрю на который день припадает первое число
    var startDay = 0;

    for (var j = 0; j < data[4].length; j++) {
      if (data[4][j] == "0.0") {
        startDay++;
        //нахожу индекс в ряде этого числа в таблице
        if (startDay == indexFirstDayInMonth) {
          indexFirstDayInMonth = j;
          break;
        }
      }
    }
    startDay = 1;
    for (var i = 0; i < data.length; i++) {
      indexFirstDayInMonth = startDay == 1 ? indexFirstDayInMonth : 0;
      //ище ячейку с датой
      if (data[i][0] == "дата") {
        for (var j = indexFirstDayInMonth; j < data[i].length; j++) {
          //если счетчик дней < дней в месяце, то вставляю дату
          if (data[i][j] == "0.0" && startDay <= lastDayInMonth) {
            data[i][j] = startDay;
            startDay++;
          }
        }
      }
    }

    for (var i = 0; i < data.length; i++) {
      if (data[i][0] == "дата") {
        for (var j = 0; j < data[i].length; j++) {
          //если остались ячейки без даты
          if (data[i][j] == "0.0") {
            data[i][j] = "";
            for (var k = j; k < data[i].length; k++) {
              //и после идет синей разделитель
              if (bgrds[i + 1][k] == "#9fc5e8") {
                //то обьеденяю ячейки
                sheet.getRange(i + 1, j + 1, 15, k + 1 - (j + 1)).merge();
                break;
              } else if (k == data[i].length - 1) {
                //то обьеденяю ячейки
                sheet.getRange(i + 1, j + 1, 15, k + 1 - (j - 1)).merge();
                break;
              }
            }
          }
        }
      }
    }

    sheet.getDataRange().setValues(data);
  }
}

function getWeekDay(date) {
  var currentDay = date.getDay() == 0 ? 7 : date.getDay();
  return currentDay;
}

//копирывание листов в табл. Расписание с откр. доступом
function copySheetsToFreeTable() {
  var ssSchedule = getSpreadsheetById(SCHEDULE_SPREADSHEET_ID),
    ssScheduleFree = getSpreadsheetById(SCHEDULE_FREE_SPREADSHEET_ID)
    sheets = ssSchedule.getSheets(),
    sheetsFree = ssScheduleFree.getSheets(),
    sheetName, //имя листа таблицы
    tempSheet, //текущий лист таблицы
    link; //ссылка на лист таблицы
  //=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1y5fVM2V1ZXuHAD4GXOdUM42wru1x0BHViryaJUies0Q/edit#gid=504215770";"'Январь 2018'!A1:O9")

  //удаляю все листы с откр. таблицы(кроме шаблона, потому что, нельзя удалить все листы таблицы)
  for (var i = 0; i < sheetsFree.length; i++) {
    if (sheetsFree[i].getName() != "шаблон") {
      ssScheduleFree.deleteSheet(sheetsFree[i]);
    }
  }
  //добавляю листы с Расписания, кроме шаблона
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() != "шаблон") {
      sheetName = sheets[i].getName(); //имя текущего листа
      sheets[i].copyTo(ssScheduleFree).setName(sheetName); //копирую в откр таблицу
      tempSheet = ssScheduleFree.getSheetByName(sheetName); //меняю название
      tempSheet.getDataRange().clearContent(); //все удаляю с листа, чтобы вставить формулу
      link = ssSchedule.getUrl() + "#gid=" + sheets[i].getSheetId(); //ссылка на лист в табл. Расписание
      tempSheet
        .getRange(1, 1)
        .setValue('=IMPORTRANGE("' + link + '";"' + sheetName + '!A1:EX")'); //добавляю формулу
    }
  }
}
