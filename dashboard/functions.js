//получение данных для модального окна
function getDataForModal() {
  const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  const SHEET = SPREADSHEET.getActiveSheet();
  const CATALOG_SHEET = SPREADSHEET.getSheetByName("Справочник");
  const CATALOG_VALUES = CATALOG_SHEET.getDataRange().getValues();
  const SHEET_VALUES = SHEET.getDataRange().getValues();
  const bgrsCatalog = CATALOG_SHEET.getDataRange().getBackgrounds();
  const colorsCatalog = CATALOG_SHEET.getDataRange().getFontColors();
  const cell = SHEET.getActiveCell();
  const row = cell.getRow();

  try {
    //формирую данные для модального окна
    const res = {
      classes: [],
      teachers: [],
      bgrs: [],
      colors: [],
      currentTeacher: CATALOG_VALUES[1][3] //текущий педагог со справочника
    };

    //номер абонемента, клиент, дисц, тип, комент, цена, кол, израсх.
    res.subscription = [
      SHEET_VALUES[row - 1][0],
      SHEET_VALUES[row - 1][1],
      SHEET_VALUES[row - 1][2],
      SHEET_VALUES[row - 1][3],
      SHEET_VALUES[row - 1][4],
      SHEET_VALUES[row - 1][5],
      SHEET_VALUES[row - 1][6]
    ];

    //добавляю все классы, всех педаг., все стиле
    for (var i = 1; i < CATALOG_VALUES.length; i++) {
      if (CATALOG_VALUES[i][0] != "") {
        res.classes.push(CATALOG_VALUES[i][0]);
      }

      if (CATALOG_VALUES[i][1] != "") {
        res.teachers.push(CATALOG_VALUES[i][2]);
        res.bgrs.push(bgrsCatalog[i][2]);
        res.colors.push(colorsCatalog[i][2]);
      }
    }

    //добавляю выбранные ранее занятия в этом абонементе
    var lessonsInfo = [];
    var countLessons = SHEET_VALUES[row - 1][6]; //кол-во занятий
    var startCellsInfo = 8; //начало колонок с иформ. о занятиях
    var countCellsInfo = 5; //кол-во колонок с иформ. о занятиях

    //иду по этому абонементу
    for (
      var j = startCellsInfo;
      j < startCellsInfo + countCellsInfo * countLessons;
      j++
    ) {
      //если это новое занятие, добавляю новый массив
      if ((j - startCellsInfo) % countCellsInfo == 0) {
        lessonsInfo.push([]);
      }

      //в него добавляю текущее поле
      lessonsInfo[lessonsInfo.length - 1].push(SHEET_VALUES[row - 1][j]);
    }
    res.lessInfo = lessonsInfo;
    res.dataTable = getDataFromSchedule();

    return JSON.stringify(res);
  } catch (e) {
    //выкидываю ошибку, если не выбран номер абонемента
    var htmlOutput = HtmlService.createHtmlOutput(
      "<p>Выберите номер абонемента!</p>"
    )
      .setWidth(250)
      .setHeight(70);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Ошибка!");

    return false;
  }
}

//бронирую время
function PassForm(allData) {
  var newData = JSON.parse(allData)[0];
  var dataFromTable = JSON.parse(allData)[1]; //даные взятые с таблицы
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
    sheet = spreadSheet.getActiveSheet(),
    dataSubscription = sheet.getDataRange().getValues(),
    cell = sheet.getActiveCell(),
    row = cell.getRow();
  var arrayResult = [];
  var passedLessons = 0;

  //заносим данные в личный кабинет
  for (var i = 0; i < newData.length; i++) {
    arrayResult.push(newData[i].date || "");
    arrayResult.push(newData[i].time || "");
    arrayResult.push(newData[i].class || "");
    arrayResult.push(newData[i].teacher || "");
    arrayResult.push(newData[i].status || "");

    if (newData[i].status != "-") {
      passedLessons++;
    }
  }

  //cтарые данные для удаления
  try {
    var lessonsForDelete = [];
    var lessDel = sheet.getRange(row, 9, 1, arrayResult.length).getValues();

    for (var i = 0; i < lessDel[0].length; i += 5) {
      var obj = {};
      obj.date = lessDel[0][0 + i];
      obj.time = lessDel[0][1 + i];
      obj.class = lessDel[0][2 + i];
      obj.teacher = lessDel[0][3 + i];
      obj.client = dataSubscription[row - 1][1];
      lessonsForDelete.push(obj);
    }

    //удаляю с расписания старые даные
    deleteLessonFromSchedule(lessonsForDelete);
  } catch (e) {
    Logger.log("Ошибка удаление старых занятий");
    Logger.log(e);
  }
  sheet.getRange(row, 9, 1, arrayResult.length).setValues([arrayResult]);
  sheet.getRange(row, 8).setValue(passedLessons);

  try {
    setLessonToSchedule(newData, dataFromTable);
  } catch (e) {
    Logger.log("Ошибка записи новых занятий");
    Logger.log(e);
  }
}

//добавляю занятия в расписание
function setLessonToSchedule(data, dataFromTable) {
  var ssSchedule = getSpreadsheetById(SCHEDULE_SPREADSHEET_ID),
    sSchedule,
    dataSchedule;

  for (var less = 0; less < data.length; less++) {
    try {
      var checkedDate = data[less].date.toString().split(".")[0];
      var sheetName = getSheetNameThisMonth(data[less]);

      sSchedule = ssSchedule.getSheetByName(sheetName);
      dataSchedule = sSchedule.getDataRange().getValues();

      for (var i = 0; i < dataSchedule.length; i++) {
        //ищу поле с датой
        if (dataSchedule[i][0] == "дата") {
          for (var j = 1; j < dataSchedule[i].length; j++) {
            //ищу выбраную дату
            if (+dataSchedule[i][j] == +checkedDate) {
              //ищу выбранный клас в этой дате
              var indexColumn;
              for (var k = j; k < dataSchedule[i].length; k++) {
                //!!
                if (dataSchedule[1][k] == data[less].class) {
                  indexColumn = k + 1;
                  break;
                }
              }
              //ищу в этой колонке выбраное время? 24-макс. кол. часов
              for (var p = i + 2; p < dataSchedule.length; p++) {
                if (
                  dataSchedule[p][0].replace(/\s/g, "") ==
                  data[less].time.replace(/\s/g, "")
                ) {
                  //проверяю точно ли пусто в таблице и не его ли там урок
                  //если нет
                  if (sSchedule.getRange(p + 1, indexColumn).getValue() != "") {
                    if (
                      sSchedule.getRange(p + 1, indexColumn).getValue() !=
                        data[less].teacher ||
                      sSchedule.getRange(p + 1, indexColumn + 1).getValue() !=
                        data[less].client
                    ) {
                      updateIfErrorSet(
                        data[less].date,
                        data[less].time,
                        data[less].class,
                        dataFromTable,
                        sSchedule.getRange(p + 1, indexColumn).getValue(),
                        sSchedule.getRange(p + 1, indexColumn + 1).getValue(),
                        data[less].teacher,
                        data[less].client
                      );
                    }
                    break;
                  } else {
                    //если да,вставляю значения в расп и в откр расписание
                    sSchedule
                      .getRange(p + 1, indexColumn)
                      .setValue(data[less].teacher)
                      .setBackground(data[less].bgr)
                      .setFontColor(data[less].color);
                    sSchedule
                      .getRange(p + 1, indexColumn + 1)
                      .setValue(data[less].client)
                      .setBackground(data[less].bgr)
                      .setFontColor(data[less].color);

                    break;
                  }
                }
              }
              break;
            }
          }
        }
      }
    } catch (e) {
      Logger.log("Ошибка записи данных");
      Logger.log(e);
    }
  }
}

//обновляю личный кабинет, если хотели добавить занятие, а там уже занято
function updateIfErrorSet(
  date,
  time,
  _class,
  dataFromTable,
  teacherFromTable,
  clientFromTable,
  tempTeacher,
  tempClient
) {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet(),
    sheet = spreadSheet.getActiveSheet(),
    dataSubscription = sheet.getDataRange().getValues(),
    cell = sheet.getActiveCell(),
    row = cell.getRow();

  //затираю то занятие, где возникла ошибка
  var day = date.split(".")[0],
    month = date.split(".")[1],
    year = date.split(".")[2];

  //прохожу по абонементу
  for (var i = 8; i < dataSubscription[row - 1].length; i += 5) {
    try {
      //если это та дата, то время и год
      if (
        dataSubscription[row - 1][i].getFullYear() == year &&
        dataSubscription[row - 1][i].getDate() == day &&
        dataSubscription[row - 1][i].getMonth() + 1 == month &&
        dataSubscription[row - 1][i + 1] == time &&
        dataSubscription[row - 1][i + 2] == _class
      ) {
        //удаляю
        sheet.getRange(row, i + 1, 1, 5).setValue("");
        break;
      }
    } catch (e) {}
  }
  //вывожу окно об ошибке
  var htmlOutput = HtmlService.createHtmlOutput(
    "<p>" +
      date +
      "<br>" +
      time +
      "<br>" +
      _class +
      "<br><br>Педагог: " +
      teacherFromTable +
      "<br>Клиент: " +
      clientFromTable +
      "</p>"
  )
    .setWidth(350)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Это место уже занято:");

  //название листа в таблице в зависимости от месяца
  var sheetName = MONTHS[month - 1] + " " + year;

  MailApp.sendEmail({
    to: "melnykkatia@gmail.com",
    subject: "Место занято",
    htmlBody:
      "<p>Таблица: " +
      spreadSheet.getName() +
      ",</p>" +
      "<p>Дата: </p>" +
      date +
      "<p>Время: </p>" +
      time +
      "<p>Класс: </p>" +
      _class +
      "<p>Св. часы, которые были: </p>" +
      dataFromTable[sheetName][day][_class] +
      "<p>Педагог с таблицы: </p>" +
      teacherFromTable +
      "<p>Клиент с таблицы: </p>" +
      clientFromTable +
      "<p>Текущий педагог: </p>" +
      tempTeacher +
      "<p>Текущий клиент: </p>" +
      tempClient
  });
}

//обновляю личный кабинет, если хотели добавить занятие, а там уже занято
function updateIfErrorDel(
  date,
  time,
  _class,
  teacherFromTable,
  clientFromTable,
  tempTeacher,
  tempClient
) {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log("удаление");

  MailApp.sendEmail({
    to: ADMINISTRATOR_EMAIL,
    subject: "Удаление урока",
    htmlBody:
      "<p>Таблица: " +
      spreadSheet.getName() +
      ",</p>" +
      "<p>Дата: </p>" +
      date +
      "<p>Время: </p>" +
      time +
      "<p>Класс:" +
      _class +
      " </p>" +
      "<p>Педагог с таблицы: </p>" +
      teacherFromTable +
      "<p>Клиент с таблицы: </p>" +
      clientFromTable +
      "<p>Текущий педагог: </p>" +
      tempTeacher +
      "<p>Текущий клиент: </p>" +
      tempClient
  });
}

//удаление занятий с расписания
function deleteLessonFromSchedule(data) {
  var ssSchedule = getSpreadsheetById(SCHEDULE_SPREADSHEET_ID),
    sSchedule,
    dataSchedule;

  for (var less = 0; less < data.length; less++) {
    try {
      var checkedDate = data[less].date.getDate();
      var sheetName = getSheetNameThisMonth(data[less]);

      sSchedule = ssSchedule.getSheetByName(sheetName);
      dataSchedule = sSchedule.getDataRange().getValues();

      for (var i = 0; i < dataSchedule.length; i++) {
        //ищу поле с датой
        if (dataSchedule[i][0] == "дата") {
          for (var j = 1; j < dataSchedule[i].length; j++) {
            //ищу выбраную дату
            if (+dataSchedule[i][j] == +checkedDate) {
              //ищу выбранный клас в этой дате
              var indexColumn;
              for (var k = j; k < dataSchedule[i].length; k++) {
                if (dataSchedule[1][k] == data[less].class) {
                  indexColumn = k + 1;
                  break;
                }
              }
              //ищу в этой колонке выбраное время
              for (var p = i + 2; p < dataSchedule.length; p++) {
                if (
                  dataSchedule[p][0].replace(/\s/g, "") ==
                  data[less].time.replace(/\s/g, "")
                ) {
                  //!!!
                  //можна удалить только если тут был твой урок
                  if (
                    sSchedule
                      .getRange(p + 1, indexColumn)
                      .getValue()
                      .toLowerCase() == data[less].teacher.toLowerCase() &&
                    sSchedule
                      .getRange(p + 1, indexColumn + 1)
                      .getValue()
                      .toLowerCase() == data[less].client.toLowerCase()
                  ) {
                    sSchedule
                      .getRange(p + 1, indexColumn)
                      .clearContent()
                      .clearFormat();
                    sSchedule
                      .getRange(p + 1, indexColumn + 1)
                      .clearContent()
                      .clearFormat();
                  } else {
                    //если хотел удалить не свой урок и там не пусто
                    if (
                      sSchedule.getRange(p + 1, indexColumn).getValue() != "" &&
                      sSchedule.getRange(p + 1, indexColumn + 1).getValue() !=
                        ""
                    ) {
                      //сообщение об ошибке на почту
                      updateIfErrorDel(
                        data[less].date,
                        data[less].time,
                        data[less].class,
                        sSchedule.getRange(p + 1, indexColumn).getValue(),
                        sSchedule.getRange(p + 1, indexColumn + 1).getValue(),
                        data[less].teacher,
                        data[less].client
                      );
                    }
                  }
                  break;
                }
              }
              break;
            }
          }
        }
      }
    } catch (e) {
      Logger.log(e);
      Logger.log("Ошибка удаления старых даних");
    }
  }
}

//формирование данных
function getDataFromSchedule() {
  const SCHEDULE_SPREADSHEET = getSpreadsheetById(SCHEDULE_SPREADSHEET_ID);
  const SCHEDULE_SHEETS = SCHEDULE_SPREADSHEET.getSheets();

  const sheets = {};
  var today = new Date();
  //назвы таблиц, за следующие 4 месяца
  var sheetNames = [
    getSheetNameThisMonth(today),
    getSheetNameThisMonth(
      new Date(today.getFullYear(), today.getMonth() + 1, 1)
    ),
    getSheetNameThisMonth(
      new Date(today.getFullYear(), today.getMonth() + 2, 1)
    ),
    getSheetNameThisMonth(
      new Date(today.getFullYear(), today.getMonth() + 3, 1)
    )
  ];

  //иду по всем листам таблицы
  for (var s = 0; s < SCHEDULE_SHEETS.length; s++) {
    //иду по листам  четырех месяцов
    for (var lastSheets = 0; lastSheets < sheetNames.length; lastSheets++) {
      //если одинаковые названия
      if (sheetNames[lastSheets] == SCHEDULE_SHEETS[s].getSheetName()) {
        //считываю все даные и бгр
        const SHEET_VALUES = SCHEDULE_SHEETS[s].getDataRange().getValues();
        const SHEET_BACKGROUNDS = SCHEDULE_SHEETS[s].getDataRange().getBackgrounds();

        //добавляю этот лист в конечный обьект
        sheets[SCHEDULE_SHEETS[s].getSheetName()] = {};
        //прохожу по даным этого листа
        for (var i = 0; i < SHEET_VALUES.length; i++) {
          //исчу рядок с полем дата
          //если нахожу
          if (SHEET_VALUES[i][0] == "дата") {
            //прохожу по этом ряду
            for (var j = 1; j < SHEET_VALUES[i].length; j++) {
              if (SHEET_VALUES[i][j] != "") {
                //добавляю все даты, тоесть не пустые поля
                sheets[SCHEDULE_SHEETS[s].getSheetName()][
                  SHEET_VALUES[i][j]
                ] = {};

                //добавляю места
                var startClassIndexInDay = j; //индекс начала даты в рядке
                //прохожу по первой строке с классами
                for (
                  var startClassIndexInDay = j;
                  startClassIndexInDay < SHEET_VALUES[1].length;
                  startClassIndexInDay++
                ) {
                  // добаляю классы в дату, пока не найду синий разделитель(разделитель между днями)
                  if (SHEET_BACKGROUNDS[1][startClassIndexInDay] == "#9fc5e8") {
                    break;
                  } else if (SHEET_VALUES[1][startClassIndexInDay] != "") {
                    //если не пусто добаляю класс
                    sheets[SCHEDULE_SHEETS[s].getSheetName()][
                      SHEET_VALUES[i][j]
                    ][SHEET_VALUES[1][startClassIndexInDay]] = [];
                    //прохожу по всем часом
                    for (var k = 0; k < 24; k++) {
                      //если нахожу cиний разделитель или конец таблицы, часы закончились, выхожу из цикла;
                      if (SHEET_BACKGROUNDS[i + 2 + k][0] == "#9fc5e8") {
                        break;
                      }

                      //если  пусто, добавляю свободный час
                      if (SHEET_VALUES[i + 2 + k][startClassIndexInDay] == "") {
                        //беру время без пробелов
                        var timeWithOutSpace = SHEET_VALUES[
                          i + 2 + k
                        ][0].replace(/\s/g, ""); // ?????
                        sheets[SCHEDULE_SHEETS[s].getSheetName()][
                          SHEET_VALUES[i][j]
                        ][SHEET_VALUES[1][startClassIndexInDay]].push(
                          timeWithOutSpace
                        );
                      }

                      if (i + 2 + k == dataSchedule.length - 1) {
                        break;
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }

  return sheets;
}
