//данные с вьюхи с выбором месяца для выписок, создаю листы для все педагогов
function PassFormStat(monthName) {
  monthName = JSON.parse(monthName);

  const ADMIN_PANEL_SPREADSHEET = getSpreadsheetById(
    ADMIN_PANEL_SPREADSHEET_ID
  );
  const TEACHERS_VALUES = ADMIN_PANEL_SPREADSHEET.getSheetByName("педагоги")
    .getDataRange()
    .getValues();
  const STATEMENT_TEMPLATE = ADMIN_PANEL_SPREADSHEET.getSheetByName(
    "шаблон выписки"
  );
  const teachers = []; //массив педагогов (name, id файла личного кабинета)

  //прохожу по колонке с id таблицы преподавателей
  for (var i = 1; i < TEACHERS_VALUES.length; i++) {
    const teacher = {
      name: TEACHERS_VALUES[i][0],
      sheetId: TEACHERS_VALUES[i][4]
    };

    teachers.push(teacher);

    try {
      //если такая таблица уже есть
      const STATEMENT_SHEET = ADMIN_PANEL_SPREADSHEET.getSheetByName(
        "Выписки " + teacher.name
      );
      STATEMENT_SHEET.getRange(1, 1, 1, STATEMENT_SHEET.getLastColumn()).clearContent();
      STATEMENT_SHEET.getRange(1, 1, 1, 2).setValues([["Месяц:", monthName]]);

      let lastRow = STATEMENT_SHEET.getLastRow();

      //удаляю все с таблицы пока не останется три рядка
      while (lastRow > 3) {
        STATEMENT_SHEET.deleteRow(lastRow);
        lastRow--;
      }
    } catch (e) {
      Logger.log("Ошибка пересоздания таблицы!");
      Logger.log(e);

      //создаю таблицу с именем педагога и добавляю туда название выбраного месяца
      STATEMENT_TEMPLATE.copyTo(ADMIN_PANEL_SPREADSHEET)
        .setName("Выписки " + teacher.name)
        .getRange(1, 1, 1, 2)
        .setValue([["Месяц:", monthName]]); //имя листа выписки+имя педагога
    }
  }

  getDataForStatement(teachers, monthName);
}

//собираю занятия со всех личных кабинетов педагогов
function getDataForStatement(teachers, monthName) {
  var monthAndYear = getMonthAndYearFromName(monthName),
    subscribeSheet = getSpreadsheetById(
      ADMIN_PANEL_SPREADSHEET_ID
    ).getSheetByName("абонементы"),
    subscData = subscribeSheet.getDataRange().getValues(), //данные абонементов
    month = monthAndYear.split("/")[0],
    year = monthAndYear.split("/")[1],
    startCellsInfo = 8, //начало колонок с иформ. о занятиях
    countCellsInfo = 5, //кол-во колонок с иформ. о занятиях
    countLessons,
    allTeachers = {},
    sum,
    countPositive,
    countNegative,
    countNeutral,
    lessonsInfo;

  for (var teacher = 0; teacher < teachers.length; teacher++) {
    const SUBSCRIPTIONS_VALUES = SpreadsheetApp.openById(
      teachers[teacher].sheetId
    )
      .getSheetByName("Абонементы")
      .getDataRange()
      .getValues();

    for (var i = 1; i < SUBSCRIPTIONS_VALUES.length; i++) {
      countLessons = SUBSCRIPTIONS_VALUES[i][6]; //кол-во занятий

      //иду по занятиям в одном абонементе
      for (
        var j = startCellsInfo;
        j < startCellsInfo + countCellsInfo * countLessons;
        j += 5
      ) {
        if (SUBSCRIPTIONS_VALUES[i][j] instanceof Date) {
          if (
            SUBSCRIPTIONS_VALUES[i][j].getMonth() == month &&
            SUBSCRIPTIONS_VALUES[i][j].getFullYear() == year
          ) {
            lessonsInfo = []; //вся информация за занятие
            countPositive = 0;
            countNegative = 0;
            countNeutral = 0;
            sum = 0;
            lessonsInfo.push(SUBSCRIPTIONS_VALUES[i][j]); // дата
            lessonsInfo.push(SUBSCRIPTIONS_VALUES[i][0]); //  абонемент
            lessonsInfo.push((j - startCellsInfo) / countCellsInfo + 1); //  занятие
            lessonsInfo.push(SUBSCRIPTIONS_VALUES[i][j + 4]); //  статус
            //счет начисления
            //идем по абонементам в админке
            for (var k = 1; k < subscData.length; k++) {
              //если нашли такой тип занятия
              if (SUBSCRIPTIONS_VALUES[i][3] == subscData[k][0]) {
                if (SUBSCRIPTIONS_VALUES[i][j + 4] == "✅") {
                  sum =
                    (SUBSCRIPTIONS_VALUES[i][5] / SUBSCRIPTIONS_VALUES[i][6]) *
                    subscData[k][2];
                  lessonsInfo.push(sum);
                  countPositive = 1;
                  break;
                } else if (SUBSCRIPTIONS_VALUES[i][j + 4] == "⛔") {
                  sum =
                    (SUBSCRIPTIONS_VALUES[i][5] / SUBSCRIPTIONS_VALUES[i][6]) *
                    subscData[k][4];
                  lessonsInfo.push(sum);
                  countNegative = 1;
                  break;
                } else {
                  lessonsInfo.push(0);
                  countNeutral = 1;
                  break;
                }
              }
            }
            try {
              //если уже был такой педагог, добавляем новое занятие
              lessonsInfo.unshift(
                allTeachers[SUBSCRIPTIONS_VALUES[i][j + 3]].length + 1
              );
              allTeachers[SUBSCRIPTIONS_VALUES[i][j + 3]].push(lessonsInfo);
              allTeachers[SUBSCRIPTIONS_VALUES[i][j + 3]][0][0] += sum;
              allTeachers[
                SUBSCRIPTIONS_VALUES[i][j + 3]
              ][0][1] += countPositive;
              allTeachers[
                SUBSCRIPTIONS_VALUES[i][j + 3]
              ][0][2] += countNegative;
              allTeachers[SUBSCRIPTIONS_VALUES[i][j + 3]][0][3] += countNeutral;
            } catch (e) {
              //если нет, добавляем педагога
              allTeachers[SUBSCRIPTIONS_VALUES[i][j + 3]] = []; // cума, статус1 , ст2, ст3
              allTeachers[SUBSCRIPTIONS_VALUES[i][j + 3]].push([
                +sum,
                +countPositive,
                +countNegative,
                +countNeutral
              ]);
              //и занятие
              lessonsInfo.unshift(1);
              allTeachers[SUBSCRIPTIONS_VALUES[i][j + 3]].push(lessonsInfo);
            }
          }
        }
      }
    }
  }

  addLessonsToAllTables(allTeachers, teachers);
}

//добавляю занятия в выписки всех педагогов
function addLessonsToAllTables(allTeachers, teachers) {
  var subscribeSheet;

  //прохожу по всем учителям с педагогов
  for (var i = 0; i < teachers.length; i++) {
    //иду по собраных с личных кабинетов данным
    for (var teacher in allTeachers) {
      //если нашла такого педагога
      if (teacher == teachers[i].name) {
        //открываю его таблицу
        subscribeSheet = getSpreadsheetById(
          ADMIN_PANEL_SPREADSHEET_ID
        ).getSheetByName("Выписки " + teachers[i].name);

        //и добавляю все его занятия
        for (var row = 1; row < allTeachers[teacher].length; row++) {
          subscribeSheet.appendRow(allTeachers[teacher][row]);
        }

        //добавляю кол-во статусов
        subscribeSheet
          .getRange(1, 4, 1, 3)
          .setValues([
            [
              "✅ - " + allTeachers[teacher][0][1],
              "⛔ - " + allTeachers[teacher][0][2],
              "? - " + allTeachers[teacher][0][3]
            ]
          ]);

        //добавляю сумму
        subscribeSheet.appendRow([
          "",
          "Сумма",
          "",
          "",
          "",
          allTeachers[teacher][0][0]
        ]);

        break;
      }
    }
  }
}

function getMonthAndYearFromName(monthName) {
  var result = "";

  for (var i = 0; i < MONTHS.length; i++) {
    if (monthName.split(" ")[0] == MONTHS[i]) {
      result += i + "/";
      break;
    }
  }

  result += monthName.split(" ")[1];

  return result;
}

//Импорт служебных данных из таблицы "Админка" из вкладок "Педагоги", "Аудитории", "Дисциплины", "Абонементы" в вкладку "Справочник" таблицы "Абонементы"...
function importData() {
  const SUBSCRIPTIONS_SHEET = getSpreadsheetById(
    SUBSCRIPTIONS_SPREADSHEET_ID
  ).getSheetByName("Справочник");

  const ADMIN_PANEL_SPREADSHEET = getSpreadsheetById(
    ADMIN_PANEL_SPREADSHEET_ID
  );

  const TEACHERS_VALUES = ADMIN_PANEL_SPREADSHEET.getSheetByName("педагоги")
    .getDataRange()
    .getValues();

  const ROOMS_VALUES = ADMIN_PANEL_SPREADSHEET.getSheetByName("аудитории")
    .getDataRange()
    .getValues();

  const LESSONS_VALUES = ADMIN_PANEL_SPREADSHEET.getSheetByName("дисциплины")
    .getDataRange()
    .getValues();

  const SUBSCRIPTIONS_VALUES = ADMIN_PANEL_SPREADSHEET.getSheetByName(
    "абонементы"
  )
    .getDataRange()
    .getValues();

  const PAY_TYPE = [["карта"], ["мих"], ["бас"]];

  SUBSCRIPTIONS_SHEET.clear();

  SUBSCRIPTIONS_SHEET.getRange(
    1,
    1,
    TEACHERS_VALUES.length,
    TEACHERS_VALUES[0].length
  ).setValues(TEACHERS_VALUES);

  SUBSCRIPTIONS_SHEET.getRange(
    1,
    8,
    ROOMS_VALUES.length,
    ROOMS_VALUES[0].length
  ).setValues(ROOMS_VALUES);

  SUBSCRIPTIONS_SHEET.getRange(
    1,
    11,
    LESSONS_VALUES.length,
    LESSONS_VALUES[0].length
  ).setValues(LESSONS_VALUES);

  SUBSCRIPTIONS_SHEET.getRange(
    1,
    14,
    SUBSCRIPTIONS_VALUES.length,
    SUBSCRIPTIONS_VALUES[0].length
  ).setValues(SUBSCRIPTIONS_VALUES);

  SUBSCRIPTIONS_SHEET.getRange(
    2,
    19,
    PAY_TYPE.length,
    PAY_TYPE[0].length
  ).setValues(PAY_TYPE);
}

// формирование данных
function getDataFromSchedule() {
  var dataSchedule,
    bgrdSchedule,
    sheets = getSpreadsheetById(SCHEDULE_SPREADSHEET_ID).getSheets();
  var allSheets = {};
  var today = new Date();
  var sheetNames = []; //назвы таблиц, за следующие 4 месяца
  sheetNames.push(getSheetNameThisMonth(today));
  sheetNames.push(
    getSheetNameThisMonth(
      new Date(today.getFullYear(), today.getMonth() + 1, 1)
    )
  );
  sheetNames.push(
    getSheetNameThisMonth(
      new Date(today.getFullYear(), today.getMonth() + 2, 1)
    )
  );
  sheetNames.push(
    getSheetNameThisMonth(
      new Date(today.getFullYear(), today.getMonth() + 3, 1)
    )
  );

  //иду по всем листам таблицы
  for (var s = 0; s < sheets.length; s++) {
    //иду по листам  четырех месяцов
    for (var lastSheets = 0; lastSheets < sheetNames.length; lastSheets++) {
      //если одинаковые названия
      if (sheetNames[lastSheets] == sheets[s].getSheetName()) {
        //считываю все даные и бгр
        dataSchedule = sheets[s].getDataRange().getValues();
        bgrdSchedule = sheets[s].getDataRange().getBackgrounds();
        //добавляю этот лист в конечный обьект
        allSheets[sheets[s].getSheetName()] = {};
        //прохожу по даным этого листа
        for (var i = 0; i < dataSchedule.length; i++) {
          //исчу рядок с полем дата
          //если нахожу
          if (dataSchedule[i][0] == "дата") {
            //прохожу по этом ряду
            for (var j = 1; j < dataSchedule[i].length; j++) {
              if (dataSchedule[i][j] != "") {
                //добавляю все даты, тоесть не пустые поля
                allSheets[sheets[s].getSheetName()][dataSchedule[i][j]] = {};
                //добавляю места
                var startClassIndexInDay = j; //индекс начала даты в рядке
                //прохожу по первой строке с классами
                for (
                  var startClassIndexInDay = j;
                  startClassIndexInDay < dataSchedule[1].length;
                  startClassIndexInDay++
                ) {
                  // добаляю классы в дату, пока не найду синий разделитель(разделитель между днями)
                  if (bgrdSchedule[1][startClassIndexInDay] == "#9fc5e8") {
                    break;
                  } else if (dataSchedule[1][startClassIndexInDay] != "") {
                    //если не пусто добаляю класс
                    allSheets[sheets[s].getSheetName()][dataSchedule[i][j]][
                      dataSchedule[1][startClassIndexInDay]
                    ] = [];
                    //прохожу по всем часом
                    for (var k = 0; k < 24; k++) {
                      //если нахожу cиний разделитель или конец таблицы, часы закончились, выхожу из цикла;
                      if (
                        bgrdSchedule[i + 2 + k][0] == "#9fc5e8" ||
                        i + 2 + k == dataSchedule.length - 1
                      ) {
                        break;
                      }
                      //если  пусто, добавляю свободный час
                      if (dataSchedule[i + 2 + k][startClassIndexInDay] == "") {
                        //беру время без пробелов
                        var timeWithOutSpace = dataSchedule[
                          i + 2 + k
                        ][0].replace(/\s/g, ""); // ?????
                        allSheets[sheets[s].getSheetName()][dataSchedule[i][j]][
                          dataSchedule[1][startClassIndexInDay]
                        ].push(timeWithOutSpace);
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

  return allSheets;
}

//бронирую время
function PassForm2(allData) {
  PropertiesService.getScriptProperties().setProperty("_flagAlert", false);
  var newData = JSON.parse(allData)[0]; //занятия с абонемента
  var dataFromTable = JSON.parse(allData)[1]; //данные взятые с таблицы (4 месяца)
  var idTeacher = JSON.parse(allData)[2]; //ид учителя
  var numberOfAbon = JSON.parse(allData)[3]; ////номер абонемента

  var spreadSheet = SpreadsheetApp.openById(idTeacher),
    sheet = spreadSheet.getSheetByName("Абонементы"),
    dataSubscription = sheet.getDataRange().getValues(),
    row;
  var arrayResult = [];
  var passedLessons = 0;
  for (var i = 1; i < dataSubscription.length; i++) {
    if (numberOfAbon == dataSubscription[i][0]) {
      row = i + 1;
    }
  }
  //заносим данные в личный кабинет
  for (var i = 0; i < newData.length; i++) {
    arrayResult.push(newData[i].date == undefined ? "" : newData[i].date);
    arrayResult.push(newData[i].time == undefined ? "" : newData[i].time);
    arrayResult.push(newData[i].class == undefined ? "" : newData[i].class);
    arrayResult.push(newData[i].teacher == undefined ? "" : newData[i].teacher);
    arrayResult.push(newData[i].status == undefined ? "" : newData[i].status);
    if (newData[i].status != "-") {
      passedLessons++;
    }
  }

  //cтарые данные для удаления
  try {
    var lessonsForDelete = [];
    var lessDel = sheet.getRange(row, 9, 1, arrayResult.length).getValues();
    for (var i = 0; i < lessDel[0].length; i += 5) {
      var lesson = {};
      lesson.date = lessDel[0][0 + i];
      lesson.time = lessDel[0][1 + i];
      lesson.class = lessDel[0][2 + i];
      lesson.teacher = lessDel[0][3 + i];
      lesson.client = dataSubscription[row - 1][1];
      lessonsForDelete.push(lesson);
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
    setLessonToSchedule(newData, dataFromTable, sheet, row);
  } catch (e) {
    Logger.log("Ошибка записи новых занятий");
    Logger.log(e);
  }
}

//добавляю занятия в расписание
function setLessonToSchedule(data, dataFromTable, sheet, row) {
  const SCHEDULE_SPREADSHEET = getSpreadsheetById(SCHEDULE_SPREADSHEET_ID);

  for (var less = 0; less < data.length; less++) {
    try {
      var checkedDate = data[less].date.toString().split(".")[0];
      var sheetName = getSheetNameThisMonth(data[less]);

      const SCHEDULE_SHEET = SCHEDULE_SPREADSHEET.getSheetByName(
        sheetName
      );
      const SCHEDULE_VALUES = SCHEDULE_SHEET.getDataRange().getValues();

      for (var i = 0; i < SCHEDULE_VALUES.length; i++) {
        //ищу поле с датой
        if (SCHEDULE_VALUES[i][0] == "дата") {
          for (var j = 1; j < SCHEDULE_VALUES[i].length; j++) {
            //ищу выбраную дату
            if (+SCHEDULE_VALUES[i][j] == +checkedDate) {
              //ищу выбранный клас в этой дате
              var indexColumn;
              for (var k = j; k < SCHEDULE_VALUES[i].length; k++) {
                if (SCHEDULE_VALUES[1][k] == data[less].class) {
                  indexColumn = k + 1;
                  break;
                }
              }

              //ищу в этой колонке выбраное время? 24-макс. кол. часов
              for (var p = i + 2; p < SCHEDULE_VALUES.length; p++) {
                if (
                  SCHEDULE_VALUES[p][0].replace(/\s/g, "") ==
                  data[less].time.replace(/\s/g, "")
                ) {
                  //проверяю точно ли пусто в таблице и не его ли там урок
                  //если нет
                  if (SCHEDULE_SHEET.getRange(p + 1, indexColumn).getValue() != "") {
                    if (
                      SCHEDULE_SHEET.getRange(p + 1, indexColumn).getValue() !=
                        data[less].teacher ||
                      SCHEDULE_SHEET.getRange(p + 1, indexColumn + 1).getValue() !=
                        data[less].client
                    ) {
                      updateIfErrorSet(
                        data[less].date,
                        data[less].time,
                        data[less].class,
                        data[less].status,
                        dataFromTable,
                        SCHEDULE_SHEET.getRange(p + 1, indexColumn).getValue(),
                        SCHEDULE_SHEET.getRange(p + 1, indexColumn + 1).getValue(),
                        data[less].teacher,
                        data[less].client,
                        sheet,
                        row
                      );
                    }
                    break;
                  } else {
                    //если да,вставляю значения в расп и в откр расписание
                    SCHEDULE_SHEET
                      .getRange(p + 1, indexColumn)
                      .setValue(data[less].teacher)
                      .setBackground(data[less].bgr)
                      .setFontColor(data[less].color);

                    SCHEDULE_SHEET
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
  status,
  dataFromTable,
  teacherFromTable,
  clientFromTable,
  tempTeacher,
  tempClient,
  sheet,
  row
) {
  var logger = getSpreadsheetById(SCHEDULE_SPREADSHEET_ID).getSheetByName(
    "logger"
  );

  logger.appendRow([
    new Date(),
    "добавление",
    date,
    time,
    _class,
    dataFromTable,
    teacherFromTable,
    clientFromTable,
    tempTeacher,
    tempClient
  ]);

  var dataSubscription = sheet.getDataRange().getValues();

  //затираю то занятие, где возникла ошибка
  var day = date.split(".")[0],
    month = date.split(".")[1],
    year = date.split(".")[2];

  //прохожу по абонементу
  if (status != "⛔") {
    for (var i = 8; i < dataSubscription[row - 1].length; i += 5) {
      try {
        //если это та дата, то время и год
        if (
          dataSubscription[row - 1][i].getFullYear() == year &&
          dataSubscription[row - 1][i].getDate() == day &&
          dataSubscription[row - 1][i].getMonth() + 1 == month &&
          dataSubscription[row - 1][i + 1] == time &&
          dataSubscription[row - 1][i + 2] == _class &&
          dataSubscription[row - 1][i + 3] == tempTeacher
        ) {
          //удаляю
          sheet.getRange(row, i + 1, 1, 5).setValue("");
          break;
        }
      } catch (e) {}
    }

    teacherFromTable = teacherFromTable == undefined ? "" : teacherFromTable;
    clientFromTable = clientFromTable == undefined ? "" : clientFromTable;

    PropertiesService.getScriptProperties().setProperty("_date", date);
    PropertiesService.getScriptProperties().setProperty("_time", time);
    PropertiesService.getScriptProperties().setProperty("_class", _class);
    PropertiesService.getScriptProperties().setProperty(
      "_teacherFromTable",
      teacherFromTable
    );
    PropertiesService.getScriptProperties().setProperty(
      "_clientFromTable",
      clientFromTable
    );
    PropertiesService.getScriptProperties().setProperty("_flagAlert", "true");

    //название листа в таблице в зависимости от месяца
    var sheetName = MONTHS[month - 1];

    //добавляю до названия год
    sheetName += " " + year;

    MailApp.sendEmail({
      to: ADMINISTRATOR_EMAIL,
      subject: "Место занято",
      htmlBody:
        "<p>Таблица: " +
        sheet.getName() +
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
  var logger = getSpreadsheetById(SCHEDULE_SPREADSHEET_ID).getSheetByName(
    "logger"
  );

  logger.appendRow([
    new Date(),
    "удаление",
    date,
    time,
    _class,
    teacherFromTable,
    clientFromTable,
    tempTeacher,
    tempClient
  ]);

  MailApp.sendEmail({
    to: ADMINISTRATOR_EMAIL,
    subject: "Удаление урока",
    htmlBody:
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
  var sSchedule, dataSchedule;

  for (var less = 0; less < data.length; less++) {
    try {
      var checkedDate = data[less].date.getDate();
      var sheetName = getSheetNameThisMonth(data[less]);
      sSchedule = getSpreadsheetById(SCHEDULE_SPREADSHEET_ID).getSheetByName(
        sheetName
      );
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
      Logger.log("Ошибка удаления старых данных");
    }
  }
}

// принимает номер id и номер абонемента
function getAllLessons(ssTeacherId, numAbon) {
  var sheetTeacher = SpreadsheetApp.openById(ssTeacherId).getSheetByName(
    "Абонементы"
  );

  var teacherAbons = sheetTeacher
    .getRange(1, 1, sheetTeacher.getLastRow(), sheetTeacher.getLastColumn())
    .getValues();
  //собираю данные о занятиях педагога
  var allLessons = {};

  var lessonsInfo;
  var countLessons; //кол-во занятий
  var startCellsInfo = 8; //начало колонок с иформ. о занятиях
  var countCellsInfo = 5; //кол-во колонок с иформ. о занятиях

  for (var i = 1; i < teacherAbons.length; i++) {
    if (teacherAbons[i][0] == numAbon) {
      countLessons = teacherAbons[i][6];
      lessonsInfo = [];
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
        lessonsInfo[lessonsInfo.length - 1].push(teacherAbons[i][j]);
      }
      allLessons[teacherAbons[i][0]] = lessonsInfo;

      break;
    }
  }

  var res = allLessons[numAbon].slice();
  return res;
}

//принимаю новые данные
function PassFormOne(newData) {
  newData = JSON.parse(newData);

  var sheetTeacher = SpreadsheetApp.openById(newData[0].id).getSheetByName(
    "Абонементы"
  );
  var teachersData = sheetTeacher.getDataRange().getValues();

  var numberClient = newData[1]["number"];
  var comment = newData[1]["comment"];

  for (var i = 0; i < teachersData.length; i++) {
    if (teachersData[i][0] == numberClient) {
      sheetTeacher.getRange(i + 1, 5).setValue(comment);
    }
  }
}
