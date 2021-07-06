//Получение данных для модального окна
function getDataForModal() {
  var clients = getSpreadsheetById(SUBSCRIPTIONS_SPREADSHEET_ID);
  var archiveSheet = clients.getSheetByName("Абонементы в работе");
  var archiveSheetData = archiveSheet.getDataRange().getValues();
  var noteData = clients
    .getSheetByName("Справочник")
    .getDataRange()
    .getValues();

  //собираю массивы из справочника для выпадающего списка выбора учителя, дисциплины, и типа
  var nameLessons = [],
    teachersNames = [],
    types = [],
    payType = [];

  for (var i = 1; i < noteData.length; i++) {
    if (noteData[i][11] !== "") {
      nameLessons.push(noteData[i][11]);
    }
    if (noteData[i][3] !== "") {
      var teacher = {};
      teacher.name = noteData[i][3];
      var twoLett = noteData[i][4][0] + noteData[i][4][1];
      teacher.twoLett = twoLett;
      teachersNames.push(teacher);
    }
    if (noteData[i][13]) {
      var type = {};
      type.typeLesson = noteData[i][13];
      type.price = noteData[i][14];
      type.days = noteData[i][16];
      types.push(type);
    }
    if (noteData[i][18] !== "") {
      payType.push(noteData[i][18]);
    }
  }

  //извлекаю последний номер абонента и +1, если его нет, захожу в архив, если и там нет, то нулевой заказ.
  var numbersClient = [];
  var lastNumber;

  if (archiveSheetData[1] !== undefined && archiveSheetData[1] !== "") {
    for (
      i = archiveSheetData.length - 1;
      i >= archiveSheetData.length - archiveSheetData.length;
      --i
    ) {
      numbersClient.push(archiveSheetData[i][1]);
    }

    for (var j = 0; j < numbersClient.length; j++) {
      var checkNum = numbersClient[j].split("/");

      var checkNumber = checkNum[1];

      if (isNaN(checkNumber) == false && checkNumber !== "") {
        lastNumber = numbersClient[j];
        break;
      }
    }
  }

  var numberDrop = lastNumber.split("/");

  var numberPlusOne = numberDrop[1];
  numberPlusOne++;

  numberDrop.splice(1, 1, numberPlusOne);
  var newNumber = numberDrop.join("/");

  //формирую данные для модального окна
  var res = {};
  res.number = newNumber;
  res.lessons = nameLessons;
  res.type = types;
  res.teachers = teachersNames;
  res.payType = payType;

  return JSON.stringify(res);
}

//Добавляем новый аобнемент
function PassForm(newData) {
  newData = JSON.parse(newData);

  var clients = getSpreadsheetById(SUBSCRIPTIONS_SPREADSHEET_ID);
  var archiveSheet = clients.getSheetByName("Абонементы в работе");

  var noteData = clients
    .getSheetByName("Справочник")
    .getDataRange()
    .getValues();

  for (var i = 0; i < noteData.length; i++) {
    if (newData[5] == noteData[i][3]) {
      var idTeacherSS = noteData[i][4];
    }
  }

  // добавляем строку в абонементах
  archiveSheet.appendRow(newData);
  // копируем формат даты с прошлой записи
  var cellformat = archiveSheet
    .getRange(archiveSheet.getLastRow() - 1, 1, 2, 1)
    .copyFormatToRange(
      0,
      1,
      1,
      archiveSheet.getLastRow(),
      archiveSheet.getLastRow() + 1
    );

  // хрен пойми что делает кусок кода ниже. Вова, допиши комменты.
  var noteData = clients
    .getSheetByName("Справочник")
    .getDataRange()
    .getValues();

  //Добавить новый абонемент в личный кабинет преподавателя и отправить ему письмо
  for (var i = 0; i < noteData.length; i++) {
    if (newData[5] == noteData[i][3]) {
      var idTeacherSS = noteData[i][4];
      var emailTeacher = noteData[i][2];
      var newSSheet = SpreadsheetApp.openById(idTeacherSS);
      var arrForTeacher = [];
      arrForTeacher.push(
        newData[1],
        newData[4],
        newData[2],
        newData[3],
        newData[12],
        newData[6],
        newData[7]
      );
      newSSheet.appendRow(arrForTeacher);
      if (newData[7] < 8) {
        var countGrayCell = 5 * (8 - newData[7]);

        var spreadSheet = SpreadsheetApp.openById(idTeacherSS),
          sCatalog = spreadSheet.getSheetByName("Абонементы");

        var cellTeach = sCatalog.getRange(
          sCatalog.getLastRow(),
          sCatalog.getMaxColumns() - countGrayCell + 1,
          1,
          countGrayCell
        );

        cellTeach.setBackground("#cfd1d0");
      }

      if (emailTeacher !== "") {
        MailApp.sendEmail({
          to: emailTeacher,
          subject: "Добавлен новый клиент.",
          htmlBody:
            "Здравствуйте," +
            newData[5] +
            '.<br/>Поздравляем! <br/> В Ваш личный кабинет добавлен новый клиент. <br/> Для просмотра подробностей перейдите по <b><a href="https://docs.google.com/spreadsheets/d/' +
            idTeacherSS +
            '" >ссылке</a><b> на личный кабинет преподавателя школы вокала. '
        });
      }
    }
  }
}
