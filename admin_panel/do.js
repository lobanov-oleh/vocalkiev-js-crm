function doGet(e) {
  var ss = getSpreadsheetById(ADMIN_PANEL_SPREADSHEET_ID);
  var loggerSheet = ss.getSheetByName("logger");
  var teachersSheet = ss.getSheetByName("педагоги");
  var teachersData = teachersSheet
    .getRange(1, 1, teachersSheet.getLastRow(), teachersSheet.getLastColumn())
    .getValues();
  var ssTeacherId = false;

  // берем ключ и по нему определяем id таблицы педагога
  var key = e.parameter.key;
  var flag = false;
  //  var key = 2;
  for (var i = 1; i < teachersData.length; i++) {
    if (teachersData[i][6] == key) {
      ssTeacherId = teachersData[i][4];
      flag = true;
      break;
    }
  }
  if (flag == false) {
    return HtmlService.createHtmlOutput(
      " <p style='text-align:  center'>Мобильная версия находится в доработке. <br>Воспользуйтесь пока таблицей личного кабинета. <br>Спасибо. <p>"
    );
  }
  // loggerSheet.appendRow([new Date(),'key',key]);

  // идем в таблицу педагога и тащим данные с листа по открытым абонементам
  var ssTeacher = SpreadsheetApp.openById(ssTeacherId);
  var sheetTeacher = ssTeacher.getSheetByName("Абонементы");
  var teacherAbons = sheetTeacher
    .getRange(1, 1, sheetTeacher.getLastRow(), sheetTeacher.getLastColumn())
    .getValues();

  // Извлекаю имя текущего преподавателя
  var note = ssTeacher.getSheetByName("Справочник"),
    nameTeacher = note.getRange(2, 4).getValue();

  var data = [];

  for (var i = 1; i < teacherAbons.length; i++) {
    // собираю занятия абонемента
    var lessons = [];
    var q = teacherAbons[i][6];
    for (var j = 1; j <= q; j++) {
      lessons.push([teacherAbons[i][3 + 5 * q]]);
    }

    // собираю данные по абонементу
    oneClient = {
      number: teacherAbons[i][0],
      client: teacherAbons[i][1],
      disc: teacherAbons[i][2],
      type: teacherAbons[i][3],
      comment: teacherAbons[i][4],
      price: teacherAbons[i][5],
      q: teacherAbons[i][6],
      count: teacherAbons[i][7],
      lessons: lessons
    };
    data.push(oneClient);
  }

  loggerSheet.appendRow([new Date(), "key", key, data, JSON.stringify(data)]);

  var template = HtmlService.createTemplateFromFile("mobileCabinet");

  // собираю данные о занятиях педагога
  var allLessons = {};

  var lessonsInfo;
  var countLessons; // кол-во занятий
  var startCellsInfo = 8; // начало колонок с иформ. о занятиях
  var countCellsInfo = 5; // кол-во колонок с иформ. о занятиях

  for (var i = 1; i < teacherAbons.length; i++) {
    if (teacherAbons[i][0] !== "") {
      countLessons = teacherAbons[i][6];
      lessonsInfo = [];
      // иду по этому абонементу
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
    }
  }
  var getDataFromReserving = getDataFromSchedule();
  template.key = key;
  template.data = JSON.stringify(data);
  template.teacherName = nameTeacher;
  template.idOfTable = ssTeacherId;
  template.allLessons = JSON.stringify(allLessons);
  template.getDataFromReserving = JSON.stringify(getDataFromReserving);

  var sCatalog = ssTeacher.getSheetByName("Справочник"),
    dataCatalog = sCatalog.getDataRange().getValues(),
    bgrsCatalog = sCatalog.getDataRange().getBackgrounds(),
    colorsCatalog = sCatalog.getDataRange().getFontColors();

  //dopost возврат чего-то затем протестировать в постман, отправить пост запрос, получить нормальный ответ, в вьюхе делать аякс запрос,

  //формирую данные для модального окна
  var classes = [];
  var teachers_ = [];
  var bgrs = [];
  var colors = [];
  //номер абонемента, клиент, дисц, тип, комент, цена, кол, израсх.
  //добавляю все классы, всех педаг., все стиле
  for (var i = 1; i < dataCatalog.length; i++) {
    if (dataCatalog[i][0] != "") {
      classes.push(dataCatalog[i][0]);
    }
    if (dataCatalog[i][1] != "") {
      teachers_.push(dataCatalog[i][2]);
      bgrs.push(bgrsCatalog[i][2]);
      colors.push(colorsCatalog[i][2]);
    }
  }

  template.classes = classes;
  template.teachers = teachers_;
  template.bgrs = bgrs;
  template.colors = colors;

  var resHtml = template.evaluate();
  resHtml.addMetaTag("viewport", "width=device-width, initial-scale=1.0");
  return resHtml;
}

function doPost(e) {
  if (e.parameter.alert == "checkIn") {
    var alert = PropertiesService.getScriptProperties().getProperty(
      "_flagAlert"
    );
    Logger.log(alert);
    if (alert) {
      var objAlert = {};
      objAlert.time = PropertiesService.getScriptProperties().getProperty(
        "_time"
      );
      objAlert.date = PropertiesService.getScriptProperties().getProperty(
        "_date"
      );
      objAlert.class = PropertiesService.getScriptProperties().getProperty(
        "_class"
      );
      objAlert.teacherFromTable = PropertiesService.getScriptProperties().getProperty(
        "_teacherFromTable"
      );
      objAlert.clientFromTable = PropertiesService.getScriptProperties().getProperty(
        "_clientFromTable"
      );
      objAlert.flagAlert = PropertiesService.getScriptProperties().getProperty(
        "_flagAlert"
      );
      PropertiesService.getScriptProperties().setProperty(
        "_flagAlert",
        "false"
      );
      return ContentService.createTextOutput(JSON.stringify(objAlert));
    }
  } else {
    var newAbon = getAllLessons(e.parameter.id, e.parameter.numAbon);
    var getDataFromReserving = getDataFromSchedule();
    var objResult = {
      newAbon: newAbon,
      dataFromReserving: getDataFromReserving
    };
    return ContentService.createTextOutput(JSON.stringify(objResult));
  }
}
