function resetTeachers() {
  const DASHBORDS_FOLDER = DriveApp.getFolderById(DASHBOARDS_FOLDER_ID);
  const TEACHERS_SHEET = getSpreadsheetById(
    ADMIN_PANEL_SPREADSHEET_ID
  ).getSheetByName("педагоги");
  const TEACHERS_VALUES = TEACHERS_SHEET.getDataRange().getValues();

  // getSpreadsheetById(DASHBOARD_TEMPLATE_SPREADSHEET_ID).setFrozenColumns(8);

  const teachers = [];
  for (let i = 1; i < TEACHERS_VALUES.length; i++) {
    const row = TEACHERS_VALUES[i];

    const teacher = {
      row: i + 1,
      name: row[0],
      phone: "" + row[1],
      email: row[2],
      style: row[3],
      sheetId: row[4],
      status: row[5]
    };

    teachers.push(teacher);
  }

  for (const teacher of teachers) {
    if (teacher.status == "Не активен") {
      var delAcces = SpreadsheetApp.openById(teacher.sheetId);
      delAcces.removeEditor(teacher.email);
    }

    if (teacher.status == "Новый" && !teacher.sheetId) {
      const DASHBOARD_FILE = DriveApp.getFileById(
        DASHBOARD_TEMPLATE_SPREADSHEET_ID
      ).makeCopy(DASHBORDS_FOLDER);
      const DASHBOARD_FILE_ID = DASHBOARD_FILE.getId();
      const DASHBOARD_SPREADSHEET = SpreadsheetApp.openById(DASHBOARD_FILE_ID);
      const KEY = "" + Math.floor(Math.random() * 1000000 + 1);

      DASHBOARD_SPREADSHEET.rename(
        "Личный кабинет преподавателя " + teacher.name
      );

      if (teacher.email) {
        DASHBOARD_SPREADSHEET.addEditor(teacher.email);
      }

      TEACHERS_SHEET.getRange(teacher.row, 5).setValue(DASHBOARD_FILE_ID);
      TEACHERS_SHEET.getRange(teacher.row, 7).setValue(KEY);
      TEACHERS_SHEET.getRange(teacher.row, 6).setValue("Активен");

      if (teacher.email) {
        MailApp.sendEmail({
          to: teacher.email,
          subject: EMAIL_TO_NEW_TEACHER.subject,
          htmlBody: EMAIL_TO_NEW_TEACHER.body(teacher.name, DASHBOARD_FILE_ID)
        });
      }
    }
  }

  importData();
  reloadNote();
}

//обновляет вкладку "справочник" во всех таблицах педагогов
function reloadNote() {
  const ADMIN_PANEL_SPREADSHEET = getSpreadsheetById(
    ADMIN_PANEL_SPREADSHEET_ID
  );
  const TEACHERS_SHEET = ADMIN_PANEL_SPREADSHEET.getSheetByName("педагоги");
  const TEACHERS_VALUES = TEACHERS_SHEET.getDataRange().getValues();
  const TEACHERS_FONTS = TEACHERS_SHEET.getRange(
    1,
    4,
    TEACHERS_SHEET.getLastRow(),
    1
  ).getFontColors();
  const TEACHERS_BACKGROUNDS = TEACHERS_SHEET.getRange(
    1,
    4,
    TEACHERS_SHEET.getLastRow(),
    1
  ).getBackgrounds();
  const ROOMS_VALUES = ADMIN_PANEL_SPREADSHEET.getSheetByName("аудитории")
    .getDataRange()
    .getValues();
  const SUBSCRIPTIONS_VALUES = ADMIN_PANEL_SPREADSHEET.getSheetByName(
    "абонементы"
  )
    .getDataRange()
    .getValues();

  const pakageOut = [];
  for (var i = 1; i < SUBSCRIPTIONS_VALUES.length; i++) {
    pakageOut.push([
      SUBSCRIPTIONS_VALUES[i][0], // type
      SUBSCRIPTIONS_VALUES[i][2], // percent
      SUBSCRIPTIONS_VALUES[i][4] // fired
    ]);
  }

  const teachersID = [];
  const outToTeacher = [];
  const result = [];

  for (var i = 1; i < TEACHERS_VALUES.length; i++) {
    outToTeacher.push([TEACHERS_VALUES[i][0], TEACHERS_VALUES[i][3]]);
    teachersID.push(TEACHERS_VALUES[i][4]);
  }

  var rooms = [];
  for (var i = 1; i < ROOMS_VALUES.length; i++) {
    rooms.push(ROOMS_VALUES[i][1]);
  }

  for (var i = 0; i < outToTeacher.length; i++) {
    result.push([rooms[i], outToTeacher[i][0], outToTeacher[i][1]]);
  }

  for (var i = 0; i < teachersID.length; i++) {
    var TEACHER_SPREADSHEET = SpreadsheetApp.openById(teachersID[i]);
    var CATALOG_SHEET = TEACHER_SPREADSHEET.getSheetByName("Справочник");

    CATALOG_SHEET.getRange(
      2,
      1,
      result.length,
      result[0].length
    ).setValues(result);

    const TEACHER_NAME = TEACHER_SPREADSHEET.getName().split(" ").slice(3).join(" ");
    CATALOG_SHEET.getRange(2, 4).setValue(TEACHER_NAME);

    var outForSymbols = CATALOG_SHEET.getRange(
      2,
      5,
      SYMBOLS.length,
      SYMBOLS[0].length
    );
    outForSymbols.setValues(SYMBOLS);

    var outForTypes = CATALOG_SHEET.getRange(
      2,
      6,
      pakageOut.length,
      pakageOut[0].length
    );
    outForTypes.setValues(pakageOut);

    // нужно подкрасить текущего преподавателя ячейку в справочнике так как в админке
    // иду по всем педагогам в админке и сверяю по id таблицы
    for (var ar = 0; ar < TEACHERS_VALUES.length; ar++) {
      //находим
      if (TEACHERS_VALUES[ar][4] == teachersID[i]) {
        // строка этого препода в админке
        var rowTeacherInAdmin = ar + 1;
        //ячейка которую с которой берем стиль
        var cell = TEACHERS_SHEET.getRange(rowTeacherInAdmin, 4);
        // берем фон цвет
        var color = cell.getBackground();
        // берем шрифта цвет
        var fontColor = cell.getFontColor();
        // ячейка которую нужно подкрасить в таблице препода
        var cellTeach = CATALOG_SHEET.getRange(2, 4);

        //красим
        cellTeach.setBackground(color);
        cellTeach.setFontColor(fontColor);

        break;
      }
    }

    // подкрашиваем стили в справочнике преподавателя
    CATALOG_SHEET
      .getRange(1, 3, TEACHERS_SHEET.getLastRow(), 1)
      .setBackgrounds(TEACHERS_BACKGROUNDS)
      .setFontColors(TEACHERS_FONTS);
  }
}

//обновление выписок
function updateAccountStatement() {
  const htmlOutputText = HtmlService.createTemplateFromFile(
    "chooseMonthForUpdateStatement"
  ).evaluate().getContent();

  const htmlOutput = HtmlService.createHtmlOutput(htmlOutputText)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(200)
    .setHeight(95);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Выберите месяц:");
}
