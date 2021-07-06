const SPREADSHEETS = {};

function getSpreadsheetById(id) {
  return SPREADSHEETS[id]
    ? SPREADSHEETS[id]
    : (SPREADSHEETS[id] = SpreadsheetApp.openById(id));
}

function createMenu(items) {
  const menu = SpreadsheetApp.getUi().createMenu("Меню");

  items.map(item => menu.addItem(item.c, item.f).addSeparator());

  menu.addToUi();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function showDialog(name, res) {
  const TEMPLATE = HtmlService.createTemplateFromFile(name);
  TEMPLATE.res = res;

  const HTML = TEMPLATE.evaluate().getContent();

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(HTML)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(960)
      .setHeight(620),
    " "
  );
}

function getSheetNameThisMonth(data) {
  var month, year;

  try {
    month = data.date.getMonth();
    year = data.date.getFullYear();
  } catch (e) {
    try {
      month = data.date.split(".")[1] - 1;
      year = data.date.split(".")[2];
    } catch (e) {
      month = data.getMonth();
      year = data.getFullYear();
    }
  }

  var nameSheet = MONTHS[month] + " " + year;

  return nameSheet;
}
