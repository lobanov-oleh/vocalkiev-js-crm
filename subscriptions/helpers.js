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
