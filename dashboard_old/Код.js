function onOpen(e) {
  SpreadsheetApp.getUi()
  .createMenu('Меню')
  .addSeparator()
   .addItem('Открыть абонемент', 'reservingSubscription')
  .addSeparator()
  .addItem('Пересчитать выписку', 'makeStatistic')
  .addSeparator()
  .addItem('Архивировать абоненменты', 'archiving')
  .addSeparator()
  .addToUi();
  
}


