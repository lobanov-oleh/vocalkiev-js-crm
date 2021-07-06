function onOpen() {
  const menuItems = [
    { c: "Добавить абонемент", f: "addClient" },
    { c: "Обновить израсход. абон", f: "checkCountLesson" },
    { c: "Обновить архив", f: "clientArchive" }
  ];

  createMenu(menuItems);
}
