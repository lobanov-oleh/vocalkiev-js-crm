function onOpen() {
  const menuItems = [
    { c: "Обновить педагогов", f: "resetTeachers" },
    { c: "Обновить справочник педагога", f: "reloadNote" },
    { c: "Пересчет выписок", f: "updateAccountStatement" }
  ];

  createMenu(menuItems);
}
