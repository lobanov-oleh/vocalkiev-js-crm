function onOpen(e) {
  const menuItems = [
    { c: "Открыть абонемент", f: "reservingSubscription" },
    { c: "Пересчитать выписку", f: "makeStatistic" },
    { c: "Архивировать абоненменты", f: "archiving" }
  ];

  createMenu(menuItems);
}
