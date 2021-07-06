const EMAIL_TO_NEW_TEACHER = {
  subject: "Школа вокала. Личный кабинет.",
  body: (teacherName, fileId) =>
    `Здравствуйте, ${teacherName}.<br/>` +
    `Коллектив школы вокала рад приветствовать Вас в нашем коллективе, и поздравляет с началом работы в школе вокала!<br/>` +
    `Перейдя по <b><a href="https://docs.google.com/spreadsheets/d/${fileId}" >ссылке</a></b>, вы откроете личный кабинет.<br/>`
};
