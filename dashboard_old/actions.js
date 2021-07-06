//название таблицы в этом месяце
function getSheetNameThisMonth(data){
  var month,
      year;
  try{
    month = data.date.getMonth();
    year = data.date.getFullYear();
    
  }
  catch(e){
    try{
      month = data.date.split('.')[1]-1;
      year = data.date.split('.')[2];
    }
    catch(e){
      month = data.getMonth();
      year = data.getFullYear();
    }
  }
  
  var monthes =[
    'Январь',
    'Февраль',
    'Март',
    'Апрель',
    'Май',
    'Июнь',
    'Июль',
    'Август',
    'Сентябрь',
    'Октябрь',
    'Ноябрь',
    'Декабрь'
  ];
  var nameSheet = "";
  for(var i = 0; i < monthes.length; i++){
    if(i == month){
      nameSheet += monthes[i];
      break;
    }
  }
  nameSheet +=" "+year;
  return nameSheet;
}
