function makeStatistic(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAbon = ss.getSheetByName('Абонементы');
  var sheetAbonArch = ss.getSheetByName('Архив абонементов');
  if(sheetAbon.getLastColumn()<=1 && sheetAbonArch.getLastColumn()<=1) {return} // если пустой журнал абонементов - закончили.
  var abonsData = [];
  var abonsArchData = [];
  if(sheetAbon.getLastRow()>1){
    abonsData = sheetAbon.getRange(2, 1, sheetAbon.getLastRow()-1, sheetAbon.getLastColumn()).getValues();
  }
  Logger.log(abonsData)
  if(sheetAbonArch.getLastRow()>1){
    abonsArchData = sheetAbonArch.getRange(2, 1, sheetAbonArch.getLastRow()-1, sheetAbonArch.getLastColumn()).getValues();
  }
  
  var allAbons = abonsData.concat(abonsArchData);
  if(allAbons.length==0){return}
  
  var sheetSpravochnik = ss.getSheetByName('Справочник');
  var nachislenia = sheetSpravochnik.getRange(2, 6, sheetSpravochnik.getLastRow()-1, 3).getValues();
  var teacherName = sheetSpravochnik.getRange(2, 4).getValue();
  var symbolPlus = sheetSpravochnik.getRange(2, 5).getValue();
  var symbolMinus = sheetSpravochnik.getRange(3, 5).getValue();
  
  var sheetVipiska = ss.getSheetByName('Выписка по урокам');
  var date = sheetVipiska.getRange(1, 4).getValue();

  var res = [];
  var plusQ = 0;
  var minusQ = 0;
  var noDataQ = 0;
  
  var plusBgr = '#D9EAD4';
  var minusBgr = '#F3CCCD';
  var noDataBgr = '#FFF1CE';
  var num = 0;
  // проверяю внесена ли дата вообще
  if(date instanceof Date) {
    var month = date.getMonth()+1;
    var year = date.getFullYear();
    var monthYear = ""+month+year;
    var needLessons = []; // занятия только в этом мес
    //прохожусь по всем абонементам - 
    for(var i=0;i<allAbons.length;i++) {
      //прохожусь по всем занятиям (исходя из кол-ва занятий)
      var q = allAbons[i][6]// кол-во занятий в абонементе
      for(var j=0;j<q;j++) {
        if(allAbons[i][11+j*5] == teacherName){
          var lessonDate = allAbons[i][8+j*5];
          //ищем стоимость занятия по этому абону
          var sumPayed = allAbons[i][5];
          var lessonsQ = allAbons[i][6];
          var oneLessonPrice = sumPayed/lessonsQ;
          var typeAbon = allAbons[i][3];
          var nachPrcnt = 0;
          var nachPrcntSgorel = 0;
          for(var a=0;a<nachislenia.length;a++) {
            if(typeAbon == nachislenia[a][0]) {
              nachPrcnt = nachislenia[a][1];
              nachPrcntSgorel = nachislenia[a][2];
              break;
            }
          }
          
          // проверяю что это вообще дата
          if(lessonDate instanceof Date) {
            var lessonMonthYear = ""+(lessonDate.getMonth()+1)+""+lessonDate.getFullYear();
            //проверяю подходит ли этот урок в этот месяц/год
            if(lessonMonthYear == monthYear) {
              //добавляем кол-во определенного статуса (сверху в выписке будет)
              if(allAbons[i][12+j*5] == symbolPlus) {
                var onPeayed = oneLessonPrice*nachPrcnt;
                plusQ++;
              } else if(allAbons[i][12+j*5] == symbolMinus) {
                var onPeayed = oneLessonPrice*nachPrcntSgorel; //другой процент, если клиент не пришел
                minusQ++
              } else {
                var onPeayed = 0; // ничего если не внес данные
                noDataQ++;
              }
              res.push(['','',lessonDate,allAbons[i][1],allAbons[i][0],j+1,allAbons[i][12+j*5],onPeayed,""])
              //            num++
            }
          }
        }
      }
    }
    
    //    stop
    res.sort(function(a, b) {
      return a[2] - b[2];
    });
    
    // раздаем цвета строкам
    var resBgr = [];
    for(var r=0;r<res.length;r++){
      if(res[r][6] == symbolPlus) {
        resBgr.push([plusBgr,plusBgr,plusBgr,plusBgr,plusBgr,plusBgr,plusBgr,plusBgr,plusBgr]);
      } else if(res[r][6] == symbolMinus) {
        resBgr.push([minusBgr,minusBgr,minusBgr,minusBgr,minusBgr,minusBgr,minusBgr,minusBgr,minusBgr]);
      } else {
        resBgr.push([noDataBgr,noDataBgr,noDataBgr,noDataBgr,noDataBgr,noDataBgr,noDataBgr,noDataBgr,noDataBgr]);
      }
      res[r][1] = num+1;
      num++;
    }
//    stop
    if(sheetVipiska.getMaxRows()>4) {
//      sheetVipiska.getRange(4, 1, sheetVipiska.getMaxRows()-4, sheetVipiska.getLastColumn()).clearContent().clearFormat();    
      sheetVipiska.deleteRows(4, sheetVipiska.getMaxRows()-4)
    }
    
    if(res.length == 0){return}
    sheetVipiska.getRange(4, 1, res.length, res[0].length).setValues(res).setBackgrounds(resBgr);
    
//    sheetVipiska.appendRow(["","","","","","","","",""]);
    sheetVipiska.appendRow(["","","","Сумма","","=count(F4:F"+(sheetVipiska.getMaxRows())+")","","=sum(H4:H"+(sheetVipiska.getMaxRows())+")",""]);
    sheetVipiska.getRange(sheetVipiska.getLastRow(), 1, 1, sheetVipiska.getMaxColumns()).setBackgrounds([["#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"]]).setFontSize("12")
    
    sheetVipiska.getRange(1, 7).setValue(symbolPlus+"  - "+plusQ);
    sheetVipiska.getRange(1, 8).setValue(symbolMinus+"  - "+minusQ);
    sheetVipiska.getRange(1, 9).setValue("???"+" - "+noDataQ);
  } else {
    return
  }
  
}