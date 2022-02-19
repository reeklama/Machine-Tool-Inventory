/** @OnlyCurrentDoc */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{name : "Отобразить",functionName : "OpenWind"},{name : "Перенести",functionName : "Trade"},{name : "СписатьОдну",functionName :   "MinusOne"}, {name : "Инвентаризация",functionName : "Invent"}, {name : "ТаблицаСписывания",functionName : "Tabl1"}, {name : "ТаблицаСписывания",functionName : "Tabl2"}, {name : "ТаблицаСписывания",functionName : "Tabl3"}, {name : "ТаблицаСписывания",functionName : "Tabl4"}, {name : "ТаблицаСписывания",functionName : "Tabl5"}, {name : "Переносщик",functionName : "Trader"}];
  sheet.addMenu("Скрипты", entries);
};

function OpenWind(){
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var range = sheet.getActiveRange(); //диапазон кнопки
    var open_close = range.getValues(); //открыть|закрыть
  
    var sheetS = sheet.getSheetByName("Склад"); 
    var rangeS = sheetS.getRange("B2:B500"); 
    var artS = rangeS.getValues();
  
    var tmp = 0;
  
    if(open_close[0] == "Закрыто"){   
      range.offset(0,0,1,1).setValue("Открыто");
      range.offset(-1,1,1,1).setValue("Наименование");
      range.offset(-1,2,1,1).setValue("Кол-во приход");
      range.offset(-1,3,1,1).setValue("Кол-во склад");
  
    for (var i=0;i <= range.getHeight();i++){
        tmp = range.offset(i-1,-10,1,1).getValue();
        range.offset(i,1,1,1).setValue(range.offset(i,-9,1,1).getValue());
  
        if(range.offset(i,-8,1,1).getBackgroundColor() != range.offset(0,-9,1,1).getBackgroundColor()){
          range.offset(i,2,1,1).setValue('0');
        } else{
          range.offset(i,2,1,1).setValue(range.offset(i,-8,1,1).getValue());
        }
        
  
        for (var j=0; j <= rangeS.getHeight();j++){
          if (tmp == artS[j])
            range.offset(i-1,3,1,1).setValue(rangeS.offset(j,9,1,1).getValue());
        }
      }  
    } 
    else {
      range.offset(0,0,1,1).setValue("Закрыто");
      for (var i=range.getHeight();i > 0;i--){
        range.offset(-1,1,1,1).setValue(null);
        range.offset(-1,2,1,1).setValue(null);
        range.offset(-1,3,1,1).setValue(null);
      
        range.offset(0,1,i,1).setValue(null);
        range.offset(0,2,i,1).setValue(null);
        range.offset(0,3,i,1).setValue(null);
      }
    }
  };

  function Tabl5(){
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var range = sheet.getActiveRange();
    var command = range.getValues();
    if (command == 'Отменить'){
      clearTabl()
    } else {
  
      for(var k = 0; k < 12; k++){
        if(range.offset(k, -3, 1, 1).getBackgroundColor() == "#ffdd88"){
          var stanok = range.offset(k, -3, 1, 1).getValue();
          for(var l = 0; l < 4; l++){
            if(range.offset(l, -2, 1, 1).getBackgroundColor() == "#ffdd88"){
              stanok = stanok + " " + range.offset(l, -2, 1, 1).getValue();
              for(var p = 0; p < 4; p++){
                if(range.offset(p, -1, 1, 1).getBackgroundColor() == "#ffdd88"){
                  stanok = stanok + " " + range.offset(p, -1, 1, 1).getValue();
                }
              }
            }
          }
        }
      }
  
      var sheetSpis = sheet.getSheetByName(stanok); 
      var rangeSpis = sheetSpis.getRange("A6:A200"); 
      var idSpis = rangeSpis.getValues();
      var rangeCount = sheetSpis.getRange("H6:H200");
      var idCount = rangeCount.getValues();
  
      for(var i=0;i < 200;i++){
        if(idSpis[i] != null){
          range.offset(i, -8, 1, 1).setValue(range.offset(i, -8, 1, 1).getValue() - (idCount[i] * command));
        }
      }
      clearTabl()
    }
  };
  
  function Tabl4(){
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var range = sheet.getActiveRange();
    var count = range.getValues();
    range.offset(0,0,1,1).setBackgroundColor("#FFDD88"); 
  
    sheet.getRange('S2').setValue('');
    sheet.getRange('S3').setValue('Отменить');
  };
  
  function Tabl3(){
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var range = sheet.getActiveRange();
    var count = range.getValues();
    range.offset(0,0,1,1).setBackgroundColor("#FFDD88"); 
  
    sheet.getRange('R2').setValue('0');
    sheet.getRange('R3').setValue('1');
  };
  
  function Tabl2(){
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var range = sheet.getActiveRange();
    var count = range.getValues();
    range.offset(0,0,1,1).setBackgroundColor("#FFDD88"); 
  
    sheet.getRange('Q2').setValue('0');
    sheet.getRange('Q3').setValue('2,5');
    sheet.getRange('Q4').setValue('5,5');
    sheet.getRange('Q5').setValue('15');
  };
  
  function Tabl1(){
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    if(sheet.getRange('P2').getValue() == sheet.getRange('O2').getValue()){
      sheet.getRange('P2').setValue('Start 1510');
      sheet.getRange('P3').setValue('Start 1510 PRO');
      sheet.getRange('P4').setValue('Start 1510 SRT');
      sheet.getRange('P5').setValue('Start 1510 SXRT');
      sheet.getRange('P6').setValue('Start 1200');
      sheet.getRange('P7').setValue('Start 1200 PRO');
      sheet.getRange('P8').setValue('Start 1200 SRT');
      sheet.getRange('P9').setValue('Start 1200 SXRT');
      sheet.getRange('P10').setValue('Start 1000');
      sheet.getRange('P11').setValue('Start 1000 PRO');
      sheet.getRange('P12').setValue('Start 1000 SRT');
      sheet.getRange('P13').setValue('Start 1000 SXRT');
    } else {
      clearTabl();
    }
  
  
  };

function clearTabl(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  sheet.getRange('P2').setValue(null);
  sheet.getRange('P3').setValue(null);
  sheet.getRange('P4').setValue(null);
  sheet.getRange('P5').setValue(null);
  sheet.getRange('P6').setValue(null);
  sheet.getRange('P7').setValue(null);
  sheet.getRange('P8').setValue(null);
  sheet.getRange('P9').setValue(null);
  sheet.getRange('P10').setValue(null);
  sheet.getRange('P11').setValue(null);
  sheet.getRange('P12').setValue(null);
  sheet.getRange('P13').setValue(null);
  sheet.getRange('Q2').setValue(null);
  sheet.getRange('Q3').setValue(null);
  sheet.getRange('Q4').setValue(null);
  sheet.getRange('Q5').setValue(null);
  sheet.getRange('R2').setValue(null);
  sheet.getRange('R3').setValue(null);
  sheet.getRange('S2').setValue(null);
  sheet.getRange('S3').setValue(null);

  sheet.getRange('P2').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('P3').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('P4').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('P5').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('P6').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('P7').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('P8').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('P9').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('P10').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('P11').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('P12').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('P13').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('Q2').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('Q3').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('Q4').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('Q5').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('R2').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('R3').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('S2').setBackgroundColor("#FFFFFF"); 
  sheet.getRange('S3').setBackgroundColor("#FFFFFF"); 
}


function Invent(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var range = sheet.getActiveRange();
  var count = range.getValues();
  var sklad = range.offset(0,-3,1,1).getValue()

  range.offset(0,0,1,1).setValue((sklad - count) + ' потери');
  range.offset(0,-3,1,1).setValue(count);
};


function MinusOne(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var range = sheet.getActiveRange();
  var count = range.getValues();

  range.offset(0,-2,1,1).setValue(range.offset(0,-2,1,1).getValue() - count);
  
};

function Trade(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var range = sheet.getActiveRange();

  var sheetS = sheet.getSheetByName("Склад"); 
  var rangeS = sheetS.getRange("B2:B1000"); 
  var artS = rangeS.getValues();
 
  var id = range.offset(0,-11,1,1).getValue();
  var count_val = range.offset(0,1,1,1).getValue();

  for (var j=rangeS.getHeight();j > -1;j--){
    if (id == artS[j]){
      range.offset(0,1,1,1).setValue('0');
      range.offset(0,2,1,1).setValue(count_val + rangeS.offset(j,9,1,1).getValue());

      range.offset(0,-9,1,1).setBackgroundColor("#FFDD88");  

      rangeS.offset(0,9,1,1).setValue(range.offset(0,2,1,1).getValue());
    }  
  }
};

