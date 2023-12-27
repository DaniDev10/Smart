/*Llamar por columna de una hoja a la otra*/
function getSitios(){
    var ss = SpreadsheetApp.openById("ID ORIGEN");
    var hoja = ss.getSheetByName("HOJA ORIGEN");
    var data = hoja.getRange("B5:B").getValues(); 
    var data2 = hoja.getRange("C5:C").getValues();
    
    var ssdest = SpreadsheetApp.getActive();
    var hojadest = ssdest.getSheetByName("HOJA DESTINO");
    hojadest.getRange(2, 1,data.length,data[0].length).setValues(data);
    hojadest.getRange(2, 2,data2.length,data2[0].length).setValues(data2);
    
}