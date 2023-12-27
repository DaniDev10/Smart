/*Llamar a toda la hoja*/
function getFac(){
    var ss = SpreadsheetApp.openById("ID ORIGEN");
    var hoja = ss.getSheetByName("HOJA ORIGEN");
    var data = hoja.getDataRange().getDisplayValues();
    var ssdest = SpreadsheetApp.getActive();
    var hojadest = ssdest.getSheetByName("HOJA DESTINO");
    hojadest.getRange(1, 1,data.length,13).setValues(data);
}