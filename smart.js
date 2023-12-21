/*Llamar por columna de una hoja a la otra*/
function getSitios(){
    var ss = SpreadsheetApp.openById("11gPnTx_lN69A44W8rLWUrzmZBa68IMksL7cy2XFFNt0");
    var hoja = ss.getSheetByName("Informacion sitios");
    var data = hoja.getRange("B5:B").getValues(); 
    var data2 = hoja.getRange("C5:C").getValues();
    
    var ssdest = SpreadsheetApp.getActive();
    var hojadest = ssdest.getSheetByName("SITIOS SMU");
    hojadest.getRange(2, 1,data.length,data[0].length).setValues(data);
    hojadest.getRange(2, 2,data2.length,data2[0].length).setValues(data2);
    
}


/*Llamar a toda la hoja*/
function getFac(){
    var ss = SpreadsheetApp.openById("1wYuvJAibQANWf5hj6lvserKOivkTcf8S944gkiVhSTA");
    var hoja = ss.getSheetByName("MATRIZ LPU");
    var data = hoja.getDataRange().getDisplayValues();
    var ssdest = SpreadsheetApp.getActive();
    var hojadest = ssdest.getSheetByName("LPU SMU");
    hojadest.getRange(1, 1,data.length,13).setValues(data);
}