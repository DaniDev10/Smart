/*Realizar una carga masiva*/
function getBaster() {
    // Origen
    var ss = SpreadsheetApp.openById("ID ORIGEN");
    var hoja = ss.getSheetByName("HOJA ORIGEN");

    // Sequimiento Cotización
var ssdest2 = SpreadsheetApp.openById("ID DESTINO");
var hojadest2 = ssdest2.getSheetByName("HOJA DESTINO");

var startRow= hojadest2.getLastRow() + 0;
    //Cotización
    var data = hoja.getRange("A2:A" + hoja.getLastRow());//OT
    var dataValues1 = data.getValues();
    var data2 = hoja.getRange("E2:E" + hoja.getLastRow());//EB
    var dataValues2 = data2.getValues();
    var data3 = hoja.getRange("D2:D" + hoja.getLastRow());//TIPO MTTO
    var dataValues3 = data3.getValues();
    var data4 = hoja.getRange("B2:B" + hoja.getLastRow()); //PROYECTADA
    var dataValues4 = data4.getValues();
    var data5 = hoja.getRange("Z2:Z" + hoja.getLastRow());//DESCRIPCIÓN
    var dataValues5 = data5.getValues();
    var data6 = hoja.getRange("C2:C" + hoja.getLastRow());//F.ASIGNACIÓN
    var dataValues6 = data6.getValues();
    
    var newData = dataValues1.filter(function(row){
    return row[0] !=="";
    });
    var newData2 = dataValues2.filter(function(row){
    return row[0] !=="";
    });

    var newData3 = dataValues3.filter(function(row){
    return row[0] !=="";
    });

    var newData4 = dataValues4.filter(function(row){
    return row[0] !=="";
    });
    var newData5 = dataValues5.filter(function(row){
    return row[0] !=="";
    });
    var newData6 = dataValues6.filter(function(row){
    return row[0] !=="";
    });
    
if (newData.length > 0 && newData[0].length > 0) {
    hojadest2.getRange(startRow + 1, 6, newData.length, newData[0].length).setValues(newData);//OT
}
if (newData2.length > 0 && newData2[0].length > 0) {
    hojadest2.getRange(startRow + 1, 7, newData2.length, newData2[0].length).setValues(newData2);//EB
}
if (newData3.length > 0 && newData3[0].length > 0) {
    hojadest2.getRange(startRow + 1, 9, newData3.length, newData3[0].length).setValues(newData3);//TIPO MTTO
}
if (newData4.length > 0 && newData4[0].length > 0) {
    hojadest2.getRange(startRow + 1, 25, newData4.length, newData4[0].length).setValues(newData4);//PROYECTADA
}
if (newData5.length > 0 && newData5[0].length > 0) {
    hojadest2.getRange(startRow + 1, 12, newData5.length, newData5[0].length).setValues(newData5);//DESCRIPCIÓN
}
if (newData6.length > 0 && newData6[0].length > 0) {
    hojadest2.getRange(startRow + 1, 18, newData6.length, newData6[0].length).setValues(newData6);//F.ASIGNACIÓN
}
    // Limpiar hoja
    hoja.getRange("A2:F").clearContent();
}