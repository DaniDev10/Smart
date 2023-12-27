/*Carga de datos, validación de existencia y confirmación de querer ejecutar el código*/
function zonaUno() {
    var respuesta = Browser.msgBox('Confirmación', '¿Estás seguro de ejecutar la función Insertar OT?', Browser.Buttons.YES_NO);
    // Si el usuario selecciona "Sí", ejecuta la función
    if (respuesta === 'yes') {

    var ss = SpreadsheetApp.getActive();
    var hoja = ss.getSheetByName("Carga de TK nuevos");
    var data = hoja.getRange("F11").getValue(); // SITIO
    var data2 = hoja.getRange("F13").getValue(); // NUM. TICKET
    var data3 = hoja.getRange("F15").getValue(); // SERVICE DESK
    var data4 = hoja.getRange("F17").getValue(); // FECHA ASIGNACIÓN
    var data5 = hoja.getRange("F19").getValue();// DESCRIPCIÓN 
    var data6 = hoja.getRange("F21").getValue();// TIPO DE ALARMA
    var data7 = hoja.getRange("F23").getValue();// ESPECIALIDAD
    var data8 = hoja.getRange("F9").getValue()// TIPO DE MTTO 
    var data9 = hoja.getRange("F25").getValue()// CANT MTTO

    if (data === "" || data2 === "" || data3 === "" || data4 === "" || data8 === "") {
      // Alguna de las celdas está vacía, mostrar mensaje y salir de la función sin hacer nada adicional
    var motivo = "";
    if (data === "") {
        motivo += "El sitio está vacío. ";
    }
    if (data2 === "") {
        motivo += "El numero de ticket está vacío. ";
    }
    if (data3 === "") {
        motivo += "service Desk está vacío. ";
        }
    if (data4 === "") {
        motivo += "La fecha de asignación está vacío. ";
        }
    if (data8 === "") {
        motivo += "Tipo de mantenimiento está vacío. ";
        }

        Browser.msgBox("El código no se ejecutó debido a los siguientes motivos: " + motivo);
    return;
    }


    var ssdest =  SpreadsheetApp.openById("1kyxEIVA5bB9COOkP0SDIHIot5LIeWMdgFZhHFLHKByo")
    var hojadest = ssdest.getSheetByName("Seguimiento");
    var destinationData = hojadest.getRange("F:F").getValues(); 
    // VERIFICAR EXISTENCIA
    for (var i = 0; i < destinationData.length; i++) {
    if (destinationData[i][0] === data2) {
        Browser.msgBox("El ticket ya existe en la hoja de destino.");
        return;
    }
    }

    var lastRow = hojadest.getLastRow();
    var targetRow = lastRow + 1;

    Logger.log("Data: " + data);
    Logger.log("Data2: " + data2);
    Logger.log("Data3: " + data3);
    Logger.log("Data4: " + data4);
    Logger.log("Data5: " + data5);
    Logger.log("Data6: " + data6);
    Logger.log("Data7: " + data7);
    Logger.log("Target Row: " + targetRow);

    hojadest.getRange(targetRow, 7).setValue(data);
    hojadest.getRange(targetRow, 6).setValue(data2);
    hojadest.getRange(targetRow, 21).setValue(data3);
    hojadest.getRange(targetRow, 18).setValue(data4);
    hojadest.getRange(targetRow, 12).setValue(data5);
    hojadest.getRange(targetRow, 13).setValue(data6);
    hojadest.getRange(targetRow, 16).setValue(data7);
    hojadest.getRange(targetRow, 9).setValue(data8);
    hojadest.getRange(targetRow, 19).setValue(data9);

    hoja.getRange("F9").clearContent();
    hoja.getRange("F11").clearContent();
    hoja.getRange("F13").clearContent();
    //hoja.getRange("F15").clearContent();
    hoja.getRange("F17").clearContent();
    hoja.getRange("F19").clearContent();
    hoja.getRange("F21").clearContent();
    hoja.getRange("F23").clearContent();
    hoja.getRange("E17").clearContent();
    hoja.getRange("F25").clearContent();
} else {
      // Si el usuario selecciona "No" o cierra el cuadro de diálogo, no se ejecuta la función.
    }
}