/*Llamar información dependiendo de un dato */
function obtenerInfo() {

    var respuesta = Browser.msgBox('Confirmación', '¿Estás seguro de ejecutar la función Buscar OT?', Browser.Buttons.YES_NO);
    // Si el usuario selecciona "Sí", ejecuta la función
    if (respuesta === 'yes') {
  
    // Obtén la hoja de cálculo activa y la celda que contiene el dato
    var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdaDato = hojaActiva.getRange("C5");
    
    // Obtén el valor de la celda
    var dato = celdaDato.getValue();
    
    // Abre el archivo externo
    var archivoExterno = SpreadsheetApp.openById("1kyxEIVA5bB9COOkP0SDIHIot5LIeWMdgFZhHFLHKByo");
    
    // Obtén la hoja de cálculo del archivo externo
    var hojaExterna = archivoExterno.getSheetByName("Seguimiento");
    
    // Busca la coincidencia en la hoja de destino (columna F)
    var data = hojaExterna.getDataRange().getValues();
    var filaCoincidencia = -1;

    for (var i = 0; i < data.length; i++) {
      if (data[i][5] === dato) {  // Columna F (índice 5)
        filaCoincidencia = i;
        break;
    }
    }

    //PRIMERA TABLA
    if (filaCoincidencia !== -1) {
    var informacion = hojaExterna.getRange("B" + (filaCoincidencia + 1)).getValue();
    hojaActiva.getRange("C8").setValue(informacion);
    }
    if (filaCoincidencia !== -1) {
    var informacion = hojaExterna.getRange("G" + (filaCoincidencia + 1)).getValue();
    hojaActiva.getRange("C9").setValue(informacion);
    }  
    if (filaCoincidencia !== -1) {
    var informacion = hojaExterna.getRange("I" + (filaCoincidencia + 1)).getValue();
    hojaActiva.getRange("C10").setValue(informacion);  
    }
    if (filaCoincidencia !== -1) {
    var informacion = hojaExterna.getRange("U" + (filaCoincidencia + 1)).getValue();
    hojaActiva.getRange("C11").setValue(informacion);  
    }
    if (filaCoincidencia !== -1) {
    var informacion = hojaExterna.getRange("V" + (filaCoincidencia + 1)).getValue();
    hojaActiva.getRange("C12").setValue(informacion);  
    }
    if (filaCoincidencia !== -1) {
    var informacion = hojaExterna.getRange("AQ" + (filaCoincidencia + 1)).getValue();
    hojaActiva.getRange("C13").setValue(informacion);  
    }
    if (filaCoincidencia !== -1) {
    var informacion = hojaExterna.getRange("P" + (filaCoincidencia + 1)).getValue();
    hojaActiva.getRange("C14").setValue(informacion);  
    }
    if (filaCoincidencia !== -1) {
    var informacion = hojaExterna.getRange("S" + (filaCoincidencia + 1)).getValue();
    hojaActiva.getRange("C15").setValue(informacion);  
    }
    
} else {
      // Si el usuario selecciona "No" o cierra el cuadro de diálogo, no se ejecuta la función.
    }
}