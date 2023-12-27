function copiarRequisicion() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getActiveSheet();
    var dataSheet = ss.getSheetByName("Data"); // Nombre de la pestaña "Data"

    // Obtener el nombre de la celda D8
    var nombre = hoja.getRange("D8").getValue();

    // Buscar el correo correspondiente al nombre en la hoja "Data"
    var dataRange = dataSheet.getRange("E:F");
    var dataValues = dataRange.getValues();
    var correo = "";
    for (var i = 0; i < dataValues.length; i++) {
    if (dataValues[i][0] === nombre) {
        correo = dataValues[i][1];
        break;
    }
    }

    // Obtener los demás datos necesarios
    var data1 = hoja.getRange("D7").getValue();
    var data2 = hoja.getRange("D8").getValue();
    var data3 = hoja.getRange("K2").getValue();
    var data4 = hoja.getRange("K4").getValue();
    var data5 = hoja.getRange("G7").getValue();
    var data6 = hoja.getRange("K6").getValue();
    var data7 = hoja.getRange("K7").getValue();
    var data8 = hoja.getRange("I11").getValue();

    // Obtener la carpeta de destino en Google Drive por ID
    var carpetaDestinoId = "15GE5dI01CqOESfurKeW9BAs9EmSxeiGZ";
    var carpetaDestino = DriveApp.getFolderById(carpetaDestinoId);

    // Verificar si las celdas están vacías antes de continuar
    if (data1 === "" || data2 === "" || data3 === "" || data4 === "" || data5 === "" || data6 === "" || data7 === "" || data8 ==="") {
      // Alguna de las celdas está vacía, mostrar mensaje y salir de la función sin hacer nada adicional
    var motivo = "";
    if (data1 === "") {
        motivo += "El solicitante está vacío. ";
    }
    if (data2 === "") {
        motivo += "El aprobador está vacío. ";
    }
    if (data3 === "") {
        motivo += "Fecha de solicitud está vacío. ";
    }
    if (data4 === "") {
        motivo += "Fecha de entrega está vacío. ";
    }
    if (data5 === "") {
        motivo += "Centro de costos está vacío. ";
    }
    if (data6 === "") {
        motivo += "Dirección de entrega está vacío. ";
    }
    if (data7 === "") {
        motivo += "Ciudad está vacío. ";
    }
    if (data8 === "") {
        motivo += "Columna cantidad no tiene datos. ";
    }
    Browser.msgBox("El código no se ejecutó debido a los siguientes motivos: " + motivo);
    return;
    }

    // Obtener el valor actual de la celda I3
    var numeroCeldaActual = hoja.getRange("I3").getValue();

    // Crear el nombre del archivo con el número de la celda I3
    var nombreArchivo = numeroCeldaActual +"_" +"REQUISICION" + "_" + data5;

    // Crear una copia del archivo en la carpeta de destino sin fórmulas
    var archivoActivo = DriveApp.getFileById(ss.getId());
    var copiaArchivo = archivoActivo.makeCopy(nombreArchivo, carpetaDestino);
    var copiaArchivoId = copiaArchivo.getId();
    var copiaArchivoSinFormulas = SpreadsheetApp.openById(copiaArchivoId);
    var hojaCopia = copiaArchivoSinFormulas.getActiveSheet();

    // Obtener el rango de origen con fórmulas
    var rangoFormulas = hoja.getDataRange();

    // Obtener los valores del rango de origen
    var valores = rangoFormulas.getValues();

    // Obtener el rango de destino
    var rangoDestino = hojaCopia.getRange(1, 1, valores.length, valores[0].length);

    // Pegar los valores en el rango de destino sin fórmulas
    rangoDestino.setValues(valores);

    // Insertar enlace y valores en la siguiente fila
    var archivoDestinoId = "1GE2jnfZh80eTEdrNxNfypGYAMkeKmD6HywXZWaEv8OY";
    var archivoDestino = SpreadsheetApp.openById(archivoDestinoId);
    var hojaDestino = archivoDestino.getActiveSheet();
    var ultimaFila = hojaDestino.getLastRow() + 1;
    var celdaEnlace = hojaDestino.getRange(ultimaFila, 9); // Columna I
    var celdaValor = hojaDestino.getRange(ultimaFila, 2); // Columna B
    var celdaCostos = hojaDestino.getRange(ultimaFila, 4); // Columna D
    var celdaSolicitante = hojaDestino.getRange(ultimaFila, 6); // Columna E
    var celdaAprobador = hojaDestino.getRange(ultimaFila, 7); // Columna F
    var fechaSoli = hojaDestino.getRange(ultimaFila, 11); // Columna K
    var fechaEntre = hojaDestino.getRange(ultimaFila, 12); // Columna L
    celdaEnlace.setValue(copiaArchivo.getUrl());
    celdaValor.setValue(numeroCeldaActual);
    celdaCostos.setValue(data5);
    celdaSolicitante.setValue(data1);
    celdaAprobador.setValue(data2);
    fechaSoli.setValue(data3);
    fechaEntre.setValue(data4);

    // Nuevo destino
  var ssdest = SpreadsheetApp.openById("17vXASKKh3797ztRZKgeluP6fGUmic6Xbev8MhcBD8CI"); // Hoja de destino
var hojadest2 = ssdest.getSheetByName("BD");

var destRow = hojadest2.getLastRow() - 1;

var dataRange = hoja.getRange("C11:O" + hoja.getLastRow());
var dataValues = dataRange.getValues();

var newData = dataValues.filter(function(row) {
    return row[0] !== "";
});

hojadest2.getRange(destRow + 2, 2, newData.length, newData[0].length).setValues(newData);

  // Desplazar I3
var datoDesplazado = hoja.getRange("I3").getValue();
var columnaA = hojadest2.getRange(destRow + 2, 1, newData.length, 1);
var valoresColumnaA = columnaA.getValues();

for (var i = 0; i < valoresColumnaA.length; i++) {
    if (valoresColumnaA[i][0] === "") {
    valoresColumnaA[i][0] = datoDesplazado;
    }
}

columnaA.setValues(valoresColumnaA);

  // Desplazar D7
var datoDesplazadoD = hoja.getRange("D7").getValue();
var columnaD = hojadest2.getRange(destRow + 2, 15, newData.length, 1);
var valoresColumnaD = columnaD.getValues();

for (var i = 0; i < valoresColumnaD.length; i++) {
    if (valoresColumnaD[i][0] === "") {
    valoresColumnaD[i][0] = datoDesplazadoD;
    }
}

columnaD.setValues(valoresColumnaD);

  // Desplazar D8
var datoDesplazadoD2 = hoja.getRange("D8").getValue();
var columnaD2 = hojadest2.getRange(destRow + 2, 16, newData.length, 1);
var valoresColumnaD2 = columnaD2.getValues();

for (var i = 0; i < valoresColumnaD2.length; i++) {
    if (valoresColumnaD2[i][0] === "") {
    valoresColumnaD2[i][0] = datoDesplazadoD2;
    }
}

columnaD2.setValues(valoresColumnaD2);

  // Desplazar G7
var datoDesplazadoG7 = hoja.getRange("G7").getValue();
var columnaG7 = hojadest2.getRange(destRow + 2, 17, newData.length, 1);
var valoresColumnaG7 = columnaG7.getValues();

for (var i = 0; i < valoresColumnaG7.length; i++) {
    if (valoresColumnaG7[i][0] === "") {
    valoresColumnaG7[i][0] = datoDesplazadoG7;
    }
}

columnaG7.setValues(valoresColumnaG7);



    // Enviar correo al destinatario encontrado
    var asunto = "REQUISICIÓN " + "#"+ numeroCeldaActual + " " +  data5;
    var mensaje = "Estimado(a), " + data2 + "\n\n 1.Copia el numero de solicitud que se encuentra en el asunto del correo" + "\n 2. Ingresa al siguiente link para aprobar la solicitud. Debes ingresar el numero copiado en la celda D2 para traer la información " +  "\nhttps://docs.google.com/spreadsheets/d/17vXASKKh3797ztRZKgeluP6fGUmic6Xbev8MhcBD8CI/edit#gid=629074020" + "\n 3. Dar clic en el botón de check para aprobar segun presupuesto"  + 
"\n 4. Dirigirse al botón de ZITRACK para enviar la aprobación" + "\n\n Atentamente " +  data1;

    MailApp.sendEmail({
    to: correo,
    subject: asunto,
    body: mensaje
    });

    // Limpiar los valores en la hoja de origen
    hoja.getRange("D7").clearContent();
    hoja.getRange("D8").clearContent();
    //hoja.getRange("K2").clearContent();
    hoja.getRange("K4").clearContent();
    hoja.getRange("G7").clearContent();
    hoja.getRange("K6").clearContent();
    hoja.getRange("K7").clearContent();
    //hoja.getRange("H11:H").clearContent();
    hoja.getRange("C11:C").clearContent();
    hoja.getRange("D11:D").clearContent();
    hoja.getRange("I11:I").clearContent();
    hoja.getRange("O11:O").clearContent();
}