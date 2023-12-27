/*Ejecutor*/
function onOpen() {
    SpreadsheetApp.getUi()
    .createMenu('Zi-Track')
    .addItem('NOMBRE VISUAL', 'NOMBRE FUNCIóN')
    .addSeparator()
    .addItem('NOMBRE VISUAL', 'NOMBRE FUNCIóN')
    .addToUi();
    Logger.log('Menu Uploaded');
}