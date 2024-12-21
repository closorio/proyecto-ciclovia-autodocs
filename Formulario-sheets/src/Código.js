// Función para agregar el menú personalizado
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Autofill Form Ciclovida')
    .addSubMenu(ui.createMenu('Cronograma de Actividades')
      .addItem('Cronograma Inicial', 'ejecutarCronogramaIncial')
      .addItem('Cronograma Final', 'ejecutarCronogramaFinal')
    )
    .addSeparator()
    .addItem('Limpiar Tabla', 'limpiarTabla')
    .addToUi();
}

// Funciones para las opciones del menú
function ejecutarCronogramaIncial() {
  AutoCronograma.main(ID_TEMPLATE_INICIAL, ID_FOLDER_DESTINO_CA_INICIAL, "Inicial");
}

function ejecutarCronogramaFinal() {
  AutoCronograma.main(ID_TEMPLATE_FINAL, ID_FOLDER_DESTINO_CA_FINAL, "Final");
}


// Función para limpiar la tabla cuando se cierra el documento
function limpiarTabla() {
  // Obtener la hoja de cálculo activa
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Obtener la hoja donde se encuentra la tabla 
  var sheet = spreadsheet.getSheetByName('Hoja 1');
  var range = sheet.getRange('A2:B49'); // Ajusta el rango

  // Limpiar el rango
  range.clearContent();
}