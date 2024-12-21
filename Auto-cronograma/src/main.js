// Función principal
function main(templateId, folderDestinoId, templateType) {
  try {
    const sourceSheet = getSourceSheet();
    const lastRow = sourceSheet.getLastRow();

    for (let i = 2; i <= lastRow; i++) {
      const rowData = getRowData(sourceSheet, i);
      const nombreCompleto = rowData[0];
      const direccion = rowData[1];

      if (!nombreCompleto) {
        console.log("El programa ha finalizado satisfactoriamente");
        return true;
      }

      const { comuna, mes } = getHeaderData(sourceSheet);
      const { observations } = getObservations(sourceSheet);

      const archivoCopia = createTemplateCopy(
        nombreCompleto,
        mes,
        templateId,
        folderDestinoId,
        templateType
      );

      const templateSpreadsheet = SpreadsheetApp.openById(archivoCopia.getId());
      processTemplate(templateSpreadsheet, sourceSheet, { nombreCompleto, direccion, comuna, mes, observations });
    }
  } catch (e) {
    Logger.log("Error en la función main: " + e.toString());
  }
}

// Procesa la plantilla con los datos necesarios
function processTemplate(templateSpreadsheet, sourceSheet, data) {
  copyDataToTemplate(sourceSheet, templateSpreadsheet);
  fillTemplate(templateSpreadsheet, data);
  fillObservations(templateSpreadsheet, data.observations);
}

// Obtiene la primera hoja de la hoja de cálculo de origen
function getSourceSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

// Obtiene los valores de una fila específica de la hoja de origen
function getRowData(sheet, row) {
  return sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
}

// Obtiene los valores de comuna y mes de la segunda fila de la hoja de origen
function getHeaderData(sheet) {
  return {
    comuna: sheet.getRange(2, 4).getValue() || "Comuna no especificada",
    mes: sheet.getRange(2, 6).getValue() || "ENERO",
  };
}

// Obtiene los valores de Observaciones de la hoja de origen
function getObservations(sheet) {
  return { observations: sheet.getRange(2, 14).getValue() || "" };
}

// Obtiene los valores de las semanas de la hoja de origen
function getSemanas(sheet) {
  const weekColumns = [8, 9, 10, 11, 12];
  return weekColumns.map((col) => sheet.getRange(2, col).getValue() || "");
}

// Crea una copia de la plantilla con el nombre específico
function createTemplateCopy(nombreCompleto, mes, templateId, folderDestinoId, templateType) {
  try {
    const fileTemplate = DriveApp.getFileById(templateId);
    const folderDestino = DriveApp.getFolderById(folderDestinoId);
    const nombreArchivo = getTemplateFileName(nombreCompleto, mes, templateType);
    return fileTemplate.makeCopy(nombreArchivo, folderDestino);
  } catch (e) {
    Logger.log("Error en createTemplateCopy: " + e.toString());
    throw e;
  }
}

// Genera el nombre del archivo basado en el tipo de plantilla
function getTemplateFileName(nombreCompleto, mes, templateType) {
  const fileTypeMap = {
    Inicial: "CRONOGRAMA E INFORME INICIAL",
    Final: "CRONOGRAMA E INFORME FINAL",
  };

  if (!fileTypeMap[templateType]) {
    throw new Error("Tipo de plantilla desconocido");
  }

  return `${nombreCompleto} - ${fileTypeMap[templateType]} - ${mes} - 2024.xlsx`;
}

// Llena los datos básicos en la plantilla
function fillTemplate(templateSpreadsheet, { nombreCompleto, direccion, comuna, mes }) {
  const templateSheet = templateSpreadsheet.getActiveSheet();
  templateSheet.getRange("M11").setValue(nombreCompleto);
  templateSheet.getRange("C15").setValue(direccion);
  templateSheet.getRange("D10").setValue(comuna);
  templateSheet.getRange("U10").setValue(mes);
}

// Llena las observaciones en la última hoja de la plantilla
function fillObservations(templateSpreadsheet, observations) {
  const lastSheet = templateSpreadsheet.getSheets().slice(-1)[0];
  lastSheet.getRange("C22").setValue(observations);
}

// Copia datos de semanas desde la hoja de origen a la plantilla
function copyDataToTemplate(sourceSheet, templateSpreadsheet) {
  const semanas = getSemanas(sourceSheet);
  const sheets = templateSpreadsheet.getSheets();

  sheets.slice(0, -1).forEach((sheet, index) => {
    if (semanas[index]) {
      sheet.getRange("C9").setValue(semanas[index]);
    }
  });
}
