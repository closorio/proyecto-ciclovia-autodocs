/**Keys for acces */

/**------------------------------Forms----------------------------------*/
/**Instancia de variables */
// ID del archivo formato-subido-prueba
const ID_ARCHIVO_ORIGEN = "1mFleBJdacDPiBZsbsYOvXfwEY5Yqdd0TS8iVsChfdew";



/** Instancias de los archivos y hojas*/
// Abrir el archivo de origen usando su ID
const sourceSpreadsheet = SpreadsheetApp.openById(ID_ARCHIVO_ORIGEN);
// Obtener la hoja de cálculo activa del archivo de origen
const sourceSheet = sourceSpreadsheet.getActiveSheet();


