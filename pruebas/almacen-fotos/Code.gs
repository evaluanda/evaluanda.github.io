/**
 * PROTOTIPO del almacén local de fotos. Proyecto Apps Script APARTE (no es la
 * app): pega este archivo e Index.html en un proyecto nuevo y despliégalo como
 * aplicación web (Ejecutar como: tú; Acceso: solo tú o tu dominio).
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Prueba almacén de fotos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
