/**
 * Plik: Code.gs
 * Główny punkt wejścia aplikacji.
 * Odpowiada tylko za serwowanie strony HTML.
 * Cała logika biznesowa (getInitialData, addTask itp.) znajduje się
 * w plikach DataRead.gs i DataWrite.gs.
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Centrum Dowodzenia Magazynem')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}