function doGet() {
  return HtmlService.createHtmlOutputFromFile('formulario');
}

function incluirNaPlanilha(codigo, peso, lote, observacao) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBobinas = ss.getSheetByName('BOBINA');
  const sheetSlitter = ss.getSheetByName('SLITER');

  const codigosBobinas = ["532862", "532863", "532864", "532866", "532859",
    "532874", "084628", "535408", "536064", "535939", "535935", "535936", "532867",
    "535937", "535938", "536239", "536240", "536147", "536148", "536149", "536150",
    "536151", "536152", "536134", "532868", "532865", "532873", "532858", "532869",
    "532860", "532861", "533745", "536507", "533207", "532882", "536306",
    "532809", "532810", "532883", "532877", "532876", "532941", "534509",
    "533418", "535436", "535437", "083624", "536530",
    "535689", "543134", "543136", "533947", "533796"];

  const codigosSlitter = ["535636", "535637", "535638", "536698"];

  let data = new Date();

  let eBobina = codigosBobinas.includes(codigo);
  let eSlitter = codigosSlitter.includes(codigo);

  if (eBobina) {
    let lastRow = sheetBobinas.getLastRow() + 1;
    sheetBobinas.getRange("A" + lastRow).setValue(data);
    sheetBobinas.getRange("B" + lastRow).setValue(codigo);
    sheetBobinas.getRange("C" + lastRow).setValue(peso);
    sheetBobinas.getRange("D" + lastRow).setValue(lote);
    sheetBobinas.getRange("E" + lastRow).setValue(observacao);
    return "Peso da Bobina registrado com sucesso!";
  } else if (eSlitter) {
    let lastRow = sheetSlitter.getLastRow() + 1;
    sheetSlitter.getRange("A" + lastRow).setValue(data);
    sheetSlitter.getRange("B" + lastRow).setValue(codigo);
    sheetSlitter.getRange("C" + lastRow).setValue(peso);
    sheetSlitter.getRange("D" + lastRow).setValue(lote);
    sheetSlitter.getRange("E" + lastRow).setValue(observacao);
    return "Peso do Slitter registrado com sucesso!";
  } else {
    return "Código não encontrado, por favor verifique valores digitados.";
  }
}
