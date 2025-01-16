function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function incluirNaPlanilha(comprador, fornecedor, origem, emissao, status, prioridade, observacao, motivo, notaFiscal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBoloro = ss.getSheetByName('BOLORO');
  const sheetEntrega = ss.getSheetByName('LANCAMENTO');

  if (!sheetBoloro || !sheetEntrega) {
    Logger.log('Erro: Uma ou ambas as planilhas não foram encontradas.');
    return 'Erro: Uma ou ambas as planilhas não foram encontradas.';
  }

  if (motivo == "Atraso entrega") {
    let lastRow = sheetEntrega.getLastRow() + 1;
    sheetEntrega.getRange("A" + lastRow).setValue(notaFiscal);
    sheetEntrega.getRange("B" + lastRow).setValue(origem);
    sheetEntrega.getRange("C" + lastRow).setValue(prioridade);
    sheetEntrega.getRange("D" + lastRow).setValue(fornecedor);
    sheetEntrega.getRange("E" + lastRow).setValue(emissao);
    sheetEntrega.getRange("F" + lastRow).setValue(status);
    sheetEntrega.getRange("G" + lastRow).setValue(comprador);
    sheetEntrega.getRange("H" + lastRow).setValue(observacao);
    Logger.log('Dados cadastrados na planilha NOTASATRASADA.');
  } else {
    let lastRow = sheetBoloro.getLastRow() + 1;
    sheetBoloro.getRange("A" + lastRow).setValue(notaFiscal);
    sheetBoloro.getRange("B" + lastRow).setValue(origem);
    sheetBoloro.getRange("C" + lastRow).setValue(prioridade);
    sheetBoloro.getRange("D" + lastRow).setValue(fornecedor);
    sheetBoloro.getRange("E" + lastRow).setValue(emissao);
    sheetBoloro.getRange("F" + lastRow).setValue(status);
    sheetBoloro.getRange("G" + lastRow).setValue(comprador);
    sheetBoloro.getRange("H" + lastRow).setValue(observacao);
    Logger.log('Dados cadastrados na planilha NOTASPARACORRIGIR.');
  }
}
