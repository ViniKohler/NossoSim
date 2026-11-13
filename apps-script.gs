// Cole este código em Extensões > Apps Script, na sua planilha do Google Sheets.

function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    // Cria o cabeçalho na primeira vez
    if (aba.getLastRow() === 0) {
      aba.appendRow(["Data/Hora", "Nome", "Presença", "Mensagem"]);
      aba.getRange("1:1").setFontWeight("bold");
    }

    aba.appendRow([
      new Date(),
      String(dados.nome || "").slice(0, 200),
      dados.presenca === "Vou" ? "Vou" : "Não vou",
      String(dados.mensagem || "").slice(0, 1000)
    ]);

    return ContentService.createTextOutput("ok");
  } catch (erro) {
    return ContentService.createTextOutput("erro");
  }
}
