// Cole este código em Extensões > Apps Script, na sua planilha do Google Sheets.
// Cada pessoa (convidado ou acompanhante) vira uma linha na planilha.

function limpa(valor, max) {
  var t = String(valor || "").trim().slice(0, max);
  // evita que o texto seja interpretado como fórmula
  if (/^[=+\-@]/.test(t)) t = "'" + t;
  return t;
}

function tipoPessoa(valor) {
  return valor === "Criança" ? "Criança" : "Adulto";
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var dados = JSON.parse(e.postData.contents);
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    if (aba.getLastRow() === 0) {
      aba.appendRow(["Data/Hora", "Nome", "Tipo", "Presença", "Acompanhante de", "Mensagem"]);
      aba.getRange("1:1").setFontWeight("bold");
    }

    var agora = new Date();
    var vai = dados.presenca === "Vou";
    var principal = limpa(dados.nome, 200);

    var linhas = [[agora, principal, tipoPessoa(dados.tipo), vai ? "Vou" : "Não vou", "", limpa(dados.mensagem, 1000)]];

    if (vai && Array.isArray(dados.acompanhantes)) {
      dados.acompanhantes.slice(0, 20).forEach(function (a) {
        var nome = limpa(a.nome, 200);
        if (nome) linhas.push([agora, nome, tipoPessoa(a.tipo), "Vou", principal, ""]);
      });
    }

    aba.getRange(aba.getLastRow() + 1, 1, linhas.length, 6).setValues(linhas);
    return ContentService.createTextOutput("ok");
  } catch (erro) {
    return ContentService.createTextOutput("erro");
  } finally {
    lock.releaseLock();
  }
}
