function extrairRelatorioLimpo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPainel = ss.getSheetByName("Painel de Trocas");
  const ui = SpreadsheetApp.getUi();
  
  if (!abaPainel) {
    ui.alert("Erro: A aba 'Painel de Trocas' não foi encontrada.");
    return;
  }
  
  // Cria uma janela HTML com os botões personalizados
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family: Arial; text-align: center; padding: 20px;">' +
    '<h3>Escolha o formato de extração:</h3>' +
    '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).processarExtracao(\'PDF\')" style="padding: 10px 20px; margin: 5px; cursor: pointer;">PDF</button>' +
    '<button onclick="google.script.run.withSuccessHandler(google.script.host.close).processarExtracao(\'CSV\')" style="padding: 10px 20px; margin: 5px; cursor: pointer;">CSV</button>' +
    '</div>'
  ).setWidth(300).setHeight(150);
  
  ui.showModalDialog(html, "Formato do Relatório");
}

// Esta função processa o que foi escolhido no HTML
function processarExtracao(formato) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPainel = ss.getSheetByName("Painel de Trocas");
  const mes = abaPainel.getRange("C3").getValue().toString().trim() || "Mes";
  const ano = abaPainel.getRange("C4").getValue().toString().trim() || "Ano";
  
  try {
    const nomeTemp = "Temp_Relatorio_" + mes + "_" + ano;
    const novoArquivo = SpreadsheetApp.create(nomeTemp);
    const abaCopiada = abaPainel.copyTo(novoArquivo);
    abaCopiada.setName("Relatorio");
    novoArquivo.deleteSheet(novoArquivo.getSheets()[0]);
    
    // Remove botões e filtros
    abaCopiada.getDrawings().forEach(d => d.remove());
    const filtro = abaCopiada.getFilter();
    if (filtro) filtro.remove();
    abaCopiada.getRange("B3:C4").clearContent();
    
    let arquivoFinal;
    const token = ScriptApp.getOAuthToken();
    const idTemp = novoArquivo.getId();
    const gid = abaCopiada.getSheetId();
    
    if (formato === 'PDF') {
      const urlConfig = "https://docs.google.com/spreadsheets/d/" + idTemp + "/export?format=pdf&size=A4&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=true&gridlines=false&gid=" + gid;
      const response = UrlFetchApp.fetch(urlConfig, { headers: {'Authorization': 'Bearer ' + token} });
      arquivoFinal = DriveApp.createFile(response.getBlob().setName("Relatorio_" + mes + "_" + ano + ".pdf"));
    } else {
      const urlCSV = "https://docs.google.com/spreadsheets/d/" + idTemp + "/export?format=csv&gid=" + gid;
      const response = UrlFetchApp.fetch(urlCSV, { headers: {'Authorization': 'Bearer ' + token} });
      arquivoFinal = DriveApp.createFile(response.getBlob().setName("Relatorio_" + mes + "_" + ano + ".csv"));
    }
    
    DriveApp.getFileById(idTemp).setTrashed(true);
    mostrarLinkSucesso(arquivoFinal.getUrl(), arquivoFinal.getName(), formato);
    
  } catch (erro) {
    SpreadsheetApp.getUi().alert("Erro: " + erro.message);
  }
}

function mostrarLinkSucesso(url, nome, tipo) {
  const htmlOutput = HtmlService.createHtmlOutput(
    '<div style="font-family: Arial; text-align: center; padding: 20px;">' +
    '<h2 style="color: #4CAF50;">' + tipo + ' Gerado!</h2>' +
    '<p>Arquivo: <b>' + nome + '</b></p>' +
    '<a href="' + url + '" target="_blank" style="background-color: #4CAF50; color: white; padding: 10px; text-decoration: none; border-radius: 5px;">ABRIR ARQUIVO</a>' +
    '</div>'
  ).setWidth(300).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Sucesso");
}
