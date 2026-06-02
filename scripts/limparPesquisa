/**
 * BOTÃO: LIMPAR
 * Apaga os dados filtrados e reseta os campos de seleção.
 */
function limparPesquisa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPainel = ss.getSheetByName("Painel de Trocas");
  
  if (!abaPainel) return;

  // 1. Apaga os dados da pesquisa (assume que começam na linha 8)
  const ultimaLinha = abaPainel.getLastRow();
  if (ultimaLinha >= 8) {
    abaPainel.getRange(8, 1, ultimaLinha - 6, abaPainel.getLastColumn()).clearContent();
  }

  // 2. Limpa os campos de mês e ano (C3 e C4)
  abaPainel.getRange("C3:C4").clearContent();
}
