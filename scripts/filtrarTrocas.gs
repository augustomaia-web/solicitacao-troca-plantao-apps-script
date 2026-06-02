/**
 * SISTEMA DE PESQUISA DO PAINEL DE TROCAS
 * Lê o Mês (C3) e o Ano (C4) e puxa os dados a partir da linha 8.
 */
function filtrarTrocas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPainel = ss.getSheetByName("Painel de Trocas");
  const abaDados = ss.getSheetByName("dados");
  const ui = SpreadsheetApp.getUi();
  
  if (!abaPainel || !abaDados) {
    ui.alert("Erro: As abas 'Painel de Trocas' ou 'dados' não foram encontradas. Verifique os nomes.");
    return;
  }
  
  // 1. Lê o Mês e o Ano selecionados nas células C3 e C4
  const mesSelecionado = abaPainel.getRange("C3").getValue().toString().trim();
  const anoSelecionado = abaPainel.getRange("C4").getValue().toString().trim();
  
  if (!mesSelecionado || !anoSelecionado) {
    ui.alert("Atenção: Selecione o Mês e o Ano antes de clicar em Pesquisar.");
    return;
  }
  
  // 2. Limpa a área de resultados antigos (da linha 8 para baixo)
  const ultimaLinhaPainel = abaPainel.getLastRow();
  const totalColunasPainel = Math.max(abaPainel.getLastColumn(), abaDados.getLastColumn());
  
  if (ultimaLinhaPainel >= 8) {
    // Apaga apenas o conteúdo, mantendo as cores e formatações que você fez
    abaPainel.getRange(8, 1, ultimaLinhaPainel - 7, totalColunasPainel).clearContent();
  }
  
  // 3. Lê o Banco de Dados inteiro (aba "dados")
  const dados = abaDados.getDataRange().getValues();
  if (dados.length <= 1) {
    ui.alert("O banco de dados (aba 'dados') está vazio.");
    return;
  }
  
  const totalColunasDados = abaDados.getLastColumn();
  const linhasFiltradas = [];
  const meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
  
  // 4. Vasculha linha por linha procurando o mês e ano exatos
  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];
    let dataSlt = linha[3]; // Coluna D (índice 3) é a Data da Solicitação
    
    if (dataSlt) {
      let dataObj = new Date(dataSlt);
      
      // Se o Sheets ler a data como texto puro "DD/MM/AAAA", fazemos a conversão
      if (isNaN(dataObj.getTime())) {
        let textoData = String(dataSlt).split(" ")[0]; 
        let partes = textoData.split("/");
        if (partes.length === 3) {
          dataObj = new Date(partes[2], partes[1] - 1, partes[0]);
        }
      }
      
      // Se for uma data válida, verifica se bate com a pesquisa do usuário
      if (!isNaN(dataObj.getTime())) {
        let nomeMesData = meses[dataObj.getMonth()];
        let anoData = dataObj.getFullYear().toString();
        
        // Ignora maiúsculas e minúsculas para evitar erros de digitação (ex: MAIO e Maio)
        if (nomeMesData.toLowerCase() === mesSelecionado.toLowerCase() && anoData === anoSelecionado) {
          linhasFiltradas.push(linha);
        }
      }
    }
  }
  
  // 5. Injeta os resultados no seu Painel a partir da linha 8
  if (linhasFiltradas.length > 0) {
    abaPainel.getRange(8, 1, linhasFiltradas.length, totalColunasDados).setValues(linhasFiltradas);
    ui.alert("Pronto! " + linhasFiltradas.length + " troca(s) encontrada(s) para " + mesSelecionado + " de " + anoSelecionado + ".");
  } else {
    ui.alert("Nenhuma troca de plantão encontrada para " + mesSelecionado + " de " + anoSelecionado + ".");
  }
}
