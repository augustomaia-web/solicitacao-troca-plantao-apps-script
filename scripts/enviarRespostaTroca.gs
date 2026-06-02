function enviarRespostaTroca(e) {
  if (!e) return;

  const sheet = e.source.getActiveSheet();
  const range = e.range;

  // --- 1. MAPEAMENTO DAS COLUNAS ---
  const colSoltr = 17;     // Coluna Q (AÇÃO DO COORDENADOR)
  const colDataRegistro = 18;   // Coluna R (REGISTRO DA RESPOSTA)
  const colEmailSoli = 3;       // Coluna C (E-MAIL DO SOLICITANTE)
  const colEmailSnt = 13;       // Coluna M (E-MAIL DO SOLICITADO)
  const colDataPlantaoSoli = 11; // Coluna K (DIA DO PLANTÃO - SOLICITADO)
  const colTurnoSoli = 12;      // Coluna L (TURNO - SOLICITADO)
  const colDataPlantaoSnt = 7;  // Coluna G (DIA DO PLANTÃO - SOLICITANTE)
  const colTurnoSnt = 8;        // Coluna H (TURNO - SOLICITANTE)
  const colNomeSoli = 9;        // Coluna I (NOME DO SOLICITADO)
  const colNomeSnt = 5;         // Coluna E (NOME DO SOLICITANTE)
  const colDataSlt = 4 // Coluna D (DATA DA SOLICITAÇÃO DE TROCA)

  // --------------------------------------------------------------

  // Verifica se a edição foi na coluna Q e ignora o cabeçalho
  if (range.getColumn() !== colSoltr || range.getRow() === 1) return;

  const linha = range.getRow();
  const status = range.getValue().toString().trim().toUpperCase(); 
  
  // Coleta nomes e e-mails
  const emailSolicitante = sheet.getRange(linha, colEmailSoli).getValue();
  const emailSolicitado = sheet.getRange(linha, colEmailSnt).getValue();
  const nomeSolicitado = sheet.getRange(linha, colNomeSoli).getValue();
  const nomeSolicitante = sheet.getRange(linha, colNomeSnt).getValue();

  // Verifica se existe ao menos um e-mail válido
  if (!emailSolicitante && !emailSolicitado) return;

  // Função para formatar a data (Evita erros de fuso horário)
  const fData = (col) => {
    const val = sheet.getRange(linha, col).getValue();
    return val instanceof Date ? Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy") : val;
  };

  // --- COLETA DE DADOS COM NOMES ÚNICOS ---
  const dSoli = fData(colDataPlantaoSoli);
  const tSoli = sheet.getRange(linha, colTurnoSoli).getValue();
  const dSnt = fData(colDataPlantaoSnt);
  const tSnt = sheet.getRange(linha, colTurnoSnt).getValue();
  const dSlt = fData(colDataSlt);
  
  let assunto = "";
  let mensagem = "";

if (status === "APROVADO") {
    assunto = "Comprovante de Troca de Plantão - APROVADA✅";
    mensagem = "Prezados(as),\n\nA solicitação de troca de plantão enviada no dia " + dSlt + " foi APROVADA ✅\n\n" +
               "--- COMPROVANTE DE TROCA ---\n" +
               "SOLICITANTE: " + nomeSolicitante + "\n" +
               "Plantão que assume: " + dSoli + " (" + tSoli + ")\n\n" +
               "SOLICITADO: " + nomeSolicitado + "\n" +
               "Plantão que admiti: " + dSnt + " (" + tSnt + ")\n" +
               "----------------------------\n\n" +
               "Atenciosamente,\nCoordenação de Enfermagem";
  } 
  else if (status === "NÃO APROVADO") {
    assunto = "Solicitação de troca de plantão - NÃO APROVADA❌";
    mensagem = "Prezado(a),\n\nA solicitação de troca de plantão enviada no dia " + dSlt + " entre,\n" + nomeSolicitante + " e " + nomeSolicitado + " NÃO foi aprovada ❌\n\n" +
               "Atenciosamente,\nCoordenação de Enfermagem";
  }

  if (assunto !== "") {
    try {
      // Monta a lista de destinatários removendo campos vazios
      let listaEmails = [emailSolicitante, emailSolicitado].filter(String).join(",");

      // --- O COMANDO DE ENVIO ---
      MailApp.sendEmail({
        to: listaEmails,
        subject: assunto,
        body: mensagem
      });
      
      // Registra a data da resposta na Coluna R
      sheet.getRange(linha, colDataRegistro).setValue(new Date());
    } catch (err) {
      Logger.log("Erro no envio: " + err.toString());
    }
  }
}
