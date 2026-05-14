function enviarAlertaEstoque() {
  
  // 🔹 Abre a planilha pelo ID e seleciona a aba específica
  var sheet = SpreadsheetApp
    .openById("") //Colocar o ID da planilha aqui, senao n funfa
    .getSheetByName("TESTE Estoque");
  
  // 🔹 Validação: garante que a aba existe
  if (!sheet) {
    throw new Error("A aba 'TESTE Estoque' não foi encontrada.");
  }

  // 🔹 Pega todos os dados da planilha (incluindo fórmulas já calculadas)
  var dados = sheet.getDataRange().getValues();

  // 🔹 Array que armazenará os itens abaixo do estoque mínimo
  var itensCriticos = [];

  // 🔹 Loop pelas linhas da planilha
  // Começa da linha 3 (índice 2), pois linhas anteriores são cabeçalho
  for (var i = 2; i < dados.length; i++) {
    
    // 🔹 Captura os dados das colunas específicas
    var nomeProduto = dados[i][1];     // Coluna B → Nome do produto
    var quantidade = dados[i][7];      // Coluna H → Quantidade atual
    var estoqueMinimo = dados[i][11];  // Coluna L → Estoque mínimo

    // 🔹 Normaliza os dados (garante tipos corretos)
    var nome = String(nomeProduto).trim(); // Converte para texto e remove espaços
    var qtd = Number(quantidade);          // Converte para número
    var min = Number(estoqueMinimo);       // Converte para número

    // 🔹 Validação:
    // Ignora a linha se:
    // - Nome estiver vazio
    // - Quantidade não for número
    // - Estoque mínimo não for número
    if (!nome || isNaN(qtd) || isNaN(min)) {
      continue; // Pula para a próxima linha
    }

    // 🔹 (Opcional)
    // Ignora produtos com quantidade 0 ou negativa
    // Útil quando fórmulas retornam vazio ("") que vira 0


    // 🔹 Regra de negócio:
    // Se a quantidade atual for menor que o estoque mínimo → item crítico
    if (qtd < min) {
      itensCriticos.push({
        nome: nome, // Nome do produto
        qtd: qtd,   // Quantidade atual
        min: min    // Estoque mínimo
      });
    }

    // 🔍 DEBUG:
    // Mostra no log os dados de cada linha processada
    Logger.log(
      "Linha " + (i+1) +
      " | Produto: " + nome +
      " | Qtd: " + qtd +
      " | Min: " + min
    );
  }

  // 🔹 Se não houver itens críticos, encerra o script
  if (itensCriticos.length === 0) {
    Logger.log("Nenhum item abaixo do estoque mínimo.");
    return;
  }
  var emailDestino = ["arthlaender@gmail.com"];
  // ============================
  // 📧 MONTAGEM DO E-MAIL
  // ============================

  // 🔹 Texto inicial do e-mail
  var mensagemHTML = `
  <p><strong>Atenção,</strong></p>
  <p>Os seguintes itens estão abaixo do estoque mínimo:</p>

  <table style="border-collapse: collapse; width: 100%; font-family: Arial;">
    <tr style="background-color: #f2f2f2;">
      <th style="border: 1px solid #ddd; padding: 8px;">Produto</th>
      <th style="border: 1px solid #ddd; padding: 8px;">Quantidade Atual</th>
      <th style="border: 1px solid #ddd; padding: 8px;">Estoque Mínimo</th>
      <th style="border: 1px solid #ddd; padding: 8px;">Status</th>
    </tr>
`;

// 🔹 Preenche a tabela
itensCriticos.forEach(function(item) {

  // Define nível de criticidade
  var status = "";
  var cor = "";

  if (item.qtd === 0) {
    status = "CRÍTICO";
    cor = "#ff4d4d";
  } else if (item.qtd <= item.min * 0.5) {
    status = "URGENTE";
    cor = "#ffa500";
  } else {
    status = "BAIXO";
    cor = "#ffd966";
  }

  mensagemHTML += `
    <tr>
      <td style="border: 1px solid #ddd; padding: 8px;">${item.nome}</td>
      <td style="border: 1px solid #ddd; padding: 8px; text-align:center;">${item.qtd}</td>
      <td style="border: 1px solid #ddd; padding: 8px; text-align:center;">${item.min}</td>
      <td style="border: 1px solid #ddd; padding: 8px; text-align:center; background-color:${cor};">
        <strong>${status}</strong>
      </td>
    </tr>
  `;
});

// 🔹 Fecha tabela
mensagemHTML += `
  </table>

  <p style="margin-top: 15px;">
    Favor verificar o almoxarifado o quanto antes.
  </p>

  <p style="color: gray; font-size: 12px;">
    Mensagem automática.
  </p>
`;

// 🔹 Envio do e-mail
MailApp.sendEmail({
  to: emailDestino.join(","),
  subject: "⚠️ Alerta de Estoque Baixo",
  htmlBody: mensagemHTML
});
}