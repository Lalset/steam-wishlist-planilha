// Adiciona um menu personalizado "Wishlist" na interface do Google Sheets(lá em cima ele vai aparecer, tinha outras opções também)
// com a opção "ATUALIZE" que executa a função atualizarSomenteNovosJogos
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Wishlist')
    .addItem('ATUALIZE', 'atualizarSomenteNovosJogos')
    .addToUi();
}

// Atualiza apenas as linhas onde o nome ou preço do jogo estão vazios
// Ideal para atualização incremental ao adicionar novos jogos
function atualizarSomenteNovosJogos() {
  const abaAtual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ultimaLinha = abaAtual.getLastRow();

  // Percorre as linhas da planilha, da última até a segunda (considerando cabeçalho na primeira)
  for (let linha = ultimaLinha; linha >= 2; linha--) {
    const nomeJogo = abaAtual.getRange(linha, 1).getValue();
    const precoJogo = abaAtual.getRange(linha, 2).getValue();

    // Executa atualização somente se o nome ou preço estiverem ausentes
    if (!nomeJogo || !precoJogo) {
      atualizarLinha(abaAtual, linha);
    }
  }
}

// Atualiza todos os jogos presentes na planilha
// Deve ser acionada via trigger programada para evitar bloqueios por chamadas em excesso
function atualizarTodosOsJogos() {
  const abaAtual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ultimaLinha = abaAtual.getLastRow();

  // Atualiza linha por linha da última até a segunda
  for (let linha = ultimaLinha; linha >= 2; linha--) {
    atualizarLinha(abaAtual, linha);
  }
}

// Atualiza os dados de uma linha específica: nome, preço e status de promoção
// Recebe a referência da aba e o número da linha a ser atualizada
function atualizarLinha(abaAtual, linha) {
  const linkSteam = abaAtual.getRange(linha, 4).getValue(); // Coluna D: Link da Steam
  const comprado = abaAtual.getRange(linha, 5).getValue().toString().trim().toUpperCase(); // Coluna E: Status de compra

  const celulaComprado = abaAtual.getRange(linha, 5);

  // Se o jogo já foi comprado, apenas formata a célula e não atualiza dados
  if (comprado === "SIM") {
    celulaComprado.setBackground("#a4c2f4"); // Cor azul para itens comprados
    return;
  } else {
    celulaComprado.setValue("NÃO"); // Garante que o status esteja como NÃO comprado
    celulaComprado.setBackground("#ea9999"); // Cor vermelha para itens não comprados
  }

  // Vê se o link é válido e corresponde à Steam
  if (!linkSteam || typeof linkSteam !== 'string' || !linkSteam.includes("store.steampowered.com")) {
    abaAtual.getRange(linha, 1, 1, 3).setValues([["Link inválido", "Preço não encontrado", "Não"]]);
    return;
  }

  // Extrai o AppID do URL da Steam via expressão regular
  const appIdExtraido = linkSteam.match(/\/app\/(\d+)/);
  if (!appIdExtraido) {
    abaAtual.getRange(linha, 1, 1, 3).setValues([["AppID não encontrado", "Preço não encontrado", "Não"]]);
    return;
  }

  const appId = appIdExtraido[1];
  const urlApiSteam = `https://store.steampowered.com/api/appdetails?appids=${appId}&cc=br&l=portuguese`;

  try {
    // Realiza a requisição para API pública da Steam para obter dados do jogo
    const resposta = UrlFetchApp.fetch(urlApiSteam);
    const dados = JSON.parse(resposta.getContentText());

    // Verifica se a resposta foi bem sucedida para o AppID solicitado
    if (!dados[appId].success) {
      abaAtual.getRange(linha, 1, 1, 3).setValues([["Produto não encontrado", "Preço não encontrado", "Não"]]);
      return;
    }

    // Extrai as informações relevantes do JSON retornado
    const infoJogo = dados[appId].data;
    const precoInfo = infoJogo.price_overview;

    const nomeDoJogo = infoJogo.name;
    const precoFinal = precoInfo.final / 100; // Converte preço para formato decimal
    const estaEmPromocao = precoInfo.discount_percent > 0 ? "SIM" : "NÃO";
    const precoFormatado = `R$ ${precoFinal.toFixed(2).replace(".", ",")}`;

    // Atualiza a planilha com nome, preço e status de promoção
    abaAtual.getRange(linha, 1).setValue(nomeDoJogo);
    abaAtual.getRange(linha, 2).setValue(precoFormatado);
    abaAtual.getRange(linha, 3).setValue(estaEmPromocao);

    // Ajusta a cor da célula de promoção conforme o status
    const celulaPromo = abaAtual.getRange(linha, 3);
    if (estaEmPromocao === "SIM") {
      celulaPromo.setBackground("#a4c2f4"); // Azul claro para promoções
    } else {
      celulaPromo.setBackground("#ea9999"); // Vermelho claro para jogos sem promoção
    }

    // Pausa para evitar excesso de requisições e bloqueio pela Steam, o que me ocorreu no primeiro dia, pois fiz muitas requisições, CUIDADO!
    Utilities.sleep(1000);
  } catch (erro) {
    // Em caso de erro na requisição, registra o problema na planilha e no log
    abaAtual.getRange(linha, 1, 1, 3).setValues([["Erro ao buscar", "Preço não encontrado", "Não"]]);
    Logger.log(`Erro na linha ${linha}: ${erro}`);
  }
}

// Cria um acionador programado para atualizar todos os jogos diariamente às 7h da manhã(pode botar o horário que quiser :) )
// Remove acionadores existentes para evitar duplicidade antes de criar o novo
// Detalhe que pode também usar a aba de acionadores do App Script, porém achei mais prático usar por aqui.
function criarAcionadorDiario() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "atualizarTodosOsJogos") {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp.newTrigger("atualizarTodosOsJogos")
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();
}
