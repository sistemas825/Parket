function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ERP Compras - Parket')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Marcador de versão para diagnosticar deploy desatualizado: se isso não
// bater com a tag "build" mostrada no topo do app depois de publicar uma
// Nova versão, o navegador (ou o Web App) ainda está servindo código antigo.
function obterVersaoBackend() {
  return "2026-07-31.8";
}

function obterDadosKanban() {
  var sheet = obterAbaFlexivel_("Kanban");
  if (!sheet) return empacotarResposta_([], "Aba 'Kanban' não encontrada.");
  var dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return empacotarResposta_(dados, "Aba '" + sheet.getName() + "' está vazia (só cabeçalho ou nada).");
  return empacotarResposta_(dados, "Aba '" + sheet.getName() + "': " + (dados.length - 1) + " cartão(ões) na planilha (cabeçalho: " + dados[0].join(" | ") + ").");
}

// CORREÇÃO: várias abas de cadastro (Produtos, Clientes, Fornecedores,
// Transportadoras) podem ter sido criadas com título/linhas em branco antes
// do cabeçalho de verdade, em posições variadas — não só o padrão fixo
// "linha 1 = título, linha 2 = em branco, linha 3 = cabeçalho". Agora
// procuramos ativamente, dentro das primeiras linhas, qual é a que realmente
// parece um cabeçalho (a primeira com pelo menos 2 células preenchidas) e
// cortamos tudo antes dela.
function localizarLinhaCabecalho_(dados) {
  var limite = Math.min(dados.length, 10);
  for (var i = 0; i < limite; i++) {
    var preenchidas = dados[i].filter(function (c) { return String(c || "").trim() !== ""; }).length;
    if (preenchidas >= 2) return i;
  }
  return 0;
}

function normalizarInicioTabela_(dados) {
  if (!dados || dados.length === 0) return dados;
  var indice = localizarLinhaCabecalho_(dados);
  return indice > 0 ? dados.slice(indice) : dados;
}

// Versão interna (objeto puro, não string) — usada por pesquisarDadosAba
// para não ter que dar JSON.parse em cima de JSON.stringify sem necessidade.
function obterDadosAba_(tipo) {
  var nomeAba = "Produtos";
  if (tipo === 'clientes') nomeAba = "Clientes";
  else if (tipo === 'produtos') nomeAba = "Produtos";
  else if (tipo === 'transportadoras') nomeAba = "Transportadoras";
  else if (tipo === 'fornecedores') nomeAba = "Fornecedores";
  else if (tipo === 'financeiro') nomeAba = "Financeiro";
  else if (tipo === 'estoque-xml' || tipo === 'estoque-manual' || tipo === 'estoque-saida') nomeAba = "Estoque";
  else if (tipo === 'rel-pedidos' || tipo === 'rel-entradas') nomeAba = "Relatorios";

  var sheet = obterAbaFlexivel_(nomeAba);
  if (!sheet) return { linhas: [], debug: "Aba '" + nomeAba + "' não encontrada na planilha." };

  var dadosBrutos = sheet.getDataRange().getValues();
  var totalLinhasBrutas = dadosBrutos.length;
  var indiceHeader = localizarLinhaCabecalho_(dadosBrutos);
  var dados = indiceHeader > 0 ? dadosBrutos.slice(indiceHeader) : dadosBrutos;

  if (dados.length === 0) {
    return { linhas: [], debug: "Aba '" + nomeAba + "' está vazia (getDataRange devolveu " + totalLinhasBrutas + " linha(s) no total)." };
  }

  var infoHeader = indiceHeader > 0 ? " — cabeçalho encontrado na linha " + (indiceHeader + 1) + " da planilha" : "";
  return {
    linhas: dados,
    debug: "Aba '" + nomeAba + "': " + totalLinhasBrutas + " linha(s) brutas na planilha" + infoHeader + ", cabeçalho: [" + dados[0].join(" | ") + "], " + (dados.length - 1) + " registro(s) de dado(s)."
  };
}

function obterDadosAba(tipo) {
  var res = obterDadosAba_(tipo);
  return empacotarResposta_(res.linhas, res.debug);
}

function pesquisarDadosAba(tipo, termo) {
  var res = obterDadosAba_(tipo);
  var dados = res.linhas;
  if (!dados || dados.length === 0) return empacotarResposta_(dados, res.debug);
  if (!termo || termo.trim() === "") return empacotarResposta_(dados, res.debug);

  var cabecalhos = dados.slice(0, 1);
  var termoBusca = String(termo).toLowerCase().trim();

  var linhasFiltradas = [];
  for (var i = 1; i < dados.length; i++) {
    var linha = dados[i];
    var atende = linha.some(function(celula) {
      return String(celula).toLowerCase().indexOf(termoBusca) !== -1;
    });
    if (atende) linhasFiltradas.push(linha);
  }
  return empacotarResposta_(cabecalhos.concat(linhasFiltradas), linhasFiltradas.length + " resultado(s) para \"" + termo + "\".");
}

// Procura uma aba pelo nome ignorando maiúsculas/minúsculas e espaços nas
// pontas — protege contra abas como "Financeiro " (com espaço) ou
// "financeiro" que getSheetByName() exato não encontraria.
function obterAbaFlexivel_(nomeAlvo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nomeAlvo);
  if (sheet) return sheet;
  var alvo = String(nomeAlvo).trim().toLowerCase();
  var todas = ss.getSheets();
  for (var i = 0; i < todas.length; i++) {
    if (todas[i].getName().trim().toLowerCase() === alvo) return todas[i];
  }
  return null;
}

function pesquisarFinanceiroComData(fornecedor, nf, projeto, dataInicio, dataFim) {
  var sheet = obterAbaFlexivel_("Financeiro");
  if (!sheet) {
    var nomesExistentes = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function (s) { return s.getName(); }).join(", ");
    return empacotarResposta_([], "Aba 'Financeiro' não encontrada. Abas existentes na planilha: " + nomesExistentes);
  }

  var dados = sheet.getDataRange().getValues();
  if (dados.length === 0) return empacotarResposta_([], "Aba '" + sheet.getName() + "' está completamente vazia (nem cabeçalho).");

  var cabecalhoPadrao = ["CNPJ Fornecedor", "Fornecedor", "Nota Fiscal", "Data Emissão", "Data Vencimento", "Valor (R$)", "Status Pagamento", "Projeto", "Histórico"];

  // CORREÇÃO CRÍTICA: a versão anterior forçava esse cabeçalho padrão em toda
  // pesquisa (mesmo que a planilha real já tivesse outro layout de colunas),
  // sem realinhar os dados abaixo. Isso fazia os índices fixos (1,2,4,7) lerem
  // a coluna errada sempre que o layout real divergia, e os lançamentos vindos
  // do XML "sumiam" da tela mesmo estando corretos na planilha.
  // Agora só criamos o cabeçalho padrão se a planilha estiver de fato vazia;
  // caso contrário, lemos os nomes de coluna que já existem na linha 1.
  var cabecalho = dados[0];
  var cabecalhoVazio = cabecalho.every(function (c) { return String(c || "").trim() === ""; });
  if (cabecalhoVazio) {
    sheet.getRange(1, 1, 1, cabecalhoPadrao.length).setValues([cabecalhoPadrao]);
    cabecalho = cabecalhoPadrao;
  }

  if (dados.length <= 1) {
    return empacotarResposta_([cabecalho], "Aba '" + sheet.getName() + "' tem só o cabeçalho (" + cabecalho.join(" | ") + "), nenhuma linha de dado abaixo dele.");
  }

  var col = obterMapaColunas_(cabecalho);
  var idxForn = col["fornecedor"] !== undefined ? col["fornecedor"] : 1;
  var idxNf = col["nota fiscal"] !== undefined ? col["nota fiscal"] : (col["documento/nf"] !== undefined ? col["documento/nf"] : 2);
  var idxVenc = col["data vencimento"] !== undefined ? col["data vencimento"] : (col["data pagamento"] !== undefined ? col["data pagamento"] : 4);
  var idxProj = col["projeto"] !== undefined ? col["projeto"] : 7;

  var linhas = dados.slice(1);
  var totalTexto = 0, totalData = 0;

  var filtradas = linhas.filter(function(row) {
    var fornRow = String(row[idxForn] || "").toLowerCase();
    var nfRow = String(row[idxNf] || "").toLowerCase();
    var dataVenc = row[idxVenc];
    var projRow = String(row[idxProj] || "").toLowerCase();

    if (projeto && projeto.trim() !== "" && projRow.indexOf(projeto.toLowerCase().trim()) === -1) return false;
    if (fornecedor && fornecedor.trim() !== "" && fornRow.indexOf(fornecedor.toLowerCase().trim()) === -1) return false;
    if (nf && nf.trim() !== "" && nfRow.indexOf(nf.toLowerCase().trim()) === -1) return false;
    totalTexto++;

    if (dataInicio || dataFim) {
      if (!dataDentroDoIntervalo_(dataVenc, dataInicio, dataFim)) return false;
    }
    totalData++;

    return true;
  });

  return empacotarResposta_(
    [cabecalho].concat(filtradas),
    "Aba '" + sheet.getName() + "': " + (dados.length - 1) + " linha(s) na planilha (cabeçalho: " + cabecalho.join(" | ") + "), " + totalTexto + " após filtros de texto, " + totalData + " após filtro de data, " + filtradas.length + " no resultado final."
  );
}

function salvarRegistroGenerico(d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = obterAbaFlexivel_(d.aba);
    if (!sheet) throw new Error("Aba '" + d.aba + "' não encontrada.");

    var range = sheet.getDataRange();
    var valores = range.getValues();
    var rowIndex = -1;

    if (d.idOuChaveAntiga) {
      for (var i = 1; i < valores.length; i++) {
        if (String(valores[i][0]).trim() === String(d.idOuChaveAntiga).trim()) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, d.dados.length).setValues([d.dados]);
      return "Registro atualizado com sucesso!";
    } else {
      sheet.appendRow(d.dados);
      return "Registro cadastrado com sucesso!";
    }
  } catch (e) {
    throw new Error("Erro ao salvar: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// Exclui um registro de qualquer aba de cadastro genérico (Clientes,
// Produtos, Fornecedores, Transportadoras) buscando pela chave da primeira
// coluna (mesmo critério usado por salvarRegistroGenerico para localizar
// a linha a atualizar).
function excluirRegistroGenerico(aba, chave) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = obterAbaFlexivel_(aba);
    if (!sheet) throw new Error("Aba '" + aba + "' não encontrada.");

    var valores = sheet.getDataRange().getValues();
    for (var i = 1; i < valores.length; i++) {
      if (String(valores[i][0]).trim() === String(chave).trim()) {
        sheet.deleteRow(i + 1);
        return "Registro excluído com sucesso!";
      }
    }
    throw new Error("Registro com chave '" + chave + "' não encontrado.");
  } catch (e) {
    throw new Error("Erro ao excluir: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ==============================================
// ESTOQUE - Saldo consolidado (refatorado em helper reutilizável)
// ==============================================
function obterMapaSaldos_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Estoque");
  var saldos = {};
  if (!sheet) return saldos;

  var dados = sheet.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    var codigo = String(row[1] || "").trim(); // Coluna B
    if (!codigo) continue;

    var descricao = row[2] || ""; // Coluna C
    var tipo = String(row[3] || "").trim(); // Coluna D
    var qtd = parseFloat(row[4]) || 0; // Coluna E
    var valorUnit = parseFloat(row[5]) || 0; // Coluna F

    if (!saldos[codigo]) {
      saldos[codigo] = { codigo: codigo, descricao: descricao, quantidade: 0, valorUnitario: valorUnit };
    }

    if (tipo === "Entrada") {
      saldos[codigo].quantidade += qtd;
      if (valorUnit > 0) saldos[codigo].valorUnitario = valorUnit;
    } else if (tipo === "Saída") {
      saldos[codigo].quantidade -= qtd;
    }
  }
  return saldos;
}

function obterSaldoEstoqueAtual() {
  var saldos = obterMapaSaldos_();
  var resultado = [];
  for (var cod in saldos) {
    var item = saldos[cod];
    var valorTotalEstoque = item.quantidade * item.valorUnitario;
    resultado.push([item.codigo, item.descricao, item.quantidade, item.valorUnitario.toFixed(2), valorTotalEstoque.toFixed(2)]);
  }
  return resultado;
}

function obterListaProjetos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Estoque");
  if (!sheet) return [];
  var dados = sheet.getDataRange().getValues();
  var projetos = {};
  for (var i = 1; i < dados.length; i++) {
    var proj = String(dados[i][9] || "").trim(); // Coluna J (Índice 9)
    if (proj) projetos[proj] = true;
  }
  return Object.keys(projetos).sort();
}

// ==============================================
// KANBAN
// Layout de colunas (1-indexed):
// 1-ID | 2-Título | 3-Solicitante | 4-Setor | 5-Etapa | 6-Prioridade | 7-Valor | 8-Data
// 9-Pedido Por | 10-Projeto | 11-Prazo Entrega | 12-Observações
// (colunas 9-12 adicionadas ao final para não quebrar índices já usados por outras funções)
// ==============================================
function adicionarCartaoKanban(d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kanban");
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Kanban");
    }
    sheet.getRange(1, 1, 1, 12).setValues([[
      "ID", "Título", "Solicitante", "Setor", "Etapa", "Prioridade", "Valor", "Data",
      "Pedido Por", "Projeto", "Prazo Entrega", "Observações"
    ]]);

    var novoId = gerarProximoIdKanban_(sheet);
    var etapaInicial = "1. Solicitado (Entrada)";
    var dataSolicitacao = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    var descMateriais = d.materiais.map(function(m) {
      return m.quantidade + "x " + m.material + " (" + m.justificativa + ")";
    }).join(" | ");

    var tituloCartao = "Solicitação: " + (d.departamento || "Geral") + " - " + (d.projeto || "Sem Projeto");

    sheet.appendRow([
      novoId,
      tituloCartao + " | Itens: " + descMateriais,
      d.seuNome || "",
      d.setor || "Geral",
      etapaInicial,
      "Normal",
      0,
      dataSolicitacao,
      d.pedidoPor || "",
      d.projeto || "",
      d.prazo || "",
      d.observacoes || ""
    ]);

    return "Solicitação " + novoId + " criada com sucesso!";
  } catch (e) {
    throw new Error("Erro ao criar solicitação: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

function atualizarEtapaCartao(id, novaEtapa) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kanban");
    if (!sheet) throw new Error("Aba 'Kanban' não encontrada.");

    var dados = sheet.getDataRange().getValues();
    var idBuscado = String(id).trim();
    for (var i = 0; i < dados.length; i++) {
      var idLinha = String(dados[i][0]).trim();
      if (idLinha === idBuscado) {
        sheet.getRange(i + 1, 5).setValue(novaEtapa);
        SpreadsheetApp.flush();
        return "Cartão " + id + " movido para: " + novaEtapa;
      }
    }
    throw new Error("Cartão com ID '" + id + "' não encontrado.");
  } catch (e) {
    throw new Error("Erro ao mover cartão: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// CORREÇÃO CRÍTICA: esta função era chamada pelo Index.html mas não existia,
// causando erro sempre que um cartão do Kanban era clicado.
function obterDetalhesCartaoKanban(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kanban");
  if (!sheet) throw new Error("Aba 'Kanban' não encontrada.");

  var dados = sheet.getDataRange().getValues();
  var idBuscado = String(id).trim();
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).trim() === idBuscado) {
      var row = dados[i];
      return {
        id: row[0],
        titulo: row[1],
        solicitante: row[2],
        setor: row[3],
        etapa: row[4],
        prioridade: row[5],
        valor: row[6],
        data: row[7],
        pedidoPor: row[8] || "",
        projeto: row[9] || "",
        prazo: row[10] || "",
        observacoes: row[11] || ""
      };
    }
  }
  return null;
}

// Atualiza campos editáveis do cartão a partir do modal de detalhes.
function atualizarCartaoCompleto(id, dados) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kanban");
    if (!sheet) throw new Error("Aba 'Kanban' não encontrada.");

    var valores = sheet.getDataRange().getValues();
    var idBuscado = String(id).trim();
    for (var i = 1; i < valores.length; i++) {
      if (String(valores[i][0]).trim() === idBuscado) {
        var linha = i + 1;
        if (dados.solicitante !== undefined) sheet.getRange(linha, 3).setValue(dados.solicitante);
        if (dados.setor !== undefined) sheet.getRange(linha, 4).setValue(dados.setor);
        if (dados.prazo !== undefined) sheet.getRange(linha, 11).setValue(dados.prazo);
        if (dados.observacoes !== undefined) sheet.getRange(linha, 12).setValue(dados.observacoes);
        SpreadsheetApp.flush();
        return "Cartão " + id + " atualizado com sucesso!";
      }
    }
    throw new Error("Cartão com ID '" + id + "' não encontrado.");
  } catch (e) {
    throw new Error("Erro ao atualizar cartão: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

function excluirCartaoKanban(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kanban");
    if (!sheet) throw new Error("Aba 'Kanban' não encontrada.");

    var dados = sheet.getDataRange().getValues();
    var idBuscado = String(id).trim();
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === idBuscado) {
        sheet.deleteRow(i + 1);
        return "Cartão " + id + " excluído com sucesso!";
      }
    }
    throw new Error("Cartão com ID '" + id + "' não encontrado.");
  } catch (e) {
    throw new Error("Erro ao excluir cartão: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// Gera um PDF simples com os dados do cartão e devolve a URL do arquivo no Drive.
function gerarPdfCartaoKanban(id) {
  var cartao = obterDetalhesCartaoKanban(id);
  if (!cartao) throw new Error("Cartão com ID '" + id + "' não encontrado.");

  var partes = String(cartao.titulo || "").split(" | Itens: ");
  var tituloPrincipal = partes[0] || "";
  var itensHtml = "Nenhum item detalhado.";
  if (partes.length > 1) {
    itensHtml = "<ul>" + partes[1].split(" | ").map(function(item) {
      return "<li>" + item + "</li>";
    }).join("") + "</ul>";
  }

  var html = ""
    + "<h2>" + tituloPrincipal + " — " + cartao.id + "</h2>"
    + "<p><b>Solicitante:</b> " + (cartao.solicitante || "-") + "</p>"
    + "<p><b>Pedido por:</b> " + (cartao.pedidoPor || "-") + "</p>"
    + "<p><b>Setor:</b> " + (cartao.setor || "-") + "</p>"
    + "<p><b>Projeto:</b> " + (cartao.projeto || "-") + "</p>"
    + "<p><b>Prazo de entrega:</b> " + (cartao.prazo || "-") + "</p>"
    + "<p><b>Etapa atual:</b> " + (cartao.etapa || "-") + "</p>"
    + "<p><b>Materiais:</b></p>" + itensHtml
    + "<p><b>Observações:</b> " + (cartao.observacoes || "-") + "</p>";

  var blob = Utilities.newBlob(html, "text/html", "cartao_" + id + ".html");
  var pdfBlob = blob.getAs("application/pdf").setName("Solicitacao_" + id + ".pdf");
  var arquivo = DriveApp.createFile(pdfBlob);
  arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return arquivo.getUrl();
}

function gerarProximoIdKanban_(sheet) {
  var dados = sheet.getDataRange().getValues();
  var maiorNumero = 0;
  for (var i = 0; i < dados.length; i++) {
    var idAtual = String(dados[i][0] || "");
    var match = idAtual.match(/SOL-(\d+)/);
    if (match) {
      var numero = parseInt(match[1], 10);
      if (numero > maiorNumero) maiorNumero = numero;
    }
  }
  return "SOL-" + ("0000" + (maiorNumero + 1)).slice(-4);
}

function processarXmlNFe(xmlContent, projeto) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!projeto) throw new Error("Informe o Projeto/Centro de Custo antes de importar.");

    var doc = XmlService.parse(xmlContent);
    var root = doc.getRootElement();
    var ns = root.getNamespace();

    var infNFe = obterInfNFe_(root, ns);
    if (!infNFe) throw new Error("Tag <infNFe> não encontrada.");

    var ide = infNFe.getChild('ide', ns);
    var emit = infNFe.getChild('emit', ns);
    var detList = infNFe.getChildren('det', ns);
    var total = infNFe.getChild('total', ns);
    var cobr = infNFe.getChild('cobr', ns);

    var numeroNF = obterTexto_(ide, 'nNF', ns);
    var dataEmissaoRaw = obterTexto_(ide, 'dhEmi', ns) || obterTexto_(ide, 'dEmi', ns);
    var dataEmissao = dataEmissaoRaw ? new Date(dataEmissaoRaw.substring(0, 10)) : new Date();

    var cnpjFornecedor = obterTexto_(emit, 'CNPJ', ns);
    var nomeFornecedor = obterTexto_(emit, 'xNome', ns);
    var fantasiaFornecedor = obterTexto_(emit, 'xFant', ns) || nomeFornecedor;
    var ie = obterTexto_(emit, 'IE', ns);

    var enderEmit = emit ? emit.getChild('enderEmit', ns) : null;
    var endereco = '';
    var telefone = '';
    if (enderEmit) {
      endereco = [
        obterTexto_(enderEmit, 'xLgr', ns),
        obterTexto_(enderEmit, 'nro', ns),
        obterTexto_(enderEmit, 'xBairro', ns),
        obterTexto_(enderEmit, 'xMun', ns),
        obterTexto_(enderEmit, 'UF', ns)
      ].filter(String).join(', ');
      telefone = obterTexto_(enderEmit, 'fone', ns);
    }

    upsertFornecedor_(cnpjFornecedor, nomeFornecedor, fantasiaFornecedor, ie, endereco, telefone);

    var itensProcessados = 0;
    detList.forEach(function (item) {
      var prod = item.getChild('prod', ns);
      var codigo = obterTexto_(prod, 'cProd', ns);
      var descricao = obterTexto_(prod, 'xProd', ns);
      var qtd = parseFloat(obterTexto_(prod, 'qCom', ns)) || 0;
      var valorUnit = parseFloat(obterTexto_(prod, 'vUnCom', ns)) || 0;
      var valorTotal = parseFloat(obterTexto_(prod, 'vProd', ns)) || 0;

      registrarEntradaEstoque_(codigo, descricao, qtd, valorUnit, valorTotal, numeroNF, nomeFornecedor, projeto, dataEmissao);
      itensProcessados++;
    });

    var duplicatas = cobr ? cobr.getChildren('dup', ns) : [];
    if (duplicatas.length > 0) {
      duplicatas.forEach(function (dup) {
        var dataVencRaw = obterTexto_(dup, 'dVenc', ns);
        var dataVenc = dataVencRaw ? new Date(dataVencRaw) : dataEmissao;
        var valorDup = parseFloat(obterTexto_(dup, 'vDup', ns)) || 0;
        registrarFinanceiro_(projeto, cnpjFornecedor, nomeFornecedor, numeroNF, dataEmissao, dataVenc, valorDup, "Em Aberto", "Duplicata NF " + numeroNF);
      });
    } else {
      var valorTotalNF = 0;
      if (total) {
        var icmsTot = total.getChild('ICMSTot', ns);
        valorTotalNF = parseFloat(obterTexto_(icmsTot, 'vNF', ns)) || 0;
      }
      registrarFinanceiro_(projeto, cnpjFornecedor, nomeFornecedor, numeroNF, dataEmissao, dataEmissao, valorTotalNF, "Em Aberto", "NF-e " + numeroNF);
    }

    return "NF-e nº " + numeroNF + " importada com sucesso! " + itensProcessados + " item(ns) lançado(s) no estoque e financeiro.";
  } catch (e) {
    throw new Error("Erro ao processar XML: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

function obterInfNFe_(root, ns) {
  var nfeElement = (root.getName() === 'nfeProc') ? root.getChild('NFe', ns) : root;
  return nfeElement ? nfeElement.getChild('infNFe', ns) : null;
}

function obterTexto_(elementoPai, tag, ns) {
  if (!elementoPai) return '';
  var filho = elementoPai.getChild(tag, ns);
  return filho ? filho.getText() : '';
}

// CORREÇÃO: antes, se o CNPJ já existisse a função simplesmente retornava sem
// atualizar nada — mudanças de endereço/telefone/razão social em NF-e futuras
// eram sempre ignoradas. Agora faz upsert de verdade.
function upsertFornecedor_(cnpj, razaoSocial, fantasia, ie, endereco, telefone) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fornecedores");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Fornecedores");
  }
  sheet.getRange(1, 1, 1, 6).setValues([["CNPJ", "Razão Social", "Nome Fantasia", "IE", "Endereço", "Telefone"]]);
  var dados = sheet.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (cnpj && String(dados[i][0]) === String(cnpj)) {
      sheet.getRange(i + 1, 1, 1, 6).setValues([[cnpj, razaoSocial, fantasia, ie, endereco, telefone]]);
      return;
    }
  }
  sheet.appendRow([cnpj, razaoSocial, fantasia, ie, endereco, telefone]);
}

function registrarEntradaEstoque_(codigo, descricao, qtd, valorUnit, valorTotal, numeroNF, fornecedor, projeto, data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Estoque");
  if (!sheet) {
    sheet = ss.insertSheet("Estoque");
  }
  // Padrão rigoroso do cabeçalho de Estoque
  sheet.getRange(1, 1, 1, 10).setValues([["Data", "Código", "Descrição", "Tipo", "Quantidade", "Valor Unitário", "Valor Total", "Documento/NF", "Fornecedor", "Projeto"]]);
  sheet.appendRow([data, codigo, descricao, "Entrada", qtd, valorUnit, valorTotal, numeroNF, fornecedor, projeto]);
}

// NOVO PADRÃO: Respeita a sua planilha física (CNPJ, Forn, NF, Emissão, Vencimento, Valor, Status) + Adiciona Projeto e Histórico no final.
// CORREÇÃO: antes o cabeçalho era forçado em toda chamada, mesmo quando a aba já
// tinha outro layout de colunas — agora só escrevemos o cabeçalho padrão quando a
// aba é nova ou a linha 1 está vazia, para não desalinhar dados já existentes.
function registrarFinanceiro_(projeto, cnpj, fornecedor, documento, dataEmissao, dataVencimento, valor, status, historico) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = obterAbaFlexivel_("Financeiro");
  var cabecalhoPadrao = ["CNPJ Fornecedor", "Fornecedor", "Nota Fiscal", "Data Emissão", "Data Vencimento", "Valor (R$)", "Status Pagamento", "Projeto", "Histórico"];
  if (!sheet) {
    sheet = ss.insertSheet("Financeiro");
    sheet.getRange(1, 1, 1, cabecalhoPadrao.length).setValues([cabecalhoPadrao]);
  } else {
    var primeiraLinha = sheet.getRange(1, 1, 1, cabecalhoPadrao.length).getValues()[0];
    var vazia = primeiraLinha.every(function (c) { return String(c || "").trim() === ""; });
    if (vazia) sheet.getRange(1, 1, 1, cabecalhoPadrao.length).setValues([cabecalhoPadrao]);
  }
  sheet.appendRow([cnpj, fornecedor, documento, dataEmissao, dataVencimento, valor, status, projeto, historico]);
}

function salvarEntradaManual(d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!d.projeto) throw new Error("Informe o Projeto/Centro de Custo.");
    if (!d.itens || d.itens.length === 0) throw new Error("Adicione ao menos um material.");

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Estoque");
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Estoque");
    }
    sheet.getRange(1, 1, 1, 10).setValues([["Data", "Código", "Descrição", "Tipo", "Quantidade", "Valor Unitário", "Valor Total", "Documento/NF", "Fornecedor", "Projeto"]]);

    var dataFinal = d.data ? new Date(d.data) : new Date();
    var proximoNumero = obterMaiorNumeroCodigoProduto_(sheet) + 1;
    var count = 0;

    d.itens.forEach(function (item) {
      var qtd = parseFloat(item.quantidade) || 0;
      if (qtd <= 0) return;
      var codigo = "PRD-" + ("0000" + proximoNumero).slice(-4);
      var valorUnit = parseFloat(item.valorUnitario) || 0;
      var valorTotal = qtd * valorUnit;

      sheet.appendRow([dataFinal, codigo, item.descricao || "", "Entrada", qtd, valorUnit, valorTotal,
          d.documento || "", d.fornecedor || "", d.projeto]);

      proximoNumero++;
      count++;
    });

    if (count === 0) throw new Error("Nenhum material válido informado.");

    if (d.pagamentos && d.pagamentos.length > 0) {
      d.pagamentos.forEach(function(pag) {
        var valPag = parseFloat(pag.valor) || 0;
        if (valPag > 0) {
          var dataVenc = pag.dataVencimento ? new Date(pag.dataVencimento) : dataFinal;
          registrarFinanceiro_(d.projeto, "MANUAL", d.fornecedor || "Diversos", d.documento || "MAN-001", dataFinal, dataVenc, valPag, "Em Aberto", "Entrada Manual");
        }
      });
    }

    SpreadsheetApp.flush();
    return count + " material(is) registrado(s) e financeiro gerado com sucesso!";
  } catch (e) {
    throw new Error("Erro ao registrar entrada: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// Gera o próximo número sequencial de "Documento" (protocolo) de saída, olhando
// os números já usados na coluna Documento/NF da aba Estoque para não repetir.
function gerarProximoNumeroProtocoloSaida_(sheet) {
  var dados = sheet.getDataRange().getValues();
  var col = obterMapaColunas_(dados[0]);
  var idxDoc = col["documento/nf"] !== undefined ? col["documento/nf"] : 7;
  var maior = 0;
  for (var i = 1; i < dados.length; i++) {
    var n = parseInt(String(dados[i][idxDoc] || "").replace(/\D/g, ""), 10);
    if (!isNaN(n) && n > maior) maior = n;
  }
  return maior + 1;
}

// CORREÇÃO: agora valida saldo disponível antes de lançar qualquer saída,
// evitando estoque negativo. A checagem é cumulativa dentro do próprio lote
// (duas linhas do mesmo código no mesmo envio são somadas antes de comparar).
// ATUALIZADO: também aceita cliente/endereço/unidade e devolve os dados prontos
// para impressão do "Protocolo de Saída" no padrão Parket.
function salvarSaidaMaterial(d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (!d.projeto) throw new Error("Informe o Projeto/Centro de Custo.");
    if (!d.itens || d.itens.length === 0) throw new Error("Adicione ao menos um material.");

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Estoque");
    if (!sheet) throw new Error("Aba 'Estoque' não encontrada.");

    sheet.getRange(1, 1, 1, 10).setValues([["Data", "Código", "Descrição", "Tipo", "Quantidade", "Valor Unitário", "Valor Total", "Documento/NF", "Fornecedor", "Projeto"]]);

    var saldos = obterMapaSaldos_();
    var itensValidos = d.itens.filter(function(item) {
      var qtd = parseFloat(item.quantidade) || 0;
      return qtd > 0 && item.codigo;
    });

    if (itensValidos.length === 0) throw new Error("Nenhum material válido informado.");

    var erros = [];
    itensValidos.forEach(function(item) {
      var qtd = parseFloat(item.quantidade) || 0;
      var disponivel = saldos[item.codigo] ? saldos[item.codigo].quantidade : 0;
      if (qtd > disponivel) {
        erros.push(item.codigo + " (disponível: " + disponivel + ", solicitado: " + qtd + ")");
      } else if (saldos[item.codigo]) {
        saldos[item.codigo].quantidade -= qtd;
      }
    });

    if (erros.length > 0) {
      throw new Error("Estoque insuficiente para: " + erros.join(", "));
    }

    var dataFinal = d.data ? new Date(d.data) : new Date();
    var numeroProtocolo = d.documento && d.documento.trim() !== "" ? d.documento.trim() : String(gerarProximoNumeroProtocoloSaida_(sheet));
    var count = 0;

    itensValidos.forEach(function (item) {
      var qtd = parseFloat(item.quantidade) || 0;
      sheet.appendRow([dataFinal, item.codigo, item.descricao || "", "Saída", qtd, 0, 0,
          numeroProtocolo, "", d.projeto]);
      count++;
    });

    SpreadsheetApp.flush();

    var agora = new Date();
    var protocolo = {
      numero: numeroProtocolo,
      dataEmissao: Utilities.formatDate(dataFinal, Session.getScriptTimeZone(), "dd/MM/yyyy"),
      dataEstoque: Utilities.formatDate(dataFinal, Session.getScriptTimeZone(), "dd/MM/yyyy"),
      horaEstoque: Utilities.formatDate(agora, Session.getScriptTimeZone(), "HH:mm"),
      centroCusto: d.projeto || "",
      cliente: d.cliente || "",
      endereco: d.enderecoCliente || "",
      observacoes: d.observacoes || "",
      itens: itensValidos.map(function(item) {
        return {
          produto: item.codigo + (item.descricao ? " - " + item.descricao : ""),
          unidade: item.unidade || "UN",
          quantidade: parseFloat(item.quantidade) || 0
        };
      })
    };

    return {
      mensagem: count + " material(is) registrado(s) na saída de estoque! Protocolo nº " + numeroProtocolo + ".",
      protocolo: protocolo
    };
  } catch (e) {
    throw new Error("Erro ao registrar saída: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

function obterProdutosCadastrados() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Estoque");
  if (!sheet) return [];
  var dados = sheet.getDataRange().getValues();
  var mapa = {};
  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    var codigo = String(row[1] || "").trim();
    var descricao = row[2];
    if (codigo) mapa[codigo] = descricao;
  }
  return Object.keys(mapa).map(function (cod) {
    return { codigo: cod, descricao: mapa[cod] };
  });
}

function obterMaiorNumeroCodigoProduto_(sheet) {
  var dados = sheet.getDataRange().getValues();
  var maior = 0;
  for (var i = 1; i < dados.length; i++) {
    var cod = String(dados[i][1] || "");
    var m = cod.match(/PRD-(\d+)/);
    if (m) {
      var n = parseInt(m[1], 10);
      if (n > maior) maior = n;
    }
  }
  return maior;
}

function relatorioMovimentacaoEstoqueFiltrado(tipo, dataInicio, dataFim, projetoFiltro) {
  var sheet = obterAbaFlexivel_("Estoque");
  if (!sheet) return empacotarResposta_([], "Aba 'Estoque' não encontrada.");
  var dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return empacotarResposta_([], "Aba '" + sheet.getName() + "' está vazia (só tem cabeçalho ou nada).");

  var col = obterMapaColunas_(dados[0]);
  var idxTipo = col["tipo"] !== undefined ? col["tipo"] : 3;
  var idxData = col["data"] !== undefined ? col["data"] : 0;
  var idxProjeto = col["projeto"] !== undefined ? col["projeto"] : 9;

  var totalTipo = 0, totalData = 0;
  var resultado = [];
  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    if (String(row[idxTipo]).trim().toLowerCase() !== String(tipo).trim().toLowerCase()) continue;
    totalTipo++;
    // CORREÇÃO: antes essa checagem rodava sempre, mesmo sem filtro de data
    // informado pelo usuário — e uma célula de data que não conseguisse ser
    // interpretada excluía a linha inteira do relatório. Agora só filtramos
    // por data quando "De" ou "Até" forem realmente preenchidos.
    if ((dataInicio || dataFim) && !dataDentroDoIntervalo_(row[idxData], dataInicio, dataFim)) continue;
    totalData++;
    if (projetoFiltro && projetoFiltro !== "" && String(row[idxProjeto]).trim().toLowerCase() !== projetoFiltro.trim().toLowerCase()) continue;
    resultado.push(row);
  }
  return empacotarResposta_(
    resultado,
    (dados.length - 1) + " linha(s) na planilha, " + totalTipo + " do tipo '" + tipo + "', " + totalData + " dentro do período, " + resultado.length + " após filtro de projeto."
  );
}

function relatorioPorProjetoFiltrado(dataInicio, dataFim, projetoFiltro) {
  var sheet = obterAbaFlexivel_("Estoque");
  if (!sheet) return empacotarResposta_([], "Aba 'Estoque' não encontrada.");
  var dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return empacotarResposta_([], "Aba '" + sheet.getName() + "' está vazia (só tem cabeçalho ou nada).");

  var col = obterMapaColunas_(dados[0]);
  var idxData = col["data"] !== undefined ? col["data"] : 0;
  var idxProjeto = col["projeto"] !== undefined ? col["projeto"] : 9;

  var totalData = 0;
  var resultado = [];
  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    if ((dataInicio || dataFim) && !dataDentroDoIntervalo_(row[idxData], dataInicio, dataFim)) continue;
    totalData++;
    if (projetoFiltro && projetoFiltro !== "" && String(row[idxProjeto]).trim().toLowerCase() !== projetoFiltro.trim().toLowerCase()) continue;
    resultado.push(row);
  }
  return empacotarResposta_(
    resultado,
    (dados.length - 1) + " linha(s) na planilha, " + totalData + " dentro do período, " + resultado.length + " após filtro de projeto."
  );
}

// ==============================================
// RELATÓRIO: Pedidos de Compra (lista completa do Kanban com status/coluna)
// ==============================================
function obterRelatorioPedidosCompra() {
  var sheet = obterAbaFlexivel_("Kanban");
  if (!sheet) return empacotarResposta_([], "Aba 'Kanban' não encontrada.");
  var dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return empacotarResposta_([], "Aba '" + sheet.getName() + "' está vazia (nenhuma solicitação cadastrada).");

  var col = obterMapaColunas_(dados[0]);
  var idx = {
    id: col["id"] !== undefined ? col["id"] : 0,
    titulo: col["título"] !== undefined ? col["título"] : (col["titulo"] !== undefined ? col["titulo"] : 1),
    solicitante: col["solicitante"] !== undefined ? col["solicitante"] : 2,
    setor: col["setor"] !== undefined ? col["setor"] : 3,
    etapa: col["etapa"] !== undefined ? col["etapa"] : 4,
    prioridade: col["prioridade"] !== undefined ? col["prioridade"] : 5,
    valor: col["valor"] !== undefined ? col["valor"] : 6,
    data: col["data"] !== undefined ? col["data"] : 7,
    projeto: col["projeto"] !== undefined ? col["projeto"] : 9,
    prazo: col["prazo entrega"] !== undefined ? col["prazo entrega"] : 10
  };

  var resultado = [];
  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    if (!row[idx.id]) continue;
    var tituloCompleto = String(row[idx.titulo] || "");
    var tituloPrincipal = tituloCompleto.split(" | Itens: ")[0];
    resultado.push({
      id: row[idx.id],
      titulo: tituloPrincipal,
      solicitante: row[idx.solicitante] || "",
      setor: row[idx.setor] || "",
      projeto: row[idx.projeto] || "",
      valor: parseFloat(row[idx.valor]) || 0,
      data: row[idx.data] || "",
      prazo: row[idx.prazo] || "",
      etapa: row[idx.etapa] || ""
    });
  }
  return empacotarResposta_(resultado, resultado.length + " pedido(s) encontrado(s) no Kanban.");
}

// CORREÇÃO: datas digitadas manualmente na planilha (ex: "19/01/2026" como texto,
// em vez de um objeto Date real) faziam "new Date(dataCelula)" retornar Invalid Date,
// o que derrubava o try/catch e excluía a linha do relatório inteiro. Isso fazia os
// relatórios de Entradas/Saídas/Projeto parecerem "não gerar nada" mesmo com dados
// corretos na planilha. Agora tentamos várias interpretações antes de desistir.
function parseDataFlexivel_(valor) {
  if (!valor) return null;
  if (valor instanceof Date) {
    return isNaN(valor.getTime()) ? null : valor;
  }
  var str = String(valor).trim();
  if (!str) return null;

  // Tenta primeiro como o próprio JS entende (ISO, objetos serializados, etc.)
  var d = new Date(str);
  if (!isNaN(d.getTime())) return d;

  // Tenta formato brasileiro dd/MM/yyyy ou dd-MM-yyyy
  var m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) {
    var dia = parseInt(m[1], 10), mes = parseInt(m[2], 10) - 1, ano = parseInt(m[3], 10);
    var d2 = new Date(ano, mes, dia);
    if (!isNaN(d2.getTime())) return d2;
  }
  return null;
}

function dataDentroDoIntervalo_(dataCelula, dataInicio, dataFim) {
  var data = parseDataFlexivel_(dataCelula);
  if (!data) return false;
  var tz = Session.getScriptTimeZone();
  var dataCelulaStr;
  try {
    dataCelulaStr = Utilities.formatDate(data, tz, "yyyy-MM-dd");
  } catch (e) {
    return false;
  }
  if (dataInicio && dataCelulaStr < dataInicio) return false;
  if (dataFim && dataCelulaStr > dataFim) return false;
  return true;
}

// Mapa "Nome do Cabeçalho" -> índice de coluna (0-based), evitando índices fixos
// que quebram silenciosamente quando alguém insere/reordena uma coluna na planilha.
function obterMapaColunas_(cabecalho) {
  var mapa = {};
  (cabecalho || []).forEach(function (nome, idx) {
    mapa[String(nome || "").trim().toLowerCase()] = idx;
  });
  return mapa;
}

// Converte qualquer objeto Date dentro de uma matriz 2D para texto ISO.
// Usado antes de devolver dados ao navegador (ver empacotarResposta_).
function sanitizarDatas_(dados) {
  if (!dados) return dados;
  return dados.map(function (linha) {
    if (!Array.isArray(linha)) return (linha instanceof Date) ? linha.toISOString() : linha;
    return linha.map(function (celula) {
      return (celula instanceof Date) ? celula.toISOString() : celula;
    });
  });
}

// CORREÇÃO IMPORTANTE: devolvemos sempre uma STRING (JSON.stringify) em vez de
// um objeto/array direto para o front-end. Isso evita um problema conhecido do
// Apps Script em que objetos Date aninhados dentro de arrays dentro de um
// objeto às vezes falham silenciosamente ao atravessar a ponte
// google.script.run (foi exatamente isso que fazia o Kanban voltar vazio).
// Uma string sempre atravessa a ponte cliente/servidor sem surpresas; o
// front-end faz JSON.parse() do resultado.
function empacotarResposta_(linhas, debug) {
  return JSON.stringify({ linhas: sanitizarDatas_(linhas), debug: debug });
}