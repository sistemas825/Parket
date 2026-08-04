/**
 * EXCLUSIVE — GESTÃO DE SUPRIMENTOS
 * Backend (Google Apps Script)
 *
 * Abas obrigatórias na planilha:
 *
 * "Itens"
 *  A:ID  B:Descricao  C:Serie  D:Marca  E:Valor  F:Qtd  G:Status
 *  H:Novo  I:FuncionarioID  J:FuncionarioNome  K:Termo  L:Empresa
 *  M:DataMovimento  N:Observacao
 *  Status válidos: "estoque" | "uso" | "manutencao" | "baixado"
 *
 * "Funcionarios"
 *  A:ID  B:Nome  C:Empresa  D:CPF  E:DataNascimento  F:Telefone
 *
 * "Movimentacoes" (log de auditoria)
 *  A:Data  B:Tipo  C:ItemID  D:Descricao  E:FuncionarioNome  F:Termo  G:Detalhes
 */

const ABA_ITENS = 'Itens';
const ABA_FUNCIONARIOS = 'Funcionarios';
const ABA_MOVIMENTACOES = 'Movimentacoes';
const STATUS_VALIDOS = ['estoque', 'uso', 'manutencao', 'baixado'];

const COL_ITEM = {
  ID: 1, DESCRICAO: 2, SERIE: 3, MARCA: 4, VALOR: 5, QTD: 6, STATUS: 7,
  NOVO: 8, FUNC_ID: 9, FUNC_NOME: 10, TERMO: 11, EMPRESA: 12, DATA_MOV: 13, OBS: 14
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Exclusive - Gestão de Suprimentos')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function aba_(nome) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!sheet) throw new Error('A aba "' + nome + '" não foi encontrada na planilha.');
  return sheet;
}

function hojeFormatado_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT-3', 'dd/MM/yyyy');
}

/**
 * Converte qualquer valor de célula em algo seguro para enviar ao navegador.
 * O Google Sheets às vezes converte texto de data (ex.: "10/11/2025") em um
 * objeto Date de verdade quando o arquivo é importado — e um objeto Date
 * dentro do retorno de uma função do Apps Script pode quebrar o
 * google.script.run silenciosamente (o navegador recebe "null").
 * Por isso, toda célula passa por aqui antes de sair do servidor.
 */
function celulaSegura_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return '';
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'GMT-3', 'dd/MM/yyyy');
  }
  return v;
}

/* =====================  LEITURA  ===================== */

function obterDados() {
  return {
    itens: obterItens_(),
    funcionarios: obterFuncionarios_()
  };
}

function obterItens_() {
  const sheet = aba_(ABA_ITENS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();

  return data
    .filter(function (row) { return row[0] !== '' && row[0] !== null; })
    .map(function (row) {
      return {
        id: String(celulaSegura_(row[0])),
        descricao: String(celulaSegura_(row[1])),
        serie: String(celulaSegura_(row[2])),
        marca: String(celulaSegura_(row[3])),
        valor: Number(row[4]) || 0,
        qtd: Number(row[5]) || 1,
        status: String(celulaSegura_(row[6])) || 'estoque',
        novo: !!row[7],
        funcionarioId: String(celulaSegura_(row[8])),
        funcionarioNome: String(celulaSegura_(row[9])),
        termo: String(celulaSegura_(row[10])),
        empresa: String(celulaSegura_(row[11])),
        dataMovimento: String(celulaSegura_(row[12])),
        observacao: String(celulaSegura_(row[13]))
      };
    });
}

function obterFuncionarios_() {
  const sheet = aba_(ABA_FUNCIONARIOS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  return data
    .filter(function (row) { return row[0] !== '' && row[0] !== null; })
    .map(function (row) {
      return {
        id: String(celulaSegura_(row[0])),
        nome: String(celulaSegura_(row[1])),
        empresa: String(celulaSegura_(row[2])),
        cpf: String(celulaSegura_(row[3])),
        nascimento: String(celulaSegura_(row[4])),
        telefone: String(celulaSegura_(row[5]))
      };
    });
}

/* =====================  HELPERS INTERNOS  ===================== */

function localizarLinhaItem_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function registrarMovimentacao_(tipo, item, funcionarioNome, termo, detalhes) {
  const sheet = aba_(ABA_MOVIMENTACOES);
  sheet.appendRow([hojeFormatado_(), tipo, item.id || '', item.descricao || '', funcionarioNome || '', termo || '', detalhes || '']);
}

/* =====================  ESCRITA  ===================== */

/** ENTRADA — adiciona novo item, entra direto em Estoque */
function adicionarFerramenta(item) {
  if (!item || !item.descricao || String(item.descricao).trim() === '') {
    throw new Error('A descrição do item é obrigatória.');
  }
  const sheet = aba_(ABA_ITENS);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const id = 'IT-' + new Date().getTime();
    sheet.appendRow([
      id,
      String(item.descricao).trim(),
      item.serie ? String(item.serie).trim() : '',
      item.marca ? String(item.marca).trim() : '',
      Number(item.valor) || 0,
      Number(item.qtd) || 1,
      'estoque',
      'SIM',
      '', '', '', '', '', ''
    ]);
    registrarMovimentacao_('entrada', { id: id, descricao: item.descricao }, '', '', 'Novo item cadastrado no estoque');
    return { success: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

/**
 * EDITAR ITEM — corrige dados cadastrais (descrição, marca, série, valor, qtd).
 * Não mexe em status, funcionário ou termo — só os dados do próprio item.
 */
function editarItem(itemId, dados) {
  if (!dados || !dados.descricao || String(dados.descricao).trim() === '') {
    throw new Error('A descrição do item é obrigatória.');
  }
  const sheet = aba_(ABA_ITENS);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const linha = localizarLinhaItem_(sheet, itemId);
    if (linha === -1) throw new Error('Item não encontrado.');

    sheet.getRange(linha, COL_ITEM.DESCRICAO).setValue(String(dados.descricao).trim());
    sheet.getRange(linha, COL_ITEM.SERIE).setValue(dados.serie ? String(dados.serie).trim() : '');
    sheet.getRange(linha, COL_ITEM.MARCA).setValue(dados.marca ? String(dados.marca).trim() : '');
    sheet.getRange(linha, COL_ITEM.VALOR).setValue(Number(dados.valor) || 0);
    sheet.getRange(linha, COL_ITEM.QTD).setValue(Number(dados.qtd) || 1);

    registrarMovimentacao_('edicao', { id: itemId, descricao: dados.descricao }, '', '', 'Dados cadastrais do item atualizados');
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

/** EMPRÉSTIMO — move de Estoque para Em Uso, vinculando um funcionário */
function emprestarItem(itemId, funcionarioId, termo) {
  if (!funcionarioId) throw new Error('Selecione um funcionário para o empréstimo.');

  const sheetItens = aba_(ABA_ITENS);
  const sheetFunc = aba_(ABA_FUNCIONARIOS);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const linha = localizarLinhaItem_(sheetItens, itemId);
    if (linha === -1) throw new Error('Item não encontrado.');

    const funcs = obterFuncionarios_();
    const func = funcs.filter(function (f) { return String(f.id) === String(funcionarioId); })[0];
    if (!func) throw new Error('Funcionário não encontrado.');

    const descricao = sheetItens.getRange(linha, COL_ITEM.DESCRICAO).getValue();

    sheetItens.getRange(linha, COL_ITEM.STATUS).setValue('uso');
    sheetItens.getRange(linha, COL_ITEM.FUNC_ID).setValue(func.id);
    sheetItens.getRange(linha, COL_ITEM.FUNC_NOME).setValue(func.nome);
    sheetItens.getRange(linha, COL_ITEM.TERMO).setValue(termo || '');
    sheetItens.getRange(linha, COL_ITEM.EMPRESA).setValue(func.empresa || '');
    sheetItens.getRange(linha, COL_ITEM.DATA_MOV).setValue(hojeFormatado_());

    registrarMovimentacao_('emprestimo', { id: itemId, descricao: descricao }, func.nome, termo, 'Empréstimo registrado');
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

/** DEVOLUÇÃO — move de Em Uso de volta para Estoque */
function devolverItem(itemId) {
  const sheet = aba_(ABA_ITENS);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const linha = localizarLinhaItem_(sheet, itemId);
    if (linha === -1) throw new Error('Item não encontrado.');

    const descricao = sheet.getRange(linha, COL_ITEM.DESCRICAO).getValue();
    const funcNome = sheet.getRange(linha, COL_ITEM.FUNC_NOME).getValue();
    const termo = sheet.getRange(linha, COL_ITEM.TERMO).getValue();

    sheet.getRange(linha, COL_ITEM.STATUS).setValue('estoque');
    sheet.getRange(linha, COL_ITEM.FUNC_ID).setValue('');
    sheet.getRange(linha, COL_ITEM.FUNC_NOME).setValue('');
    sheet.getRange(linha, COL_ITEM.TERMO).setValue('');
    sheet.getRange(linha, COL_ITEM.EMPRESA).setValue('');
    sheet.getRange(linha, COL_ITEM.DATA_MOV).setValue('');

    registrarMovimentacao_('devolucao', { id: itemId, descricao: descricao }, funcNome, termo, 'Item devolvido ao estoque');
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

/** MANUTENÇÃO — move o item (de Estoque ou Em Uso) para Manutenção */
function moverParaManutencao(itemId, motivo) {
  const sheet = aba_(ABA_ITENS);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const linha = localizarLinhaItem_(sheet, itemId);
    if (linha === -1) throw new Error('Item não encontrado.');

    const descricao = sheet.getRange(linha, COL_ITEM.DESCRICAO).getValue();

    sheet.getRange(linha, COL_ITEM.STATUS).setValue('manutencao');
    sheet.getRange(linha, COL_ITEM.OBS).setValue(motivo || '');
    sheet.getRange(linha, COL_ITEM.DATA_MOV).setValue(hojeFormatado_());

    registrarMovimentacao_('manutencao', { id: itemId, descricao: descricao }, '', '', motivo || 'Enviado para manutenção');
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

/** CONCLUIR REPARO — volta da Manutenção para o Estoque */
function concluirReparo(itemId) {
  const sheet = aba_(ABA_ITENS);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const linha = localizarLinhaItem_(sheet, itemId);
    if (linha === -1) throw new Error('Item não encontrado.');

    const descricao = sheet.getRange(linha, COL_ITEM.DESCRICAO).getValue();

    sheet.getRange(linha, COL_ITEM.STATUS).setValue('estoque');
    sheet.getRange(linha, COL_ITEM.OBS).setValue('');
    sheet.getRange(linha, COL_ITEM.FUNC_ID).setValue('');
    sheet.getRange(linha, COL_ITEM.FUNC_NOME).setValue('');
    sheet.getRange(linha, COL_ITEM.TERMO).setValue('');
    sheet.getRange(linha, COL_ITEM.DATA_MOV).setValue('');

    registrarMovimentacao_('reparo_concluido', { id: itemId, descricao: descricao }, '', '', 'Reparo concluído, retornou ao estoque');
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

/** BAIXA — remove definitivamente do patrimônio ativo (extravio, roubo, descarte) */
function darBaixa(itemId, motivo) {
  if (!motivo || String(motivo).trim() === '') {
    throw new Error('Informe o motivo da baixa.');
  }
  const sheet = aba_(ABA_ITENS);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const linha = localizarLinhaItem_(sheet, itemId);
    if (linha === -1) throw new Error('Item não encontrado.');

    const descricao = sheet.getRange(linha, COL_ITEM.DESCRICAO).getValue();
    const funcNome = sheet.getRange(linha, COL_ITEM.FUNC_NOME).getValue();

    sheet.getRange(linha, COL_ITEM.STATUS).setValue('baixado');
    sheet.getRange(linha, COL_ITEM.OBS).setValue(motivo);
    sheet.getRange(linha, COL_ITEM.DATA_MOV).setValue(hojeFormatado_());

    registrarMovimentacao_('baixa', { id: itemId, descricao: descricao }, funcNome, '', motivo);
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

/* =====================================================================
 * GERAÇÃO DE PDF — TERMOS DE EMPRÉSTIMO E DEVOLUÇÃO
 * Layout fiel aos modelos "termo_emprestimo_final.pdf" e
 * "termo_devolucao_final.pdf" fornecidos pelo usuário em 2026-08-03.
 * ===================================================================== */

// Paleta Parket — mesma paleta do ERP de Compras (nogueira + âmbar + creme).
// Layout dos termos refeito em 2026-08-03 para bater exatamente com os
// modelos "termo_emprestimo_final.pdf" / "termo_devolucao_final.pdf".
const COR_NAVY = '#332720';
const COR_DOURADO = '#A9722E';
const COR_CINZA_CLARO = '#F3ECE0';
const COR_BORDA = '#E3D9C9';
const COR_CREME = '#FAF6EE';
const COR_CINZA_TEXTO = '#8A7B6C';
const COR_PRETO_QUENTE = '#2B2420';
const COR_AMBAR_ESCURO = '#8C5D24';
const FONTE_DISPLAY = 'Fraunces';
const FONTE_MONO = 'IBM Plex Mono';
const EMPRESA_NOME = 'EXCLUSIVE INSTALAÇÕES DE MOBILIÁRIOS LTDA';
const EMPRESA_CNPJ = 'CNPJ: 51.992.934/0001-98';
const LARGURA_UTIL = 522; // Letter 612pt - margens 45+45

function dataPorExtenso_() {
  const meses = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
  const d = new Date();
  return d.getDate() + ' de ' + meses[d.getMonth()] + ' de ' + d.getFullYear();
}

/** Gera um número sequencial "N/AAAA" persistido entre chamadas, por tipo de documento */
function proximoNumeroDocumento_(prefixoChave) {
  const props = PropertiesService.getScriptProperties();
  const ano = new Date().getFullYear();
  const chave = prefixoChave + '_' + ano;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const atual = Number(props.getProperty(chave) || '0') + 1;
    props.setProperty(chave, String(atual));
    return atual + '/' + ano;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Gera o próximo número de Termo de Empréstimo automaticamente.
 * Sequência contínua (não reinicia por ano) começando em 200 —
 * o primeiro número gerado será "200/AAAA".
 */
function proximoNumeroTermoEmprestimo_() {
  const props = PropertiesService.getScriptProperties();
  const chave = 'TERMO_EMPRESTIMO_SEQ';
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let atual = Number(props.getProperty(chave));
    if (!atual || atual < 199) atual = 199;
    atual += 1;
    props.setProperty(chave, String(atual));
    return atual + '/' + new Date().getFullYear();
  } finally {
    lock.releaseLock();
  }
}

/** Exposta ao front-end: retorna o próximo número de termo antes de confirmar o empréstimo */
function gerarProximoNumeroTermo() {
  return proximoNumeroTermoEmprestimo_();
}

/** Cria uma barrinha fina e cheia (usada como linha divisória colorida). */
function barraFina_(body, corFundo, larguraPt) {
  const barra = body.appendTable([[' ']]);
  barra.setBorderWidth(0);
  barra.setColumnWidth(0, larguraPt);
  const cel = barra.getCell(0, 0);
  cel.setBackgroundColor(corFundo);
  cel.setPaddingTop(0.3).setPaddingBottom(0.3).setPaddingLeft(0).setPaddingRight(0);
  cel.editAsText().setFontSize(1);
  cel.getChild(0).asParagraph().setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(0.3);
  return barra;
}

/**
 * Espaçador com altura controlada, em pt. Um parágrafo vazio comum herda o
 * tamanho de fonte do texto anterior (às vezes 11-13pt), então cada "linha
 * em branco" podia adicionar bem mais espaço do que parecia — isso empurrava
 * o termo pra uma segunda página. Fixando a fonte bem pequena aqui, a altura
 * fica sob controle.
 */
function espacador_(body, alturaPt) {
  const p = body.appendParagraph('');
  p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
  p.editAsText().setFontSize(Math.max(1, alturaPt));
  return p;
}

/** Cabeçalho: nome da empresa + CNPJ em texto simples, linha fina, e barra escura com o título. */
function cabecalhoDocumento_(body, tituloDocumento) {
  const pNome = body.appendParagraph(EMPRESA_NOME);
  pNome.setSpacingBefore(0).setSpacingAfter(2);
  pNome.editAsText().setBold(true).setFontSize(18).setForegroundColor(COR_NAVY).setFontFamily(FONTE_DISPLAY);

  const pCnpj = body.appendParagraph(EMPRESA_CNPJ);
  pCnpj.setSpacingBefore(0).setSpacingAfter(8);
  pCnpj.editAsText().setFontSize(9).setForegroundColor(COR_CINZA_TEXTO).setFontFamily(FONTE_MONO).setBold(false);

  barraFina_(body, COR_PRETO_QUENTE, LARGURA_UTIL);
  espacador_(body, 5);

  const barraTitulo = body.appendTable([['']]);
  barraTitulo.setBorderWidth(0);
  const celTitulo = barraTitulo.getCell(0, 0);
  celTitulo.setBackgroundColor(COR_NAVY);
  celTitulo.setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(14).setPaddingRight(14);
  celTitulo.clear();
  const pTitulo = celTitulo.appendParagraph(tituloDocumento);
  pTitulo.setSpacingBefore(0).setSpacingAfter(0);
  pTitulo.editAsText().setBold(true).setFontSize(13).setForegroundColor('#FFFFFF').setFontFamily(FONTE_DISPLAY);

  espacador_(body, 5);
}

/** Tabela de identificação (funcionário / documento / número de controle). */
function tabelaIdentificacao_(body, colunas) {
  // colunas: [{ rotulo, valor }, ...] — sempre 3: Funcionário / Documento / Controle.
  // O nome do funcionário é o campo mais variável (pode ser bem longo), por
  // isso ganha a maior largura — evita quebrar em duas linhas na célula.
  const larguras = [250, 150, 122]; // soma 522pt
  const rotulos = colunas.map(function (c) { return c.rotulo; });
  const valores = colunas.map(function (c) { return c.valor || '-'; });
  const table = body.appendTable([rotulos, valores]);
  table.setBorderColor(COR_BORDA);
  table.setBorderWidth(1);
  for (let c = 0; c < colunas.length; c++) {
    table.setColumnWidth(c, larguras[c] || Math.floor(LARGURA_UTIL / colunas.length));
    const cRotulo = table.getCell(0, c);
    cRotulo.setPaddingTop(6).setPaddingBottom(3).setPaddingLeft(10).setPaddingRight(10);
    cRotulo.getChild(0).asParagraph().setSpacingBefore(0).setSpacingAfter(0);
    cRotulo.editAsText().setFontSize(8).setForegroundColor(COR_CINZA_TEXTO).setBold(false).setFontFamily(FONTE_MONO);

    const cValor = table.getCell(1, c);
    cValor.setPaddingTop(2).setPaddingBottom(8).setPaddingLeft(10).setPaddingRight(10);
    cValor.getChild(0).asParagraph().setSpacingBefore(0).setSpacingAfter(0);
    // Nomes muito compridos encolhem um pouco a fonte pra nunca quebrar linha.
    const textoValor = String(valores[c] || '');
    let tamanhoFonte = 10;
    if (textoValor.length > 34) tamanhoFonte = 9;
    if (textoValor.length > 40) tamanhoFonte = 8;
    cValor.editAsText().setFontSize(tamanhoFonte).setBold(true).setForegroundColor(COR_PRETO_QUENTE).setFontFamily(FONTE_MONO);
  }
  espacador_(body, 5);
  return table;
}

/**
 * Caixa com borda fina contendo o parágrafo de abertura, com o nome da
 * empresa em negrito no meio da frase (textoAntes + negrito + textoDepois).
 */
/** Caixa com barra colorida grossa na borda esquerda + borda fina ao redor. */
function caixaComAcento_(body, corAcento) {
  const table = body.appendTable([['', '']]);
  table.setBorderColor(COR_BORDA);
  table.setBorderWidth(1);
  table.setColumnWidth(0, 4);
  table.setColumnWidth(1, LARGURA_UTIL - 4);

  const cAcento = table.getCell(0, 0);
  cAcento.setBackgroundColor(corAcento);
  cAcento.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
  cAcento.clear();
  cAcento.appendParagraph(' ').setSpacingBefore(0).setSpacingAfter(0);

  const cConteudo = table.getCell(0, 1);
  cConteudo.setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(12).setPaddingRight(12);
  cConteudo.clear();
  return cConteudo;
}

function caixaIntro_(body, textoAntes, negrito, textoDepois) {
  const cConteudo = caixaComAcento_(body, COR_NAVY);
  const textoCompleto = textoAntes + negrito + textoDepois;
  const p = cConteudo.appendParagraph(textoCompleto);
  p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY).setSpacingBefore(0).setSpacingAfter(0);
  const t = p.editAsText();
  t.setFontSize(10.5).setForegroundColor(COR_PRETO_QUENTE).setFontFamily('Arial').setBold(false);
  t.setBold(textoAntes.length, textoAntes.length + negrito.length - 1, true);
  espacador_(body, 5);
  return cConteudo;
}

/** Título de seção com uma barrinha de destaque à esquerda. */
function tituloSecaoComBarra_(body, texto) {
  const tab = body.appendTable([['', texto]]);
  tab.setBorderWidth(0);
  tab.setColumnWidth(0, 4);
  tab.setColumnWidth(1, LARGURA_UTIL - 4);
  const cBarra = tab.getCell(0, 0);
  cBarra.setBackgroundColor(COR_NAVY);
  cBarra.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
  cBarra.clear();
  cBarra.appendParagraph(' ').setSpacingBefore(0).setSpacingAfter(0);
  const cTexto = tab.getCell(0, 1);
  cTexto.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(8).setPaddingRight(0);
  const pTexto = cTexto.getChild(0).asParagraph().setSpacingBefore(0).setSpacingAfter(0);
  cTexto.editAsText().setBold(true).setFontSize(11.5).setForegroundColor(COR_NAVY).setFontFamily(FONTE_DISPLAY);
  espacador_(body, 5);
}

/** Lista de cláusulas numeradas (selo escuro com o número + texto), sem grade nem zebra. */
function tabelaClausulas_(body, tituloSecao, clausulas) {
  tituloSecaoComBarra_(body, tituloSecao);

  const linhas = clausulas.map(function (texto, idx) { return [(idx + 1) + '.', texto]; });
  const table = body.appendTable(linhas);
  table.setBorderWidth(0);
  table.setColumnWidth(0, 22);
  table.setColumnWidth(1, LARGURA_UTIL - 22);
  for (let r = 0; r < linhas.length; r++) {
    const cNum = table.getCell(r, 0);
    cNum.setPaddingTop(2).setPaddingBottom(6).setPaddingLeft(0).setPaddingRight(4);
    cNum.getChild(0).asParagraph().setSpacingBefore(0).setSpacingAfter(0);
    cNum.editAsText().setBold(true).setForegroundColor(COR_NAVY).setFontSize(10).setFontFamily('Arial');

    const cTexto = table.getCell(r, 1);
    cTexto.setPaddingTop(2).setPaddingBottom(6).setPaddingLeft(0).setPaddingRight(0);
    cTexto.getChild(0).asParagraph().setSpacingBefore(0).setSpacingAfter(0);
    cTexto.editAsText().setFontSize(10).setForegroundColor(COR_PRETO_QUENTE).setFontFamily('Arial').setBold(false);
  }
  espacador_(body, 5);
  return table;
}

function formatarMoedaBRL_(valor) {
  const n = Number(valor) || 0;
  const partes = n.toFixed(2).split('.');
  let inteiro = partes[0].replace('-', '');
  let comMilhar = '';
  while (inteiro.length > 3) {
    comMilhar = '.' + inteiro.slice(-3) + comMilhar;
    inteiro = inteiro.slice(0, -3);
  }
  comMilhar = inteiro + comMilhar;
  return (n < 0 ? '-R$ ' : 'R$ ') + comMilhar + ',' + partes[1];
}

/**
 * Tabela de ferramentas: ITEM (numerado) | DESCRIÇÃO | VALOR (rótulo variável),
 * seguida da linha de total (rótulo variável) — sem grade, igual ao modelo
 * aprovado. rotuloValor/rotuloTotal mudam entre empréstimo e devolução.
 */
function tabelaItens_(body, itensLista, rotuloValor, rotuloTotal) {
  if (!itensLista || itensLista.length === 0) return;

  const cabecalho = ['ITEM', 'DESCRIÇÃO DA FERRAMENTA / EQUIPAMENTO', rotuloValor];
  const larguras = [45, 295, 182]; // soma 522pt

  // ---- Cabeçalho ----
  const tabelaCab = body.appendTable([cabecalho]);
  tabelaCab.setBorderWidth(0);
  for (let c = 0; c < cabecalho.length; c++) {
    tabelaCab.setColumnWidth(c, larguras[c]);
    const cCab = tabelaCab.getCell(0, c);
    cCab.setPaddingTop(0).setPaddingBottom(3).setPaddingLeft(8).setPaddingRight(8);
    const pCab = cCab.getChild(0).asParagraph().setSpacingBefore(0).setSpacingAfter(0);
    cCab.editAsText().setBold(true).setFontSize(8.5).setForegroundColor(COR_PRETO_QUENTE).setFontFamily(FONTE_MONO);
    if (c === 2) pCab.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  }
  barraFina_(body, COR_BORDA, LARGURA_UTIL);

  // ---- Itens (numeração 01, 02... + listrado leve) ----
  const linhasItens = itensLista.map(function (it, idx) {
    return [String(idx + 1).padStart(2, '0'), it.descricao || '-', formatarMoedaBRL_(it.valor)];
  });
  const total = itensLista.reduce(function (s, it) { return s + (Number(it.valor) || 0); }, 0);

  const tabelaItensTbl = body.appendTable(linhasItens);
  tabelaItensTbl.setBorderWidth(0);
  for (let c = 0; c < cabecalho.length; c++) tabelaItensTbl.setColumnWidth(c, larguras[c]);

  for (let r = 0; r < linhasItens.length; r++) {
    const bgLinha = (r % 2 === 1) ? COR_CINZA_CLARO : '#FFFFFF';
    for (let c = 0; c < cabecalho.length; c++) {
      const cel = tabelaItensTbl.getCell(r, c);
      cel.setBackgroundColor(bgLinha);
      cel.setPaddingTop(5).setPaddingBottom(5).setPaddingLeft(8).setPaddingRight(8);
      const pCel = cel.getChild(0).asParagraph().setSpacingBefore(0).setSpacingAfter(0);
      cel.editAsText().setFontSize(9.5).setFontFamily(FONTE_MONO).setForegroundColor(COR_PRETO_QUENTE).setBold(false);
      if (c === 2) pCel.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    }
  }

  // ---- Linha de Total: tabela própria de 2 colunas (evita mesclar células,
  // que já causou desalinhamento antes) — larguras batem com as de cima. ----
  const totalTable = body.appendTable([['', '']]);
  totalTable.setBorderWidth(0);
  totalTable.setColumnWidth(0, larguras[0] + larguras[1]);
  totalTable.setColumnWidth(1, larguras[2]);

  const celTotalLabel = totalTable.getCell(0, 0);
  celTotalLabel.setBackgroundColor(COR_CINZA_CLARO);
  celTotalLabel.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(8).setPaddingRight(8);
  celTotalLabel.clear();
  const pTotalLabel = celTotalLabel.appendParagraph(rotuloTotal);
  pTotalLabel.setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingBefore(0).setSpacingAfter(0);
  pTotalLabel.editAsText().setBold(true).setFontSize(9.5).setForegroundColor(COR_PRETO_QUENTE).setFontFamily(FONTE_MONO);

  const celTotalValor = totalTable.getCell(0, 1);
  celTotalValor.setBackgroundColor(COR_CINZA_CLARO);
  celTotalValor.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(8).setPaddingRight(8);
  celTotalValor.clear();
  const pTotalValor = celTotalValor.appendParagraph(formatarMoedaBRL_(total));
  pTotalValor.setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingBefore(0).setSpacingAfter(0);
  pTotalValor.editAsText().setBold(true).setFontSize(10.5).setForegroundColor(COR_AMBAR_ESCURO).setFontFamily(FONTE_MONO);

  espacador_(body, 5);
  return tabelaItensTbl;
}

/** Caixa de declaração final, com o rótulo em negrito seguido do texto normal. */
function caixaDeclaracao_(body, tituloNegrito, texto) {
  const cConteudo = caixaComAcento_(body, COR_DOURADO);
  cConteudo.setBackgroundColor(COR_CREME);
  const textoCompleto = tituloNegrito + ' ' + texto;
  const p = cConteudo.appendParagraph(textoCompleto);
  p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY).setSpacingBefore(0).setSpacingAfter(0);
  const t = p.editAsText();
  t.setFontSize(10).setForegroundColor(COR_PRETO_QUENTE).setFontFamily('Arial').setBold(false);
  t.setBold(0, tituloNegrito.length - 1, true);
  espacador_(body, 5);
  return cConteudo;
}

/** Bloco de assinaturas — linha, nome e cargo centralizados em cada coluna. */
function blocoAssinatura_(body, nomeFuncionario) {
  const table = body.appendTable([['', '']]);
  table.setBorderWidth(0);

  const esquerda = table.getCell(0, 0);
  esquerda.clear();
  const pLinha1 = esquerda.appendParagraph('_______________________________');
  pLinha1.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0).setSpacingAfter(2);
  pLinha1.editAsText().setFontSize(10).setForegroundColor(COR_BORDA).setBold(false);
  const pNome = esquerda.appendParagraph((nomeFuncionario || '').toUpperCase());
  pNome.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0).setSpacingAfter(0);
  pNome.editAsText().setBold(true).setFontSize(10).setForegroundColor(COR_PRETO_QUENTE).setFontFamily('Arial');
  const pSub = esquerda.appendParagraph('Funcionário / Colaborador');
  pSub.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0).setSpacingAfter(0);
  pSub.editAsText().setFontSize(9).setForegroundColor(COR_CINZA_TEXTO).setFontFamily('Arial').setBold(false);

  const direita = table.getCell(0, 1);
  direita.clear();
  const pLinha2 = direita.appendParagraph('_______________________________');
  pLinha2.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0).setSpacingAfter(2);
  pLinha2.editAsText().setFontSize(10).setForegroundColor(COR_BORDA).setBold(false);
  const pEmpresa = direita.appendParagraph('EXCLUSIVE INSTALAÇÕES DE MOBILIÁRIOS');
  pEmpresa.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0).setSpacingAfter(0);
  pEmpresa.editAsText().setBold(true).setFontSize(10).setForegroundColor(COR_PRETO_QUENTE).setFontFamily('Arial');
  const pRh = direita.appendParagraph('Departamento de Recursos Humanos');
  pRh.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setSpacingBefore(0).setSpacingAfter(0);
  pRh.editAsText().setFontSize(9).setForegroundColor(COR_CINZA_TEXTO).setFontFamily('Arial').setBold(false);
}

function docParaPdfBase64_(doc) {
  doc.saveAndClose();
  const id = doc.getId();
  const pdfBlob = DriveApp.getFileById(id).getAs('application/pdf');
  DriveApp.getFileById(id).setTrashed(true);
  return Utilities.base64Encode(pdfBlob.getBytes());
}

/** TERMO DE EMPRÉSTIMO — chamar com o funcionário, o nº do termo e a lista de ferramentas selecionadas */
function gerarTermoEmprestimo(funcionarioId, termo, itensLista) {
  const func = obterFuncionarios_().filter(function (f) { return String(f.id) === String(funcionarioId); })[0];
  if (!func) throw new Error('Funcionário não encontrado para gerar o termo.');

  const doc = DocumentApp.create('Termo_Emprestimo_temp_' + new Date().getTime());
  const body = doc.getBody();
  body.setMarginTop(30).setMarginBottom(30).setMarginLeft(45).setMarginRight(45);

  cabecalhoDocumento_(body, 'TERMO DE RESPONSABILIDADE E EMPRÉSTIMO DE FERRAMENTAS');

  tabelaIdentificacao_(body, [
    { rotulo: 'FUNCIONÁRIO / COLABORADOR', valor: func.nome },
    { rotulo: 'DOCUMENTO (CPF)', valor: func.cpf },
    { rotulo: 'Nº CONTROLE', valor: termo || '-' }
  ]);

  caixaIntro_(body,
    'Eu, acima identificado, declaro que recebi da ',
    EMPRESA_NOME,
    ', a título de empréstimo para uso exclusivo no desempenho de minhas atividades profissionais, as ferramentas e equipamentos descritos na tabela abaixo.'
  );

  tabelaItens_(body, itensLista, 'VALOR DE REPOSIÇÃO', 'VALOR TOTAL SOB GUARDA');

  tabelaClausulas_(body, 'COMPROMISSOS E OBRIGAÇÕES DO COLABORADOR', [
    'Zelar pela correta utilização, guarda e conservação de todas as ferramentas sob minha responsabilidade, utilizando-as estritamente para os fins profissionais da empresa.',
    'Comunicar imediatamente ao setor responsável qualquer indício de defeito, dano, avaria ou irregularidade identificada no equipamento.',
    'Reconhecer que perdas, extravios, danos decorrentes de mau uso, negligência, roubo ou furto decorrentes de falta de cuidado na guarda serão de minha responsabilidade.',
    'Em caso de ocorrência das hipóteses previstas na cláusula anterior, autorizo o ressarcimento à empresa mediante desconto em folha ou acerto rescisório, conforme o Art. 462, §1º da CLT.',
    'Devolver todas as ferramentas recebidas em perfeitas condições de funcionamento ao término das atividades ou imediatamente quando solicitado pela empresa.'
  ]);

  caixaDeclaracao_(body, 'DECLARAÇÃO DE RECEBIMENTO:', 'Confirmo que recebi os itens acima relacionados em perfeito estado de funcionamento e conservação, estando ciente de todas as regras relativas à guarda e responsabilidade sob estes equipamentos.');

  const pData = body.appendParagraph('São Paulo, ' + dataPorExtenso_() + '.');
  pData.setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingBefore(0).setSpacingAfter(0);
  pData.editAsText().setFontSize(10).setForegroundColor(COR_CINZA_TEXTO).setFontFamily('Arial').setBold(false);
  espacador_(body, 14);

  blocoAssinatura_(body, func.nome);

  const base64 = docParaPdfBase64_(doc);
  return { filename: 'Termo_Emprestimo_' + func.nome.replace(/\s+/g, '_') + '.pdf', base64: base64 };
}

/**
 * TERMO DE DEVOLUÇÃO — grupos: [{ funcionarioNome, cpf, termo, itens }]
 * (itens: lista de {descricao, valor} devolvidos por esse funcionário/termo)
 * Gera um PDF com uma página por grupo (funcionário + termo original),
 * para cobrir devoluções em lote de pessoas/termos diferentes.
 */
function gerarTermoDevolucao(grupos) {
  if (!grupos || grupos.length === 0) throw new Error('Nenhum item para gerar termo de devolução.');

  const doc = DocumentApp.create('Termo_Devolucao_temp_' + new Date().getTime());
  const body = doc.getBody();
  body.setMarginTop(30).setMarginBottom(30).setMarginLeft(45).setMarginRight(45);

  grupos.forEach(function (grupo, idx) {
    cabecalhoDocumento_(body, 'TERMO DE DEVOLUÇÃO DE FERRAMENTAS E EQUIPAMENTOS');

    tabelaIdentificacao_(body, [
      { rotulo: 'FUNCIONÁRIO / COLABORADOR', valor: grupo.funcionarioNome },
      { rotulo: 'DOCUMENTO (CPF)', valor: grupo.cpf || '-' },
      { rotulo: 'CÓD. CONTROLE', valor: grupo.termo || '-' }
    ]);

    caixaIntro_(body,
      'Eu, acima identificado, declaro que estou devolvendo à ',
      EMPRESA_NOME,
      ' todas as ferramentas e equipamentos listados abaixo, anteriormente entregues para execução das minhas atividades profissionais.'
    );

    tabelaItens_(body, grupo.itens, 'VALOR DE AVALIAÇÃO', 'VALOR TOTAL EQUIPAMENTOS DEVOLVIDOS');

    tabelaClausulas_(body, 'DECLARAÇÕES E CONDIÇÕES GERAIS', [
      'Todas as ferramentas estão sendo devolvidas em perfeitas condições de funcionamento, salvo observações registradas no ato da devolução.',
      'Fui orientado a informar previamente qualquer defeito, dano, perda, extravio, roubo ou furto ocorrido durante o período de uso.',
      'Reconheço que eventuais danos por mau uso, perdas ou irregularidades já identificadas ou ainda a identificar, caso decorrentes de minha responsabilidade, poderão resultar em cobrança de ressarcimento conforme avaliação da empresa.',
      'Autorizo que valores referentes a ressarcimentos pendentes, caso existam, possam ser descontados dos pagamentos mensais ou, se aplicável, do acerto/rescisão contratual.',
      'Confirmo que todos os itens devolvidos foram conferidos em conjunto com o responsável da empresa no ato da entrega.'
    ]);

    caixaDeclaracao_(body, 'CIÊNCIA E CONCORDÂNCIA:', 'Declaro estar ciente de todas as condições acima e que este termo possui validade legal para fins de comprovação da devolução dos equipamentos.');

    const pData = body.appendParagraph('São Paulo, ' + dataPorExtenso_() + '.');
    pData.setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingBefore(0).setSpacingAfter(0);
    pData.editAsText().setFontSize(10).setForegroundColor(COR_CINZA_TEXTO).setFontFamily('Arial').setBold(false);
    espacador_(body, 14);

    blocoAssinatura_(body, grupo.funcionarioNome);

    if (idx < grupos.length - 1) {
      body.appendPageBreak();
    }
  });

  const base64 = docParaPdfBase64_(doc);
  return { filename: 'Termo_Devolucao.pdf', base64: base64 };
}

/* =====================================================================
 * ETAPA 2 — FUNCIONÁRIOS (CRUD)
 * ===================================================================== */

function adicionarFuncionario(dados) {
  if (!dados || !dados.nome || String(dados.nome).trim() === '') {
    throw new Error('O nome do funcionário é obrigatório.');
  }
  const sheet = aba_(ABA_FUNCIONARIOS);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const id = 'FN-' + new Date().getTime();
    sheet.appendRow([
      id,
      String(dados.nome).trim(),
      dados.empresa ? String(dados.empresa).trim() : '',
      dados.cpf ? String(dados.cpf).trim() : '',
      dados.nascimento ? String(dados.nascimento).trim() : '',
      dados.telefone ? String(dados.telefone).trim() : ''
    ]);
    return { success: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function editarFuncionario(id, dados) {
  if (!dados || !dados.nome || String(dados.nome).trim() === '') {
    throw new Error('O nome do funcionário é obrigatório.');
  }
  const sheet = aba_(ABA_FUNCIONARIOS);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Funcionário não encontrado.');
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        const linha = i + 2;
        const nomeAntigo = sheet.getRange(linha, 2).getValue();
        sheet.getRange(linha, 2).setValue(String(dados.nome).trim());
        sheet.getRange(linha, 3).setValue(dados.empresa ? String(dados.empresa).trim() : '');
        sheet.getRange(linha, 4).setValue(dados.cpf ? String(dados.cpf).trim() : '');
        sheet.getRange(linha, 5).setValue(dados.nascimento ? String(dados.nascimento).trim() : '');
        sheet.getRange(linha, 6).setValue(dados.telefone ? String(dados.telefone).trim() : '');

        // Mantém o nome do funcionário sincronizado nos itens que ele tem em uso no momento
        if (String(nomeAntigo) !== String(dados.nome).trim()) {
          const itensSheet = aba_(ABA_ITENS);
          const lastRowItens = itensSheet.getLastRow();
          if (lastRowItens >= 2) {
            const funcIds = itensSheet.getRange(2, COL_ITEM.FUNC_ID, lastRowItens - 1, 1).getValues();
            for (let j = 0; j < funcIds.length; j++) {
              if (String(funcIds[j][0]) === String(id)) {
                itensSheet.getRange(j + 2, COL_ITEM.FUNC_NOME).setValue(String(dados.nome).trim());
              }
            }
          }
        }
        return { success: true };
      }
    }
    throw new Error('Funcionário não encontrado.');
  } finally {
    lock.releaseLock();
  }
}

/** Remove o funcionário da lista de seleção. NÃO apaga o histórico de movimentações já registrado. */
function removerFuncionario(id) {
  const sheet = aba_(ABA_FUNCIONARIOS);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Funcionário não encontrado.');
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    throw new Error('Funcionário não encontrado.');
  } finally {
    lock.releaseLock();
  }
}

/* =====================================================================
 * ETAPA 2 — DADOS E EXCEL (exportação CSV, backup e restauração)
 * ===================================================================== */

function celulaCSV_(v) {
  const s = (v === null || v === undefined) ? '' : String(celulaSegura_(v));
  if (/[",\n;]/.test(s)) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function linhasParaCSV_(linhas) {
  return linhas.map(function (linha) {
    return linha.map(celulaCSV_).join(',');
  }).join('\r\n');
}

/** Inventário atual (Itens), pronto para abrir no Excel */
function exportarInventarioCSV() {
  const itens = obterItens_();
  const cabecalho = ['ID', 'Descricao', 'Serie', 'Marca', 'Valor', 'Qtd', 'Status', 'Novo', 'FuncionarioNome', 'Termo', 'Empresa', 'DataMovimento', 'Observacao'];
  const linhas = [cabecalho].concat(itens.map(function (it) {
    return [it.id, it.descricao, it.serie, it.marca, it.valor, it.qtd, it.status, it.novo ? 'SIM' : '', it.funcionarioNome, it.termo, it.empresa, it.dataMovimento, it.observacao];
  }));
  return linhasParaCSV_(linhas);
}

/** Auditoria completa de movimentações (log) */
function exportarMovimentacoesCSV() {
  const sheet = aba_(ABA_MOVIMENTACOES);
  const lastRow = sheet.getLastRow();
  const cabecalho = ['Data', 'Tipo', 'ItemID', 'Descricao', 'FuncionarioNome', 'Termo', 'Detalhes'];
  if (lastRow < 2) return linhasParaCSV_([cabecalho]);
  const dados = sheet.getRange(2, 1, lastRow - 1, 7).getValues().map(function (row) {
    return row.map(celulaSegura_);
  });
  return linhasParaCSV_([cabecalho].concat(dados));
}

/** Backup completo em JSON — Itens + Funcionarios + Movimentacoes */
function obterBackupCompleto() {
  const sheet = aba_(ABA_MOVIMENTACOES);
  const lastRow = sheet.getLastRow();
  const movimentacoes = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, 7).getValues().map(function (row) {
    return row.map(celulaSegura_);
  });

  return JSON.stringify({
    versao: 1,
    geradoEm: hojeFormatado_(),
    itens: obterItens_(),
    funcionarios: obterFuncionarios_(),
    movimentacoes: movimentacoes
  });
}

/** Restaura a planilha inteira a partir de um backup gerado por obterBackupCompleto() */
function restaurarBackup(jsonTexto) {
  let dados;
  try {
    dados = JSON.parse(jsonTexto);
  } catch (e) {
    throw new Error('Arquivo de backup inválido (JSON malformado).');
  }
  if (!dados || !dados.itens || !dados.funcionarios) {
    throw new Error('Arquivo de backup em formato inesperado.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheetItens = aba_(ABA_ITENS);
    limparAbaMantendoCabecalho_(sheetItens);
    if (dados.itens.length > 0) {
      const linhasItens = dados.itens.map(function (it) {
        return [it.id, it.descricao, it.serie, it.marca, it.valor, it.qtd, it.status, it.novo ? 'SIM' : '', it.funcionarioId, it.funcionarioNome, it.termo, it.empresa, it.dataMovimento, it.observacao];
      });
      sheetItens.getRange(2, 1, linhasItens.length, 14).setValues(linhasItens);
    }

    const sheetFunc = aba_(ABA_FUNCIONARIOS);
    limparAbaMantendoCabecalho_(sheetFunc);
    if (dados.funcionarios.length > 0) {
      const linhasFunc = dados.funcionarios.map(function (f) {
        return [f.id, f.nome, f.empresa, f.cpf, f.nascimento, f.telefone];
      });
      sheetFunc.getRange(2, 1, linhasFunc.length, 6).setValues(linhasFunc);
    }

    if (dados.movimentacoes) {
      const sheetMov = aba_(ABA_MOVIMENTACOES);
      limparAbaMantendoCabecalho_(sheetMov);
      if (dados.movimentacoes.length > 0) {
        sheetMov.getRange(2, 1, dados.movimentacoes.length, 7).setValues(dados.movimentacoes);
      }
    }

    return { success: true, itens: dados.itens.length, funcionarios: dados.funcionarios.length };
  } finally {
    lock.releaseLock();
  }
}

function limparAbaMantendoCabecalho_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
}

/** Apaga todos os dados operacionais (mantém cabeçalhos). Ação destrutiva e irreversível. */
function limparTodoSistema(confirmacao) {
  if (confirmacao !== 'CONFIRMAR') {
    throw new Error('Confirmação inválida.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    limparAbaMantendoCabecalho_(aba_(ABA_ITENS));
    limparAbaMantendoCabecalho_(aba_(ABA_FUNCIONARIOS));
    limparAbaMantendoCabecalho_(aba_(ABA_MOVIMENTACOES));
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}