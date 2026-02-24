/**
 * Núcleo Asset - Backend Google Apps Script Full Stack
 * Arquitetura de Dados | Regras de Negócio - Revisado conforme Requisitos Bancários
 */

const PLANILHA_ID = '1mM8zs3zUUd2V2HCItqjWCf2W2SN4vohza6moytrPdPI';

function doGet(e) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  criarAbasBanco(ss);
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Bem-vindo ao Gerenciador de Ativos - Núcleo Asset');
}

function criarAbasBanco(ss) {
  const abas = [
    {nome:'DB_Cadastro', cols:['ID','Nome','Estoque_Minimo','Prazo_Maximo','Data_Criacao']},
    {nome:'DB_Movimentacoes', cols:['Data','ID','Nome','Tipo','Qtd_Requisitada','Qtd_Recebida','Prazo','Status','Usuario','Incidente']},
    {nome:'DB_Estoque_Nucleo', cols:['ID','Nome','Saldo_Atual']},
    {nome:'DB_Controle_Emails', cols:['Data_Hora','Item_ID','Tipo_Alerta','Hash_Status']},
    {nome:'DB_Logs', cols:['Timestamp','Usuario','Ação','Justificativa']},
    {nome:'Espelho', cols:['ITEM','QUANT. COSUP','QUANT. DTVM','QUANT. DTVM']}
  ];
  for (let aba of abas) {
    let sheet = ss.getSheetByName(aba.nome);
    if (!sheet) {
      sheet = ss.insertSheet(aba.nome);
      sheet.getRange(1,1,1,aba.cols.length).setValues([aba.cols]);
    }
  }
}

function getActiveUserEmail() {
  try {
    var email = Session.getActiveUser().getEmail();
    return email || "Visitante";
  } catch(e) {
    return "Visitante";
  }
}

function getAppData() {
  try {
    Logger.log("INÍCIO getAppData");

    const PLANILHA_ID = '1mM8zs3zUUd2V2HCItqjWCf2W2SN4vohza6moytrPdPI';
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    Logger.log("Planilha aberta OK");

    // Lê DB_Cadastro
    let brindes = [], cadastroRows = [];
    let cadastro = ss.getSheetByName('DB_Cadastro');
    if (cadastro && cadastro.getLastRow() > 1) {
      cadastroRows = cadastro.getRange(2, 1, cadastro.getLastRow() - 1, 5).getValues();
      brindes = cadastroRows.map(row => ({
        id: row[0], nome: row[1]
      }));
      Logger.log("Linhas em DB_Cadastro: " + brindes.length);
    } else {
      Logger.log("DB_Cadastro vazio ou não encontrado");
    }

    // Lê DB_Estoque_Nucleo
    let estoqueRows = [];
    let estoqueN = ss.getSheetByName('DB_Estoque_Nucleo');
    if (estoqueN && estoqueN.getLastRow() > 1) {
      estoqueRows = estoqueN.getRange(2, 1, estoqueN.getLastRow() - 1, 3).getValues();
      Logger.log("Linhas em DB_Estoque_Nucleo: " + estoqueRows.length);
    } else {
      Logger.log("DB_Estoque_Nucleo vazio ou não encontrado");
    }

    // Lê Espelho (COSUP)
    let cosupRows = [];
    let espelho = ss.getSheetByName('Espelho');
    if (espelho && espelho.getLastRow() > 2) {
      cosupRows = espelho.getRange(3, 1, espelho.getLastRow() - 2, 2).getValues()
        .map(row => ({
          nome: String(row[0] || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim(),
          cosup: Number(row[1]) || 0
        }));
      Logger.log("Linhas em Espelho (COSUP): " + cosupRows.length);
    } else {
      Logger.log("Espelho vazio ou não encontrado");
    }

    // DASHBOARD (comparativo)
    let dashboard = [];
    for (let cad of cadastroRows) {
      const id = cad[0], nome = cad[1], min = Number(cad[2]), prazo = cad[3], data_criacao = cad[4];
      const buscaNome = String(nome || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();

      // Saldo núcleo
      const saldo_nucleo = (() => {
        let found = estoqueRows.find(row => String(row[0]) == String(id));
        return found ? Number(found[2]) : 0;
      })();
      // Saldo COSUP
      const saldo_cosup = (() => {
        let found = cosupRows.find(row => row.nome == buscaNome);
        return found ? Number(found.cosup) : 0;
      })();
      dashboard.push({ id, nome, saldo_nucleo, saldo_cosup, min, prazo, data_criacao });
    }

    // Lê DB_Movimentacoes
    let movimentos = [];
    let movs = ss.getSheetByName('DB_Movimentacoes');
    if (movs && movs.getLastRow() > 1) {
      movimentos = movs.getRange(2, 1, movs.getLastRow() - 1, 10).getValues().map(row => ({
        data: row[0],
        id: row[1],
        nome: row[2],
        tipo: row[3],
        qtd_rq: row[4],
        qtd_rc: row[5],
        prazo: row[6],
        status: row[7],
        usuario: row[8],
        incidente: row[9]
      }));
      Logger.log("Linhas em DB_Movimentacoes: " + movimentos.length);
    } else {
      Logger.log("DB_Movimentacoes vazio ou não encontrado");
    }

    // Lê DB_Logs
    let logList = [];
    let logs = ss.getSheetByName('DB_Logs');
    if (logs && logs.getLastRow() > 1) {
      logList = logs.getRange(2, 1, logs.getLastRow() - 1, 4).getValues().map(row => ({
        timestamp: row[0],
        usuario: row[1],
        acao: row[2],
        justificativa: row[3]
      }));
      Logger.log("Linhas em DB_Logs: " + logList.length);
    } else {
      Logger.log("DB_Logs vazio ou não encontrado");
    }

    // KPIs
    let estoqueBaixo = dashboard.filter(d => d.saldo_nucleo <= d.min).length;
    let pedidosPendentes = movimentos.filter(m => m.status === 'Encomenda').length;
    let divergencias = dashboard.filter(d => d.saldo_nucleo !== d.saldo_cosup).length;

    // Usuário (com fallback)
    let usuarioEmail = "";
    try {
      usuarioEmail = Session.getActiveUser().getEmail();
      if (!usuarioEmail) throw "Usuário não detectado";
    } catch (e) {
      usuarioEmail = "Usuário não detectado";
    }

    Logger.log("FINAL getAppData - retorno pronto");
    return {
      brindes,
      dashboard,
      movimentos,
      logs: logList,
      usuario: usuarioEmail, // ou "Usuário não detectado" em fallback seguro
      kpis: { estoqueBaixo, pedidosPendentes, divergencias },
      timestamp: Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd HH:mm:ss"),
      debug: {
        brindesLen: brindes.length,
        dashboardLen: dashboard.length,
        movimentosLen: movimentos.length,
        logsLen: logList.length,
        usuarioEmail: usuarioEmail
      }
    };
  } catch (err) {
    Logger.log("ERRO FATAL EM getAppData: " + err.toString());
    return {
      brindes: brindes || [],
      dashboard: dashboard || [],
      movimentos: movimentos || [],
      logs: logList || [],
      usuario: usuarioEmail || "Visitante",
      kpis: { estoqueBaixo: estoqueBaixo || 0, pedidosPendentes: pedidosPendentes || 0, divergencias: divergencias || 0 },
      timestamp: Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd HH:mm:ss"),
      debug: { brindesLen: brindes.length, dashboardLen: dashboard.length, movimentosLen: movimentos.length, logsLen: logList.length, usuarioEmail: usuarioEmail }
    };
  }
}

/** -------- ENCOMENDA / TRANSFERÊNCIA / REGULARIZAÇÃO ----------- */
function registrarEncomenda(dados) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  criarAbasBanco(ss);
  const cadastro = ss.getSheetByName('DB_Cadastro');
  const estoqueN = ss.getSheetByName('DB_Estoque_Nucleo');
  const movs = ss.getSheetByName('DB_Movimentacoes');
  let emailUser = getActiveUserEmail();

  let novoId = dados.isNovo && dados.nome ? gerarNovoId(cadastro) : dados.id;
  if (dados.isNovo && dados.nome) {
    cadastro.appendRow([
      novoId, dados.nome, dados.minimo, dados.prazo||'', Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd")
    ]);
    estoqueN.appendRow([novoId, dados.nome, 0]);
    logAcao('Cadastro de brinde', emailUser, `Novo brinde: ${dados.nome} (${novoId})`);
  }
  movs.appendRow([
    Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd HH:mm:ss"),
    novoId,
    dados.nome,
    'Encomenda',
    dados.qtd,
    '',
    dados.prazo||'',
    'Encomenda',
    emailUser,
    ''
  ]);
  logAcao('Nova encomenda', emailUser, `Brinde: ${dados.nome}, Qtd: ${dados.qtd}`);
  dispararEmail('Encomenda', {nome:dados.nome, qtd:dados.qtd, prazo:dados.prazo});
  return {sucesso:true, mensagem:"Encomenda registrada com sucesso."};
}

function gerarNovoId(cadastro) {
  const used = cadastro.getLastRow()>1 ? cadastro.getRange(2,1,cadastro.getLastRow()-1,1).getValues().map(r=>r[0]) : [];
  let tenta = 100+Math.floor(Math.random()*900);
  while (used.includes(String(tenta))) tenta = 100+Math.floor(Math.random()*900);
  return String(tenta);
}

function registrarTransferencia(dados) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  criarAbasBanco(ss);
  const movs = ss.getSheetByName('DB_Movimentacoes');
  const espelho = ss.getSheetByName('Espelho');
  let emailUser = getActiveUserEmail();

  let cosupDisp = 0;
  if (espelho.getLastRow()>1) {
    let vals = espelho.getRange(2,1,espelho.getLastRow()-1,2).getValues()
      .map(row => ({
        nome: String(row[0]).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim(),
        cosup: Number(row[1]) || 0
      }));
    let nomeCosup = String(dados.nome).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
    let found = vals.find(r=>r.nome==nomeCosup);
    if (found) cosupDisp = Number(found.cosup);
    if (!found) throw new Error('Item não disponível na COSUP.');
    if (Number(dados.qtd) > cosupDisp) {
      logAcao('Tentativa transferência acima do COSUP', emailUser, `Brinde: ${nomeCosup}, Qtd: ${dados.qtd}, Disponível: ${cosupDisp}`);
      throw new Error('Quantidade excede o estoque disponível na COSUP.');
    }
  }
  movs.appendRow([
    Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd HH:mm:ss"),
    dados.id,
    dados.nome,
    'Transferência',
    dados.qtd,
    '',
    dados.prazo||'',
    'Transferência',
    emailUser,
    ''
  ]);
  logAcao('Nova transferência', emailUser, `Brinde: ${dados.nome}, Qtd: ${dados.qtd}`);
  dispararEmail('Transferência', {nome:dados.nome, qtd:dados.qtd});
  return {sucesso:true, mensagem:"Transferência registrada com sucesso."};
}

/** -- Confirmação de Recebimento com INCIDENTE obrigatório para status Desconformidade e Alinhamento -- */
function confirmarRecebimento(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  criarAbasBanco(ss);
  const movs = ss.getSheetByName('DB_Movimentacoes');
  const estoqueN = ss.getSheetByName('DB_Estoque_Nucleo');
  let emailUser = getActiveUserEmail();

  const vals = movs.getRange(2,1,movs.getLastRow()-1,10).getValues();
  let idx = -1;
  for(let i=0;i<vals.length;i++) {
    if(vals[i][1]==dados.id && vals[i][0]==dados.data && vals[i][2]==dados.nome) { idx=i+2; break;}
  }
  if(idx<0) throw new Error('Registro de movimentação não localizado.');

  const qtdReq = Number(vals[idx-2][4]);
  const qtdRec = Number(dados.qtdRecebida);

  let statusReceb = '';
  let incidenteMsg = dados.incidente || '';
  if(qtdRec<qtdReq){
    statusReceb='Desconformidade';
    if(!incidenteMsg || incidenteMsg.trim()=='') throw new Error("É obrigatório descrever o incidente para recebimento com divergência.");
    if(!existeEmailDesconformidade(dados.nome, qtdRec, ss)) {
      dispararEmail('Desconformidade', {nome:dados.nome, saldo_nucleo:qtdRec, saldo_cosup:'-', incidente:incidenteMsg});
      registrarControleEmail(dados.id, 'Desconformidade', dados.nome+qtdRec, ss);
    }
  }
  else if(qtdRec>qtdReq) {
    statusReceb='Alinhamento';
    dispararEmail('Alinhamento', {nome:dados.nome, qtdRequisitada: qtdReq, qtdRecebida: qtdRec});
  }
  else {
    statusReceb='Confirmado';
  }
  movs.getRange(idx,6).setValue(qtdRec);
  movs.getRange(idx,8).setValue(statusReceb);
  movs.getRange(idx,10).setValue(incidenteMsg);

  let estoqueRows=estoqueN.getLastRow()>1?estoqueN.getRange(2,1,estoqueN.getLastRow()-1,3).getValues():[];
  let estIdx=estoqueRows.findIndex(row=>row[1]==dados.nome);
  if(estIdx>=0){
    let saldoAtual = Number(estoqueRows[estIdx][2]) + qtdRec;
    estoqueN.getRange(estIdx+2,3).setValue(saldoAtual);
  }
  logAcao('Confirmação de recebimento', emailUser, 
    `Brinde: ${dados.nome}, Qtd Req: ${qtdReq}, Qtd Rec: ${qtdRec}, Status: ${statusReceb}, Incidente: ${incidenteMsg}`);
  return {sucesso:true, mensagem:"Recebimento confirmado."};
}

function existeEmailDesconformidade(nome, saldo, ss) {
  const controle = ss.getSheetByName('DB_Controle_Emails');
  let hash = Utilities.base64Encode(nome+'-'+saldo);
  let exists = controle.getLastRow()>1 
    ? controle.getRange(2,4,controle.getLastRow()-1,1).getValues().some(row=>row[0]===hash)
    : false;
  return exists;
}
function registrarControleEmail(item_id, tipo, chave, ss) {
  const controle = ss.getSheetByName('DB_Controle_Emails');
  let hash = Utilities.base64Encode(chave);
  controle.appendRow([
    Utilities.formatDate(new Date(),"GMT-3","yyyy-MM-dd HH:mm:ss"),
    item_id, tipo, hash
  ]);
}

/** ---- Sincronização e Divergência ---- */
function checarDivergencias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  criarAbasBanco(ss);
  const estoqueN = ss.getSheetByName('DB_Estoque_Nucleo');
  const espelho = ss.getSheetByName('Espelho');
  const movs = ss.getSheetByName('DB_Movimentacoes');
  const controle_emails = ss.getSheetByName('DB_Controle_Emails');

  let valsN = estoqueN.getLastRow()>1 ? estoqueN.getRange(2,1,estoqueN.getLastRow()-1,3).getValues() : [];
  let valsE = espelho.getLastRow()>1 ? espelho.getRange(2,1,espelho.getLastRow()-1,2).getValues().map(row => ({
    nome: String(row[0]).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim(),
    cosup: Number(row[1]) || 0
  })) : [];
  let divergencias = [];
  for(let n of valsN){
    const buscaNome = String(n[1]).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
    let e = valsE.find(row=>row.nome === buscaNome);
    if(e && Number(n[2])!=Number(e.cosup)) {
      divergencias.push({id:n[0], nome:n[1], saldo_nucleo:n[2], saldo_cosup:e.cosup});
      let movVals = movs.getRange(2,1,movs.getLastRow()-1,10).getValues();
      for(let row=movVals.length-1;row>=0;row--){
        if(movVals[row][1]==n[0] && movVals[row][2]==n[1]) {
          movs.getRange(row+2,8).setValue('Desconformidade');
          break;
        }
      }
      let hash = Utilities.base64Encode(n[1]+'-'+e.cosup);
      let existeNotificacao = controle_emails.getLastRow()>1 
        ? controle_emails.getRange(2,4,controle_emails.getLastRow()-1,1).getValues().some(row=>row[0]==hash) : false;
      if(!existeNotificacao) {
        controle_emails.appendRow([
          Utilities.formatDate(new Date(),"GMT-3","yyyy-MM-dd HH:mm:ss"),
          n[0], 'Divergencia', hash
        ]);
        dispararEmail('Desconformidade',{id:n[0],nome:n[1],saldo_nucleo:n[2],saldo_cosup:e.cosup});
      }
    }
  }
  logAcao('Checagem de divergências', getActiveUserEmail(), divergencias.length ? JSON.stringify(divergencias) : 'Sem divergências');
  return divergencias;
}

/** ------ REGULARIZAÇÃO MANUAL ------ */
function regularizarManual(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  criarAbasBanco(ss);
  const estoqueN = ss.getSheetByName('DB_Estoque_Nucleo');
  const movs = ss.getSheetByName('DB_Movimentacoes');
  let emailUser = getActiveUserEmail();
  let lastRow = estoqueN.getLastRow();
  let rows = lastRow>1 ? estoqueN.getRange(2,1,lastRow-1,3).getValues() : [];
  let idx = rows.findIndex(row=>row[1]==dados.nome && row[0]==dados.id);
  if(idx>=0) estoqueN.getRange(idx+2,3).setValue(dados.novoSaldo);

  movs.appendRow([
    Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd HH:mm:ss"),
    dados.id,
    dados.nome,
    'Ajuste Manual',
    '',
    dados.novoSaldo,
    '',
    'Regularizacao',
    emailUser,
    dados.justificativa||''
  ]);
  logAcao('Regularização manual', emailUser, `Brinde: ${dados.nome}, Novo saldo: ${dados.novoSaldo}, Just: ${dados.justificativa||''}`);
  return {sucesso:true, mensagem:"Regularização realizada."};
}

/** ------ PEDIDO ESPECIAL (diretoria) ----- */
function registrarPedidoEspecial(dados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  criarAbasBanco(ss);
  dispararEmail('Requisito Diretoria',dados);
  logAcao('Pedido especial diretoria', getActiveUserEmail(), `Brinde: ${dados.nome}, Qtd: ${dados.qtd}, Obs: ${dados.obs||''}`);
  return {sucesso:true, mensagem:"Pedido especial enviado."};
}

/** -------- LOG -------- */
function logAcao(acao, usuario, justificativa) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName('DB_Logs');
  log.appendRow([
    Utilities.formatDate(new Date(),"GMT-3","yyyy-MM-dd HH:mm:ss"),
    usuario||getActiveUserEmail(),
    acao,
    justificativa||''
  ]);
}

/** -------- SISTEMA DE EMAILS -------- */
function dispararEmail(tipo, dados) {
  const emails = [
    'gadian@banestes.com.br', 'asmoreira@banestes.com.br',
    'csdamasceno@banestes.com.br', 'spandrade@banestes.com.br'
  ];
  let subject='',body='';
  if(tipo==='Encomenda'){
    subject=`Nova Encomenda - Brinde ${dados.nome}`;
    body=`Solicitação de encomenda: <b>${dados.nome}</b> (${dados.qtd} unidades)<br>PRAZO: ${dados.prazo||''}<br>Solicitante: ${getActiveUserEmail()}`;
  } else if(tipo==='Transferência') {
    subject=`Nova Transferência - Brinde ${dados.nome}`;
    body=`Solicitação de transferência: <b>${dados.nome}</b> (${dados.qtd} unidades)<br>Solicitante: ${getActiveUserEmail()}`;
  } else if(tipo==='Desconformidade') {
    subject=`Desconformidade detectada - Brinde ${dados.nome}`;
    body=`<b>Brinde:</b> ${dados.nome} | Núcleo: ${dados.saldo_nucleo} | COSUP: ${dados.saldo_cosup}<br>Incidente: ${dados.incidente||''}<br>Executor: ${getActiveUserEmail()}`;
  } else if(tipo==='Alinhamento') {
    subject=`Alinhamento necessário - Brinde ${dados.nome}`;
    body=`Recebida quantidade acima do solicitado para <b>${dados.nome}</b>:<br>Recebida: ${dados.qtdRecebida} vs Requisitada: ${dados.qtdRequisitada}.<br>Executor: ${getActiveUserEmail()}`;
  } else if(tipo==='Requisito Diretoria') {
    subject=`Pedido especial - Diretoria [${dados.nome}]`;
    body=`<b>Pedido especial:</b> Brinde: ${dados.nome} | Qtd: ${dados.qtd}.<br>Data: ${dados.data}<br>Observação: ${dados.obs||''}<br>Executor: ${getActiveUserEmail()}`;
  } else {
    subject=`Alerta sistema - Brinde ${dados.nome||''}`;
    body=`Detalhes: ${JSON.stringify(dados)}`;
  }
  MailApp.sendEmail({
    to: emails.join(","), subject: subject, htmlBody: body
  });
}

/** -------- SEGURANÇA -------- */
function getUserEmail() {
  return getActiveUserEmail();
}

function normalizeNome(nome) {
  if (!nome) return "";
  return nome
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9 ]+/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function debugGetSheets() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const abas = [
    'DB_Cadastro',
    'DB_Estoque_Nucleo',
    'Espelho',
    'DB_Movimentacoes',
    'DB_Logs',
    'DB_Controle_Emails'
  ];
  abas.forEach(function(aba) {
    let sh = ss.getSheetByName(aba);
    if (sh) {
      Logger.log("Aba %s encontrada. Linhas: %d, Colunas: %d", aba, sh.getLastRow(), sh.getLastColumn());
    } else {
      Logger.log("ERRO: Aba %s NÃO encontrada!!! (verifique o nome exato)", aba);
    }
  });
}

function debugGetRows() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  const abas = [
    { nome: 'DB_Cadastro', minCols: 5 },
    { nome: 'DB_Estoque_Nucleo', minCols: 3 },
    { nome: 'Espelho', minCols: 2 },
    { nome: 'DB_Movimentacoes', minCols: 10 },
    { nome: 'DB_Logs', minCols: 4 }
  ];
  abas.forEach(function(aba) {
    let sh = ss.getSheetByName(aba.nome);
    if (!sh) {
      Logger.log("FALHA ABA: %s", aba.nome);
      return;
    }
    let rows = sh.getLastRow();
    let cols = sh.getLastColumn();
    if (rows < 2) {
      Logger.log("Aba %s NÃO TEM DADOS (apenas cabeçalho ou está vazia)", aba.nome);
      return;
    }
    try {
      let data = sh.getRange(2,1,rows-1,aba.minCols).getValues();
      Logger.log("Aba %s OK - Linhas de dados: %d", aba.nome, data.length);
    } catch(e) {
      Logger.log("ERRO ao ler range da aba %s: %s", aba.nome, e);
    }
  });
}

/**
 * Função universal para extrair dados da aba "DB_Cadastro" (ou qualquer outra)
 * Retorna array de objetos para exibição no front.
 */
function getCadastroList() {
  try {
    var ss = SpreadsheetApp.openById(PLANILHA_ID);  // Use a constante já definida
    var sheet = ss.getSheetByName('DB_Cadastro');   // Troque para a aba desejada
    if (!sheet) {
      return { success: false, error: "Aba 'DB_Cadastro' não encontrada", list: [] };
    }
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, list: [] }; // só cabeçalho, nada de dados
    }
    var values = sheet.getRange(2, 1, lastRow-1, 5).getValues(); // 5 colunas: ID, Nome, Min, Prazo, Data
    var list = values.map(row => ({
      id: row[0],
      nome: row[1],
      estoqueMin: row[2],
      prazoMax: row[3],
      dataCriacao: row[4]
    }));
    return { success: true, list: list };
  } catch (e) {
    return { success: false, error: e.toString(), list: [] };
  }
}
