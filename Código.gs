/**
 * Núcleo Asset - Backend Google Apps Script Full Stack
 * Arquitetura de Dados | Regras de Negócio - Revisado conforme Requisitos Bancários
 */

const PLANILHA_ID = '1mM8zs3zUUd2V2HCItqjWCf2W2SN4vohza6moytrPdPI';

function doGet(e) {
  // const ss = SpreadsheetApp.openById(PLANILHA_ID);
  // criarAbasBanco(ss);
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Bem-vindo ao Gerenciador de Ativos - Núcleo Asset')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function criarAbasBanco(ss) {
  const abas = [
    {nome:'DB_Cadastro', cols:['ID','Nome','Estoque_Minimo','Prazo_Maximo','Data_Criacao']},
    {nome:'DB_Movimentacoes', cols:['Data','ID','Nome','Tipo','Qtd_Requisitada','Qtd_Recebida','Prazo','Status','Usuario','Incidente']},
    {nome:'DB_Estoque_Nucleo', cols:['ID','Nome','Saldo_Atual']},
    {nome:'DB_Controle_Emails', cols:['Data_Hora','Item_ID','Tipo_Alerta','Hash_Status']},
    {nome:'DB_Logs', cols:['Timestamp','Usuario','Ação','Justificativa']},
    {nome:'Espelho', cols:['ITEM','QUANT.COSUP']}
  ];
  for (let aba of abas) {
    let sheet = ss.getSheetByName(aba.nome);
    if (!sheet) {
      sheet = ss.insertSheet(aba.nome);
      if (aba.nome === 'Espelho') {
        sheet.getRange(1,1).setValue(aba.nome);
        sheet.getRange(2,1,1,aba.cols.length).setValues([aba.cols]);
      } else {
        sheet.getRange(1,1,1,aba.cols.length).setValues([aba.cols]);
      }
    }
  }
}

function formatDateSafe(v) {
  if (!v && v !== 0) return '';
  if (v instanceof Date) return Utilities.formatDate(v, "GMT-3", "yyyy-MM-dd HH:mm:ss");
  return String(v);
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
        id: row[0],
        nome: row[1],
        min: row[2],
        prazo: row[3],
        data_criacao: formatDateSafe(row[4])
      }));
      Logger.log("Linhas em DB_Cadastro: " + brindes.length);
    }

    // Lê DB_Estoque_Nucleo
    let estoqueRows = [];
    let estoqueN = ss.getSheetByName('DB_Estoque_Nucleo');
    if (estoqueN && estoqueN.getLastRow() > 1) {
      estoqueRows = estoqueN.getRange(2, 1, estoqueN.getLastRow() - 1, 3).getValues().map(row => ({
        id: String(row[0]),
        nome: String(row[1]),
        saldo_nucleo: Number(row[2]) || 0
      }));
    }

    // Lê Espelho (COSUP) COM ID
    let cosupRows = [];
    let espelho = ss.getSheetByName('Espelho');
    if (espelho && espelho.getLastRow() > 2) {
      // ATENÇÃO: Colunas: ITEM(B), QUANT COSUP(C), ID(D) => indexes: 0(B), 1(C), 2(D)
      cosupRows = espelho.getRange(3, 1, espelho.getLastRow() - 2, 3).getValues().map(row => ({
        nome: String(row[0] || '').trim(),
        cosup: Number(row[1]) || 0,
        id: String(row[2] || '').trim()
      })).filter(r => r.id); // Só pega linhas com ID preenchido
    }

    // DASHBOARD (comparativo unificado por ID)
    let dashboard = [];
    for (let est of estoqueRows) {
      // Procura COSUP com o mesmo id
      const cosup = cosupRows.find(c => c.id === est.id);
      const saldo_cosup = cosup ? cosup.cosup : 0;
      // Ache info de cadastro/limite minimo, se quiser (opcional)
      const cad = cadastroRows.find(row => String(row[0]) === est.id) || [];
      const min = Number(cad[2]) || 0;
      const prazo = cad[3] || '';
      const data_criacao = cad[4] || '';

      // STATUS do estoque núcleo
      let status = '';
      if (est.saldo_nucleo === 0) status = 'Zero';
      else if (est.saldo_nucleo <= min) status = 'Baixo';
      else status = 'Ok';

      dashboard.push({
        id: est.id,
        nome: est.nome,
        saldo_nucleo: est.saldo_nucleo,
        saldo_cosup,
        total_estoque: est.saldo_nucleo + saldo_cosup,
        min,
        prazo: String(prazo || ''),
        data_criacao: formatDateSafe(data_criacao),
        status
      });
    }

    // Lê DB_Movimentacoes
    let movimentos = [];
    let movs = ss.getSheetByName('DB_Movimentacoes');
    if (movs && movs.getLastRow() > 1) {
      movimentos = movs.getRange(2, 1, movs.getLastRow() - 1, 10).getValues().map(row => ({
        data: formatDateSafe(row[0]),
        id: String(row[1] || ''),
        nome: String(row[2] || ''),
        tipo: String(row[3] || ''),
        qtd_rq: Number(row[4]) || 0,
        qtd_rc: Number(row[5]) || 0,
        prazo: formatDateSafe(row[6]),
        status: String(row[7] || ''),
        usuario: String(row[8] || ''),
        incidente: String(row[9] || '')
      }));
    }

    // Lê DB_Logs
    let logList = [];
    let logs = ss.getSheetByName('DB_Logs');
    if (logs && logs.getLastRow() > 1) {
      logList = logs.getRange(2, 1, logs.getLastRow() - 1, 4).getValues().map(row => ({
        timestamp: formatDateSafe(row[0]),
        usuario: String(row[1] || ''),
        acao: String(row[2] || ''),
        justificativa: String(row[3] || '')
      }));
    }

    // ================================
    // KPIs NOVOS PARA OS QUATRO CARDS:
    // ================================
    // KPIs baseados em movimentos
    let estoqueBaixo = dashboard.filter(d => d.status === 'Baixo').length;
    let pedidosEncomenda = movimentos.filter(m => m.status === 'Encomenda').length;
    let pedidosTransferencia = movimentos.filter(m => m.status === 'Transferência').length;
    let divergencias = movimentos.filter(m => m.status === 'Desconformidade').length;

    // Usuário (com fallback)
    let usuarioEmail = "";
    try {
      usuarioEmail = Session.getActiveUser().getEmail();
      if (!usuarioEmail) throw "Usuário não detectado";
    } catch (e) {
      usuarioEmail = "Usuário não detectado";
    }

    return {
      brindes,
      dashboard,
      movimentos,
      logs: logList,
      usuario: usuarioEmail,
      kpis: {
        estoqueBaixo,
        pedidosEncomenda,
        pedidosTransferencia,
        divergencias
      },
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
      brindes: [],
      dashboard: [],
      movimentos: [],
      logs: [],
      usuario: "Erro",
      kpis: { pedidosTransferencia: 0, pedidosEncomenda: 0, pedidosDesconformidade: 0, estoqueBaixo: 0 },
      timestamp: "",
      debug: { error: err.toString() }
    };
  }
}


/** -------- CADASTRO SIMPLES DE BRINDE (sem movimentação) ----------- */
/**
 * Cadastra um novo brinde no sistema sem criar qualquer registro de movimentação.
 * Use esta função quando o item está sendo apenas adicionado ao catálogo,
 * sem necessidade de encomenda ou transferência imediata.
 * Para cadastrar + criar encomenda, use registrarEncomenda({ isNovo: true, ...}).
 */
function cadastrarBrinde(dados) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  criarAbasBanco(ss);
  const cadastro = ss.getSheetByName('DB_Cadastro');
  const estoqueN = ss.getSheetByName('DB_Estoque_Nucleo');
  let emailUser = getActiveUserEmail();

  if (!dados.nome || String(dados.nome).trim() === '') {
    throw new Error('O nome do brinde é obrigatório.');
  }

  let novoId = gerarNovoId(cadastro);
  cadastro.appendRow([
    novoId,
    dados.nome,
    dados.minimo || 0,
    '',
    Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd")
  ]);
  estoqueN.appendRow([novoId, dados.nome, 0]);
  logAcao('Cadastro de brinde', emailUser, `Novo brinde: ${dados.nome} (${novoId})`);
  return { sucesso: true, mensagem: `Brinde "${dados.nome}" cadastrado com sucesso (ID: ${novoId}).` };
}

/** -------- ENCOMENDA / TRANSFERÊNCIA ----------- */
/**
 * Lógica de status das movimentações:
 *
 * ENCOMENDA (tipo = 'Encomenda', status inicial = 'Encomenda'):
 *   Representa uma solicitação de compra/reposição de item ao fornecedor.
 *   Fluxo: Encomenda → (ao confirmar recebimento) → Confirmado | Desconformidade | Alinhamento
 *
 * TRANSFERÊNCIA (tipo = 'Transferência', status inicial = 'Transferência'):
 *   Representa uma transferência de estoque da COSUP para o Núcleo Asset.
 *   Fluxo: Transferência → (ao confirmar recebimento) → Confirmado | Desconformidade | Alinhamento
 *
 * NÃO HÁ transição automática entre Encomenda e Transferência — são fluxos independentes.
 *
 * DESCONFORMIDADE: Qtd. recebida < solicitada. Requer incidente. Botão "Regularizar" disponível.
 * ALINHAMENTO: Qtd. recebida > solicitada (excedente).
 * CONFIRMADO: Qtd. recebida = solicitada.
 * EM ESTOQUE: Após regularização manual de uma Desconformidade.
 */
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
  if (espelho.getLastRow()>2) {
    let vals = espelho.getRange(3,1,espelho.getLastRow()-2,3).getValues()
      .map(row => ({
        nome: String(row[0]).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim(),
        cosup: Number(row[1]) || 0,
        id: String(row[2] || '').trim()
      }));
    let nomeCosup = String(dados.nome).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
    let found = vals.find(r => r.id && r.id === String(dados.id).trim());
    if (!found) found = vals.find(r => r.nome === nomeCosup);
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
  dispararEmail('Transferência', {nome:dados.nome, qtd:dados.qtd, prazo:dados.prazo});
  return {sucesso:true, mensagem:"Transferência registrada com sucesso."};
}

/** -- Confirmação de Recebimento com INCIDENTE obrigatório para status Desconformidade e Alinhamento -- */
function confirmarRecebimento(dados) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  criarAbasBanco(ss);
  const movs = ss.getSheetByName('DB_Movimentacoes');
  const estoqueN = ss.getSheetByName('DB_Estoque_Nucleo');
  let emailUser = getActiveUserEmail();

  const vals = movs.getRange(2,1,movs.getLastRow()-1,10).getValues();
  let idx = -1;
  for(let i=0;i<vals.length;i++) {
    if(String(vals[i][1])==String(dados.id) && formatDateSafe(vals[i][0])==dados.data && vals[i][2]==dados.nome) { idx=i+2; break;}
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
      dispararEmail('Desconformidade', {nome:dados.nome, qtdRequisitada:qtdReq, qtdRecebida:qtdRec, incidente:incidenteMsg});
      registrarControleEmail(dados.id, 'Desconformidade', dados.nome+qtdRec, ss);
    }
  }
  else if(qtdRec>qtdReq) {
    statusReceb='Alinhamento';
    dispararEmail('Alinhamento', {id:dados.id, nome:dados.nome, qtdRequisitada: qtdReq, qtdRecebida: qtdRec});
  }
  else {
    statusReceb='Confirmado';
    dispararEmail('Confirmado', {id:dados.id, nome:dados.nome, qtdRequisitada: qtdReq, qtdRecebida: qtdRec});
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

/** ------ SAÍDA DE ITEM DO ESTOQUE NÚCLEO ------ */
function registrarSaida(dados) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  criarAbasBanco(ss);
  const movs = ss.getSheetByName('DB_Movimentacoes');
  const estoqueN = ss.getSheetByName('DB_Estoque_Nucleo');
  let emailUser = getActiveUserEmail();

  if (!dados.retirante || String(dados.retirante).trim() === '') {
    throw new Error('O nome do retirante é obrigatório.');
  }

  let estoqueRows = estoqueN.getLastRow() > 1
    ? estoqueN.getRange(2, 1, estoqueN.getLastRow() - 1, 3).getValues()
    : [];
  let estIdx = estoqueRows.findIndex(row => String(row[0]) === String(dados.id));
  if (estIdx < 0) throw new Error('Item não encontrado no estoque núcleo.');

  let saldoAtual = Number(estoqueRows[estIdx][2]) || 0;
  let qtdSaida = Number(dados.qtd);
  if (qtdSaida <= 0) throw new Error('Quantidade deve ser maior que zero.');
  if (qtdSaida > saldoAtual) throw new Error(`Quantidade solicitada (${qtdSaida}) excede o estoque disponível (${saldoAtual}).`);

  let dataHora = Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd HH:mm:ss");
  movs.appendRow([
    dataHora,
    dados.id,
    dados.nome,
    'Saída',
    qtdSaida,
    qtdSaida,
    '',
    'Saída',
    emailUser,
    dados.retirante  // Para registros do tipo 'Saída', o campo Incidente armazena o nome do retirante
  ]);

  let novoSaldo = saldoAtual - qtdSaida;
  estoqueN.getRange(estIdx + 2, 3).setValue(novoSaldo);

  // Verifica alertas de estoque após saída
  try {
    let cad = ss.getSheetByName('DB_Cadastro');
    let cadRows = cad && cad.getLastRow() > 1
      ? cad.getRange(2, 1, cad.getLastRow() - 1, 3).getValues()
      : [];
    let cadItem = cadRows.find(row => String(row[0]) === String(dados.id));
    let minimo = cadItem ? Number(cadItem[2]) || 0 : 0;
    if (novoSaldo === 0) {
      enviarAlertaEstoque('zero', { id: dados.id, nome: dados.nome });
    } else if (minimo > 0 && novoSaldo <= minimo) {
      enviarAlertaEstoque('minimo', { id: dados.id, nome: dados.nome, saldo_nucleo: novoSaldo, min: minimo });
    }
  } catch(alertErr) {
    Logger.log('Aviso: Falha ao enviar alerta de estoque: ' + alertErr);
  }

  logAcao('Saída de item', emailUser,
    `Brinde: ${dados.nome}, Qtd: ${qtdSaida}, Retirante: ${dados.retirante}`);

  dispararEmail('Saída', {
    nome: dados.nome,
    qtd: qtdSaida,
    retirante: dados.retirante,
    dataHora: dataHora,
    usuario: emailUser
  });

  return { sucesso: true, mensagem: 'Saída registrada com sucesso.' };
}

/** ---- Sincronização e Divergência ---- */
function checarDivergencias() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  criarAbasBanco(ss);
  const estoqueN = ss.getSheetByName('DB_Estoque_Nucleo');
  const espelho = ss.getSheetByName('Espelho');
  const movs = ss.getSheetByName('DB_Movimentacoes');
  const controle_emails = ss.getSheetByName('DB_Controle_Emails');

  let valsN = estoqueN.getLastRow()>1 ? estoqueN.getRange(2,1,estoqueN.getLastRow()-1,3).getValues() : [];
  let valsE = espelho.getLastRow()>2 ? espelho.getRange(3,1,espelho.getLastRow()-2,2).getValues().map(row => ({
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
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  criarAbasBanco(ss);
  const estoqueN = ss.getSheetByName('DB_Estoque_Nucleo');
  const movs = ss.getSheetByName('DB_Movimentacoes');
  let emailUser = getActiveUserEmail();
  let lastRow = estoqueN.getLastRow();
  let rows = lastRow>1 ? estoqueN.getRange(2,1,lastRow-1,3).getValues() : [];
  let idx = rows.findIndex(row=>row[1]==dados.nome && row[0]==dados.id);
  if(idx>=0)
    estoqueN.getRange(idx+2,3).setValue(dados.novoSaldo);

  // Atualizar movimento específico para status "Em estoque"
  let movRows = movs.getRange(2,1,movs.getLastRow()-1,10).getValues();
  for(let i=0;i<movRows.length;i++){
    if(movRows[i][1]==dados.id && movRows[i][2]==dados.nome && movRows[i][0]==dados.data){
      movs.getRange(i+2,8).setValue('Em estoque'); // status
      movs.getRange(i+2,6).setValue(dados.novoSaldo); // qtd recebida
      break;
    }
  }

  dispararEmail('Regularizacao', {
    id: dados.id,
    nome: dados.nome,
    novoSaldo: dados.novoSaldo,
    justificativa: dados.justificativa
  });
  logAcao('Regularização manual', emailUser, `Brinde: ${dados.nome}, Novo saldo: ${dados.novoSaldo}, Just: ${dados.justificativa||''}`);
  return {sucesso:true, mensagem:"Regularização realizada com sucesso."};
}

/** ------ PEDIDO ESPECIAL (diretoria) ----- */
/*function registrarPedidoEspecial(dados) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  criarAbasBanco(ss);

  // Disparar e-mail com o template correto
  let emails = [
    'spandrade@banestes.com.br',
    //'gadian@banestes.com.br',
    //'asmoreira@banestes.com.br',
    //'csdamasceno@banestes.com.br'
  ];
  let subject = `Pedido Especial de Diretoria - Brinde ${dados.nome}`;
  let htmlBody = renderEmailPedidoDiretoria(dados);

  MailApp.sendEmail({
    to: emails.join(","),
    subject: subject,
    htmlBody: htmlBody
  });

  logAcao('Pedido especial diretoria', getActiveUserEmail(), `Brinde: ${dados.nome}, Qtd: ${dados.qtd}, Obs: ${dados.obs||''}`);
  return {sucesso:true, mensagem:"Pedido especial enviado."};
}*/

/** -------- LOG -------- */
function logAcao(acao, usuario, justificativa) {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
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
  let emails = [];
  let subject = '';
  let body = '';

  if(tipo === 'Encomenda') {
    subject = `Nova Encomenda - Brinde ${dados.nome}`;
    body = renderEmailEncomenda(dados);
    emails = [
      //'spandrade@banestes.com.br',
      //'gadian@banestes.com.br',
      //'asmoreira@banestes.com.br',
      'csdamasceno@banestes.com.br'
    ];
  } else if(tipo === 'Transferência') {
    subject = `Transferência Efetuada - Brinde ${dados.nome}`;
    body = renderEmailTransferencia(dados);
    emails = [
      //'spandrade@banestes.com.br',
      //'gadian@banestes.com.br',
      //'asmoreira@banestes.com.br',
      'csdamasceno@banestes.com.br'
    ];
  } else if(tipo === 'Desconformidade') {
    subject = `Divergência de Recebimento - Brinde ${dados.nome}`;
    body = renderEmailDivergenciaCosup(dados);
    emails = [
      //'spandrade@banestes.com.br',
      //'gadian@banestes.com.br',
      //'asmoreira@banestes.com.br',
      'csdamasceno@banestes.com.br'
    ];
  } else if(tipo === 'Alinhamento') {
    subject = `Recebimento Excedente - Brinde ${dados.nome}`;
    body = renderEmailExcedente(dados);
    emails = [
      //'spandrade@banestes.com.br',
      //'gadian@banestes.com.br',
      //'asmoreira@banestes.com.br',
      'csdamasceno@banestes.com.br'
    ];
  } else if(tipo === 'Confirmado') {
    subject = `Recebimento Confirmado - Brinde ${dados.nome}`;
    body = renderEmailConfirmado(dados);
    emails = [
      //'spandrade@banestes.com.br',
      //'gadian@banestes.com.br',
      //'asmoreira@banestes.com.br',
      'csdamasceno@banestes.com.br'
    ];
  } else if(tipo === 'Regularizacao') {
    subject = `Formalização de Ajuste Manual - Brinde ${dados.nome}`;
    body = renderEmailRegularizacao(dados);
    emails = [
      //'spandrade@banestes.com.br',
      //'gadian@banestes.com.br',
      //'asmoreira@banestes.com.br',
      'csdamasceno@banestes.com.br'
    ];
  } else if(tipo === 'Saída') {
    subject = `Saída de Item - ${dados.nome}`;
    body = renderEmailSaida(dados);
    emails = [
      //'spandrade@banestes.com.br',
      //'gadian@banestes.com.br',
      //'asmoreira@banestes.com.br',
      'csdamasceno@banestes.com.br'
    ];
  } else {
    subject = `Alerta sistema - Brinde ${dados.nome||''}`;
    body = `<pre>${JSON.stringify(dados)}</pre>`;
    emails = [
      //'spandrade@banestes.com.br',
      //'gadian@banestes.com.br',
      //'asmoreira@banestes.com.br',
      'csdamasceno@banestes.com.br'
    ];
  }

  MailApp.sendEmail({
    to: emails.join(","),
    subject: subject,
    htmlBody: body
  });
}

// ============================================================
// TEMPLATES DE E-MAIL — Layout padronizado Núcleo Asset
// ============================================================

/** Helper: linha de detalhe da tabela padrão */
function _emailRow(label, value, last) {
  var border = last ? '' : 'border-bottom:1px solid #eef0f3;';
  return `<tr style="${border}">
    <td style="padding:10px 14px;font-weight:600;color:#003366;width:42%;white-space:nowrap;">${label}</td>
    <td style="padding:10px 14px;color:#444;">${value}</td>
  </tr>`;
}

/** Helper: wrapper completo do e-mail com cabeçalho, corpo e rodapé padronizados */
function _emailWrapper(headerBg, headerText, title, subtitle, bodyHtml) {
  var ts = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm:ss');
  return `<div style="font-family:'Segoe UI',Arial,sans-serif;color:#333;max-width:600px;margin:0 auto;border:1px solid #e4e8ee;border-radius:8px;overflow:hidden;">
  <div style="background:${headerBg};color:${headerText};padding:20px 25px;text-align:center;">
    <h2 style="margin:0;font-size:18px;letter-spacing:1px;text-transform:uppercase;">${title}</h2>
    <p style="margin:7px 0 0;font-size:12px;opacity:0.88;">${subtitle}</p>
  </div>
  <div style="padding:24px 25px;background:#ffffff;line-height:1.6;">
    ${bodyHtml}
  </div>
  <div style="background:#f4f7f9;padding:13px 25px;text-align:center;font-size:11px;color:#7a8fa6;border-top:1px solid #e4e8ee;">
    E-mail automático — Gerenciador de Brindes · Núcleo Asset &nbsp;|&nbsp; Gerado em: <strong>${ts}</strong>
  </div>
</div>`;
}

/** 1. NOVA ENCOMENDA */
function renderEmailEncomenda(dados) {
  var rows =
    _emailRow('Brinde', dados.nome) +
    _emailRow('Quantidade Solicitada', dados.qtd) +
    (dados.prazo ? _emailRow('Prazo Máximo', dados.prazo + ' dias') : '') +
    _emailRow('Solicitante', getActiveUserEmail(), true);

  var body = `<p>Prezados,</p>
<p>Uma nova <strong>encomenda de brinde</strong> foi registrada no sistema e aguarda providências de aquisição:</p>
<table style="width:100%;border-collapse:collapse;margin:16px 0;border:1px solid #e4e8ee;border-radius:6px;overflow:hidden;">
  ${rows}
</table>
<p style="font-size:13px;color:#555;">Por favor, providenciar o pedido conforme os dados acima e registrar o recebimento no sistema após a chegada do material.</p>`;

  return _emailWrapper('#003366', '#ffffff', 'Nova Encomenda de Brinde', 'Gerenciador de Brindes — Núcleo Asset', body);
}

/** 2. TRANSFERÊNCIA DE ESTOQUE (COSUP → NÚCLEO) */
function renderEmailTransferencia(dados) {
  var rows =
    _emailRow('Brinde', dados.nome) +
    _emailRow('Quantidade Transferida', dados.qtd) +
    (dados.prazo ? _emailRow('Prazo Previsto', dados.prazo + ' dias') : '') +
    _emailRow('Responsável', getActiveUserEmail(), true);

  var body = `<p>Prezados,</p>
<p>Uma <strong>transferência de estoque</strong> foi registrada no sistema. Os itens serão movimentados do estoque COSUP para o Núcleo Asset:</p>
<table style="width:100%;border-collapse:collapse;margin:16px 0;border:1px solid #e4e8ee;border-radius:6px;overflow:hidden;">
  ${rows}
</table>
<div style="background:#eef4ff;border-left:4px solid #003366;padding:12px 16px;border-radius:0 6px 6px 0;margin:16px 0;font-size:13px;">
  Aguardar confirmação do recebimento físico para atualização do saldo no sistema.
</div>`;

  return _emailWrapper('#003366', '#ffffff', 'Transferência de Estoque', 'Gerenciador de Brindes — Núcleo Asset', body);
}

/** 3. RECEBIMENTO CONFIRMADO */
function renderEmailConfirmado(dados) {
  var rows =
    _emailRow('Brinde', dados.nome + (dados.id ? ' (ID: ' + dados.id + ')' : '')) +
    _emailRow('Qtd. Solicitada', dados.qtdRequisitada) +
    _emailRow('Qtd. Recebida', dados.qtdRecebida) +
    _emailRow('Responsável pela conferência', getActiveUserEmail(), true);

  var body = `<p>Prezados,</p>
<p>O recebimento do item abaixo foi <strong>confirmado com sucesso</strong> no Núcleo Asset. A quantidade recebida confere exatamente com a solicitada:</p>
<table style="width:100%;border-collapse:collapse;margin:16px 0;border:1px solid #e4e8ee;border-radius:6px;overflow:hidden;">
  ${rows}
</table>
<div style="background:#f0fff4;border-left:4px solid #276749;padding:12px 16px;border-radius:0 6px 6px 0;margin:16px 0;font-size:13px;">
  <strong>Ação tomada:</strong> O saldo foi incorporado ao estoque atual do Núcleo Asset.
</div>`;

  return _emailWrapper('#276749', '#ffffff', 'Recebimento Confirmado', 'Confirmação de Entrada — Núcleo Asset', body);
}

/** 4. SAÍDA DE ITEM (MOVIMENTAÇÃO INTERNA) */
function renderEmailSaida(dados) {
  var rows =
    _emailRow('Brinde', dados.nome) +
    _emailRow('Quantidade Retirada', dados.qtd) +
    _emailRow('Retirante', dados.retirante) +
    _emailRow('Data / Hora', dados.dataHora) +
    _emailRow('Operador', dados.usuario, true);

  var body = `<p>Prezado(a),</p>
<p>Informamos que foi realizada uma <strong>saída de item</strong> do estoque do Núcleo Asset:</p>
<table style="width:100%;border-collapse:collapse;margin:16px 0;border:1px solid #e4e8ee;border-radius:6px;overflow:hidden;">
  ${rows}
</table>
<p style="font-size:13px;color:#555;">Esta movimentação já foi descontada do saldo do estoque núcleo no sistema.</p>`;

  return _emailWrapper('#003366', '#ffffff', 'Saída de Item — Movimentação Interna', 'Gerenciador de Brindes — Núcleo Asset', body);
}

/** 4. DIVERGÊNCIA DE RECEBIMENTO */
function renderEmailDivergenciaCosup(dados) {
  var qtdEsperada = dados.qtdRequisitada !== undefined ? dados.qtdRequisitada : dados.saldo_cosup;
  var qtdEfetiva  = dados.qtdRecebida   !== undefined ? dados.qtdRecebida   : dados.saldo_nucleo;
  var diff = Number(qtdEsperada) - Number(qtdEfetiva);

  var rows =
    _emailRow('Brinde', dados.nome) +
    _emailRow('Qtd. Esperada', qtdEsperada) +
    _emailRow('Qtd. Recebida', qtdEfetiva) +
    _emailRow('Diferença (a menor)', '- ' + diff, !dados.incidente);

  var incidenteBloco = dados.incidente
    ? `<div style="background:#fff5f5;border-left:4px solid #c53030;padding:12px 16px;border-radius:0 6px 6px 0;margin:12px 0;font-size:13px;">
        <strong>Incidente Registrado:</strong><br><i>"${dados.incidente}"</i>
      </div>`
    : '';

  var body = `<p>Prezados,</p>
<p>Identificamos uma <strong>desconformidade entre a guia de remessa e o material físico</strong> recebido no Núcleo:</p>
<table style="width:100%;border-collapse:collapse;margin:16px 0;border:1px solid #e4e8ee;border-radius:6px;overflow:hidden;">
  ${rows}
</table>
${incidenteBloco}
<p style="font-size:13px;color:#555;">Solicitamos a verificação no estoque central (COSUP) para alinhamento de saldos e regularização do lote.</p>`;

  return _emailWrapper('#c53030', '#ffffff', 'Divergência de Recebimento', 'Alerta de Desconformidade — Núcleo Asset', body);
}

/** 5. ALINHAMENTO — RECEBIMENTO A MAIS */
function renderEmailExcedente(dados) {
  var diff = Number(dados.qtdRecebida) - Number(dados.qtdRequisitada);

  var rows =
    _emailRow('Brinde', dados.nome + (dados.id ? ' (ID: ' + dados.id + ')' : '')) +
    _emailRow('Qtd. Solicitada', dados.qtdRequisitada) +
    _emailRow('Qtd. Recebida', dados.qtdRecebida) +
    _emailRow('Diferença (a maior)', '+ ' + diff) +
    _emailRow('Responsável pela conferência', getActiveUserEmail(), true);

  var body = `<p>Prezados,</p>
<p>Durante a conferência de entrada no Núcleo Asset, identificamos o recebimento de uma <strong>quantidade superior</strong> à solicitada originalmente:</p>
<table style="width:100%;border-collapse:collapse;margin:16px 0;border:1px solid #e4e8ee;border-radius:6px;overflow:hidden;">
  ${rows}
</table>
<div style="background:#f0fff4;border-left:4px solid #276749;padding:12px 16px;border-radius:0 6px 6px 0;margin:16px 0;font-size:13px;">
  <strong>Ação tomada:</strong> O saldo excedente foi incorporado ao estoque atual após validação física.
</div>`;

  return _emailWrapper('#276749', '#ffffff', 'Recebimento com Excedente', 'Alinhamento de Saldo — Núcleo Asset', body);
}

/** 6. FORMALIZAÇÃO DE AJUSTE MANUAL (REGULARIZAÇÃO) */
function renderEmailRegularizacao(dados) {
  var rows =
    _emailRow('Brinde', dados.nome + (dados.id ? ' (ID: ' + dados.id + ')' : '')) +
    _emailRow('Saldo Final Ajustado', dados.novoSaldo + ' unidades') +
    _emailRow('Responsável', getActiveUserEmail(), !dados.justificativa);

  var justBloco = dados.justificativa
    ? `<div style="background:#f0f7ff;border-left:4px solid #1a6bb5;padding:12px 16px;border-radius:0 6px 6px 0;margin:16px 0;font-size:13px;">
        <strong>Justificativa:</strong><br>${dados.justificativa}
      </div>`
    : '';

  var body = `<p>Prezados,</p>
<p>Informamos que o item abaixo foi <strong>regularizado manualmente</strong> no sistema, encerrando o incidente de desconformidade para este lote:</p>
<table style="width:100%;border-collapse:collapse;margin:16px 0;border:1px solid #e4e8ee;border-radius:6px;overflow:hidden;">
  ${rows}
</table>
${justBloco}
<p style="font-size:12px;color:#888;">*Este registro substitui a pendência anterior e consolida o saldo no estoque núcleo.</p>`;

  return _emailWrapper('#1a6bb5', '#ffffff', 'Formalização de Ajuste Manual', 'Regularização de Estoque — Núcleo Asset', body);
}

/** 7. ESTOQUE MÍNIMO ATINGIDO */
function renderEstoqueMinEmail(dados) {
  var rows =
    _emailRow('Brinde', dados.nome) +
    _emailRow('ID', dados.id) +
    _emailRow('Saldo Atual', dados.saldo_nucleo + ' unidades') +
    _emailRow('Estoque Mínimo de Segurança', dados.min + ' unidades', true);

  var body = `<p>Prezados,</p>
<p>O item abaixo atingiu o <strong>nível mínimo de segurança</strong> no estoque do Núcleo e necessita de reposição imediata:</p>
<table style="width:100%;border-collapse:collapse;margin:16px 0;border:1px solid #e4e8ee;border-radius:6px;overflow:hidden;">
  ${rows}
</table>
<div style="background:#fffbeb;border-left:4px solid #b87c00;padding:12px 16px;border-radius:0 6px 6px 0;margin:16px 0;font-size:13px;">
  Sugerimos a abertura de uma nova <strong>Encomenda</strong> ou <strong>Transferência COSUP</strong> através do sistema para evitar a ruptura total do estoque.
</div>`;

  return _emailWrapper('#b87c00', '#ffffff', '⚠️ Estoque Mínimo Atingido', 'Alerta de Reposição — Núcleo Asset', body);
}

/** 8. ESTOQUE ZERADO */
function renderEstoqueZeroEmail(dados) {
  var rows =
    _emailRow('Brinde', dados.nome) +
    _emailRow('ID', dados.id) +
    _emailRow('Saldo Atual', '0 unidades') +
    _emailRow('Status', 'INDISPONÍVEL', true);

  var body = `<p>Prezados,</p>
<p style="color:#c53030;font-weight:600;">Atenção: O item abaixo está com saldo <u>zerado</u> no estoque físico do Núcleo.</p>
<table style="width:100%;border-collapse:collapse;margin:16px 0;border:1px solid #e4e8ee;border-radius:6px;overflow:hidden;">
  ${rows}
</table>
<div style="background:#fff5f5;border-left:4px solid #c53030;padding:12px 16px;border-radius:0 6px 6px 0;margin:16px 0;font-size:13px;">
  Qualquer nova solicitação de saída para este brinde será <strong>negada</strong> até que uma nova entrada ou regularização seja processada no sistema.
</div>`;

  return _emailWrapper('#c53030', '#ffffff', '🚨 Estoque Zerado — Ação Urgente', 'Alerta Crítico — Núcleo Asset', body);
}


function enviarAlertaEstoque(status, dados) {
  var subject = '';
  var htmlBody = '';
  if (status === 'minimo') {
    subject = `⚠️ Estoque Mínimo Atingido - ${dados.nome}`;
    htmlBody = renderEstoqueMinEmail(dados);
  } else if (status === 'zero') {
    subject = `🚨 Estoque Zerado - ${dados.nome}`;
    htmlBody = renderEstoqueZeroEmail(dados);
  }
  if (!subject) return;
  var emails = [
    //'spandrade@banestes.com.br',
    //'gadian@banestes.com.br',
    //'asmoreira@banestes.com.br',
    'csdamasceno@banestes.com.br'
  ];
  MailApp.sendEmail({ to: emails.join(','), subject: subject, htmlBody: htmlBody });
}

/*function renderEmailPedidoDiretoria(dados) {
  return `<div style="font-family: 'Segoe UI', Arial, sans-serif; color: #00284d; max-width: 600px; border: 1px solid #003366; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.1);">
    <div style="background-color: #003366; color: #ffffff; padding: 20px; text-align: center;">
      <h2 style="margin: 0; font-size: 18px; letter-spacing: 1px; text-transform: uppercase;">Solicitação de Pedido - Diretoria</h2>
    </div>
    <div style="padding: 25px; line-height: 1.6; background-color: #ffffff;">
      <p>Prezada equipe do <strong>Núcleo Asset</strong>,</p>
      <p>Uma nova requisição prioritária foi registrada via portal da Diretoria. Favor organizar o atendimento conforme os detalhes abaixo:</p>
      <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; padding: 20px; border-radius: 6px; margin: 20px 0;">
        <table style="width: 100%; border-collapse: collapse;">
          <tr>
            <td style="padding: 8px 0; border-bottom: 1px solid #edf2f7; font-weight: bold; color: #003366;">Item Solicitado:</td>
            <td style="padding: 8px 0; border-bottom: 1px solid #edf2f7;">${dados.nome}</td>
          </tr>
          <tr>
            <td style="padding: 8px 0; border-bottom: 1px solid #edf2f7; font-weight: bold; color: #003366;">Quantidade:</td>
            <td style="padding: 8px 0; border-bottom: 1px solid #edf2f7;">${dados.qtd} unidades</td>
          </tr>
          <tr>
            <td style="padding: 8px 0; border-bottom: 1px solid #edf2f7; font-weight: bold; color: #003366;">Data Máxima:</td>
            <td style="padding: 8px 0; border-bottom: 1px solid #edf2f7;">${dados.data}</td>
          </tr>
          <tr>
            <td style="padding: 12px 0 4px 0; font-weight: bold; color: #003366;" colspan="2">Justificativa / Observação:</td>
          </tr>
          <tr>
            <td style="padding: 8px; background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 4px; font-style: italic; color: #4a5568;" colspan="2">
              "${dados.obs || 'Nenhuma observação informada.'}"
            </td>
          </tr>
        </table>
      </div>
      <p style="font-size: 13px; color: #718096;">
        <strong>Solicitante:</strong> ${getActiveUserEmail()} <br>
        <strong>Sistema:</strong> Gerenciador de Ativos - Núcleo Asset
      </p>
    </div>
    <div style="background-color: #f4f7f9; padding: 12px; text-align: center; font-size: 11px; color: #a0aec0; border-top: 1px solid #e2e8f0;">
      Este é um e-mail automático de alta prioridade.
    </div>
  </div>`;
}*/

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

function debugReturnAppData() {
  let resultado = {};
  try {
    // Rode aqui a MESMA lógica de getAppData:
    const ss = SpreadsheetApp.openById(PLANILHA_ID);
    let cadastro = ss.getSheetByName('DB_Cadastro');
    let brindes = [];
    if (cadastro && cadastro.getLastRow() > 1) {
      brindes = cadastro.getRange(2,1,cadastro.getLastRow()-1,5).getValues()
        .map(row => ({ id: row[0], nome: row[1] }));
    }
    let usuarioEmail = "";
    try {
      usuarioEmail = Session.getActiveUser().getEmail();
      if (!usuarioEmail) throw "Usuário não detectado";
    } catch (e) {
      usuarioEmail = "Usuário não detectado";
    }

    resultado = {
      brindes: brindes,
      usuario: usuarioEmail,
      kpis: { estoqueBaixo: 1, pedidosPendentes: 2, divergencias: 3 },
      debug: {
        brindesLen: brindes.length,
        usuarioEmail: usuarioEmail
      }
    };
  } catch(e) {
    resultado = { brindes: [], usuario:"Erro", kpis:{}, debug:{error:e.toString()} };
  }
  Logger.log("==== RETORNO debugReturnAppData ====");
  Logger.log(JSON.stringify(resultado, null, 2));
  return resultado;
}

function debugReturnCompleto() {
  try {
    let resp = getAppData();
    Logger.log("==== RETORNO DE getAppData ====");
    Logger.log(JSON.stringify(resp, null, 2));
    return resp;
  } catch(e) {
    Logger.log("ERRO: " + e.toString());
    return {debug:{error:e.toString()}};
  }
}

function debugEstrutura() {
  let resp = getAppData();
  Logger.log("Tipo brindes: " + typeof resp.brindes + " / Array? " + Array.isArray(resp.brindes));
  Logger.log("Tipo usuario: " + typeof resp.usuario);
  Logger.log("Tipo kpis: " + typeof resp.kpis);
  Logger.log("Tipo dashboard: " + typeof resp.dashboard + " / Array? " + Array.isArray(resp.dashboard));
  Logger.log("Tipo movimentos: " + typeof resp.movimentos + " / Array? " + Array.isArray(resp.movimentos));
  Logger.log("Tipo logs: " + typeof resp.logs + " / Array? " + Array.isArray(resp.logs));
  Logger.log("Tipo debug: " + typeof resp.debug);
  Logger.log("==== DADOS: " + JSON.stringify(resp));
  return resp;
}

function testarTodosEmailsSistema() {
  var destinatario = 'spandrade@banestes.com.br';

  // 1. Nova Encomenda
  var dadosEncomenda = {
    nome: "Brinde Teste Encomenda",
    qtd: 10,
    prazo: "7",
    id: "999"
  };
  MailApp.sendEmail({
    to: destinatario,
    subject: '[TESTE] Nova Encomenda',
    htmlBody: renderEmailEncomenda(dadosEncomenda)
  });

  // 1b. Transferência
  var dadosTransferencia = {
    nome: "Brinde Teste Transferência",
    qtd: 5,
    prazo: "3",
    id: "999"
  };
  MailApp.sendEmail({
    to: destinatario,
    subject: '[TESTE] Transferência de Estoque',
    htmlBody: renderEmailTransferencia(dadosTransferencia)
  });

  // 2. Divergência à COSUP
  var dadosDivergencia = {
    nome: "Brinde Divergente",
    qtdRequisitada: 15,
    qtdRecebida: 10,
    incidente: "Material faltante - avaria.",
    id: "998"
  };
  MailApp.sendEmail({
    to: destinatario,
    subject: '[TESTE] Notificação de Divergência à COSUP',
    htmlBody: renderEmailDivergenciaCosup(dadosDivergencia)
  });

  // 3. Regularização Manual
  var dadosRegularizacao = {
    nome: "Brinde Regularização",
    id: "997",
    novoSaldo: 20,
    justificativa: "Aceite por ajuste de inventário"
  };
  MailApp.sendEmail({
    to: destinatario,
    subject: '[TESTE] Regularização Manual',
    htmlBody: renderEmailRegularizacao(dadosRegularizacao)
  });

  // 4. Recebimento Excedente (A Mais)
  var dadosExcedente = {
    nome: "Brinde Excedente",
    id: "996",
    qtdRequisitada: 7,
    qtdRecebida: 12
  };
  MailApp.sendEmail({
    to: destinatario,
    subject: '[TESTE] Recebimento Excedente (A Mais)',
    htmlBody: renderEmailExcedente(dadosExcedente)
  });

  // 5. Estoque Mínimo
  var dadosMinimo = {
    nome: "Brinde Minimo",
    id: "995",
    saldo_nucleo: 2,
    min: 5
  };
  MailApp.sendEmail({
    to: destinatario,
    subject: '[TESTE] Alerta de Estoque Mínimo',
    htmlBody: renderEstoqueMinEmail(dadosMinimo)
  });

  // 6. Estoque Zerado
  var dadosZero = {
    nome: "Brinde Zerado",
    id: "994"
  };
  MailApp.sendEmail({
    to: destinatario,
    subject: '[TESTE] Alerta de Estoque Zerado',
    htmlBody: renderEstoqueZeroEmail(dadosZero)
  });

  // 7. Saída de Item
  var dadosSaida = {
    nome: "Brinde Saída Teste",
    qtd: 3,
    retirante: "João da Silva",
    dataHora: Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd HH:mm:ss"),
    usuario: "spandrade@banestes.com.br"
  };
  MailApp.sendEmail({
    to: destinatario,
    subject: '[TESTE] Saída de Item',
    htmlBody: renderEmailSaida(dadosSaida)
  });

  Logger.log('[TESTE] Todos os modelos de e-mail enviados para ' + destinatario);
}

function adicionarItensEstoqueLista() {
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
  criarAbasBanco(ss);
  const cadastro = ss.getSheetByName('DB_Cadastro');
  const estoqueN = ss.getSheetByName('DB_Estoque_Nucleo');

  // Lista de itens [nome, qtd]
  const itens = [
    ["BOLSA TÉRMICA AZUL", 5],
    ["BONÉ BORDADO AZUL", 120],
    ["BONÉ BORDADO BRANCO", 20],
    ["BONÉ TRUCKER", 0],
    ["CADEIRA DE PRAIA RECLINÁVEL (NOVA)", 1],
    ["CADEIRA DE PRAIA RECLINÁVEL (ANTIGA)", 1],
    ["GARRAFA TÉRMICA ECOLÓGICA", 3],
    ["GUARDA-SOL PEQUENO", 3],
    ["GUARDA-SOL GRANDE", 3],
    ["KIT CHURRASCO", 11],
    ["MOCHILA TÉRMICA CINZA", 5],
    ["MALA DE VIAGEM", 4],
    ["KIT CANETA E LAPISEIRA", 78]
  ];

  itens.forEach(item => {
    // Gera ID único (100~999 que não esteja em uso)
    let used = cadastro.getLastRow() > 1 ? cadastro.getRange(2,1,cadastro.getLastRow()-1,1).getValues().map(r=>r[0]) : [];
    let id;
    do {
      id = String(100 + Math.floor(Math.random() * 900));
    } while (used.includes(id));

    // Cadastro padrão: [ID, Nome, Estoque_Minimo (0), Prazo_Maximo (""), Data_Criacao]
    cadastro.appendRow([id, item[0], 0, "", Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd")]);
    estoqueN.appendRow([id, item[0], item[1]]);
  });
}
