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
      estoqueRows = estoqueN.getRange(2, 1, estoqueN.getLastRow() - 1, 3).getValues();
      Logger.log("Linhas em DB_Estoque_Nucleo: " + estoqueRows.length);
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
    }

    // DASHBOARD (comparativo)
    let dashboard = [];
    for (let cad of cadastroRows) {
      const id = cad[0], nome = cad[1], min = Number(cad[2]), prazo = cad[3], data_criacao = cad[4];
      const buscaNome = String(nome || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();

      // Saldo núcleo e COSUP
      const saldo_nucleo = (() => {
        let found = estoqueRows.find(row => String(row[0]) == String(id));
        return found ? Number(found[2]) : 0;
      })();
      const saldo_cosup = (() => {
        let found = cosupRows.find(row => row.nome == buscaNome);
        return found ? Number(found.cosup) : 0;
      })();

      // STATUS do estoque núcleo
      let status = '';
      if (saldo_nucleo === 0) status = 'Zero';
      else if (saldo_nucleo <= min) status = 'Baixo';
      else status = 'Ok';

      dashboard.push({
        id: String(id || ''),
        nome: String(nome || ''),
        saldo_nucleo,
        saldo_cosup,
        total_estoque: (Number(saldo_nucleo) + Number(saldo_cosup)),
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
  if (espelho.getLastRow()>2) {
    let vals = espelho.getRange(3,1,espelho.getLastRow()-2,2).getValues()
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
  const ss = SpreadsheetApp.openById(PLANILHA_ID);
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
  dispararEmail('Regularizacao', {
    id: dados.id,
    nome: dados.nome,
    novoSaldo: dados.novoSaldo,
    justificativa: dados.justificativa
  });
  logAcao('Regularização manual', emailUser, `Brinde: ${dados.nome}, Novo saldo: ${dados.novoSaldo}, Just: ${dados.justificativa||''}`);
  return {sucesso:true, mensagem:"Regularização realizada."};
}

/** ------ PEDIDO ESPECIAL (diretoria) ----- */
function registrarPedidoEspecial(dados) {
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
}

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
    body = renderEmailFormalizacaoInterna(tipo, dados);
    emails = [
      'spandrade@banestes.com.br',
      //'gadian@banestes.com.br',
      //'asmoreira@banestes.com.br',
      //'csdamasceno@banestes.com.br'
    ];
  } else if(tipo === 'Transferência') {
    subject = `Transferência Efetuada - Brinde ${dados.nome}`;
    body = renderEmailFormalizacaoInterna(tipo, dados);
    emails = [
      'spandrade@banestes.com.br',
      //'gadian@banestes.com.br',
      //'asmoreira@banestes.com.br',
      //'csdamasceno@banestes.com.br'
    ];
  } else if(tipo === 'Desconformidade') {
    subject = `Divergência Recebida - Brinde ${dados.nome}`;
    body = renderEmailDivergenciaCosup(dados);
    emails = [
      'spandrade@banestes.com.br',
      //'enviocosup@banestes.com.br'
    ];
  } else if(tipo === 'Alinhamento') {
    subject = `Recebimento Excedente - Brinde ${dados.nome}`;
    body = renderEmailExcedente(dados);
    emails = [
      'spandrade@banestes.com.br',
      //'gadian@banestes.com.br',
      //'asmoreira@banestes.com.br'
    ];
  } else if(tipo === 'Regularizacao') {
    subject = `Regularização Manual`;
    body = renderEmailRegularizacao(dados);
    emails = [
      'spandrade@banestes.com.br',
      //'gadian@banestes.com.br',
      //'asmoreira@banestes.com.br'
    ];
  } else {
    subject = `Alerta sistema - Brinde ${dados.nome||''}`;
    body = `<pre>${JSON.stringify(dados)}</pre>`;
    emails = [
      'spandrade@banestes.com.br',
      //'gadian@banestes.com.br'
    ];
  }

  MailApp.sendEmail({
    to: emails.join(","),
    subject: subject,
    htmlBody: body
  });
}

// Templates de email (copie do seu exemplo e use string template)
function renderEmailFormalizacaoInterna(tipo, dados) {
  return `<div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #00284d; max-width: 600px; border: 1px solid #e4e8ee; border-radius: 8px; overflow: hidden;">
  <div style="background-color: #003366; color: #ffffff; padding: 20px; text-align: center;">
    <h2 style="margin: 0; font-size: 18px; letter-spacing: 1px;">RELATÓRIO DE MOVIMENTAÇÃO INTERNA</h2>
  </div>
  <div style="padding: 25px; line-height: 1.6;">
    <p>Prezados,</p>
    <p>Informamos que uma nova ação foi registrada no <strong>Gerenciador de Ativos - Núcleo Asset</strong> e requer sua ciência:</p>
    <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
      <tr>
        <td style="padding: 8px; border-bottom: 1px solid #f4f7f9; font-weight: bold;">Evento:</td>
        <td style="padding: 8px; border-bottom: 1px solid #f4f7f9;">${tipo}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border-bottom: 1px solid #f4f7f9; font-weight: bold;">Item:</td>
        <td style="padding: 8px; border-bottom: 1px solid #f4f7f9;">${dados.nome}</td>
      </tr>
      <tr>
        <td style="padding: 8px; border-bottom: 1px solid #f4f7f9; font-weight: bold;">Solicitante:</td>
        <td style="padding: 8px; border-bottom: 1px solid #f4f7f9;">${getActiveUserEmail()}</td>
      </tr>
    </table>
  </div>
  <div style="background-color: #f4f7f9; padding: 15px; text-align: center; font-size: 12px; color: #7a8fa6;">
    Este é um e-mail automático do Sistema de Gestão de Brindes - Núcleo Asset.
  </div>
</div>`; // coloque o HTML fornecido no seu exemplo!
}
function renderEmailDivergenciaCosup(dados) {
  return `<div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; border: 2px solid #fa4444; border-radius: 8px;">
  <div style="background-color: #fa4444; color: #ffffff; padding: 15px;">
    <h3 style="margin: 0;">ALERTA: Divergência de Recebimento</h3>
  </div>
  <div style="padding: 20px;">
    <p>Prezada equipe <strong>COSUP</strong>,</p>
    <p>Identificamos uma desconformidade entre a guia de remessa e o material físico recebido no Núcleo:</p>
    <div style="background-color: #fff4f4; border-left: 4px solid #fa4444; padding: 15px; margin: 15px 0;">
      <strong>Item:</strong> ${dados.nome}<br>
      <strong>Qtd. Esperada:</strong> ${dados.qtdRequisitada}<br>
      <strong>Qtd. Efetiva:</strong> ${dados.qtdRecebida}<br>
      <br>
      <strong>Incidente Registrado:</strong><br>
      <i style="color: #555;">"${dados.incidente}"</i>
    </div>
    <p>Solicitamos a verificação no estoque central para alinhamento de saldos.</p>
  </div>
</div>`;
}
function renderEmailRegularizacao(dados) {
  return `<div style="font-family: 'Segoe UI', sans-serif; color: #00284d; max-width: 600px; border: 1px solid #2684ff;">
  <div style="background-color: #2684ff; color: #ffffff; padding: 15px;">
    <h3 style="margin: 0;">Formalização de Ajuste Manual (Regularização)</h3>
  </div>
  <div style="padding: 20px;">
    <p>Informamos que o item abaixo foi <strong>regularizado manualmente</strong> no sistema, alterando o status de <i>Desconformidade</i> para <i>Em Estoque</i> conforme acordado entre as partes.</p>
    
    <table style="width: 100%; background: #f0f7ff; border-radius: 4px; padding: 15px;">
      <tr><td><strong>Brinde:</strong></td><td>${dados.nome} (ID: ${dados.id})</td></tr>
      <tr><td><strong>Saldo Final Ajustado:</strong></td><td>${dados.novoSaldo} unidades</td></tr>
      <tr><td><strong>Responsável:</strong></td><td>${getActiveUserEmail()}</td></tr>
    </table>

    <div style="margin-top: 20px; border-top: 1px dashed #2684ff; padding-top: 10px;">
      <strong>Justificativa da Aceite/Troca:</strong><br>
      <p style="background: #ffffff; border: 1px solid #e4e8ee; padding: 10px; border-radius: 4px;">
        ${dados.justificativa}
      </p>
    </div>
    <p style="font-size: 11px; color: #7a8fa6;">*Este registro substitui a pendência anterior e encerra o incidente de desconformidade para este lote.</p>
  </div>
</div>`;
}
function renderEmailExcedente(dados) {
  return `<div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #00284d; max-width: 600px; border: 1px solid #36b37e; border-radius: 8px; overflow: hidden;">
  <div style="background-color: #36b37e; color: #ffffff; padding: 20px; text-align: center;">
    <h2 style="margin: 0; font-size: 18px; letter-spacing: 1px;">NOTIFICAÇÃO DE EXCEDENTE DE RECEBIMENTO</h2>
  </div>
  <div style="padding: 25px; line-height: 1.6;">
    <p>Prezados,</p>
    <p>Durante a conferência de entrada no <strong>Núcleo Asset</strong>, identificamos o recebimento de uma <strong>quantidade superior</strong> à solicitada originalmente.</p>
    
    <div style="background-color: #f0fff4; border: 1px solid #c6f6d5; padding: 15px; border-radius: 5px; margin: 15px 0;">
      <table style="width: 100%; border-collapse: collapse;">
        <tr>
          <td style="padding: 5px; font-weight: bold;">Item:</td>
          <td style="padding: 5px;">${dados.nome} (ID: ${dados.id})</td>
        </tr>
        <tr>
          <td style="padding: 5px; font-weight: bold;">Qtd. Solicitada:</td>
          <td style="padding: 5px; color: #555;">${dados.qtdRequisitada}</td>
        </tr>
        <tr>
          <td style="padding: 5px; font-weight: bold; color: #276749;">Qtd. Recebida:</td>
          <td style="padding: 5px; font-weight: bold; color: #276749;">${dados.qtdRecebida}</td>
        </tr>
        <tr>
          <td style="padding: 5px; font-weight: bold; color: #c53030;">Diferença (A maior):</td>
          <td style="padding: 5px; font-weight: bold; color: #c53030;">+ ${dados.qtdRecebida - dados.qtdRequisitada}</td>
        </tr>
      </table>
    </div>

    <p><strong>Ação tomada:</strong> O saldo excedente foi incorporado ao estoque atual após validação física.</p>
    
    <p style="font-size: 13px; color: #666;">
      <strong>Responsável pela conferência:</strong> ${getActiveUserEmail()}<br>
      <strong>Data/Hora:</strong> ${new Date().toLocaleString('pt-BR')}
    </p>
  </div>
  <div style="background-color: #f4f7f9; padding: 15px; text-align: center; font-size: 12px; color: #7a8fa6;">
    Este é um registro automático de conformidade do Sistema Núcleo Asset.
  </div>
</div>`;
}

function enviarAlertaEstoque(status, dados) {
  var subject = '';
  var htmlBody = '';
  if (status === 'minimo') {
    subject = `AVISO: Estoque Mínimo Atingido - ${dados.nome}`;
    htmlBody = renderEstoqueMinEmail(dados);
  } else if (status === 'zero') {
    subject = `ALERTA: Estoque Zerado - ${dados.nome}`;
    htmlBody = renderEstoqueZeroEmail(dados);
  }
  var emails = [
    'spandrade@banestes.com.br',
    //'gadian@banestes.com.br',
    //'asmoreira@banestes.com.br',
    //'csdamasceno@banestes.com.br'
  ];
  MailApp.sendEmail({ to: emails.join(","), subject: subject, htmlBody: htmlBody });
}
function renderEstoqueMinEmail(dados) {
  return `<div style="font-family: 'Segoe UI', Arial, sans-serif; color: #00284d; max-width: 600px; border: 1px solid #ffd166; border-radius: 8px; overflow: hidden;">
  <div style="background-color: #ffd166; color: #003366; padding: 20px; text-align: center;">
    <h2 style="margin: 0; font-size: 18px; letter-spacing: 1px;">⚠️ ALERTA: ESTOQUE MÍNIMO ATINGIDO</h2>
  </div>
  <div style="padding: 25px; line-height: 1.6;">
    <p>Prezados,</p>
    <p>O item abaixo atingiu o <strong>nível crítico de segurança</strong> no estoque do Núcleo e necessita de reposição imediata:</p>
    
    <div style="background-color: #fffbeb; border-left: 4px solid #ffd166; padding: 15px; margin: 15px 0;">
      <strong>Brinde:</strong> ${dados.nome} <br>
      <strong>ID:</strong> ${dados.id} <br>
      <hr style="border: 0; border-top: 1px solid #eee; margin: 10px 0;">
      <strong>Saldo Atual:</strong> <span style="color: #d69e2e; font-weight: bold;">${dados.saldo_nucleo}</span><br>
      <strong>Estoque Mínimo:</strong> ${dados.min}
    </div>

    <p>Sugerimos a abertura de uma nova <strong>Encomenda</strong> ou <strong>Transferência COSUP</strong> através do sistema para evitar a ruptura total.</p>
  </div>
  <div style="background-color: #f4f7f9; padding: 15px; text-align: center; font-size: 11px; color: #7a8fa6;">
    Gerenciado por Núcleo Asset - Sistema de Controle de Ativos.
  </div>
</div>`;
}
function renderEstoqueZeroEmail(dados) {
  return `<div style="font-family: 'Segoe UI', Arial, sans-serif; color: #00284d; max-width: 600px; border: 2px solid #fa4444; border-radius: 8px; overflow: hidden;">
  <div style="background-color: #fa4444; color: #ffffff; padding: 20px; text-align: center;">
    <h2 style="margin: 0; font-size: 18px; letter-spacing: 1px;">🚨 CRÍTICO: ESTOQUE ZERADO</h2>
  </div>
  <div style="padding: 25px; line-height: 1.6;">
    <p>Prezados,</p>
    <p style="color: #fa4444; font-weight: bold;">Atenção: O item abaixo está com saldo zerado no estoque físico do Núcleo.</p>
    
    <div style="background-color: #fff5f5; border: 1px solid #feb2b2; padding: 15px; border-radius: 5px; margin: 15px 0;">
      <table style="width: 100%;">
        <tr>
          <td><strong>Item:</strong></td>
          <td>${dados.nome}</td>
        </tr>
        <tr>
          <td><strong>ID:</strong></td>
          <td>${dados.id}</td>
        </tr>
        <tr>
          <td><strong>Status:</strong></td>
          <td><span style="background: #fa4444; color: white; padding: 2px 6px; border-radius: 4px; font-size: 12px;">INDISPONÍVEL</span></td>
        </tr>
      </table>
    </div>

    <p>Qualquer nova solicitação de saída para este brinde será negada até que a regularização ou nova entrada seja processada.</p>
  </div>
  <div style="border-top: 1px solid #eee; padding: 15px; font-size: 12px; color: #666; font-style: italic;">
    Relatório gerado em: ${new Date().toLocaleString('pt-BR')}
  </div>
</div>`;
}

function renderEmailPedidoDiretoria(dados) {
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
