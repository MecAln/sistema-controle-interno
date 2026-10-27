// ==========================================
// CONFIGURAÇÕES GERAIS E SEGURANÇA
// ==========================================
const SPREADSHEET_ID = '1qVLCvjaKJiAidxKjtfq17D40XR3a5Vr3kUBhQ9ytjww';
const MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Controle Interno')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getVerdadeiraUltimaLinha(sheet) {
  const maxRow = sheet.getLastRow();
  if (maxRow === 0) return 1;
  const data = sheet.getRange(1, 1, maxRow, 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== "" && data[i][0] != null) return i + 1;
  }
  return 1; 
}

// ==========================================
// LOGS E AUDITORIA (NOVO)
// ==========================================
function registrarLog(usuario, acao, detalhes) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let aba = ss.getSheetByName("LOGS_SISTEMA");
    if (!aba) {
      aba = ss.insertSheet("LOGS_SISTEMA");
      aba.appendRow(["DATA/HORA", "USUÁRIO", "AÇÃO", "DETALHES"]);
      aba.getRange("A1:D1").setBackground("#000000").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
      aba.setFrozenRows(1);
    }
    const novaLinha = getVerdadeiraUltimaLinha(aba) + 1;
    aba.getRange(novaLinha, 1, 1, 4).setValues([[new Date(), usuario, acao, detalhes]]);
  } catch (e) {
    Logger.log("Erro ao registrar log: " + e.message);
  }
}

function validarAdmin(usuarioLogado) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaUsuarios = ss.getSheetByName("USUARIOS");
  if (!abaUsuarios) throw new Error("Base de usuários não encontrada.");
  
  const dados = abaUsuarios.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).trim().toLowerCase() === String(usuarioLogado).trim().toLowerCase()) {
      if (String(dados[i][2]).trim() === "Admin") return true;
      throw new Error("Acesso Negado: Privilégios de Administrador requeridos.");
    }
  }
  throw new Error("Usuário não encontrado.");
}

// ==========================================
// SISTEMA DE LOGIN E AUTENTICAÇÃO (ATUALIZADO)
// ==========================================
function autenticarUsuario(usuario, senha) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let abaUsuarios = ss.getSheetByName("USUARIOS");
  
  if (!abaUsuarios) {
    abaUsuarios = ss.insertSheet("USUARIOS");
    abaUsuarios.appendRow(["NOME", "SENHA", "NIVEL", "ACESSOS"]);
    abaUsuarios.appendRow(["Admin", "1234", "Admin", "tab-reembolso,tab-revertido,tab-bloqueio,tab-plp,tab-procedimentos,tab-relatorios,tab-admin"]);
  }
  
  const dados = abaUsuarios.getDataRange().getValues();
  
  for (let i = 1; i < dados.length; i++) {
    const nomePlanilha = String(dados[i][0]).trim();
    const senhaPlanilha = String(dados[i][1]).trim();
    const nivelAcesso = String(dados[i][2] || "Analista").trim();
    const acessos = dados[i][3] ? String(dados[i][3]).trim() : "tab-reembolso,tab-revertido,tab-bloqueio,tab-procedimentos";
    
    if (nomePlanilha.toLowerCase() === usuario.toLowerCase() && senhaPlanilha === senha) {
      registrarLog(nomePlanilha, "LOGIN", "Usuário acessou o sistema.");
      return { nome: nomePlanilha, nivel: nivelAcesso, acessos: acessos };
    }
  }
  throw new Error("Usuário ou senha incorretos.");
}

// ==========================================
// ADMINISTRAÇÃO - GESTÃO DE USUÁRIOS
// ==========================================
function getUsuarios(adminLogado) {
  validarAdmin(adminLogado);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const abaUsuarios = ss.getSheetByName("USUARIOS");
  const dados = abaUsuarios.getDataRange().getValues();
  const usuarios = [];
  
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0]) {
      usuarios.push({
        nome: String(dados[i][0]).trim(),
        senha: String(dados[i][1]).trim(),
        nivel: String(dados[i][2] || "Analista").trim(),
        acessos: String(dados[i][3] || "").trim()
      });
    }
  }
  return usuarios;
}

function salvarUsuario(adminLogado, dadosUsuario) {
  validarAdmin(adminLogado);
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { throw new Error("Sistema ocupado."); }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = ss.getSheetByName("USUARIOS");
    const dados = aba.getDataRange().getValues();
    
    // Atualização se o usuário já existe
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim().toLowerCase() === dadosUsuario.nome.toLowerCase()) {
        aba.getRange(i + 1, 2, 1, 3).setValues([[dadosUsuario.senha, dadosUsuario.nivel, dadosUsuario.acessos]]);
        registrarLog(adminLogado, "EDITAR_USUARIO", `Alterou permissões/senha de ${dadosUsuario.nome}.`);
        return "Usuário atualizado com sucesso!";
      }
    }
    
    const novaLinha = getVerdadeiraUltimaLinha(aba) + 1;
    aba.getRange(novaLinha, 1, 1, 4).setValues([[dadosUsuario.nome, dadosUsuario.senha, dadosUsuario.nivel, dadosUsuario.acessos]]);
    registrarLog(adminLogado, "CRIAR_USUARIO", `Criou o usuário ${dadosUsuario.nome}.`);
    return "Novo usuário cadastrado com sucesso!";
  } finally {
    lock.releaseLock();
  }
}

function deletarUsuario(adminLogado, nomeUsuario) {
  validarAdmin(adminLogado);
  if(String(adminLogado).trim().toLowerCase() === String(nomeUsuario).trim().toLowerCase()) {
    throw new Error("Você não pode excluir seu próprio usuário.");
  }
  
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { throw new Error("Sistema ocupado."); }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = ss.getSheetByName("USUARIOS");
    const dados = aba.getDataRange().getValues();
    
    for (let i = dados.length - 1; i >= 1; i--) {
      if (String(dados[i][0]).trim().toLowerCase() === String(nomeUsuario).trim().toLowerCase()) {
        aba.deleteRow(i + 1);
        registrarLog(adminLogado, "DELETAR_USUARIO", `Deletou o usuário ${nomeUsuario}.`);
        return "Usuário excluído com sucesso!";
      }
    }
    throw new Error("Usuário não encontrado.");
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// ADMINISTRAÇÃO - GESTÃO DE TRANSPORTADORAS
// ==========================================
function garantirAbasTransporte() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let abaConfig = ss.getSheetByName("CONFIG_TRANSPORTADORAS");
  if (!abaConfig) {
    abaConfig = ss.insertSheet("CONFIG_TRANSPORTADORAS");
    abaConfig.appendRow(["NOME TRANSPORTADORA", "EMAILS DESTINO"]);
    abaConfig.getRange("A1:B1").setBackground("#000000").setFontColor("#FFFFFF").setFontWeight("bold");
    abaConfig.setFrozenRows(1);
  }
  let abaRegistro = ss.getSheetByName("REGISTRO_BARRADOS");
  if (!abaRegistro) {
    abaRegistro = ss.insertSheet("REGISTRO_BARRADOS");
    abaRegistro.appendRow(["DATA/HORA", "ANALISTA", "TRANSPORTADORA", "NF", "TICKET", "ASSUNTO ENVIADO", "EMAILS DESTINO"]);
    abaRegistro.getRange("A1:G1").setBackground("#000000").setFontColor("#FFFFFF").setFontWeight("bold");
  }
}

function getTodasTransportadoras() {
  garantirAbasTransporte();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dados = ss.getSheetByName("CONFIG_TRANSPORTADORAS").getDataRange().getValues();
  const lista = [];
  for(let i = 1; i < dados.length; i++) {
    if(dados[i][0]) lista.push({ nome: String(dados[i][0]).trim().toUpperCase(), emails: String(dados[i][1]).trim() });
  }
  return lista.sort((a,b) => a.nome.localeCompare(b.nome));
}

function salvarTransportadora(adminLogado, dadosTransportadora) {
  validarAdmin(adminLogado);
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { throw new Error("Sistema ocupado."); }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = ss.getSheetByName("CONFIG_TRANSPORTADORAS");
    const dados = aba.getDataRange().getValues();
    const nomeBusca = dadosTransportadora.nomeOriginal || dadosTransportadora.nome;
    
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim().toUpperCase() === String(nomeBusca).toUpperCase()) {
        aba.getRange(i + 1, 1, 1, 2).setValues([[dadosTransportadora.nome.toUpperCase(), dadosTransportadora.emails]]);
        registrarLog(adminLogado, "EDITAR_TRANSPORTADORA", `Atualizou a transportadora ${dadosTransportadora.nome}.`);
        return "Transportadora atualizada com sucesso!";
      }
    }
    
    const novaLinha = getVerdadeiraUltimaLinha(aba) + 1;
    aba.getRange(novaLinha, 1, 1, 2).setValues([[dadosTransportadora.nome.toUpperCase(), dadosTransportadora.emails]]);
    registrarLog(adminLogado, "CRIAR_TRANSPORTADORA", `Adicionou a transportadora ${dadosTransportadora.nome}.`);
    return "Nova transportadora cadastrada!";
  } finally {
    lock.releaseLock();
  }
}

function deletarTransportadora(adminLogado, nomeTransp) {
  validarAdmin(adminLogado);
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { throw new Error("Sistema ocupado."); }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const aba = ss.getSheetByName("CONFIG_TRANSPORTADORAS");
    const dados = aba.getDataRange().getValues();
    
    for (let i = dados.length - 1; i >= 1; i--) {
      if (String(dados[i][0]).trim().toUpperCase() === String(nomeTransp).trim().toUpperCase()) {
        aba.deleteRow(i + 1);
        registrarLog(adminLogado, "DELETAR_TRANSPORTADORA", `Removeu a transportadora ${nomeTransp}.`);
        return "Transportadora excluída com sucesso!";
      }
    }
    throw new Error("Transportadora não encontrada.");
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// BASE DE CONHECIMENTO (PROCEDIMENTOS)
// ==========================================
function garantirAbaProcedimentos() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let aba = ss.getSheetByName("BASE_PROCEDIMENTOS");
  if (!aba) {
    aba = ss.insertSheet("BASE_PROCEDIMENTOS");
    aba.appendRow(["ID", "CATEGORIA", "TITULO", "CONTEUDO", "AUTOR", "DATA_ATUALIZACAO"]);
    aba.getRange("A1:F1").setBackground("#000000").setFontColor("#FFFFFF").setFontWeight("bold");
    aba.setFrozenRows(1);
  }
  return aba;
}

function getProcedimentos() {
  const aba = garantirAbaProcedimentos();
  const lastRow = getVerdadeiraUltimaLinha(aba);
  if (lastRow < 2) return [];
  const dados = aba.getRange(2, 1, lastRow - 1, 6).getValues();
  const procs = [];
  for (let i = 0; i < dados.length; i++) {
    if (dados[i][0]) {
      procs.push({
        id: String(dados[i][0]),
        categoria: String(dados[i][1]),
        titulo: String(dados[i][2]),
        conteudo: String(dados[i][3]),
        autor: String(dados[i][4]),
        data: Utilities.formatDate(new Date(dados[i][5]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
      });
    }
  }
  return procs.reverse(); 
}

function salvarProcedimento(dados, analistaLogado) {
  validarAdmin(analistaLogado); 
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { throw new Error("Sistema ocupado."); }

  try {
    const aba = garantirAbaProcedimentos();
    const agora = new Date();
    
    if (dados.id) {
      const lastRow = getVerdadeiraUltimaLinha(aba);
      const grid = aba.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < grid.length; i++) {
        if (String(grid[i][0]) === dados.id) {
          aba.getRange(i + 2, 2, 1, 5).setValues([[dados.categoria, dados.titulo, dados.conteudo, analistaLogado, agora]]);
          registrarLog(analistaLogado, "EDITAR_PROCEDIMENTO", `Editou o procedimento: ${dados.titulo}`);
          return "Procedimento atualizado com sucesso!";
        }
      }
      throw new Error("Procedimento não encontrado.");
    } else {
      const novoId = Utilities.getUuid();
      const novaLinha = getVerdadeiraUltimaLinha(aba) + 1;
      aba.getRange(novaLinha, 1, 1, 6).setValues([[novoId, dados.categoria, dados.titulo, dados.conteudo, analistaLogado, agora]]);
      registrarLog(analistaLogado, "CRIAR_PROCEDIMENTO", `Criou o procedimento: ${dados.titulo}`);
      return "Novo procedimento criado com sucesso!";
    }
  } finally {
    lock.releaseLock();
  }
}

function deletarProcedimento(id, analistaLogado) {
  validarAdmin(analistaLogado);
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { throw new Error("Sistema ocupado."); }

  try {
    const aba = garantirAbaProcedimentos();
    const lastRow = getVerdadeiraUltimaLinha(aba);
    const grid = aba.getRange(2, 1, lastRow - 1, 1).getValues();
    
    for (let i = grid.length - 1; i >= 0; i--) {
      if (String(grid[i][0]) === id) {
        aba.deleteRow(i + 2);
        registrarLog(analistaLogado, "DELETAR_PROCEDIMENTO", `Deletou o procedimento ID: ${id}`);
        return "Procedimento excluído com sucesso!";
      }
    }
    throw new Error("Procedimento não encontrado.");
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// GESTÃO DE REVERSÕES
// ==========================================
function getReversoesAtuais(analistaLogado) {
  const now = new Date();
  const sheetName = 'rev' + MESES[now.getMonth()] + now.getFullYear();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const result = [];
  for(let i = 1; i < data.length; i++) {
    if(data[i][2] && String(data[i][4]).trim() === String(analistaLogado).trim()) {
      result.push({
        dataStr: Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        tipo: data[i][1],
        pedido: data[i][2],
        analista: data[i][4],
        valor: data[i][7]
      });
    }
  }
  return result.reverse(); 
}

function deletarReversaoSegura(pedido, analistaLogado) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { throw new Error("Sistema ocupado."); }
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const now = new Date();
    const sheetName = 'rev' + MESES[now.getMonth()] + now.getFullYear();
    const sheet = ss.getSheetByName(sheetName);
    
    if(!sheet) throw new Error("Planilha do mês atual não encontrada.");

    const textFinder = sheet.createTextFinder(String(pedido).trim()).matchEntireCell(true);
    const result = textFinder.findNext();
    
    if(!result) throw new Error("Pedido não encontrado na base.");

    const row = result.getRow();
    const donoDoRegistro = String(sheet.getRange(row, 5).getValue()).trim(); // Coluna 5 é o Analista

    if(donoDoRegistro !== String(analistaLogado).trim()) {
       throw new Error("Você só pode excluir suas próprias reversões.");
    }

    sheet.deleteRow(row); 
    registrarLog(analistaLogado, "DELETAR_REVERSAO", `Excluiu a reversão do pedido ${pedido}.`);
    return "Registro excluído com sucesso!";
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// FORMULÁRIOS: REEMBOLSO E REVERTIDOS
// ==========================================
function salvarFormulario(formData, analistaLogado) {
  let anexoUrl = "";
  if (formData.fileBase64) {
    try {
      let blob = Utilities.newBlob(Utilities.base64Decode(formData.fileBase64), formData.mimeType, formData.fileName);
      let file = DriveApp.createFile(blob);
      anexoUrl = file.getUrl();
    } catch(e) { Logger.log("Erro Drive: " + e); }
  }
  delete formData.fileBase64;
  formData.anexoGeradoUrl = anexoUrl;

  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { throw new Error("Sistema com alto tráfego. Tente novamente."); }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const now = new Date();
    const sheetName = 'dev' + MESES[now.getMonth()] + now.getFullYear();

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headerValues = ['DATA', 'PEDIDO', 'CC', 'NF', 'DEV', 'ANALISTA', 'TICKET', 'NOME', 'VALOR TOTAL', 'FORMA DE PAGAMENTO', 'MOTIVO GERAL', 'MOTIVO', 'STATUS', 'CONDIÇÃO DO PRODUTO', 'DADOS BANCÁRIOS', 'LINK ANEXO'];
      sheet.getRange(1, 1, 1, headerValues.length).setValues([headerValues]).setBackground('#000000').setFontColor('#FFFFFF').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const row = [
      now, formData.pedido, formData.cc, formData.nf, formData.dev,
      analistaLogado, formData.ticket, formData.nome, formatarValorContabil(formData.valorTotal), 
      formData.formaPagamento, formData.motivoGeral, formData.motivo, 'REEMBOLSO PENDENTE', formData.condicaoProduto,
      formData.dadosBancarios || "", anexoUrl
    ];

    const novaLinha = getVerdadeiraUltimaLinha(sheet) + 1;
    sheet.getRange(novaLinha, 1, 1, row.length).setValues([row]).setFontWeight('bold').setBackground('#fff2cc');
    
    return formData;
  } finally {
    lock.releaseLock();
  }
}

function criarValidarDadosEEnviarEmail(formData, analistaLogado) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { throw new Error("Sistema ocupado."); }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. Atualizar Status na aba do mês
    const now = new Date();
    const sheetName = 'dev' + MESES[now.getMonth()] + now.getFullYear();
    const sheet = ss.getSheetByName(sheetName);
    
    if (sheet) {
      const sheetData = sheet.getDataRange().getValues();
      for (let i = sheetData.length - 1; i >= 1; i--) {
        if (String(sheetData[i][1]).trim() === String(formData.pedido).trim()) {
          sheet.getRange(i + 1, 13).setValue('SOLICITADO'); // Status
          sheet.getRange(i + 1, 1, 1, 16).setBackground('#d9ead3'); // Cor Verde
          break;
        }
      }
    }

    // 2. Disparar o Email
    const dadosBanc = formData.dadosBancarios || 'Não informados';
    const corpo = `PEDIDO Nº: ${formData.pedido}\nNOME DO CLIENTE: ${formData.nome}\nNF: ${formData.nf}\nNF DEV: ${formData.dev}\nCC: ${formData.cc}\nVALOR TOTAL: R$ ${formatarValorContabil(formData.valorTotal)}\nMOTIVO GERAL: ${formData.motivoGeral}\nMOTIVO: ${formData.motivo}\nFORMA DE PAGAMENTO: ${formData.formaPagamento}\nCONDIÇÃO DO PRODUTO: ${formData.condicaoProduto}\n\nDADOS BANCÁRIOS:\n${dadosBanc}\n\nE-mail encaminhado automaticamente pelo Controle Interno! (Analista: ${analistaLogado})`;

    let anexosParaEmail = [];
    if (formData.anexoGeradoUrl) {
      try {
        const fileId = formData.anexoGeradoUrl.match(/[-\w]{25,}/);
        if (fileId) anexosParaEmail.push(DriveApp.getFileById(fileId[0]).getBlob());
      } catch(e) {}
    }

    MailApp.sendEmail({
      to: "xx", // coloque o email a ser enviado aqui
      cc: "xx", // email que fica em copia
      subject: `DEVOLUÇÃO - Pedido Nº: ${formData.pedido}`, // titulo do email
      body: corpo.trim(),
      attachments: anexosParaEmail
    });
    
    registrarLog(analistaLogado, "SOLICITAR_REEMBOLSO", `Enviou devolução/e-mail para pedido ${formData.pedido}`);
    return "Processo de reembolso validado e e-mail disparado!";
  } finally {
    lock.releaseLock();
  }
}

function salvarRevertido(formData, analistaLogado) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { throw new Error("Sistema ocupado. Tente novamente."); }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const now = new Date();
    const sheetName = 'rev' + MESES[now.getMonth()] + now.getFullYear();

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const header = ['DATA', 'TIPO', 'PEDIDO', 'TICKET', 'ANALISTA', 'SOLICITAÇÃO', 'TIPO DE REVERSÃO', 'VALOR', 'OBSERVAÇÃO'];
      sheet.getRange(1, 1, 1, header.length).setValues([header]).setBackground('#000000').setFontColor('#FFFFFF').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const row = [now, formData.tipo, formData.pedido, formData.ticket, analistaLogado, formData.solicitacao, formData.tipoReversao, formatarValorContabil(formData.valor), formData.observacao];
    const novaLinha = getVerdadeiraUltimaLinha(sheet) + 1;
    sheet.getRange(novaLinha, 1, 1, row.length).setValues([row]).setFontWeight('bold').setBackground('#d9ead3');
    
    registrarLog(analistaLogado, "SALVAR_REVERSAO", `Registrou reversão para pedido ${formData.pedido}`);
    return row;
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// BARRADOS (TRANSPORTE)
// ==========================================
function dispararBloqueioTransporte(dadosForm, analistaLogado) {
  garantirAbasTransporte();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const dadosConfig = ss.getSheetByName("CONFIG_TRANSPORTADORAS").getDataRange().getValues();
  let emailsDestino = "";
  
  for (let i = 1; i < dadosConfig.length; i++) {
    if (String(dadosConfig[i][0]).trim().toUpperCase() === String(dadosForm.transportadora).toUpperCase().trim()) {
      emailsDestino = String(dadosConfig[i][1]).trim();
      break;
    }
  }

  if (!emailsDestino) throw new Error("E-mails da transportadora não localizados na base.");
  if (!dadosForm.fileBase64) throw new Error("A NF de remessa é obrigatória para o bloqueio."); // Trava de segurança no backend

  const assuntoOficial = `DEVOLUÇÃO / ${dadosForm.nf} / ${dadosForm.transportadora.toUpperCase()}`;
  const corpoEmail = `Prezados,\n\nSolicitamos o bloqueio de entrega e o retorno imediato da mercadoria referente à Nota Fiscal nº ${dadosForm.nf} para o CD.\n\nIdentificador do Ticket: ${dadosForm.ticket || "N/A"}\nResponsável pela Solicitação: ${analistaLogado}\n\nSegue em anexo a NF de Remessa referente ao retorno.\n\nPor favor, nos confirme o recebimento e o status deste bloqueio.\n\nAtenciosamente,\nControle Interno | Atendimento`;

  let anexosParaEmail = [];
  try {
    let blob = Utilities.newBlob(Utilities.base64Decode(dadosForm.fileBase64), dadosForm.mimeType, dadosForm.fileName);
    anexosParaEmail.push(blob);
  } catch(e) {
    throw new Error("Erro ao processar o anexo da NF. Tente novamente.");
  }

  MailApp.sendEmail({
    to: emailsDestino,
    cc: "xx",
    subject: assuntoOficial,
    body: corpoEmail.trim(),
    attachments: anexosParaEmail
  });

  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { throw new Error("Fila cheia ao registrar o log."); }

  try {
    const abaReg = ss.getSheetByName("REGISTRO_BARRADOS");
    const novaLinha = getVerdadeiraUltimaLinha(abaReg) + 1;
    abaReg.getRange(novaLinha, 1, 1, 7).setValues([[
      new Date(), analistaLogado, dadosForm.transportadora.toUpperCase(), dadosForm.nf, dadosForm.ticket || "N/A", assuntoOficial, emailsDestino
    ]]).setFontWeight('bold').setBackground('#fff2cc');
    
    registrarLog(analistaLogado, "BARRADO_TRANSPORTE", `Bloqueio NF ${dadosForm.nf} via ${dadosForm.transportadora.toUpperCase()} com Anexo`);
    return "E-mail de bloqueio com anexo disparado com sucesso à " + dadosForm.transportadora.toUpperCase() + "!";
  } finally {
    lock.releaseLock();
  }
}

function carregarHistoricoBarrados() {
  garantirAbasTransporte();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dados = ss.getSheetByName("REGISTRO_BARRADOS").getDataRange().getValues();
  const resultado = [];
  
  for (let i = dados.length - 1; i >= 1; i--) {
    if (dados[i][3]) { 
      resultado.push({
        dataStr: Utilities.formatDate(new Date(dados[i][0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
        analista: dados[i][1],
        transportadora: dados[i][2],
        nf: dados[i][3],
        ticket: dados[i][4],
        destinatarios: dados[i][6]
      });
    }
  }
  return resultado;
}

// ==========================================
// RELATÓRIOS (Não alterado)
// ==========================================
function formatarValorContabil(valor) {
  return (Number(valor) || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function getSheetByNamePrefix(prefix, month) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(prefix + month);
}
function getDevolucaoQuantidade(month) {
  const sheet = getSheetByNamePrefix('dev', month);
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const mGeral = String(data[i][10]).trim();
    if (!mGeral || isNaN(new Date(data[i][0]))) continue;
    result[mGeral] = (result[mGeral] || 0) + 1;
  }
  return Object.entries(result).map(([m, q]) => [m, Number(q)]);
}
function getDevolucaoValores(month) {
  const sheet = getSheetByNamePrefix('dev', month);
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const mGeral = String(data[i][10]).trim();
    if (!mGeral || isNaN(new Date(data[i][0]))) continue;
    let v = typeof data[i][8] === 'number' ? data[i][8] : parseFloat(String(data[i][8]).replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.')) || 0;
    result[mGeral] = (result[mGeral] || 0) + v;
  }
  return Object.entries(result).map(([m, v]) => [m, Number(v)]);
}
function getReversoes(month) {
  const sheet = getSheetByNamePrefix('rev', month);
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][4] || isNaN(new Date(data[i][0]))) continue;
    let v = typeof data[i][7] === 'number' ? data[i][7] : parseFloat(String(data[i][7]).replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.')) || 0;
    result.push([data[i][4], 1, v]);
  }
  return result;
}

function gerarRelatorioCancelamentos(month) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = "Relatorio_" + month.toUpperCase();
  const sheetRel = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sheetRel.clear(); 
  sheetRel.setHiddenGridlines(true);

  const qtds = getDevolucaoQuantidade(month) || [];
  const valores = getDevolucaoValores(month) || [];
  const reversoes = getReversoes(month) || [];
  
  const rel = {};
  const norm = t => String(t).normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();
  
  qtds.forEach(([m, q]) => { rel[norm(m)] = { qtd: Number(q) || 0, valor: 0 }; });
  valores.forEach(([m, v]) => { 
    let k = norm(m); 
    if (!rel[k]) rel[k] = { qtd: 0, valor: 0 }; 
    rel[k].valor = Number(v) || 0; 
  });

  const corpoQtd = Object.entries(rel).map(([m, v]) => [m, v.qtd]); 
  const totalQtd = corpoQtd.reduce((s, x) => s + x[1], 0);
  corpoQtd.push(["TOTAL", totalQtd]);
  
  const corpoVal = Object.entries(rel).map(([m, v]) => [m, v.valor]); 
  const totalVal = corpoVal.reduce((s, x) => s + x[1], 0);
  corpoVal.push(["TOTAL", totalVal]);

  const qtdRevertida = reversoes.reduce((s, x) => s + (x[1] || 0), 0);
  const valorRevertido = reversoes.reduce((s, x) => s + (x[2] || 0), 0);
  const desist = rel["DESISTENCIA"] || { qtd: 0, valor: 0 };

  const totalPedidosSite = 97160; 
  const valorTotalSite = 104971486.20;

  sheetRel.setColumnWidth(1, 320);
  sheetRel.setColumnWidth(2, 120);
  sheetRel.setColumnWidth(3, 20);
  sheetRel.setColumnWidth(4, 320);
  sheetRel.setColumnWidth(5, 140);

  const bgBlueDark = "#4a86e8";
  const bgBlueLight = "#9fc5e8";
  const bgGray = "#f3f3f3";
  const borderStyle = SpreadsheetApp.BorderStyle.SOLID;

  sheetRel.getRange("A2:E2").merge()
    .setValue("RELATÓRIO DE CANCELAMENTOS - " + month.toUpperCase())
    .setBackground(bgBlueDark).setFontColor("#ffffff")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontWeight("bold").setFontSize(11)
    .setBorder(true, true, true, true, false, false, "black", borderStyle);

  sheetRel.getRange("A4:B4").merge().setValue("QUANTIDADE DE PEDIDOS")
    .setBackground(bgBlueLight).setFontWeight("bold").setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, "black", borderStyle);
    
  sheetRel.getRange("D4:E4").merge().setValue("VALORES TOTAIS (R$)")
    .setBackground(bgBlueLight).setFontWeight("bold").setHorizontalAlignment("center")
    .setBorder(true, true, true, true, false, false, "black", borderStyle);

  sheetRel.getRange("A5:B5").setValues([["Motivo", "Quantidade"]])
    .setBackground(bgGray).setFontWeight("bold").setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, "black", borderStyle);
    
  sheetRel.getRange("D5:E5").setValues([["Motivo", "Valor Total (R$)"]])
    .setBackground(bgGray).setFontWeight("bold").setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, "black", borderStyle);

  if(corpoQtd.length > 0) {
    sheetRel.getRange(6, 1, corpoQtd.length, 2).setValues(corpoQtd)
      .setBorder(true, true, true, true, true, true, "black", borderStyle)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
      
    sheetRel.getRange(6, 4, corpoVal.length, 2).setValues(corpoVal)
      .setBorder(true, true, true, true, true, true, "black", borderStyle)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
      
    sheetRel.getRange(6, 5, corpoVal.length, 1).setNumberFormat("R$ #,##0.00");
  }

  const linhaResumo = Math.max(corpoQtd.length, corpoVal.length) + 8;
  
  sheetRel.getRange(linhaResumo, 1, 1, 5).merge()
    .setValue("RESUMO GERAL")
    .setBackground(bgBlueDark).setFontColor("#ffffff")
    .setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, "black", borderStyle);

  sheetRel.getRange(linhaResumo+1, 1, 1, 3).merge().setValue("Cancelamentos Revertidos").setFontWeight("bold").setHorizontalAlignment("center");
  sheetRel.getRange(linhaResumo+1, 4).setValue(qtdRevertida).setHorizontalAlignment("center");
  sheetRel.getRange(linhaResumo+1, 5).setValue(valorRevertido).setNumberFormat("R$ #,##0.00").setHorizontalAlignment("center");

  sheetRel.getRange(linhaResumo+2, 1, 1, 3).merge().setValue("% sobre total de cancelamentos").setFontWeight("bold").setHorizontalAlignment("center");
  sheetRel.getRange(linhaResumo+2, 4, 1, 2).merge().setValue(totalQtd ? (qtdRevertida / totalQtd) : 0).setNumberFormat("0.00%").setHorizontalAlignment("center");

  sheetRel.getRange(linhaResumo+3, 1, 1, 3).merge().setValue("% sobre reversíveis").setFontWeight("bold").setHorizontalAlignment("center");
  sheetRel.getRange(linhaResumo+3, 4, 1, 2).merge().setValue(desist.qtd ? (qtdRevertida / desist.qtd) : 0).setNumberFormat("0.00%").setHorizontalAlignment("center");

  sheetRel.getRange(linhaResumo+4, 1, 1, 3).merge().setValue("% cancelamento vs total de pedidos do site").setFontWeight("bold").setHorizontalAlignment("center");
  sheetRel.getRange(linhaResumo+4, 4).setValue(totalPedidosSite ? (totalQtd / totalPedidosSite) : 0).setNumberFormat("0.00%").setHorizontalAlignment("center");
  sheetRel.getRange(linhaResumo+4, 5).setValue(totalPedidosSite.toLocaleString('pt-BR') + " Pedidos").setHorizontalAlignment("center");

  sheetRel.getRange(linhaResumo+5, 1, 1, 3).merge().setValue("% cancelamento vs total de pedidos do site - VALORES").setFontWeight("bold").setHorizontalAlignment("center");
  sheetRel.getRange(linhaResumo+5, 4).setValue(valorTotalSite ? (totalVal / valorTotalSite) : 0).setNumberFormat("0.00%").setHorizontalAlignment("center");
  sheetRel.getRange(linhaResumo+5, 5).setValue(valorTotalSite).setNumberFormat("R$ #,##0.00").setHorizontalAlignment("center");

  sheetRel.getRange(linhaResumo, 1, 6, 5).setBorder(true, true, true, true, true, true, "black", borderStyle);

  if(corpoQtd.length > 1) { 
    const chartRange = sheetRel.getRange(6, 1, corpoQtd.length - 1, 2); 
    const chart = sheetRel.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(chartRange)
      .setPosition(linhaResumo + 7, 1, 10, 0)
      .setOption('title', 'Distribuição de Cancelamentos por Motivo')
      .setOption('width', 550)
      .setOption('height', 350)
      .setOption('is3D', false)
      .setOption('legend', {position: 'right', textStyle: {fontSize: 11, color: '#555555'}})
      .setOption('titleTextStyle', {color: '#777777', fontSize: 16})
      .setOption('pieSliceText', 'value')
      .build();
    sheetRel.insertChart(chart);
  }

  const footerRow = linhaResumo + 7;
  const now = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");
  
  sheetRel.getRange(footerRow, 4, 1, 2).merge()
    .setValue("Relatório gerado em " + now)
    .setFontSize(9).setFontColor("#777777")
    .setHorizontalAlignment("right").setFontStyle("italic");
    
  sheetRel.getRange(footerRow + 1, 4, 1, 2).merge()
    .setValue("Este documento é confidencial e destinado apenas para fins internos.")
    .setFontSize(9).setFontColor("#777777")
    .setHorizontalAlignment("right").setFontStyle("italic");

  return "Relatório gerado com sucesso na aba: " + sheetName;
}

function formatarValorContabil(valor) {
  return (Number(valor) || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function salvarImagemNoDrive(dataUrl) {
  try {
    const partes = dataUrl.split(',');
    const mimeType = partes[0].match(/:(.*?);/)[1];
    const base64Data = partes[1];

    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, "Img_Proc_" + new Date().getTime());

    // Se preferir, pode trocar o nome da pasta no Drive
    const nomePasta = "imagens_drive_controleint";
    let pasta;
    const pastas = DriveApp.getFoldersByName(nomePasta);
    
    if (pastas.hasNext()) {
      pasta = pastas.next();
    } else {
      pasta = DriveApp.createFolder(nomePasta);
      pasta.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    const arquivo = pasta.createFile(blob);
    
    // Retorna link direto da thumbnail em alta resolução
    return "https://drive.google.com/thumbnail?id=" + arquivo.getId() + "&sz=w1200";
  } catch(e) {
    throw new Error("Falha no upload para o Drive: " + e.message);
  }
}


// ==========================================
// UPLOAD DE ARQUIVOS GRANDES (CHUNKING)
// ==========================================
function salvarChunkDrive(fileIdTemp, chunkIndex, chunkData) {
  DriveApp.createFile(fileIdTemp + '_' + chunkIndex + '.txt', chunkData);
}

function finalizarUploadChunk(fileIdTemp, totalChunks, fileName, mimeType) {
  let fullBase64 = "";
  for(let i = 0; i < totalChunks; i++) {
    const files = DriveApp.searchFiles("title = '" + fileIdTemp + "_" + i + ".txt'");
  
    if(files.hasNext()) {
      const file = files.next();
      fullBase64 += file.getBlob().getDataAsString();
      file.setTrashed(true); // Exclui fragmento após remontar
    }
  }
  const resultBlob = Utilities.newBlob(Utilities.base64Decode(fullBase64), mimeType, fileName);
  return DriveApp.createFile(resultBlob).getUrl();
}

// ==========================================
// INTEGRAÇÃO: BARRADOS PLP (SISTEMA EXTERNO)
// ==========================================
const SPREADSHEET_PLP_ID = '1qVLCvjaKJiAidxKjtfq17D40XR3a5Vr3kUBhQ9ytjww'; 
const MESES_PLP = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];

function garantirAbaMesPLP(dataObjeto, ssPLP) {
  const nomeAba = MESES_PLP[dataObjeto.getMonth()] + "-" + dataObjeto.getFullYear();
  let sheet = ssPLP.getSheetByName(nomeAba);
  if (!sheet) {
    sheet = ssPLP.insertSheet(nomeAba);
    const cabecalhos = ["ID", "DATA CRIAÇÃO", "ANALISTA", "SETOR ORIGEM", "NF", "TRANSPORTADORA", "TICKET", "ITEM", "SERIAL", "AÇÃO", "STATUS ATUAL", "HORA BARRAGEM"];
    sheet.appendRow(cabecalhos);
    sheet.getRange(1, 1, 1, cabecalhos.length).setBackground("#8B0000").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return nomeAba;
}

function buscarDadosPLP(analistaLogado) {
  const ssPLP = SpreadsheetApp.openById(SPREADSHEET_PLP_ID);
  let hoje = new Date();
  
  let abasParaBuscar = [
    garantirAbaMesPLP(hoje, ssPLP)
  ];
  
  // Pega o mês anterior também
  let mesAnterior = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
  let abaAnterior = MESES_PLP[mesAnterior.getMonth()] + "-" + mesAnterior.getFullYear();
  if(ssPLP.getSheetByName(abaAnterior)) {
    abasParaBuscar.push(abaAnterior);
  }

  let dadosCombinados = [];
  
  abasParaBuscar.forEach(nomeAba => {
    let sheet = ssPLP.getSheetByName(nomeAba);
    if (sheet) {
      let dados = sheet.getDataRange().getValues();
      if (dados.length > 1) {
        let cabecalho = dados.shift();
        for (let i = 0; i < dados.length; i++) {
          let row = dados[i];
          let obj = {};
          cabecalho.forEach((col, idx) => {
            // Formatação rápida de data
            let val = row[idx];
            if(val instanceof Date) {
              const p = (n) => String(n).padStart(2, '0');
              val = `${p(val.getDate())}/${p(val.getMonth()+1)}/${val.getFullYear()} ${p(val.getHours())}:${p(val.getMinutes())}`;
            }
            obj[col] = val;
          });
          obj["MES_REF"] = nomeAba; 
          dadosCombinados.push(obj);
        }
      }
    }
  });

  return dadosCombinados.reverse(); 
}

function salvarSolicitacaoPLP(obj, analistaLogado) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    const ssPLP = SpreadsheetApp.openById(SPREADSHEET_PLP_ID);
    
    let hoje = new Date();
    const nomeAbaAtual = garantirAbaMesPLP(hoje, ssPLP);
    const sheetAtual = ssPLP.getSheetByName(nomeAbaAtual);
    const dadosAtual = sheetAtual.getDataRange().getValues();

    // Valida duplicidade NF mês atual
    for(let i = 1; i < dadosAtual.length; i++) {
      if(String(dadosAtual[i][4]).trim().toLowerCase() === String(obj.nf).trim().toLowerCase()) {
        throw new Error(`A NF ${obj.nf} já está registrada no sistema (Mês Atual)!`);
      }
    }

    // Pega o maior ID
    let maxId = 999;
    for(let i = 1; i < dadosAtual.length; i++) {
       let currentId = parseInt(dadosAtual[i][0]);
       if (!isNaN(currentId) && currentId > maxId) maxId = currentId;
    }
    let novoId = maxId + 1;
    
    const dh = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
    sheetAtual.appendRow([novoId, dh, analistaLogado, "Atendimento", obj.nf, obj.transportadora, obj.ticket, obj.item, "", "BARRAR", "PENDENTE", ""]);
    
    registrarLog(analistaLogado, "SOLICITAR_BARRAGEM", `Solicitou barragem PLP para NF ${obj.nf}`);
    return "Solicitação de Barragem enviada com sucesso!";
  } finally {
    lock.releaseLock();
  }
}

function marcarComoVistoPLP(id, abaExata, analistaLogado) {
  const ssPLP = SpreadsheetApp.openById(SPREADSHEET_PLP_ID);
  const sheet = ssPLP.getSheetByName(abaExata);
  if (!sheet) return;
  const dados = sheet.getDataRange().getValues();
  for(let i = 1; i < dados.length; i++) {
    if(String(dados[i][0]) === String(id)) {
      sheet.getRange(i + 1, 10).setValue("VISTO"); 
      return "Marcado como lido.";
    }
  }
}

function excluirSolicitacaoPLP(id, abaExata, analistaLogado) {
  const ssPLP = SpreadsheetApp.openById(SPREADSHEET_PLP_ID);
  const sheet = ssPLP.getSheetByName(abaExata);
  if (!sheet) throw new Error("Aba não encontrada.");
  const dados = sheet.getDataRange().getValues();
  for(let i = 1; i < dados.length; i++) {
    if(String(dados[i][0]) === String(id)) {
      if(dados[i][10] === "PENDENTE" || dados[i][10] === "VERIFICADO") {
        sheet.deleteRow(i + 1); 
        registrarLog(analistaLogado, "EXCLUIR_BARRAGEM", `Excluiu solicitação PLP ID ${id}`);
        return "Solicitação excluída.";
      } else {
        throw new Error("Não é possível excluir uma solicitação já processada.");
      }
    }
  }
}
