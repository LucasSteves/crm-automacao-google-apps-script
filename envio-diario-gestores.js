function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Relatório")
    .addItem("1. Preparar Base de Dados", "prepararBase")
    .addItem("2. Enviar Relatórios em Lote", "enviarRelatoriosEmLote")
    .addToUi();
}

function padronizarTexto(texto) {
  if (!texto) return "";
  return String(texto)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // Remove acentos
    .trim()
    .toUpperCase();
}

function prepararBase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName("BASE");

  base.getDataRange().breakApart();
  SpreadsheetApp.flush(); 

  const dados = base.getDataRange().getDisplayValues();
  if (dados.length <= 1) return SpreadsheetApp.getUi().alert("Não há dados para organizar.");
  
  for (let col = 0; col < 3; col++) {
    for (let row = 1; row < dados.length; row++) {
      let valor = String(dados[row][col]).trim();
      if (valor === "" || valor === "null") {
        dados[row][col] = dados[row - 1][col];
      }
    }
  }

  const cabecalho = dados[0];
  const cabecalhoLimpo = cabecalho.map(padronizarTexto);
  
  const cursoIndex = cabecalhoLimpo.indexOf("CURSO");
  const dataIndex = cabecalhoLimpo.indexOf("DATA MATRICULA") !== -1 ? cabecalhoLimpo.indexOf("DATA MATRICULA") : cabecalhoLimpo.indexOf("DATA DE MATRICULA");

  if (cursoIndex === -1 || dataIndex === -1) {
    SpreadsheetApp.getUi().alert("Erro: Colunas 'CURSO' ou 'DATA MATRICULA' não encontradas.\nVerifique os cabeçalhos da sua planilha.");
    return;
  }

  let linhas = dados.slice(1).filter(linha => {
    const nomeCurso = padronizarTexto(linha[cursoIndex]);
    return !nomeCurso.includes("MEDICINA");
  });

  const parseDateBr = (dStr) => {
    if (!dStr) return 0;
    const parts = String(dStr).split(' ')[0].split('/');
    if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]).getTime();
    return new Date(dStr).getTime() || 0;
  };

  linhas.sort((a, b) => {
    const cursoA = padronizarTexto(a[cursoIndex]);
    const cursoB = padronizarTexto(b[cursoIndex]);
    if (cursoA < cursoB) return -1;
    if (cursoA > cursoB) return 1;
    return parseDateBr(a[dataIndex]) - parseDateBr(b[dataIndex]);
  });

  const colsSensiveis = [];
  cabecalhoLimpo.forEach((nome, index) => {
    if (nome.includes("TELEFONE") || nome.includes("CELULAR") || nome.includes("FONE") || nome.includes("RA") || nome.includes("CPF") || nome.includes("RG")) {
      colsSensiveis.push(index);
    }
  });

  linhas.forEach(linha => {
    colsSensiveis.forEach(idx => {
      let valor = String(linha[idx] || "").trim();
      if (valor === "31/12/1969" || valor === "30/12/1899") valor = "";
      if (valor !== "" && !valor.startsWith("'")) {
        linha[idx] = "'" + valor;
      } else {
        linha[idx] = valor;
      }
    });
  });

  base.clear();
  base.appendRow(cabecalho);
  base.getRange(1, 1, 1, cabecalho.length).setFontWeight("bold").setBackground("#d9d9d9");
  
  if (linhas.length > 0) {
    base.getRange(2, 1, linhas.length, cabecalho.length).setValues(linhas);
    base.getRange(2, dataIndex + 1, linhas.length, 1).setNumberFormat("dd/MM/yyyy");
  }

  base.getDataRange().setHorizontalAlignment("center").setVerticalAlignment("middle");
  base.autoResizeColumns(1, cabecalho.length);

  SpreadsheetApp.getUi().alert("✅ Base preparada!");
}

function enviarRelatoriosEmLote() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName("BASE");
  const grupos = ss.getSheetByName("GRUPOS_ENVIO");

  const dadosBase = base.getDataRange().getValues();
  if (dadosBase.length <= 1) return ui.alert("A base está vazia.");

  const cabecalho = dadosBase[0];
  const cabecalhoLimpo = cabecalho.map(padronizarTexto);

  const cursoIndex = cabecalhoLimpo.indexOf("CURSO");
  const statusIndex = cabecalhoLimpo.indexOf("STATUS");
  
  let tipoIndex = cabecalhoLimpo.indexOf("TIPO INGRESSO");
  if (tipoIndex === -1) tipoIndex = cabecalhoLimpo.indexOf("TIPO DE INGRESSO");
  
  let dataIndex = cabecalhoLimpo.indexOf("DATA MATRICULA");
  if (dataIndex === -1) dataIndex = cabecalhoLimpo.indexOf("DATA DE MATRICULA");

  if (cursoIndex === -1 || statusIndex === -1 || tipoIndex === -1) {
    return ui.alert("❌ ERRO GRAVE: Não encontrei as colunas 'CURSO', 'STATUS' ou 'TIPO INGRESSO'.\nOs nomes exatos na linha 1 da BASE estão corretos?");
  }

  const dadosGestores = grupos.getDataRange().getValues();

  let enviosPorGestor = {}; 
  let gestoresNaoEncontrados = new Set();
  let alunosIgnorados = 0; 

  dadosBase.slice(1).forEach(linha => {
    let cursoAtual = String(linha[cursoIndex] || "").trim();
    if (!cursoAtual || cursoAtual.includes("=====")) return;

    let gestor = buscarGestor(cursoAtual, dadosGestores);
    if (!gestor || !gestor.email) {
      gestoresNaoEncontrados.add(cursoAtual);
      return;
    }

    let email = gestor.email;
    if (!enviosPorGestor[email]) {
      enviosPorGestor[email] = { cc: gestor.cc, palavrasChave: gestor.palavras, cursos: {} };
    }

    if (!enviosPorGestor[email].cursos[cursoAtual]) {
      enviosPorGestor[email].cursos[cursoAtual] = { ingressantes: [], evasao: [], prouni: [] };
    }

    let status = padronizarTexto(linha[statusIndex]);
    let tipo = padronizarTexto(linha[tipoIndex]);

    if (tipo.includes("PROUNI")) {
      enviosPorGestor[email].cursos[cursoAtual].prouni.push(linha);
    } 
    else if (status.includes("CURSANDO")) {
      enviosPorGestor[email].cursos[cursoAtual].ingressantes.push(linha);
    } 
    else if (status.includes("CANCELADO") || status.includes("TRANCADO") || status.includes("TRANSFERENCIA")) {
      enviosPorGestor[email].cursos[cursoAtual].evasao.push(linha);
    } 
    else {
      alunosIgnorados++; 
    }
  });

  const emailsAgrupados = Object.keys(enviosPorGestor);
  if (emailsAgrupados.length === 0) return ui.alert("Nenhum curso associado a um gestor válido.");

  let totalAlunosLidos = 0;
  for (const email of emailsAgrupados) {
    for (const curso of Object.keys(enviosPorGestor[email].cursos)) {
      totalAlunosLidos += enviosPorGestor[email].cursos[curso].ingressantes.length;
      totalAlunosLidos += enviosPorGestor[email].cursos[curso].evasao.length;
      totalAlunosLidos += enviosPorGestor[email].cursos[curso].prouni.length;
    }
  }

  if (totalAlunosLidos === 0) {
    return ui.alert("❌ ALERTA: O script processou a planilha, mas encontrou ZERO alunos para enviar.\nVerifique se as palavras na coluna STATUS são realmente: Cursando, Cancelado, Trancado ou Transferência.");
  }

  const resposta = ui.alert(
    "Confirmar Envio", 
    `O script encontrou ${totalAlunosLidos} aluno(s) válidos.\n\nPreparando envio para ${emailsAgrupados.length} gestor(es).\nDeseja continuar?`, 
    ui.ButtonSet.YES_NO
  );
  
  if (resposta !== ui.Button.YES) return;

  const dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  let relatorioFinal = { sucesso: 0, erro: [] };

  for (const emailGestor of emailsAgrupados) {
    const dadosDoGestor = enviosPorGestor[emailGestor];
    const nomesCursos = Object.keys(dadosDoGestor.cursos);
    const pChave = padronizarTexto(dadosDoGestor.palavrasChave);
    
    let assuntoEmail = "";
    let nomeArquivo = "";
    let avisoAbas = "";

    const ehEng = pChave.includes("ENGENHARIA") || nomesCursos.some(c => padronizarTexto(c).includes("ENGENHARIA"));
    const ehTec = pChave.includes("TECNOLOG") || nomesCursos.some(c => padronizarTexto(c).includes("TECNOLOG"));

    if (nomesCursos.length === 1) {
      assuntoEmail = `Planilha Ingressantes - ${nomesCursos[0]} - ${dataHoje}`;
      nomeArquivo = `Planilha Ingressantes - ${nomesCursos[0]}.xlsx`;
    } else {
      if (ehEng) {
        assuntoEmail = `Planilha Ingressantes - Engenharias - ${dataHoje}`;
        nomeArquivo = `Planilha Ingressantes - Engenharias.xlsx`;
      } else if (ehTec) {
        assuntoEmail = `Planilha Ingressantes - Tecnólogos - ${dataHoje}`;
        nomeArquivo = `Planilha Ingressantes - Tecnólogos.xlsx`;
      } else {
        assuntoEmail = `Planilha Ingressantes Consolidada (${nomesCursos.length} Cursos) - ${dataHoje}`;
        nomeArquivo = `Planilha Consolidada Ingressantes.xlsx`;
      }
    }

    const tempSS = SpreadsheetApp.create("TEMP_" + emailGestor);

    nomesCursos.forEach(curso => {
      const nomeAba = curso.substring(0, 31).replace(/[:\\/?*\[\]]/g, "").trim();
      const aba = tempSS.insertSheet(nomeAba);

      aba.appendRow(cabecalho);
      aba.getRange(1, 1, 1, cabecalho.length).setFontWeight("bold").setBackground("#d9d9d9");

      const dCurso = dadosDoGestor.cursos[curso];
      
      const addCat = (titulo, lista) => {
        if (lista.length > 0) {
          aba.appendRow([titulo]);
          aba.getRange(aba.getLastRow(), 1, 1, cabecalho.length).merge().setFontWeight("bold").setBackground("#d9d9d9");
          aba.getRange(aba.getLastRow() + 1, 1, lista.length, cabecalho.length).setValues(lista);
          aba.getRange(aba.getLastRow() - lista.length + 1, dataIndex + 1, lista.length, 1).setNumberFormat("dd/MM/yyyy");
        }
      };

      addCat("INGRESSANTES", dCurso.ingressantes);
      addCat("EVASÃO", dCurso.evasao);
      addCat("PROUNI", dCurso.prouni);

      aba.getDataRange().setHorizontalAlignment("center").setVerticalAlignment("middle");
      aba.autoResizeColumns(1, cabecalho.length);
      
      SpreadsheetApp.flush(); 
    });

    const sheet1 = tempSS.getSheets()[0];
    if (sheet1.getName() === "Página1" || sheet1.getName() === "Sheet1") {
      tempSS.deleteSheet(sheet1);
    }
    
    SpreadsheetApp.flush();
    Utilities.sleep(8000); 

    const url = `https://docs.google.com/spreadsheets/d/${tempSS.getId()}/export?format=xlsx`;
    const token = ScriptApp.getOAuthToken();
    
    try {
      const response = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + token } });
      const blob = response.getBlob().setName(nomeArquivo);

      MailApp.sendEmail({
        to: emailGestor,
        cc: dadosDoGestor.cc,
        subject: assuntoEmail, 
        body: `${avisoAbas}Segue em anexo a planilha de ingressantes, evasão e Prouni.`,
        attachments: [blob]
      });
      relatorioFinal.sucesso++;
    } catch (e) {
      console.log("Erro ao enviar: " + e.message);
      relatorioFinal.erro.push(emailGestor);
    }

    try {
      Utilities.sleep(1000);
      DriveApp.getFileById(tempSS.getId()).setTrashed(true);
    } catch (e) {
      console.log("Aviso: A planilha não pôde ser apagada. " + e.message);
    }
  }

  base.clear();
  base.appendRow(cabecalho);
  base.getRange(1, 1, 1, cabecalho.length).setFontWeight("bold").setBackground("#d9d9d9");

  let msg = `✅ Processo concluído!\nE-mails enviados: ${relatorioFinal.sucesso}`;
  if (alunosIgnorados > 0) msg += `\n\n⚠️ Atenção: ${alunosIgnorados} alunos não foram enviados pois seus status não eram Cursando, Cancelado, Trancado ou Prouni.`;
  if (gestoresNaoEncontrados.size > 0) msg += `\n\n⚠️ Cursos na base sem gestor cadastrado:\n- ${Array.from(gestoresNaoEncontrados).join("\n- ")}`;
  if (relatorioFinal.erro.length > 0) msg += `\n\n❌ Erro ao enviar para:\n- ${relatorioFinal.erro.join("\n- ")}`;
  
  ui.alert("Resumo de Envios", msg, ui.ButtonSet.OK);
}

function buscarGestor(curso, dadosGestores) {
  const cabecalho = dadosGestores[0].map(padronizarTexto);
  const pIndex = cabecalho.indexOf("PALAVRAS_CHAVE");
  const eIndex = cabecalho.indexOf("EMAIL_PRINCIPAL");
  const cIndex = cabecalho.indexOf("EMAIL_CC");

  for (let i = 1; i < dadosGestores.length; i++) {
    const palavras = String(dadosGestores[i][pIndex] || "");
    if (!palavras) continue;
    const lista = palavras.split(",");
    for (let p of lista) {
      let termo = padronizarTexto(p);
      if (!termo) continue;
      
      let regex = new RegExp("(^|\\s|-)" + termo.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + "(\\s|-|$)", "i");
      if (regex.test(padronizarTexto(curso))) {
        return {
          email: String(dadosGestores[i][eIndex] || "").trim(),
          cc: String(dadosGestores[i][cIndex] || "").trim(),
          palavras: palavras
        };
      }
    }
  }
  return null;
}
