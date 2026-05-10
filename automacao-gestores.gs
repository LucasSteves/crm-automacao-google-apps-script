// =========================================================
// 1. MENU UNIFICADO (Ferramentas + Relatório)
// =========================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Automação')
      .addItem('1. Corrigir Mesclagem (A, B, C)', 'corrigirMesclagem')
      .addItem('2. Preencher E-mails (Toda a Planilha)', 'preencherEmailsMassivo')
      .addSeparator()
      .addItem('📊 3. Gerar Relatório Formatado', 'atualizarRelatorio')
      .addItem('🚀 4. Enviar Relatório por E-mail', 'enviarRelatorioExcel')
      .addToUi();
}

// =========================================================
// 2. FUNÇÕES DE HIGIENIZAÇÃO E PREENCHIMENTO
// =========================================================

function corrigirMesclagem() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, 3); 
  range.breakApart(); 
  
  const values = range.getValues();
  let alterou = false;
  
  for (let col = 0; col < 3; col++) {
    for (let row = 0; row < values.length; row++) {
      const valor = values[row][col];
      if (row > 0 && (valor === "" || valor === null)) {
        values[row][col] = values[row - 1][col];
        alterou = true;
      }
    }
  }
  
  if (alterou) {
    range.setValues(values);
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert('✅ Colunas A, B e C corrigidas!');
  } else {
    SpreadsheetApp.getUi().alert('⚠️ Nenhuma correção necessária.');
  }
}

function preencherEmailsMassivo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const rangeCursos = sheet.getRange(2, 3, lastRow - 1, 1); 
  const rangeEmails = sheet.getRange(2, 12, lastRow - 1, 1);
  const valoresCursos = rangeCursos.getValues();
  const valoresEmails = rangeEmails.getValues();
  let alterou = false;

  const { abreviacoes, emailsGestores } = getTabelas();

  for (let i = 0; i < valoresCursos.length; i++) {
    let curso = valoresCursos[i][0];
    let emailAtual = valoresEmails[i][0];
    if (!curso) continue;

    if (abreviacoes[curso]) {
      curso = abreviacoes[curso];
      valoresCursos[i][0] = curso;
    }

    if (emailsGestores[curso] && emailAtual === "") {
      valoresEmails[i][0] = emailsGestores[curso];
      alterou = true;
    }
  }

  if (alterou) {
    rangeCursos.setValues(valoresCursos);
    rangeEmails.setValues(valoresEmails);
    SpreadsheetApp.getUi().alert('✅ E-mails preenchidos!');
  }
}

function onEdit(e) {
  if (!e) return;
  const range = e.range;
  if (range.getColumn() === 3 && range.getRow() > 1) {
    let curso = range.getValue();
    const { abreviacoes, emailsGestores } = getTabelas();
    if (abreviacoes[curso]) {
      curso = abreviacoes[curso];
      range.setValue(curso);
    }
    if (emailsGestores[curso]) {
      range.offset(0, 9).setValue(emailsGestores[curso]);
    }
  }
}

// =========================================================
// 3. GERAÇÃO DE RELATÓRIO PROFISSIONAL
// =========================================================

function atualizarRelatorio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName("BASE");
  const dados = base.getDataRange().getValues();
  const cabecalho = dados[0];

  const statusIndex = cabecalho.indexOf("STATUS");
  const tipoIndex = cabecalho.indexOf("TIPO INGRESSO");
  const dataIndex = cabecalho.indexOf("DATA MATRICULA");
  const cursoIndex = cabecalho.indexOf("CURSO");

  const { abreviacoes } = getTabelas();

  let ingressantes = [];
  let evasao = [];
  let prouni = [];
  
  // Tratamento de memória para preencher cursos de células agrupadas/vazias
  let cursoAtualMemorizado = "";

  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];
    let cursoLinha = String(linha[cursoIndex] || "").trim();
    let status = String(linha[statusIndex] || "").trim();
    let tipo = String(linha[tipoIndex] || "").trim();

    // Se a linha tem um curso escrito, atualizamos a memória
    if (cursoLinha !== "" && !cursoLinha.includes("===") && cursoLinha !== "INGRESSANTES") {
      cursoAtualMemorizado = cursoLinha;
    }

    // Se o curso está vazio (devido ao agrupamento), mas a linha tem aluno, usamos a memória
    if (cursoLinha === "" && cursoAtualMemorizado !== "" && status !== "") {
      cursoLinha = cursoAtualMemorizado;
    }

    // Aplica abreviação se necessário
    if (abreviacoes[cursoLinha]) {
      cursoLinha = abreviacoes[cursoLinha];
    }
    linha[cursoIndex] = cursoLinha;

    // Filtra e ignora lixo/títulos
    if (cursoLinha === "INGRESSANTES" || cursoLinha === "EVASÃO" || cursoLinha === "PROUNI" || status === "") continue;

    if (tipo === "PROUNI") {
      prouni.push(linha);
    } else if (status === "Cursando" && tipo !== "PROUNI") {
      ingressantes.push(linha);
    } else if (status === "Cancelado" || status === "Trancado" || status === "Transferência para outra IES") {
      evasao.push(linha);
    }
  }

  // Ordenação por curso e data
  const ordenar = (lista) => lista.sort((a, b) => {
    const cA = String(a[cursoIndex]).toUpperCase();
    const cB = String(b[cursoIndex]).toUpperCase();
    if (cA < cB) return -1;
    if (cA > cB) return 1;
    return new Date(a[dataIndex]) - new Date(b[dataIndex]);
  });

  ingressantes = ordenar(ingressantes);
  evasao = ordenar(evasao);
  prouni = ordenar(prouni);

  // Re-escrita formatada
  base.clear();
  base.appendRow(cabecalho);
  base.getRange(1, 1, 1, cabecalho.length).setFontWeight("bold").setBackground("#d9d9d9");

  const formatarBloco = (titulo, lista) => {
    if (lista.length > 0) {
      base.appendRow([titulo]);
      base.getRange(base.getLastRow(), 1, 1, cabecalho.length).merge().setFontWeight("bold").setBackground("#d9d9d9");
      base.getRange(base.getLastRow() + 1, 1, lista.length, cabecalho.length).setValues(lista);
      base.getRange(base.getLastRow() - lista.length + 1, dataIndex + 1, lista.length, 1).setNumberFormat("dd/MM/yyyy");
    }
  };

  formatarBloco("INGRESSANTES", ingressantes);
  formatarBloco("EVASÃO", evasao);
  formatarBloco("PROUNI", prouni);

  base.getDataRange().setHorizontalAlignment("center").setVerticalAlignment("middle");
  base.autoResizeColumns(1, cabecalho.length);
  SpreadsheetApp.getUi().alert("Relatório Atualizado!");
}

function enviarRelatorioExcel() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName("BASE");
  const grupos = ss.getSheetByName("GRUPOS_ENVIO");

  const dados = base.getDataRange().getValues();
  const cursoIndex = dados[0].indexOf("CURSO");

  let cursoDetectado = "";
  for (let i = 1; i < dados.length; i++) {
    let v = String(dados[i][cursoIndex]);
    if (v && v !== "INGRESSANTES" && v !== "EVASÃO" && v !== "PROUNI") {
      cursoDetectado = v;
      break;
    }
  }

  if (!cursoDetectado) return ui.alert("Curso não encontrado.");

  // Busca o gestor na aba de grupos
  const gestores = grupos.getDataRange().getValues();
  let email = ""; let cc = "";
  for (let i = 1; i < gestores.length; i++) {
    const palavras = String(gestores[i][0]).split(",");
    for (let p of palavras) {
      const regex = new RegExp("(^|\\s|-)" + p.trim().replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + "(\\s|-|$)", "i");
      if (regex.test(cursoDetectado)) {
        email = gestores[i][1]; cc = gestores[i][2];
        break;
      }
    }
    if (email) break;
  }

  if (!email) return ui.alert("Gestor não encontrado para: " + cursoDetectado);

  const token = ScriptApp.getOAuthToken();
  const blob = UrlFetchApp.fetch("https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx", {headers: {Authorization: "Bearer " + token}}).getBlob().setName("Relatório.xlsx");

  MailApp.sendEmail({
    to: email, cc: cc,
    subject: "Planilha de Retenção - " + Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy"),
    body: "Segue em anexo o relatório automatizado atualizado.",
    attachments: [blob]
  });

  ui.alert("Relatório enviado!");
}

// =========================================================
// 4. DICIONÁRIO DE DADOS E CONFIGURAÇÃO
// =========================================================
function getTabelas() {
  const abreviacoes = {
      "Medicina Veterinária": "Med Vet",
      "Engenharia Mecânica": "Eng Mecânica",
      "Engenharia de Produção": "Eng Produção",
      "Engenharia de Controle e Automação": "Eng Aut",
      "Engenharia Civil": "Eng Civil",
      "Arquitetura e Urbanismo": "Arquitetura",
      "Ciência da Computação": "BCC",
      "Gestão de Recursos Humanos": "RH",
      "Educação Física": "Ed Física",
      "Ciências Contábeis": "Contábeis",
      "Administração": "ADM",
      "Inteligência Artificial": "IA"
  };

  const emailsGestores = {
      "ADM": "gestor.adm@instituicao.edu.br",
      "Med Vet": "gestor.medvet@instituicao.edu.br",
      "Eng Mecânica": "gestor.eng@instituicao.edu.br",
      "Eng Produção": "gestor.eng@instituicao.edu.br",
      "Eng Aut": "gestor.eng@instituicao.edu.br",
      "Eng Civil": "gestor.eng@instituicao.edu.br",
      "Arquitetura": "gestor.arq@instituicao.edu.br",
      "BCC": "gestor.tec@instituicao.edu.br",
      "RH": "gestor.rh@instituicao.edu.br",
      "Ed Física": "gestor.edfisica@instituicao.edu.br",
      "Contábeis": "gestor.contabeis@instituicao.edu.br",
      "IA": "coordenador.ia@instituicao.edu.br",
      "Biomedicina": "gestor.saude@instituicao.edu.br",
      "Direito": "gestor.direito@instituicao.edu.br",
      "Enfermagem": "gestor.saude@instituicao.edu.br",
      "Farmácia": "gestor.saude@instituicao.edu.br",
      "Fisioterapia": "gestor.saude@instituicao.edu.br",
      "Nutrição": "gestor.saude@instituicao.edu.br",
      "Pedagogia": "gestor.pedagogia@instituicao.edu.br",
      "Psicologia": "gestor.psico@instituicao.edu.br",
      "Marketing": "gestor.mkt@instituicao.edu.br",
      "Logística": "gestor.logistica@instituicao.edu.br"
  };

  return { abreviacoes, emailsGestores };
}
