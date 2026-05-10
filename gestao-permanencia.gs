function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Gestão de Permanência')
    .addItem('1. Atualizar Planilha (Farol e Cursos)', 'atualizarPlanilha')
    .addItem('2. Gerar Relatório de Cursos', 'gerarRelatorio')
    .addItem('3. Enviar E-mails de Contato', 'enviarEmails')
    .addToUi();
}

function atualizarPlanilha() {
  const folha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = folha.getDataRange();
  const dados = range.getValues();
  const cabecalho = dados[0];
  
  const idxFarol = cabecalho.indexOf("Farol");
  const idxCurso = cabecalho.indexOf("Curso");
  const idxReqFin = cabecalho.indexOf("Req. Financeiro");
  const idxAusencia = cabecalho.indexOf("Ausência >= 20%");
  const idxReqHist = cabecalho.indexOf("Req. Histórico");
  const idxMens = cabecalho.indexOf("Mens. Vencidas");
  
  if (idxFarol === -1 || idxCurso === -1) {
    SpreadsheetApp.getUi().alert("Erro: Colunas 'Farol' ou 'Curso' não encontradas.");
    return;
  }

  // Dicionário com o de-para dos nomes dos cursos
  const mapaCursos = {
    "Administração": "ADM",
    "Arquitetura e Urbanismo": "Arquitetura",
    "Ciência da Computação": "Computação",
    "Ciências Contábeis": "Contábeis",
    "Educação Física": "Ed. Física",
    "Engenharia Civil": "Eng Civil",
    "Engenharia de Controle e Automação": "Eng Cont Aut",
    "Engenharia de Produção": "Eng Prod",
    "Engenharia Mecânica": "Eng Mec",
    "Gestão de Recursos Humanos": "RH",
    "Inteligência Artificial": "IA",
    "Medicina Veterinária": "Med Vet"
  };

  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];
    
    // 1. ATUALIZAR NOME DO CURSO
    let cursoAtual = linha[idxCurso];
    let cursoLimpo = cursoAtual ? cursoAtual.toString().trim() : "";
    
    // Se o curso limpo existir na nossa lista, substitui pelo nome curto
    if (mapaCursos[cursoLimpo]) {
      dados[i][idxCurso] = mapaCursos[cursoLimpo];
    }
    
    // 2. ATUALIZAR FAROL
    let reqFin = linha[idxReqFin] === 1 ? 1 : 0;
    let ausencia = linha[idxAusencia] === 1 ? 1 : 0;
    let reqHist = linha[idxReqHist] === 1 ? 1 : 0;
    let mensVencidas = linha[idxMens] > 0 ? 1 : 0;
    
    let qtdRiscos = reqFin + ausencia + reqHist + mensVencidas;
    let farol = "";
    
    if (qtdRiscos === 1) farol = "🟣";
    else if (qtdRiscos === 2) farol = "🟡";
    else if (qtdRiscos >= 3) farol = "🟠";
    
    dados[i][idxFarol] = farol;
  }
  
  range.setValues(dados);
  SpreadsheetApp.getUi().alert("✅ Planilha atualizada: Farol calculado e nomes dos cursos padronizados!");
}

function gerarRelatorio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folhaDados = ss.getSheets()[0]; 
  const dados = folhaDados.getDataRange().getValues();
  const cabecalho = dados[0];
  
  const idxCurso = cabecalho.indexOf("Curso");
  const idxReqFin = cabecalho.indexOf("Req. Financeiro");
  const idxAusencia = cabecalho.indexOf("Ausência >= 20%");
  const idxMens = cabecalho.indexOf("Mens. Vencidas");
  
  if (idxCurso === -1) {
    SpreadsheetApp.getUi().alert("Erro: Coluna 'Curso' não encontrada.");
    return;
  }

  let resumo = {};
  
  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];
    let curso = linha[idxCurso];
    
    if (!curso) continue; 
    
    if (!resumo[curso]) {
      resumo[curso] = { alunos: 0, faltas: 0, financeiro: 0, inadimplencia: 0, totalProblemas: 0 };
    }
    
    resumo[curso].alunos += 1;
    if (linha[idxAusencia] === 1) {
      resumo[curso].faltas += 1;
      resumo[curso].totalProblemas += 1;
    }
    if (linha[idxReqFin] === 1) {
      resumo[curso].financeiro += 1;
      resumo[curso].totalProblemas += 1;
    }
    if (linha[idxMens] > 0) {
      resumo[curso].inadimplencia += 1;
      resumo[curso].totalProblemas += 1;
    }
  }
  
  let tabelaRelatorio = [
    ["Curso", "Total de Alunos (Em Risco)", "Alunos com Faltas (>= 20%)", "Req. Financeiro", "Inadimplência (> 0)", "Pontuação Crítica (Total de Problemas)"]
  ];
  
  let linhasCursos = Object.keys(resumo).map(curso => [
    curso,
    resumo[curso].alunos,
    resumo[curso].faltas,
    resumo[curso].financeiro,
    resumo[curso].inadimplencia,
    resumo[curso].totalProblemas
  ]);
  
  linhasCursos.sort((a, b) => b[5] - a[5]);
  tabelaRelatorio = tabelaRelatorio.concat(linhasCursos);
  
  let nomeAbaRelatorio = "Resumo de Riscos";
  let abaRelatorio = ss.getSheetByName(nomeAbaRelatorio);
  
  if (abaRelatorio) {
    abaRelatorio.clear();
  } else {
    abaRelatorio = ss.insertSheet(nomeAbaRelatorio);
  }
  
  abaRelatorio.getRange(1, 1, tabelaRelatorio.length, tabelaRelatorio[0].length).setValues(tabelaRelatorio);
  abaRelatorio.getRange("A1:F1").setFontWeight("bold").setBackground("#d9ead3");
  abaRelatorio.autoResizeColumns(1, 6);
  
  SpreadsheetApp.getUi().alert("📊 Relatório gerado com sucesso! Verifique a aba 'Resumo de Riscos'.");
}

function enviarEmails() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folha = ss.getSheets()[0]; 
  const dados = folha.getDataRange().getValues();
  const cabecalho = dados[0];
  
  const idxNome = cabecalho.indexOf("Nome");
  const idxEmail = cabecalho.indexOf("E-mail"); 
  const idxEmailInst = cabecalho.indexOf("E-mail Institucional"); 
  const idxAusencia = cabecalho.indexOf("Ausência >= 20%");
  const idxReqHist = cabecalho.indexOf("Req. Histórico");
  const idxMens = cabecalho.indexOf("Mens. Vencidas");
  const idxReqFin = cabecalho.indexOf("Req. Financeiro");
  const idxObs = cabecalho.indexOf("Observações"); 
  
  if (idxObs === -1) {
    ui.alert("Erro: Coluna 'Observações' não encontrada na planilha.");
    return;
  }
  if (idxEmail === -1 && idxEmailInst === -1) {
    ui.alert("Erro: As colunas 'E-mail' e 'E-mail Institucional' não foram encontradas.");
    return;
  }
  
  let emailsEnviados = 0;

  for (let i = 1; i < dados.length; i++) {
    let linha = dados[i];
    let nome = linha[idxNome];
    let obsAtual = linha[idxObs];
    
    let emailPessoal = idxEmail !== -1 ? linha[idxEmail] : "";
    let emailInst = idxEmailInst !== -1 ? linha[idxEmailInst] : "";
    
    if (obsAtual !== "") continue;

    let emailDestino = emailPessoal;
    let emailCopia = emailInst;
    
    if (!emailDestino || emailDestino === "") {
      emailDestino = emailInst;
      emailCopia = "";
    }
    
    if (!emailDestino || emailDestino === "") continue;

    let faltas = linha[idxAusencia] === 1;
    let historico = linha[idxReqHist] === 1;
    let parcelasVencidas = linha[idxMens] >= 2; 
    let reqFin = linha[idxReqFin] >= 1; 
    
    let assunto = "";
    let corpo = "";
    
    if (faltas && parcelasVencidas) {
      assunto = "Temos algo importante para conversar com você!";
      corpo = `Olá, ${nome}!\n\nEstamos entrando em contato para entender como tem sido sua rotina acadêmica e verificar se está precisando de algum tipo de suporte.\n\nPercebemos que houve algumas ausências recentes e também identificamos uma pendência financeira.\n\nGostaríamos de saber se você está enfrentando alguma dificuldade que esteja impactando sua participação no curso.\n\nNosso objetivo é te apoiar da melhor forma possível e encontrar juntos uma solução para que você possa continuar sua jornada com tranquilidade.\n\nPodemos conversar?\n\nConte conosco!`;
    } 
    else if (faltas && historico) {
      assunto = "Estamos aqui para te apoiar!";
      corpo = `Olá, ${nome}!\n\nNotamos que você se ausentou em algumas aulas nas últimas semanas e também percebemos sua solicitação recente de histórico/conteúdo programático.\n\nGostaríamos de saber como você está e entender melhor como podemos te apoiar neste momento.\n\nEstá enfrentando alguma dificuldade para frequentar as aulas ou precisa de alguma orientação específica?\n\nNosso objetivo é garantir que você tenha todo o suporte necessário para seguir bem nos seus estudos.\n\nPodemos conversar?\n\nConte conosco!\n\nAtenciosamente,\nEquipe de Retenção Acadêmica`;
    }
    else if (faltas) {
      assunto = "Estamos aqui para te apoiar!";
      corpo = `Olá, ${nome}!\n\nNotamos que você se ausentou em algumas aulas nas últimas semanas e gostaríamos de saber como você está.\n\nEstá enfrentando alguma dificuldade para frequentar as aulas? Se precisar de apoio, queremos te ajudar da melhor forma possível!\n\nPodemos conversar?\n\nConte conosco!\n\nAtenciosamente,\nEquipe de Retenção Acadêmica`;
    }
    else if (parcelasVencidas) {
      assunto = "Precisamos conversar com você!";
      corpo = `Olá, ${nome}!\n\nEstamos entrando em contato para entender como tem sido sua rotina acadêmica e verificar se está precisando de algum tipo de suporte. Se houver algo dificultando a continuidade do curso, especialmente relacionado à parte financeira, queremos te ajudar a encontrar uma alternativa.\n\nNosso objetivo é te apoiar!\n\nAtenciosamente,\nEquipe de Retenção Acadêmica`;
    }
    else if (reqFin) {
      assunto = "Podemos conversar sobre sua vida acadêmica?";
      corpo = `Oi, ${nome}!\n\nEsperamos que esteja tudo bem com você!\n\nGostaríamos de conversar e entender se há algo no seu dia a dia acadêmico que possamos te ajudar. Se for algo relacionado à parte financeira, podemos explorar juntos algumas possibilidades de apoio.\n\nFico à disposição,\nEquipe de Retenção Acadêmica`;
    }
    else if (historico) {
      assunto = "Gostaríamos de falar com você";
      corpo = `Oi, ${nome}!\n\nPercebemos que você solicitou histórico/conteúdo programático recentemente e gostaríamos de conversar para entender como podemos te auxiliar neste momento. Nosso objetivo é garantir que você tenha todo o suporte necessário para seguir bem nos seus estudos.\n\nVamos conversar?\n\nAtenciosamente,\nEquipe de Retenção Acadêmica`;
    }

    if (assunto !== "") {
      let opcoes = {};
      
      if (emailCopia !== "") {
        opcoes.cc = emailCopia;
      }
      
      // MailApp.sendEmail(emailDestino, assunto, corpo, opcoes); // Descomentar em produção
      
      folha.getRange(i + 1, idxObs + 1).setValue("E-mail Enviado"); 
      emailsEnviados++;
    }
  }
  
  ui.alert(`Processo concluído! ${emailsEnviados} e-mails processados e marcados na coluna 'Observações'.`);
}
