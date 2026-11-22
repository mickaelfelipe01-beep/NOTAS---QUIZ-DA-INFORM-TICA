
const PLANILHA_ID = "https://script.google.com/macros/s/AKfycbybcOGXGkpeI8QZUAsiRHBVVIchzHnw4tRsbp18U6_bOJHx4qKwXJDn-VWY_8oTfC5dEA/exec";

// Troque por uma senha forte, conhecida apenas pelo professor.
const SENHA_PAINEL = "TROQUE_POR_UMA_SENHA_FORTE";

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Painel")
    .setTitle("Painel do Professor - Quiz da Informática");
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    const dados = JSON.parse(e.postData.contents);

    const nome = String(dados.nome || "").trim().slice(0, 80);
    const equipe = String(dados.equipe || "").trim().slice(0, 80);
    const pontos = Number(dados.pontos);
    const acertos = Number(dados.acertos);
    const total = Number(dados.total);

    if (
      !nome || !equipe ||
      !Number.isInteger(pontos) ||
      !Number.isInteger(acertos) ||
      !Number.isInteger(total) ||
      total !== 50 ||
      acertos < 0 || acertos > total ||
      pontos < 0 || pontos > 100 ||
      pontos !== acertos * 2
    ) {
      throw new Error("Dados inválidos.");
    }

    lock.waitLock(10000);

    try {
      const planilha = SpreadsheetApp.openById(PLANILHA_ID);
      let aba = planilha.getSheetByName("Resultados");

      if (!aba) {
        aba = planilha.insertSheet("Resultados");
        aba.appendRow([
          "Data e hora",
          "Nome do aluno",
          "Equipe",
          "Pontuação",
          "Acertos",
          "Total de perguntas",
          "Aproveitamento (%)"
        ]);

        aba.setFrozenRows(1);
        aba.getRange("A1:G1")
          .setFontWeight("bold")
          .setBackground("#6558E8")
          .setFontColor("#FFFFFF");
      }

      aba.appendRow([
        new Date(),
        nome,
        equipe,
        pontos,
        acertos,
        total,
        Math.round((acertos / total) * 100)
      ]);
    } finally {
      lock.releaseLock();
    }

    return resposta({ sucesso: true });

  } catch (erro) {
    return resposta({
      sucesso: false,
      erro: String(erro.message)
    });
  }
}

function resposta(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

// Esta função é chamada pelo painel do professor.
function getDashboardData(senha) {
  if (String(senha || "") !== SENHA_PAINEL) {
    throw new Error("Senha incorreta.");
  }

  const planilha = SpreadsheetApp.openById(PLANILHA_ID);
  const aba = planilha.getSheetByName("Resultados");

  if (!aba || aba.getLastRow() < 2) {
    return {
      alunos: [],
      equipes: [],
      totalAlunos: 0,
      totalEquipes: 0,
      mediaGeral: 0,
      melhorEquipe: "-"
    };
  }

  const dados = aba.getRange(
    2, 1, aba.getLastRow() - 1, 7
  ).getValues();

  const alunos = dados.map(linha => ({
    data: linha[0] instanceof Date
      ? Utilities.formatDate(
          linha[0],
          Session.getScriptTimeZone(),
          "dd/MM/yyyy HH:mm"
        )
      : String(linha[0]),
    nome: String(linha[1]),
    equipe: String(linha[2]),
    pontos: Number(linha[3]) || 0,
    acertos: Number(linha[4]) || 0,
    total: Number(linha[5]) || 50,
    percentual: Number(linha[6]) || 0
  }));

  const grupos = {};

  alunos.forEach(aluno => {
    const nomeEquipe = aluno.equipe.trim() || "Sem equipe";

    if (!grupos[nomeEquipe]) {
      grupos[nomeEquipe] = {
        nome: nomeEquipe,
        participantes: 0,
        somaPontos: 0,
        somaPercentuais: 0
      };
    }

    grupos[nomeEquipe].participantes++;
    grupos[nomeEquipe].somaPontos += aluno.pontos;
    grupos[nomeEquipe].somaPercentuais += aluno.percentual;
  });

  const equipes = Object.values(grupos).map(equipe => ({
    nome: equipe.nome,
    participantes: equipe.participantes,
    pontos: equipe.somaPontos,
    media: Math.round(
      equipe.somaPercentuais / equipe.participantes
    )
  }));

  equipes.sort((a, b) =>
    b.media - a.media ||
    b.pontos - a.pontos ||
    a.nome.localeCompare(b.nome)
  );

  const mediaGeral = alunos.length
    ? Math.round(
        alunos.reduce((s, a) => s + a.percentual, 0) /
        alunos.length
      )
    : 0;

  return {
    alunos: alunos.reverse(),
    equipes: equipes,
    totalAlunos: alunos.length,
    totalEquipes: equipes.length,
    mediaGeral: mediaGeral,
    melhorEquipe: equipes.length ? equipes[0].nome : "-"
  };
}
