/************************************************************
 * AUTOMAÇÃO DE VENCIMENTO DE DOCUMENTOS DE VEÍCULOS
 * Google Apps Script + Gemini + Google Drive + Google Sheets
 ************************************************************/


/*******************************
 * CONFIGURAÇÕES PRINCIPAIS
 *******************************/

// IMPORTANTE:
// Não recomendo deixar a chave exposta no código.
// Por enquanto, cole sua nova chave aqui.
// Depois podemos migrar para PropertiesService.
const GEMINI_API_KEY = "sua_chave_api_gemini_ou_outra_ia";


const GEMINI_MODEL = "gemini-3.1-flash-lite";
const PAUSA_ENTRE_CHAMADAS_MS = 5000;

const PASTA_MAE_ID = "1WsD8JgYc5QmAQq37wJtyJzMMUyaA0k7f";
const EMAIL_NOTIFICACAO = "seuemail@seuemail.com.br";

const NOME_PLANILHA_CONTROLE = "CONTROLE_IA_DOCUMENTOS_VEICULOS";

const ABA_LOG = "LOG_EXECUCAO";
const ABA_DB = "ARQUIVOS_PROCESSADOS";
const ABA_REGRAS = "REGRAS_VENCIMENTO";
const ABA_PENDENCIAS = "PENDENCIAS_REVISAO";
const ABA_CALENDARIO = "CALENDARIO_LICENCIAMENTO";

const DIAS_ALERTA = 60;


const MAX_ARQUIVOS_PROCESSAR_POR_EXECUCAO = 25;

// Somente estas pastas de primeiro nível serão processadas.
const PASTAS_RAIZ_PERMITIDAS = [
  "PASTA_RAIZ_GOOGLEDRIVE1",
  "PASTA_RAIZ_GOOGLEDRIVE1"
];

// Estas pastas serão ignoradas em qualquer nível e registradas em log.
const PASTAS_IGNORADAS = [
  "VENCIDOS",
  "VENCIDAS",
  "Z - AGREGADOS ANTIGOS",
  "Z - EX SOCIOS",
  "Z - SÓCIOS ANTIGOS",
  "VENCIDO",
  "DOCUMENTOS VENCIDOS",
  "DOC. ANTIGOS",
  "DOC ANTIGOS",
  "DOC VENCIDOS",
  "VENDIDOS",
  "VENDIDAS"
];


/*******************************
 * FUNÇÃO PRINCIPAL
 *******************************/

function verificarVencimentosDeVeiculos() {
  const contexto = inicializarAmbiente();
  const inicio = new Date();

  registrarLog(contexto, "INFO", "PROCESSO", "🚀 Iniciando verificação de documentos de veículos.");

let contador = {
  encontrados: 0,
  processados: 0,
  ignorados: 0,
  pulados: 0,
  erros: 0,
  pastasProcessadas: 0,
  pastasIgnoradas: 0,
  pastasForaEscopo: 0,
  limiteAtingido: false
};

  try {
    const pastaMae = DriveApp.getFolderById(PASTA_MAE_ID);

    registrarLog(
      contexto,
      "INFO",
      "PASTA_RAIZ",
      `📂 Pasta inicial localizada: ${pastaMae.getName()}`
    );

    registrarLog(
      contexto,
      "INFO",
      "ESCOPO",
      `🎯 Somente estas pastas serão processadas: ${PASTAS_RAIZ_PERMITIDAS.join(", ")}`
    );

    registrarLog(
      contexto,
      "INFO",
      "ESCOPO",
      `⛔ Pastas ignoradas em qualquer nível: ${PASTAS_IGNORADAS.join(", ")}`
    );

    const dbMap = carregarBancoArquivos(contexto);
    const regras = carregarRegrasVencimento(contexto);
    const calendarios = carregarCalendariosLicenciamento(contexto);

    const totalReclassificados = reclassificarRegistrosAntigosCRLVeRNC(contexto);

    if (totalReclassificados > 0) {
      registrarLog(
        contexto,
        "INFO",
        "RECLASSIFICACAO",
        `♻️ ${totalReclassificados} registro(s) antigo(s) de CRLV/RNC foram reclassificados automaticamente.`
      );
    }

    processarSomentePastasPermitidasDaRaiz(
      pastaMae,
      contexto,
      dbMap,
      regras,
      calendarios,
      contador
    );

    const totalPendencias = atualizarPendenciasRevisao(contexto);

    registrarLog(
      contexto,
      "INFO",
      "PENDENCIAS",
      `📌 Pendências de revisão atualizadas. Total: ${totalPendencias}.`
    );

    if (totalPendencias > 0) {
      registrarLog(
        contexto,
        "WARN",
        "PENDENCIAS",
        `⚠️ Existem ${totalPendencias} documento(s) pendente(s) de revisão humana. Consulte a aba ${ABA_PENDENCIAS}.`
      );
    }

    const alertas = gerarAlertasAPartirDoBanco(contexto);

    if (alertas.length > 0) {
      enviarEmailAlertas(alertas, totalPendencias);

      registrarLog(
        contexto,
        "INFO",
        "EMAIL",
        `✉️ E-mail de alerta enviado com ${alertas.length} documento(s) vencido(s) ou próximo(s) ao vencimento.`
      );
    } else {
      registrarLog(
        contexto,
        "INFO",
        "EMAIL",
        "🎉 Nenhum documento vencido ou próximo do vencimento encontrado."
      );

      if (totalPendencias > 0) {
        enviarEmailSomentePendencias(totalPendencias);
      }
    }

  } catch (erro) {
    contador.erros++;

    registrarLog(
      contexto,
      "ERRO",
      "PROCESSO",
      `❌ Erro crítico na execução principal: ${erro.message}`
    );
  }

  const fim = new Date();
  const duracaoSegundos = Math.round((fim.getTime() - inicio.getTime()) / 1000);

  registrarLog(
    contexto,
    "INFO",
    "RESUMO",
    `🏁 Processo concluído em ${duracaoSegundos}s. Pastas processadas: ${contador.pastasProcessadas}, Pastas ignoradas: ${contador.pastasIgnoradas}, Pastas fora do escopo: ${contador.pastasForaEscopo}, Arquivos encontrados: ${contador.encontrados}, Processados: ${contador.processados}, Pulados: ${contador.pulados}, Ignorados: ${contador.ignorados}, Erros: ${contador.erros}.`
  );
}


/*******************************
 * INICIALIZAÇÃO DA PLANILHA
 *******************************/

function inicializarAmbiente() {
  const pastaMae = DriveApp.getFolderById(PASTA_MAE_ID);
  const planilha = obterOuCriarPlanilhaControle(pastaMae);

  const abaLog = obterOuCriarAba(planilha, ABA_LOG, [
    "data_hora",
    "nivel",
    "categoria",
    "mensagem"
  ]);

  const abaDb = obterOuCriarAba(planilha, ABA_DB, [
    "file_id",
    "nome_arquivo",
    "caminho_pasta",
    "url",
    "mime_type",
    "tamanho_bytes",
    "ultima_modificacao_drive",
    "assinatura_arquivo",
    "data_processamento",
    "status_processamento",
    "tipo_documento",
    "placa",
    "uf",
    "data_emissao",
    "data_vencimento",
    "vencimento_calculado",
    "fonte_vencimento",
    "dias_para_vencer",
    "situacao",
    "observacao",
    "json_ia"
  ]);

  const abaRegras = obterOuCriarAba(planilha, ABA_REGRAS, [
    "tipo_documento",
    "uf",
    "criterio",
    "quantidade",
    "observacao"
  ]);

  const abaPendencias = obterOuCriarAba(planilha, ABA_PENDENCIAS, [
    "data_atualizacao",
    "motivo",
    "file_id",
    "nome_arquivo",
    "caminho_pasta",
    "url",
    "tipo_documento",
    "placa",
    "uf",
    "data_emissao",
    "data_vencimento",
    "fonte_vencimento",
    "situacao",
    "observacao",
    "acao_recomendada"
  ]);

  const abaCalendario = obterOuCriarAba(planilha, ABA_CALENDARIO, [
    "uf",
    "ano_calendario",
    "final_placa",
    "data_limite",
    "tipo_data",
    "fonte",
    "url_fonte",
    "observacao",
    "ativo"
  ]);

  garantirColunasAba(abaDb, ["exercicio_crlv", "ano_calendario_crlv"]);
  inicializarCalendariosNacionaisSeNecessario(abaCalendario);

  inicializarRegrasPadraoSeVazio(abaRegras);
  garantirRegraCRLV(abaRegras);
  garantirRegraRNC(abaRegras);

  return {
    pastaMae: pastaMae,
    planilha: planilha,
    abaLog: abaLog,
    abaDb: abaDb,
    abaRegras: abaRegras,
    abaPendencias: abaPendencias,
    abaCalendario: abaCalendario
  };
}

function garantirRegraCRLV(abaRegras) {
  garantirOuAtualizarRegra(
    abaRegras,
    "CRLV",
    "PADRAO",
    "CONTROLAR_POR_EXERCICIO",
    "",
    "CRLV é controlado pelo exercício anual. Não exigir data de vencimento impressa."
  );
}


function garantirOuAtualizarRegra(abaRegras, tipo, uf, criterio, quantidade, observacao) {
  const dados = abaRegras.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const tipoAtual = String(dados[i][0] || "").trim().toUpperCase();
    const ufAtual = String(dados[i][1] || "PADRAO").trim().toUpperCase();

    if (tipoAtual === tipo && ufAtual === uf) {
      abaRegras.getRange(i + 1, 3, 1, 3).setValues([[
        criterio,
        quantidade,
        observacao
      ]]);
      return;
    }
  }

  abaRegras.appendRow([tipo, uf, criterio, quantidade, observacao]);
}


function garantirRegraRNC(abaRegras) {
  garantirOuAtualizarRegra(
    abaRegras,
    "RNC",
    "PADRAO",
    "NAO_APLICAVEL",
    "",
    "Registro de Não-Conformidade vinculado à inspeção. Não possui vencimento próprio."
  );
}



function garantirColunasAba(aba, colunasObrigatorias) {
  const ultimaColuna = Math.max(aba.getLastColumn(), 1);
  const cabecalho = aba.getRange(1, 1, 1, ultimaColuna).getValues()[0];

  colunasObrigatorias.forEach(coluna => {
    if (cabecalho.indexOf(coluna) === -1) {
      aba.getRange(1, aba.getLastColumn() + 1).setValue(coluna);
      cabecalho.push(coluna);
    }
  });
}


function inicializarCalendariosNacionaisSeNecessario(abaCalendario) {
  const fonteCompilacao = "CALENDARIO_NACIONAL_2026";
  const urlCompilacao = "https://www.mobills.com.br/tabelas/licenciamento-detran/";
  const observacao = "Calendário 2026 consultado por UF. Confirmar anualmente nas fontes oficiais antes da renovação do exercício.";

  const grupos = {
    "AC": {"1":"31/03/2026","2":"31/03/2026","3":"30/04/2026","4":"30/04/2026","5":"29/05/2026","6":"30/06/2026","7":"31/07/2026","8":"31/08/2026","9":"30/09/2026","0":"30/10/2026"},
    "AL": {"1":"27/02/2026","2":"27/02/2026","3":"31/03/2026","4":"31/03/2026","5":"30/04/2026","6":"30/04/2026","7":"29/05/2026","8":"29/05/2026","9":"30/06/2026","0":"30/06/2026"},
    "AM": {"1":"28/03/2026","2":"30/04/2026","3":"31/05/2026","4":"28/06/2026","5":"31/07/2026","6":"31/08/2026","7":"30/09/2026","8":"31/10/2026","9":"30/11/2026","0":"31/12/2026"},
    "BA": {"1":"30/07/2026","2":"31/07/2026","3":"29/08/2026","4":"30/08/2026","5":"29/09/2026","6":"30/09/2026","7":"30/10/2026","8":"31/10/2026","9":"28/11/2026","0":"28/11/2026"},
    "CE": {"1":"10/03/2026","2":"10/04/2026","3":"12/05/2026","4":"10/06/2026","5":"10/07/2026","6":"11/08/2026","7":"10/09/2026","8":"10/10/2026","9":"10/11/2026","0":"10/12/2026"},
    "DF": {"1":"31/05/2026","2":"31/05/2026","3":"30/06/2026","4":"30/06/2026","5":"30/06/2026","6":"30/06/2026","7":"31/07/2026","8":"31/07/2026","9":"31/07/2026","0":"31/07/2026"},
    "ES": {"1":"09/09/2026","2":"09/09/2026","3":"10/09/2026","4":"10/09/2026","5":"11/09/2026","6":"11/09/2026","7":"14/09/2026","8":"14/09/2026","9":"15/09/2026","0":"15/09/2026"},
    "GO": {"1":"15/09/2026","2":"15/09/2026","3":"15/10/2026","4":"15/10/2026","5":"15/10/2026","6":"15/10/2026","7":"15/10/2026","8":"15/10/2026","9":"15/10/2026","0":"15/10/2026"},
    "MA": {"1":"06/05/2026","2":"06/05/2026","3":"13/05/2026","4":"13/05/2026","5":"20/05/2026","6":"20/05/2026","7":"27/05/2026","8":"27/05/2026","9":"04/06/2026","0":"04/06/2026"},
    "MT": {"1":"31/03/2026","2":"31/03/2026","3":"31/03/2026","4":"30/04/2026","5":"29/05/2026","6":"30/06/2026","7":"31/07/2026","8":"31/08/2026","9":"30/09/2026","0":"30/10/2026"},
    "MS": {"1":"30/04/2026","2":"30/04/2026","3":"29/05/2026","4":"30/06/2026","5":"30/06/2026","6":"31/07/2026","7":"31/08/2026","8":"31/08/2026","9":"30/09/2026","0":"30/10/2026"},
    "MG": {"1":"31/03/2026","2":"31/03/2026","3":"31/03/2026","4":"31/03/2026","5":"31/03/2026","6":"31/03/2026","7":"31/03/2026","8":"31/03/2026","9":"31/03/2026","0":"31/03/2026"},
    "PB": {"1":"31/03/2026","2":"30/04/2026","3":"29/05/2026","4":"30/06/2026","5":"31/07/2026","6":"31/08/2026","7":"30/09/2026","8":"30/10/2026","9":"30/11/2026","0":"30/12/2026"},
    "PE": {"1":"10/11/2026","2":"10/11/2026","3":"13/11/2026","4":"13/11/2026","5":"19/11/2026","6":"19/11/2026","7":"24/11/2026","8":"24/11/2026","9":"27/11/2026","0":"27/11/2026"},
    "PI": {"1":"31/03/2026","2":"31/03/2026","3":"31/03/2026","4":"31/03/2026","5":"31/03/2026","6":"31/03/2026","7":"31/03/2026","8":"31/03/2026","9":"31/03/2026","0":"31/03/2026"},
    "RJ": {"1":"31/07/2026","2":"31/07/2026","3":"31/08/2026","4":"31/08/2026","5":"31/08/2026","6":"30/09/2026","7":"30/09/2026","8":"30/09/2026","9":"30/09/2026","0":"31/07/2026"},
    "RS": {"1":"31/07/2026","2":"31/07/2026","3":"31/07/2026","4":"31/07/2026","5":"31/07/2026","6":"31/07/2026","7":"31/07/2026","8":"31/07/2026","9":"31/07/2026","0":"31/07/2026"},
    "RO": {"1":"31/03/2026","2":"31/03/2026","3":"31/03/2026","4":"30/04/2026","5":"29/05/2026","6":"30/06/2026","7":"31/07/2026","8":"31/08/2026","9":"30/09/2026","0":"30/10/2026"},
    "RR": {"1":"31/03/2026","2":"31/03/2026","3":"31/03/2026","4":"31/03/2026","5":"31/03/2026","6":"31/03/2026","7":"31/03/2026","8":"31/03/2026","9":"31/03/2026","0":"31/03/2026"},
    "SC": {"1":"31/03/2026","2":"30/04/2026","3":"31/05/2026","4":"30/06/2026","5":"31/07/2026","6":"31/08/2026","7":"30/09/2026","8":"31/10/2026","9":"30/11/2026","0":"31/12/2026"},
    "SP": {"1":"31/07/2026","2":"31/07/2026","3":"31/08/2026","4":"31/08/2026","5":"30/09/2026","6":"30/09/2026","7":"30/10/2026","8":"30/10/2026","9":"30/11/2026","0":"31/12/2026"},
    "SE": {"1":"30/06/2026","2":"30/06/2026","3":"31/07/2026","4":"31/07/2026","5":"31/08/2026","6":"30/09/2026","7":"31/10/2026","8":"30/11/2026","9":"31/12/2026","0":"31/01/2027"},
    "TO": {"1":"30/10/2026","2":"30/10/2026","3":"30/10/2026","4":"30/10/2026","5":"30/10/2026","6":"30/10/2026","7":"30/10/2026","8":"30/10/2026","9":"30/10/2026","0":"30/10/2026"}
  };

  const existentes = {};
  const dados = abaCalendario.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    const chave = `${String(dados[i][0] || "").toUpperCase()}|${Number(dados[i][1] || 0)}|${String(dados[i][2] || "")}`;
    existentes[chave] = true;
  }

  const linhas = [];
  Object.keys(grupos).forEach(uf => {
    Object.keys(grupos[uf]).forEach(finalPlaca => {
      const chave = `${uf}|2026|${finalPlaca}`;
      if (!existentes[chave]) {
        linhas.push([uf, 2026, finalPlaca, grupos[uf][finalPlaca], "DATA_LIMITE_LICENCIAMENTO", fonteCompilacao, urlCompilacao, observacao, "SIM"]);
      }
    });
  });

  // Calendário oficial de Rondônia 2025, necessário para CRLVs do exercício 2024.
  const ro2025 = {"1":"31/03/2025","2":"31/03/2025","3":"31/03/2025","4":"30/04/2025","5":"30/05/2025","6":"30/06/2025","7":"31/07/2025","8":"29/08/2025","9":"30/09/2025","0":"31/10/2025"};
  Object.keys(ro2025).forEach(finalPlaca => {
    const chave = `RO|2025|${finalPlaca}`;
    if (!existentes[chave]) {
      linhas.push(["RO", 2025, finalPlaca, ro2025[finalPlaca], "DATA_LIMITE_LICENCIAMENTO", "DETRAN-RO", "https://www.detran.ro.gov.br/", "Calendário 2025 usado para classificar CRLV do exercício 2024.", "SIM"]);
    }
  });

  // Estados com regra que exige tabela especial ou confirmação oficial detalhada.
  ["AP", "PA", "PR", "RN"].forEach(uf => {
    const chave = `${uf}|2026|PENDENTE`;
    if (!existentes[chave]) {
      linhas.push([uf, 2026, "PENDENTE", "", "CALENDARIO_ESPECIAL", `DETRAN-${uf}`, "", "Calendário estadual exige regra especial ou tabela oficial detalhada. Não usar data estimada.", "NAO"]);
    }
  });

  if (linhas.length > 0) {
    abaCalendario.getRange(abaCalendario.getLastRow() + 1, 1, linhas.length, linhas[0].length).setValues(linhas);
  }
}

function carregarCalendariosLicenciamento(contexto) {
  const dados = contexto.abaCalendario.getDataRange().getValues();
  const mapa = {};

  if (dados.length <= 1) {
    return mapa;
  }

  const cabecalho = dados[0];
  const idxUf = cabecalho.indexOf("uf");
  const idxAno = cabecalho.indexOf("ano_calendario");
  const idxFinal = cabecalho.indexOf("final_placa");
  const idxData = cabecalho.indexOf("data_limite");
  const idxTipo = cabecalho.indexOf("tipo_data");
  const idxFonte = cabecalho.indexOf("fonte");
  const idxUrl = cabecalho.indexOf("url_fonte");
  const idxObservacao = cabecalho.indexOf("observacao");
  const idxAtivo = cabecalho.indexOf("ativo");

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const uf = String(linha[idxUf] || "").trim().toUpperCase();
    const ano = Number(linha[idxAno] || 0);
    const finalPlaca = String(linha[idxFinal] || "").trim();
    const dataLimite = normalizarDataTexto(linha[idxData] || "");
    const ativo = String(linha[idxAtivo] || "SIM").trim().toUpperCase();

    if (!uf || !ano || !/^[0-9]$/.test(finalPlaca) || !dataLimite || ativo === "NAO") {
      continue;
    }

    mapa[`${uf}|${ano}|${finalPlaca}`] = {
      uf: uf,
      ano_calendario: ano,
      final_placa: finalPlaca,
      data_limite: dataLimite,
      tipo_data: String(linha[idxTipo] || "DATA_LIMITE_LICENCIAMENTO"),
      fonte: String(linha[idxFonte] || "CALENDARIO_CADASTRADO"),
      url_fonte: String(linha[idxUrl] || ""),
      observacao: String(linha[idxObservacao] || "")
    };
  }

  return mapa;
}


function consultarCalendarioLicenciamento(calendarios, uf, anoCalendario, placa) {
  const estado = String(uf || "").trim().toUpperCase();
  const finalPlaca = extrairFinalNumericoPlaca(placa);

  if (!estado || estado === "PADRAO" || !anoCalendario || finalPlaca === "") {
    return null;
  }

  return calendarios[`${estado}|${anoCalendario}|${finalPlaca}`] || null;
}


function extrairFinalNumericoPlaca(placa) {
  const numeros = String(placa || "").toUpperCase().replace(/[^0-9]/g, "");
  return numeros ? numeros.slice(-1) : "";
}


function obterOuCriarPlanilhaControle(pastaMae) {
  const arquivos = pastaMae.getFilesByName(NOME_PLANILHA_CONTROLE);

  if (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    return SpreadsheetApp.openById(arquivo.getId());
  }

  const planilha = SpreadsheetApp.create(NOME_PLANILHA_CONTROLE);
  const arquivoPlanilha = DriveApp.getFileById(planilha.getId());

  pastaMae.addFile(arquivoPlanilha);

  try {
    DriveApp.getRootFolder().removeFile(arquivoPlanilha);
  } catch (e) {
    // Pode falhar em alguns ambientes sem prejudicar o funcionamento.
  }

  return planilha;
}


function obterOuCriarAba(planilha, nomeAba, cabecalho) {
  let aba = planilha.getSheetByName(nomeAba);

  if (!aba) {
    aba = planilha.insertSheet(nomeAba);
  }

  if (aba.getLastRow() === 0) {
    aba.appendRow(cabecalho);
  }

  return aba;
}


function inicializarRegrasPadraoSeVazio(abaRegras) {
  if (abaRegras.getLastRow() > 1) {
    return;
  }

  const regrasPadrao = [
    [
      "CRLV",
      "PADRAO",
      "CONTROLAR_POR_EXERCICIO",
      "",
      "CRLV é controlado pelo exercício anual. Não exigir data de vencimento impressa."
    ],
    [
      "AET",
      "PADRAO",
      "ANOS_APOS_EMISSAO",
      "1",
      "Regra genérica. Validar conforme órgão emissor."
    ],
    [
      "CIV",
      "PADRAO",
      "ANOS_APOS_EMISSAO",
      "1",
      "Regra genérica. Validar regra oficial aplicável."
    ],
    [
      "CIPP",
      "PADRAO",
      "ANOS_APOS_EMISSAO",
      "1",
      "Regra genérica. Validar regra oficial aplicável."
    ],
    [
      "TACOGRAFO",
      "PADRAO",
      "ANOS_APOS_EMISSAO",
      "2",
      "Regra genérica. Validar conforme certificado."
    ],
    [
      "SEGURO",
      "PADRAO",
      "USAR_VALIDADE_DOCUMENTO",
      "",
      "Preferencialmente usar vigência final da apólice."
    ],
    [
      "ANTT",
      "PADRAO",
      "USAR_VALIDADE_DOCUMENTO",
      "",
      "Preferencialmente usar vencimento encontrado no documento."
    ],
    [
      "IBAMA",
      "PADRAO",
      "USAR_VALIDADE_DOCUMENTO",
      "",
      "Preferencialmente usar vencimento encontrado no documento."
    ],
    [
      "OUTRO",
      "PADRAO",
      "USAR_VALIDADE_DOCUMENTO",
      "",
      "Usa somente validade lida no documento."
    ]
  ];

  abaRegras
    .getRange(2, 1, regrasPadrao.length, regrasPadrao[0].length)
    .setValues(regrasPadrao);
}


/*******************************
 * LOG EM PLANILHA
 *******************************/

function registrarLog(contexto, nivel, categoria, mensagem) {
  const linha = [
    formatarDataHora(new Date()),
    nivel,
    categoria,
    mensagem
  ];

  contexto.abaLog.appendRow(linha);
  Logger.log(`[${nivel}] [${categoria}] ${mensagem}`);
}


/*******************************
 * CONTROLE DE ESCOPO DE PASTAS
 *******************************/

function processarSomentePastasPermitidasDaRaiz(
  pastaMae,
  contexto,
  dbMap,
  regras,
  calendarios,
  contador
) {
  const subpastas = pastaMae.getFolders();

  registrarLog(
    contexto,
    "INFO",
    "RAIZ",
    `🔎 Verificando subpastas diretas da raiz: ${pastaMae.getName()}`
  );

  while (subpastas.hasNext()) {
    if (contador.limiteAtingido) {
      break;
    }

    const subpasta = subpastas.next();
    const nomeOriginal = subpasta.getName();

    if (ehPastaIgnorada(nomeOriginal)) {
      contador.pastasIgnoradas++;

      registrarLog(
        contexto,
        "INFO",
        "PASTA_IGNORADA",
        `⛔ Pasta ignorada por regra: ${nomeOriginal}`
      );

      continue;
    }

    if (!ehPastaRaizPermitida(nomeOriginal)) {
      contador.pastasForaEscopo++;

      registrarLog(
        contexto,
        "INFO",
        "PASTA_FORA_ESCOPO",
        `⏭️ Pasta fora do escopo, não será processada: ${nomeOriginal}`
      );

      continue;
    }

    contador.pastasProcessadas++;

    registrarLog(
      contexto,
      "INFO",
      "PASTA_PERMITIDA",
      `✅ Pasta permitida para processamento: ${nomeOriginal}`
    );

    const interromper = varrerPastaPermitida(
      subpasta,
      pastaMae.getName(),
      contexto,
      dbMap,
      regras,
      calendarios,
      contador
    );

    if (interromper) {
      break;
    }
  }

  if (contador.limiteAtingido) {
    registrarLog(
      contexto,
      "INFO",
      "LIMITE",
      `⏭️ Limite de ${MAX_ARQUIVOS_PROCESSAR_POR_EXECUCAO} arquivo(s) processados atingido. A varredura foi encerrada. O fluxo continuará para pendências, alertas e envio de e-mail.`
    );
  }
}


function varrerPastaPermitida(
  pasta,
  caminhoPai,
  contexto,
  dbMap,
  regras,
  calendarios,
  contador
) {
  if (contador.limiteAtingido) {
    return true;
  }

  const nomePasta = pasta.getName();
  const caminhoAtual = caminhoPai
    ? `${caminhoPai}/${nomePasta}`
    : nomePasta;

  if (ehPastaIgnorada(nomePasta)) {
    contador.pastasIgnoradas++;

    registrarLog(
      contexto,
      "INFO",
      "PASTA_IGNORADA",
      `⛔ Pasta ignorada por regra: ${caminhoAtual}`
    );

    return false;
  }

  registrarLog(
    contexto,
    "INFO",
    "PASTA",
    `🔍 Varrendo pasta permitida: ${caminhoAtual}`
  );

  const arquivos = pasta.getFiles();

  while (arquivos.hasNext()) {
    /*
     * O limite é verificado antes de obter o próximo arquivo.
     * Ao atingir o limite, a função retorna imediatamente e
     * nenhuma outra pasta ou arquivo é procurado.
     */
    if (
      contador.processados >=
      MAX_ARQUIVOS_PROCESSAR_POR_EXECUCAO
    ) {
      contador.limiteAtingido = true;
      return true;
    }

    const arquivo = arquivos.next();
    contador.encontrados++;

    processarArquivoSeNecessario(
      arquivo,
      caminhoAtual,
      contexto,
      dbMap,
      regras,
      calendarios,
      contador
    );

    /*
     * Um arquivo pode ter elevado o contador até o limite.
     * Nesse caso, não abrimos nem procuramos o próximo arquivo.
     */
    if (
      contador.processados >=
      MAX_ARQUIVOS_PROCESSAR_POR_EXECUCAO
    ) {
      contador.limiteAtingido = true;
      return true;
    }
  }

  const subpastas = pasta.getFolders();

  while (subpastas.hasNext()) {
    if (contador.limiteAtingido) {
      return true;
    }

    const subpasta = subpastas.next();

    if (ehPastaIgnorada(subpasta.getName())) {
      contador.pastasIgnoradas++;

      registrarLog(
        contexto,
        "INFO",
        "PASTA_IGNORADA",
        `⛔ Pasta ignorada por regra: ${caminhoAtual}/${subpasta.getName()}`
      );

      continue;
    }

    const interromper = varrerPastaPermitida(
      subpasta,
      caminhoAtual,
      contexto,
      dbMap,
      regras,
      calendarios,
      contador
    );

    if (interromper) {
      return true;
    }
  }

  return false;
}


function ehPastaRaizPermitida(nomePasta) {
  const nome = normalizarNomePasta(nomePasta);

  return PASTAS_RAIZ_PERMITIDAS
    .map(normalizarNomePasta)
    .includes(nome);
}


function ehPastaIgnorada(nomePasta) {
  const nome = normalizarNomePasta(nomePasta);

  return PASTAS_IGNORADAS
    .map(normalizarNomePasta)
    .includes(nome);
}


function normalizarNomePasta(nome) {
  return String(nome || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}


/*******************************
 * PROCESSAMENTO DE ARQUIVOS
 *******************************/

function processarArquivoSeNecessario(arquivo, caminhoPasta, contexto, dbMap, regras, calendarios, contador) {
  const mime = arquivo.getMimeType();

  if (!(mime.includes("pdf") || mime.includes("image"))) {
    contador.ignorados++;

    registrarLog(
      contexto,
      "INFO",
      "ARQUIVO_IGNORADO",
      `⏭️ Ignorando arquivo não PDF/imagem: ${arquivo.getName()}`
    );

    return;
  }

  const assinaturaAtual = gerarAssinaturaArquivo(arquivo);
  const registroExistente = dbMap[arquivo.getId()];

  if (registroExistente && registroExistente.assinatura_arquivo === assinaturaAtual) {
    contador.pulados++;

    registrarLog(
      contexto,
      "INFO",
      "ARQUIVO_PULADO",
      `✅ Arquivo já processado e sem alteração: ${arquivo.getName()}`
    );

    return;
  }

  registrarLog(
    contexto,
    "INFO",
    "ARQUIVO",
    `📄 Processando arquivo: ${arquivo.getName()}`
  );

  try {
    const analiseIA = chamarGeminiParaAnalise(arquivo);

    Utilities.sleep(PAUSA_ENTRE_CHAMADAS_MS);

    const dadosNormalizados = normalizarResultadoIA(
      analiseIA,
      arquivo,
      caminhoPasta,
      regras,
      calendarios
    );

    salvarOuAtualizarRegistroArquivo(
      contexto,
      arquivo,
      caminhoPasta,
      assinaturaAtual,
      dadosNormalizados,
      analiseIA,
      dbMap
    );

    contador.processados++;

    registrarLog(
      contexto,
      "INFO",
      "ARQUIVO_PROCESSADO",
      `✅ Processado: ${arquivo.getName()} | Tipo: ${dadosNormalizados.tipo_documento} | Placa: ${dadosNormalizados.placa || "-"} | UF: ${dadosNormalizados.uf || "-"} | Emissão: ${dadosNormalizados.data_emissao || "-"} | Vencimento: ${dadosNormalizados.data_vencimento || "-"} | Situação: ${dadosNormalizados.situacao}`
    );

  } catch (erro) {
    contador.erros++;

    registrarLog(
      contexto,
      "ERRO",
      "ARQUIVO",
      `❌ Erro ao processar ${arquivo.getName()}: ${erro.message}`
    );
  }
}


function gerarAssinaturaArquivo(arquivo) {
  return [
    arquivo.getId(),
    arquivo.getName(),
    arquivo.getSize(),
    arquivo.getLastUpdated().toISOString()
  ].join("|");
}


/*******************************
 * BANCO DE DADOS EM PLANILHA
 *******************************/

function carregarBancoArquivos(contexto) {
  const aba = contexto.abaDb;
  const dados = aba.getDataRange().getValues();

  const mapa = {};

  if (dados.length <= 1) {
    return mapa;
  }

  const cabecalho = dados[0];

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const registro = {};

    for (let c = 0; c < cabecalho.length; c++) {
      registro[cabecalho[c]] = linha[c];
    }

    registro.numero_linha = i + 1;

    if (registro.file_id) {
      mapa[registro.file_id] = registro;
    }
  }

  return mapa;
}


function salvarOuAtualizarRegistroArquivo(
  contexto,
  arquivo,
  caminhoPasta,
  assinaturaArquivo,
  dados,
  analiseIA,
  dbMap
) {
  const aba = contexto.abaDb;
  const cabecalho = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  const valores = {
    file_id: arquivo.getId(),
    nome_arquivo: arquivo.getName(),
    caminho_pasta: caminhoPasta,
    url: arquivo.getUrl(),
    mime_type: arquivo.getMimeType(),
    tamanho_bytes: arquivo.getSize(),
    ultima_modificacao_drive: formatarDataHora(arquivo.getLastUpdated()),
    assinatura_arquivo: assinaturaArquivo,
    data_processamento: formatarDataHora(new Date()),
    status_processamento: dados.status_processamento,
    tipo_documento: dados.tipo_documento,
    placa: dados.placa,
    uf: dados.uf,
    data_emissao: dados.data_emissao,
    data_vencimento: dados.data_vencimento,
    vencimento_calculado: dados.vencimento_calculado,
    fonte_vencimento: dados.fonte_vencimento,
    dias_para_vencer: dados.dias_para_vencer,
    situacao: dados.situacao,
    observacao: dados.observacao,
    json_ia: JSON.stringify(analiseIA),
    exercicio_crlv: dados.exercicio_crlv || "",
    ano_calendario_crlv: dados.ano_calendario_crlv || ""
  };
  const linha = cabecalho.map(coluna => Object.prototype.hasOwnProperty.call(valores, coluna) ? valores[coluna] : "");
  const existente = dbMap[arquivo.getId()];

  if (existente && existente.numero_linha) {
    aba.getRange(existente.numero_linha, 1, 1, linha.length).setValues([linha]);
  } else {
    aba.appendRow(linha);
  }

  dbMap[arquivo.getId()] = {
    file_id: arquivo.getId(),
    assinatura_arquivo: assinaturaArquivo,
    numero_linha: existente && existente.numero_linha ? existente.numero_linha : aba.getLastRow()
  };
}


/*******************************
 * REGRAS DE VENCIMENTO
 *******************************/

function carregarRegrasVencimento(contexto) {
  const aba = contexto.abaRegras;
  const dados = aba.getDataRange().getValues();

  const regras = [];

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];

    if (!linha[0]) {
      continue;
    }

    regras.push({
      tipo_documento: String(linha[0]).toUpperCase().trim(),
      uf: String(linha[1] || "PADRAO").toUpperCase().trim(),
      criterio: String(linha[2] || "").toUpperCase().trim(),
      quantidade: linha[3],
      observacao: linha[4] || ""
    });
  }

  return regras;
}


function encontrarRegra(regras, tipoDocumento, uf) {
  const tipo = String(tipoDocumento || "OUTRO").toUpperCase().trim();
  const estado = String(uf || "PADRAO").toUpperCase().trim();

  let regra = regras.find(r => r.tipo_documento === tipo && r.uf === estado);

  if (regra) {
    return regra;
  }

  regra = regras.find(r => r.tipo_documento === tipo && r.uf === "PADRAO");

  if (regra) {
    return regra;
  }

  return regras.find(r => r.tipo_documento === "OUTRO") || null;
}


function calcularVencimentoPorRegra(dataEmissao, tipoDocumento, uf, regras) {
  const emissao = converterTextoParaData(dataEmissao);

  if (!emissao) {
    return {
      data: "",
      calculado: "NAO",
      fonte: "SEM_DATA_EMISSAO",
      observacao: "Não foi possível calcular vencimento porque a data de emissão não foi identificada."
    };
  }

  const regra = encontrarRegra(regras, tipoDocumento, uf);

  if (!regra) {
    return {
      data: "",
      calculado: "NAO",
      fonte: "SEM_REGRA",
      observacao: "Nenhuma regra de vencimento encontrada."
    };
  }

  const criterio = regra.criterio;
  const quantidade = Number(regra.quantidade || 0);

  let vencimento = new Date(emissao.getTime());

  if (criterio === "DIAS_APOS_EMISSAO") {
    vencimento.setDate(vencimento.getDate() + quantidade);
  } else if (criterio === "MESES_APOS_EMISSAO") {
    vencimento.setMonth(vencimento.getMonth() + quantidade);
  } else if (criterio === "ANOS_APOS_EMISSAO") {
    vencimento.setFullYear(vencimento.getFullYear() + quantidade);
  } else if (criterio === "FIM_ANO_EMISSAO") {
    vencimento = new Date(emissao.getFullYear(), 11, 31);
  } else if (criterio === "FIM_ANO_SEGUINTE") {
    vencimento = new Date(emissao.getFullYear() + 1, 11, 31);
  } else if (criterio === "NAO_APLICAVEL") {
  return {
    data: "",
    calculado: "NAO",
    fonte: "DOCUMENTO_COMPLEMENTAR",
    observacao: regra.observacao
  };
  } else if (criterio === "USAR_VALIDADE_DOCUMENTO") {
    return {
      data: "",
      calculado: "NAO",
      fonte: "REGRA_EXIGE_VALIDADE_NO_DOCUMENTO",
      observacao: regra.observacao
    };
  } else {
    return {
      data: "",
      calculado: "NAO",
      fonte: "CRITERIO_DESCONHECIDO",
      observacao: `Critério não reconhecido: ${criterio}`
    };
  }

  return {
    data: formatarData(vencimento),
    calculado: "SIM",
    fonte: `REGRA_${criterio}`,
    observacao: regra.observacao
  };
}


/*******************************
 * NORMALIZAÇÃO DA RESPOSTA DA IA
 *******************************/

function normalizarResultadoIA(analiseIA, arquivo, caminhoPasta, regras, calendarios) {
  const statusProcessamento = String(analiseIA.status_processamento || "OK").toUpperCase().trim();
  const tipoDocumento = String(analiseIA.tipo_documento || "OUTRO").toUpperCase().trim();
  const placa = String(analiseIA.placa || extrairPlacaDoNome(arquivo.getName()) || "").toUpperCase().trim();
  const uf = String(analiseIA.uf || "PADRAO").toUpperCase().trim();
  const dataEmissao = normalizarDataTexto(analiseIA.data_emissao || "");

  if (tipoDocumento === "RNC") {
    return classificarRNC(statusProcessamento, tipoDocumento, placa, uf, dataEmissao, analiseIA.observacao);
  }

  if (tipoDocumento === "CRLV") {
    const exercicio = extrairExercicioCRLV(
      analiseIA.exercicio,
      analiseIA.observacao,
      arquivo.getName()
    );

    return classificarCRLV(
      statusProcessamento,
      placa,
      uf,
      dataEmissao,
      exercicio,
      analiseIA.observacao,
      calendarios
    );
  }

  let dataVencimento = normalizarDataTexto(analiseIA.data_vencimento || "");
  let vencimentoCalculado = "NAO";
  let fonteVencimento = dataVencimento ? "DOCUMENTO" : "";
  let observacao = analiseIA.observacao || "";

  if (!dataVencimento && dataEmissao) {
    const calculo = calcularVencimentoPorRegra(dataEmissao, tipoDocumento, uf, regras);
    dataVencimento = calculo.data || "";
    vencimentoCalculado = calculo.calculado;
    fonteVencimento = calculo.fonte;
    observacao = combinarObservacoes(observacao, calculo.observacao);
  }

  if (!dataVencimento && !dataEmissao) {
    fonteVencimento = "SEM_DATA_EMISSAO";
  }

  const diasParaVencer = calcularDiasParaVencer(dataVencimento);
  const situacao = definirSituacao(statusProcessamento, diasParaVencer, dataVencimento);

  return {
    status_processamento: statusProcessamento,
    tipo_documento: tipoDocumento,
    placa: placa,
    uf: uf,
    data_emissao: dataEmissao,
    data_vencimento: dataVencimento,
    vencimento_calculado: vencimentoCalculado,
    fonte_vencimento: fonteVencimento,
    dias_para_vencer: diasParaVencer,
    situacao: situacao,
    observacao: observacao
  };
}


function classificarRNC(statusProcessamento, tipoDocumento, placa, uf, dataEmissao, observacaoIA) {
  return {
    status_processamento: statusProcessamento,
    tipo_documento: tipoDocumento,
    placa: placa,
    uf: uf,
    data_emissao: dataEmissao,
    data_vencimento: "",
    vencimento_calculado: "NAO",
    fonte_vencimento: "DOCUMENTO_COMPLEMENTAR",
    dias_para_vencer: "",
    situacao: statusProcessamento === "ERRO" ? "ERRO" : "NAO_APLICAVEL",
    observacao: combinarObservacoes(
      observacaoIA,
      "Registro de Não-Conformidade vinculado à inspeção. Não possui vencimento próprio."
    )
  };
}


function classificarCRLV(statusProcessamento, placa, uf, dataEmissao, exercicio, observacaoIA, calendarios) {
  const base = {
    status_processamento: statusProcessamento,
    tipo_documento: "CRLV",
    placa: placa,
    uf: uf,
    data_emissao: dataEmissao,
    data_vencimento: "",
    vencimento_calculado: "NAO",
    fonte_vencimento: "EXERCICIO_CRLV",
    dias_para_vencer: "",
    situacao: "",
    observacao: observacaoIA || "",
    exercicio_crlv: exercicio || "",
    ano_calendario_crlv: exercicio ? Number(exercicio) + 1 : ""
  };

  if (statusProcessamento === "ERRO") {
    base.situacao = "ERRO";
    base.observacao = observacaoIA || "Falha na leitura do CRLV.";
    return base;
  }

  if (!exercicio) {
    base.situacao = "SEM_EXERCICIO";
    base.observacao = combinarObservacoes(observacaoIA, "Não foi possível identificar o exercício do CRLV.");
    return base;
  }

  const anoAtual = new Date().getFullYear();
  const anoCalendario = Number(exercicio) + 1;
  const finalPlaca = extrairFinalNumericoPlaca(placa);
  const calendario = consultarCalendarioLicenciamento(calendarios || {}, uf, anoCalendario, placa);

  if (calendario) {
    const dias = calcularDiasParaVencer(calendario.data_limite);
    base.data_vencimento = calendario.data_limite;
    base.vencimento_calculado = "SIM";
    base.fonte_vencimento = `CALENDARIO_${normalizarIdentificador(calendario.fonte)}_${anoCalendario}`;
    base.dias_para_vencer = dias;

    if (dias < 0) {
      const anosAtraso = calcularAnosCompletosAtraso(calendario.data_limite);
      base.situacao = anosAtraso >= 1
        ? "VENCIDO_HA_MAIS_DE_UM_ANO"
        : "LICENCIAMENTO_ATRASADO";
    } else if (dias <= DIAS_ALERTA) {
      base.situacao = "LICENCIAMENTO_PROXIMO";
    } else {
      base.situacao = "REGULAR_ATE_CALENDARIO";
    }

    base.observacao = combinarObservacoes(
      observacaoIA,
      `CRLV do exercício ${exercicio}. Data-limite calculada pelo calendário ${anoCalendario} da UF ${uf}, placa final ${finalPlaca}: ${calendario.data_limite}. ${calendario.observacao || ""}`
    );
    return base;
  }

  if (Number(exercicio) <= anoAtual - 2) {
    base.situacao = "VENCIDO_HA_MAIS_DE_UM_ANO";
    base.fonte_vencimento = "EXERCICIO_CRLV_DESATUALIZADO";
    base.observacao = combinarObservacoes(
      observacaoIA,
      `CRLV do exercício ${exercicio}. O documento está defasado em dois ou mais exercícios e foi marcado como vencido há mais de um ano, mesmo sem data estadual disponível.`
    );
    return base;
  }

  if (Number(exercicio) >= anoAtual) {
    base.situacao = "REGULAR_POR_EXERCICIO";
    base.observacao = combinarObservacoes(
      observacaoIA,
      `CRLV do exercício ${exercicio}. O calendário ${anoCalendario} ainda não está cadastrado; o documento permanece regular pelo exercício atual/futuro.`
    );
    return base;
  }

  if (!uf || uf === "PADRAO") {
    base.situacao = "SEM_UF_CRLV";
    base.observacao = combinarObservacoes(observacaoIA, `CRLV do exercício ${exercicio}, mas a UF de registro não foi identificada.`);
    return base;
  }

  if (finalPlaca === "") {
    base.situacao = "SEM_FINAL_PLACA_CRLV";
    base.observacao = combinarObservacoes(observacaoIA, `CRLV do exercício ${exercicio}, mas não foi possível obter o final numérico da placa.`);
    return base;
  }

  base.situacao = "CALENDARIO_NAO_CADASTRADO";
  base.fonte_vencimento = "CALENDARIO_CRLV_NAO_CADASTRADO";
  base.observacao = combinarObservacoes(
    observacaoIA,
    `CRLV do exercício ${exercicio}. Cadastre na aba ${ABA_CALENDARIO} o calendário ${anoCalendario} da UF ${uf} para a placa final ${finalPlaca}.`
  );
  return base;
}


function calcularAnosCompletosAtraso(dataLimiteTexto) {
  const dataLimite = converterTextoParaData(dataLimiteTexto);
  if (!dataLimite) return 0;

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  dataLimite.setHours(0, 0, 0, 0);

  let anos = hoje.getFullYear() - dataLimite.getFullYear();
  const aniversario = new Date(hoje.getFullYear(), dataLimite.getMonth(), dataLimite.getDate());
  if (hoje < aniversario) anos--;
  return Math.max(0, anos);
}


function normalizarIdentificador(valor) {
  return String(valor || "CALENDARIO")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9]+/gi, "_")
    .replace(/^_+|_+$/g, "")
    .toUpperCase();
}


function extrairExercicioCRLV(valorIA, observacaoIA, nomeArquivo) {
  const candidatos = [valorIA, observacaoIA, nomeArquivo];
  const anoAtual = new Date().getFullYear();

  for (let i = 0; i < candidatos.length; i++) {
    const texto = String(candidatos[i] || "");
    const padroes = [
      /EXERC[ÍI]CIO\s*[:\-]?\s*(20\d{2})/i,
      /ANO\s+DE\s+EXERC[ÍI]CIO\s*[:\-]?\s*(20\d{2})/i,
      /LICENCIAMENTO\s*[:\-]?\s*(20\d{2})/i,
      /CRLV[^0-9]{0,20}(20\d{2})/i,
      /\b(20\d{2})\b/
    ];

    for (let p = 0; p < padroes.length; p++) {
      const match = texto.match(padroes[p]);
      if (!match) continue;

      const ano = Number(match[1]);
      if (ano >= 2000 && ano <= anoAtual + 2) {
        return ano;
      }
    }
  }

  return "";
}


function combinarObservacoes(principal, adicional) {
  const partes = [principal, adicional]
    .map(v => String(v || "").trim())
    .filter(Boolean);

  return [...new Set(partes)].join(" ");
}


function definirSituacao(statusProcessamento, diasParaVencer, dataVencimento) {
  if (statusProcessamento === "ERRO") {
    return "ERRO";
  }

  if (!dataVencimento) {
    return "SEM_VENCIMENTO";
  }

  if (diasParaVencer === "") {
    return "DATA_INVALIDA";
  }

  if (diasParaVencer < 0) {
    return "VENCIDO";
  }

  if (diasParaVencer <= DIAS_ALERTA) {
    return "URGENTE";
  }

  return "REGULAR";
}


/*******************************
 * GEMINI
 *******************************/

function chamarGeminiParaAnalise(arquivo) {
  try {
    if (!GEMINI_API_KEY) {
      throw new Error('Chave Gemini não configurada. Execute configurarChaveGemini() ou defina GEMINI_API_KEY nas propriedades do script.');
    }
    const blob = arquivo.getBlob();
    const base64Data = Utilities.base64Encode(blob.getBytes());
    const mimeType = arquivo.getMimeType();

    const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${GEMINI_API_KEY}`;

    const prompt = `
Você é um analista de documentos de veículos.

Analise o arquivo enviado e extraia as informações abaixo.

Responda estritamente em JSON válido, sem markdown, sem texto adicional e sem explicações.

Formato obrigatório:
{
  "status_processamento": "OK",
  "tipo_documento": "CRLV",
  "placa": "ABC1D23",
  "uf": "RO",
  "data_emissao": "DD/MM/AAAA",
  "data_vencimento": "DD/MM/AAAA",
  "exercicio": "AAAA",
  "observacao": "observação curta"
}

Campos:
- status_processamento: use "OK" se conseguir analisar; use "ERRO" se o documento estiver ilegível ou não for possível interpretar.
- tipo_documento: escolha apenas uma opção: "CRLV", "AET", "CIV", "CIPP", "RNC", "TACOGRAFO", "SEGURO", "ANTT", "IBAMA", "OUTRO".
- placa: informe a placa se encontrada. Se não encontrar, deixe vazio.
- uf: informe a UF do documento se encontrada, exemplo "RO", "MT", "SP", "GO", "AM", "RR". Se não encontrar, use "PADRAO".
- data_emissao: informe no formato DD/MM/AAAA. Se não encontrar, deixe vazio.
- data_vencimento: informe no formato DD/MM/AAAA. Se não encontrar, deixe vazio.
- exercicio: preencha somente para CRLV, usando o ano do campo EXERCÍCIO/licenciamento. Se não encontrar, deixe vazio.
- observacao: observação curta e objetiva.

Regras importantes:
- Se existir data de validade, vencimento, término, fim de vigência, validade até ou data limite, preencha data_vencimento.
- Se não existir vencimento claro, mas existir emissão, preencha data_emissao e deixe data_vencimento vazio.
- Não invente datas.
- Não use a data atual como vencimento.
- Para CRLV, procure especificamente o campo EXERCÍCIO, ANO DE EXERCÍCIO, LICENCIAMENTO ou expressão equivalente.
- Para CRLV, nunca invente data de vencimento e nunca transforme a data de emissão em validade.
- Para CRLV, preencha o campo exercicio com quatro dígitos quando estiver claramente identificado.
- Para documentos RNC, identifique o tipo como RNC.
- RNC é documento complementar vinculado à inspeção e não possui vencimento próprio para esta automação.
- Para RNC, deixe data_vencimento vazia.
- Se houver mais de uma data, diferencie emissão de vencimento.
- Se houver vigência de seguro, a data final da vigência é o vencimento.
- Se o documento trouxer apenas exercício, mas não trouxer uma data clara, deixe data_vencimento vazio e explique na observação.
`;

    const payload = {
      contents: [{
        parts: [
          { text: prompt },
          {
            inlineData: {
              mimeType: mimeType,
              data: base64Data
            }
          }
        ]
      }],
      generationConfig: {
        responseMimeType: "application/json",
        temperature: 0
      }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();

    const json = JSON.parse(responseText);

    if (!json.candidates || json.candidates.length === 0) {
      return {
        status_processamento: "ERRO",
        tipo_documento: "OUTRO",
        placa: "",
        uf: "PADRAO",
        data_emissao: "",
        data_vencimento: "",
        observacao: "Sem resposta válida da API Gemini."
      };
    }

    const respostaTexto = json.candidates[0].content.parts[0].text;
    return JSON.parse(respostaTexto);

  } catch (erro) {
    return {
      status_processamento: "ERRO",
      tipo_documento: "OUTRO",
      placa: "",
      uf: "PADRAO",
      data_emissao: "",
      data_vencimento: "",
      observacao: `Falha na análise: ${erro.message}`
    };
  }
}


/*******************************
 * RECLASSIFICAÇÃO DE REGISTROS ANTIGOS
 *******************************/

function reclassificarRegistrosAntigosCRLVeRNC(contexto) {
  const aba = contexto.abaDb;
  const dados = aba.getDataRange().getValues();

  if (dados.length <= 1) {
    return 0;
  }

  const cabecalho = dados[0];
  const idxNome = cabecalho.indexOf("nome_arquivo");
  const idxTipo = cabecalho.indexOf("tipo_documento");
  const idxStatus = cabecalho.indexOf("status_processamento");
  const idxPlaca = cabecalho.indexOf("placa");
  const idxUf = cabecalho.indexOf("uf");
  const idxEmissao = cabecalho.indexOf("data_emissao");
  const idxVencimento = cabecalho.indexOf("data_vencimento");
  const idxCalculado = cabecalho.indexOf("vencimento_calculado");
  const idxFonte = cabecalho.indexOf("fonte_vencimento");
  const idxDias = cabecalho.indexOf("dias_para_vencer");
  const idxSituacao = cabecalho.indexOf("situacao");
  const idxObservacao = cabecalho.indexOf("observacao");
  const idxJson = cabecalho.indexOf("json_ia");
  const idxExercicio = cabecalho.indexOf("exercicio_crlv");
  const idxAnoCalendario = cabecalho.indexOf("ano_calendario_crlv");
  const calendarios = carregarCalendariosLicenciamento(contexto);

  let alterados = 0;

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const tipo = String(linha[idxTipo] || "").toUpperCase().trim();

    if (tipo !== "CRLV" && tipo !== "RNC") {
      continue;
    }

    const status = String(linha[idxStatus] || "OK").toUpperCase().trim();
    const placa = String(linha[idxPlaca] || "");
    const uf = String(linha[idxUf] || "PADRAO");
    const emissao = normalizarDataTexto(linha[idxEmissao] || "");
    const observacao = String(linha[idxObservacao] || "");
    let novo;

    if (tipo === "RNC") {
      novo = classificarRNC(status, tipo, placa, uf, emissao, observacao);
    } else {
      let jsonIA = {};
      try {
        jsonIA = JSON.parse(String(linha[idxJson] || "{}"));
      } catch (e) {
        jsonIA = {};
      }

      const exercicio = extrairExercicioCRLV(
        jsonIA.exercicio,
        combinarObservacoes(jsonIA.observacao, observacao),
        linha[idxNome]
      );

      novo = classificarCRLV(status, placa, uf, emissao, exercicio, observacao, calendarios);
    }

    const mudou =
      String(linha[idxVencimento] || "") !== String(novo.data_vencimento || "") ||
      String(linha[idxCalculado] || "") !== String(novo.vencimento_calculado || "") ||
      String(linha[idxFonte] || "") !== String(novo.fonte_vencimento || "") ||
      String(linha[idxDias] || "") !== String(novo.dias_para_vencer || "") ||
      String(linha[idxSituacao] || "") !== String(novo.situacao || "") ||
      String(linha[idxObservacao] || "") !== String(novo.observacao || "") ||
      (idxExercicio >= 0 && String(linha[idxExercicio] || "") !== String(novo.exercicio_crlv || "")) ||
      (idxAnoCalendario >= 0 && String(linha[idxAnoCalendario] || "") !== String(novo.ano_calendario_crlv || ""));

    if (!mudou) {
      continue;
    }

    linha[idxVencimento] = novo.data_vencimento;
    linha[idxCalculado] = novo.vencimento_calculado;
    linha[idxFonte] = novo.fonte_vencimento;
    linha[idxDias] = novo.dias_para_vencer;
    linha[idxSituacao] = novo.situacao;
    linha[idxObservacao] = novo.observacao;
    if (idxExercicio >= 0) linha[idxExercicio] = novo.exercicio_crlv || "";
    if (idxAnoCalendario >= 0) linha[idxAnoCalendario] = novo.ano_calendario_crlv || "";
    alterados++;
  }

  if (alterados > 0) {
    aba.getRange(2, 1, dados.length - 1, cabecalho.length).setValues(dados.slice(1));
  }

  return alterados;
}


/*******************************
 * ALERTAS
 *******************************/

function gerarAlertasAPartirDoBanco(contexto) {
  const aba = contexto.abaDb;
  const dados = aba.getDataRange().getValues();

  if (dados.length <= 1) {
    return [];
  }

  const cabecalho = dados[0];

  const idxNome = cabecalho.indexOf("nome_arquivo");
  const idxPasta = cabecalho.indexOf("caminho_pasta");
  const idxUrl = cabecalho.indexOf("url");
  const idxTipo = cabecalho.indexOf("tipo_documento");
  const idxPlaca = cabecalho.indexOf("placa");
  const idxUf = cabecalho.indexOf("uf");
  const idxEmissao = cabecalho.indexOf("data_emissao");
  const idxVencimento = cabecalho.indexOf("data_vencimento");
  const idxDias = cabecalho.indexOf("dias_para_vencer");
  const idxSituacao = cabecalho.indexOf("situacao");
  const idxFonte = cabecalho.indexOf("fonte_vencimento");

  const alertas = [];

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const situacao = linha[idxSituacao];

    if (situacao === "VENCIDO" || situacao === "URGENTE" || situacao === "LICENCIAMENTO_ATRASADO" || situacao === "VENCIDO_HA_MAIS_DE_UM_ANO" || situacao === "LICENCIAMENTO_PROXIMO") {
      alertas.push({
        nome: linha[idxNome],
        pasta: linha[idxPasta],
        url: linha[idxUrl],
        tipo: linha[idxTipo],
        placa: linha[idxPlaca],
        uf: linha[idxUf],
        emissao: linha[idxEmissao],
        vencimento: linha[idxVencimento],
        dias: linha[idxDias],
        situacao: situacao,
        fonte: linha[idxFonte]
      });
    }
  }

  return alertas;
}


function enviarEmailAlertas(alertas, totalPendencias) {
  let html = `
    <h2>Alerta de documentos de veículos</h2>
    <p>Foram encontrados documentos vencidos, próximos ao vencimento ou CRLVs próximos/atrasados no calendário estadual.</p>
  `;

  if (totalPendencias > 0) {
    html += `
      <p><strong>Atenção:</strong> também existem ${totalPendencias} documento(s) pendente(s) de revisão humana na aba ${ABA_PENDENCIAS}.</p>
    `;
  }

  html += `
    <table border="1" cellpadding="6" cellspacing="0">
      <tr>
        <th>Situação</th>
        <th>Tipo</th>
        <th>Placa</th>
        <th>UF</th>
        <th>Arquivo</th>
        <th>Pasta</th>
        <th>Emissão</th>
        <th>Vencimento</th>
        <th>Dias</th>
        <th>Fonte</th>
      </tr>
  `;

  alertas.forEach(a => {
    html += `
      <tr>
        <td>${a.situacao}</td>
        <td>${a.tipo || ""}</td>
        <td>${a.placa || ""}</td>
        <td>${a.uf || ""}</td>
        <td><a href="${a.url}">${a.nome}</a></td>
        <td>${a.pasta}</td>
        <td>${a.emissao || ""}</td>
        <td>${a.vencimento || ""}</td>
        <td>${a.dias}</td>
        <td>${a.fonte || ""}</td>
      </tr>
    `;
  });

  html += `</table>`;

  MailApp.sendEmail({
    to: EMAIL_NOTIFICACAO,
    subject: "⚠️ ALERTA: Documentos de Veículos com Alerta",
    htmlBody: html
  });
}


function enviarEmailSomentePendencias(totalPendencias) {
  MailApp.sendEmail({
    to: EMAIL_NOTIFICACAO,
    subject: "📌 Pendências de Revisão em Documentos de Veículos",
    htmlBody: `
      <h2>Pendências de revisão</h2>
      <p>Não foram encontrados documentos vencidos, próximos ao vencimento ou CRLVs próximos/atrasados no calendário estadual.</p>
      <p>Porém existem <strong>${totalPendencias}</strong> documento(s) pendente(s) de revisão humana.</p>
      <p>Consulte a aba <strong>${ABA_PENDENCIAS}</strong> na planilha <strong>${NOME_PLANILHA_CONTROLE}</strong>.</p>
    `
  });
}


/*******************************
 * PENDÊNCIAS DE REVISÃO
 *******************************/

function atualizarPendenciasRevisao(contexto) {
  const abaDb = contexto.abaDb;
  const abaPendencias = contexto.abaPendencias;

  limparAbaMantendoCabecalho(abaPendencias);

  const dados = abaDb.getDataRange().getValues();

  if (dados.length <= 1) {
    return 0;
  }

  const cabecalho = dados[0];

  const idxFileId = cabecalho.indexOf("file_id");
  const idxNome = cabecalho.indexOf("nome_arquivo");
  const idxPasta = cabecalho.indexOf("caminho_pasta");
  const idxUrl = cabecalho.indexOf("url");
  const idxStatus = cabecalho.indexOf("status_processamento");
  const idxTipo = cabecalho.indexOf("tipo_documento");
  const idxPlaca = cabecalho.indexOf("placa");
  const idxUf = cabecalho.indexOf("uf");
  const idxEmissao = cabecalho.indexOf("data_emissao");
  const idxVencimento = cabecalho.indexOf("data_vencimento");
  const idxFonte = cabecalho.indexOf("fonte_vencimento");
  const idxSituacao = cabecalho.indexOf("situacao");
  const idxObservacao = cabecalho.indexOf("observacao");

  const pendencias = [];

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];

    const statusProcessamento = String(linha[idxStatus] || "").toUpperCase();
    const situacao = String(linha[idxSituacao] || "").toUpperCase();
    const fonte = String(linha[idxFonte] || "").toUpperCase();
    const emissao = linha[idxEmissao];
    const vencimento = linha[idxVencimento];

    const motivo = identificarMotivoPendencia(
      String(linha[idxTipo] || "").toUpperCase(),
      statusProcessamento,
      situacao,
      fonte,
      emissao,
      vencimento
    );

    if (!motivo) {
      continue;
    }

    const acaoRecomendada = recomendarAcaoPendencia(motivo);

    pendencias.push([
      formatarDataHora(new Date()),
      motivo,
      linha[idxFileId],
      linha[idxNome],
      linha[idxPasta],
      linha[idxUrl],
      linha[idxTipo],
      linha[idxPlaca],
      linha[idxUf],
      linha[idxEmissao],
      linha[idxVencimento],
      linha[idxFonte],
      linha[idxSituacao],
      linha[idxObservacao],
      acaoRecomendada
    ]);
  }

  if (pendencias.length > 0) {
    abaPendencias
      .getRange(2, 1, pendencias.length, pendencias[0].length)
      .setValues(pendencias);
  }

  return pendencias.length;
}


function limparAbaMantendoCabecalho(aba) {
  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha <= 1) {
    return;
  }

  aba
    .getRange(2, 1, ultimaLinha - 1, aba.getLastColumn())
    .clearContent();
}


function identificarMotivoPendencia(
  tipoDocumento,
  statusProcessamento,
  situacao,
  fonte,
  emissao,
  vencimento
) {
  if (
    situacao === "NAO_APLICAVEL" ||
    fonte === "DOCUMENTO_COMPLEMENTAR"
  ) {
    return "";
  }

  if (tipoDocumento === "CRLV") {
    if (situacao === "SEM_EXERCICIO") return "SEM_EXERCICIO_CRLV";
    if (situacao === "SEM_UF_CRLV") return "SEM_UF_CRLV";
    if (situacao === "SEM_FINAL_PLACA_CRLV") return "SEM_FINAL_PLACA_CRLV";
    if (situacao === "CALENDARIO_NAO_CADASTRADO") return "CALENDARIO_CRLV_NAO_CADASTRADO";
    if (situacao === "REGULAR_POR_EXERCICIO" || situacao === "REGULAR_ATE_CALENDARIO" || situacao === "LICENCIAMENTO_PROXIMO" || situacao === "LICENCIAMENTO_ATRASADO") return "";
  }
  if (statusProcessamento === "ERRO") {
    return "ERRO_PROCESSAMENTO";
  }

  if (situacao === "ERRO") {
    return "ERRO_PROCESSAMENTO";
  }

  if (situacao === "SEM_VENCIMENTO") {
    return "SEM_VENCIMENTO";
  }

  if (situacao === "DATA_INVALIDA") {
    return "DATA_INVALIDA";
  }

  if (!emissao && !vencimento) {
    return "SEM_DATA_EMISSAO_E_SEM_VENCIMENTO";
  }

  if (fonte === "SEM_DATA_EMISSAO") {
    return "SEM_DATA_EMISSAO";
  }

  if (fonte === "SEM_REGRA") {
    return "SEM_REGRA";
  }

  if (fonte === "REGRA_EXIGE_VALIDADE_NO_DOCUMENTO" && !vencimento) {
    return "REGRA_EXIGE_VALIDADE_NO_DOCUMENTO";
  }

  if (fonte === "CRITERIO_DESCONHECIDO") {
    return "CRITERIO_DESCONHECIDO";
  }

  return "";
}


function recomendarAcaoPendencia(motivo) {
  const mapa = {
    "ERRO_PROCESSAMENTO": "Abrir o arquivo manualmente e verificar se está legível. Se estiver correto, reprocessar ou ajustar o prompt.",
    "SEM_VENCIMENTO": "Verificar manualmente se o documento possui vencimento. Se não possuir, cadastrar regra de cálculo por tipo/UF.",
    "DATA_INVALIDA": "Conferir a data extraída pela IA e corrigir manualmente no banco de controle.",
    "SEM_DATA_EMISSAO": "Verificar se o documento possui data de emissão. Se possuir, corrigir manualmente ou melhorar o prompt.",
    "SEM_REGRA": "Cadastrar regra de vencimento na aba REGRAS_VENCIMENTO para esse tipo de documento e UF.",
    "SEM_DATA_EMISSAO_E_SEM_VENCIMENTO": "Conferir documento manualmente. Pode ser arquivo errado, imagem ilegível ou documento sem datas úteis.",
    "REGRA_EXIGE_VALIDADE_NO_DOCUMENTO": "Conferir se existe validade explícita no documento. Se não existir, definir regra de cálculo.",
    "CRITERIO_DESCONHECIDO": "Corrigir o critério informado na aba REGRAS_VENCIMENTO.",
    "SEM_EXERCICIO_CRLV": "Abrir o CRLV e confirmar o campo EXERCÍCIO. Ajustar o nome do arquivo ou reprocessar após melhorar a leitura.",
    "SEM_UF_CRLV": "Confirmar a UF de registro do veículo no CRLV e corrigir o registro ou reprocessar o documento.",
    "SEM_FINAL_PLACA_CRLV": "Confirmar a placa do veículo; o calendário depende do último dígito numérico.",
    "CALENDARIO_CRLV_NAO_CADASTRADO": `Cadastrar na aba ${ABA_CALENDARIO} a data oficial para a UF, ano do calendário e final da placa.`
  };

  return mapa[motivo] || "Revisar manualmente.";
}


/*******************************
 * FUNÇÕES DE DATA
 *******************************/

function normalizarDataTexto(valor) {
  if (!valor) {
    return "";
  }

  if (Object.prototype.toString.call(valor) === "[object Date]") {
    return formatarData(valor);
  }

  const texto = String(valor).trim();

  const match = texto.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})/);

  if (!match) {
    return "";
  }

  let dia = match[1].padStart(2, "0");
  let mes = match[2].padStart(2, "0");
  let ano = match[3];

  if (ano.length === 2) {
    ano = "20" + ano;
  }

  return `${dia}/${mes}/${ano}`;
}


function converterTextoParaData(texto) {
  const normalizada = normalizarDataTexto(texto);

  if (!normalizada) {
    return null;
  }

  const partes = normalizada.split("/");
  const dia = Number(partes[0]);
  const mes = Number(partes[1]) - 1;
  const ano = Number(partes[2]);

  const data = new Date(ano, mes, dia);

  if (
    data.getFullYear() !== ano ||
    data.getMonth() !== mes ||
    data.getDate() !== dia
  ) {
    return null;
  }

  return data;
}


function calcularDiasParaVencer(dataVencimentoTexto) {
  const vencimento = converterTextoParaData(dataVencimentoTexto);

  if (!vencimento) {
    return "";
  }

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  vencimento.setHours(0, 0, 0, 0);

  const diffMs = vencimento.getTime() - hoje.getTime();
  return Math.ceil(diffMs / (1000 * 60 * 60 * 24));
}


function formatarData(data) {
  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy"
  );
}


function formatarDataHora(data) {
  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm:ss"
  );
}


/*******************************
 * UTILITÁRIOS
 *******************************/

function extrairPlacaDoNome(nomeArquivo) {
  const texto = String(nomeArquivo).toUpperCase();

  const matchMercosul = texto.match(/[A-Z]{3}[0-9][A-Z][0-9]{2}/);
  if (matchMercosul) {
    return matchMercosul[0];
  }

  const matchAntiga = texto.match(/[A-Z]{3}[0-9]{4}/);
  if (matchAntiga) {
    return matchAntiga[0];
  }

  return "";
}


/*******************************
 * FUNÇÕES AUXILIARES MANUAIS
 *******************************/

// Use esta função quando quiser forçar reprocessamento de tudo.
// Ela apaga somente a aba ARQUIVOS_PROCESSADOS.
// Não apaga logs, regras ou pendências.
function limparBancoDeArquivosProcessados() {
  const contexto = inicializarAmbiente();
  limparAbaMantendoCabecalho(contexto.abaDb);

  registrarLog(
    contexto,
    "WARN",
    "MANUTENCAO",
    "🧹 Banco de arquivos processados foi limpo manualmente."
  );
}


// Use esta função quando quiser limpar apenas logs antigos.
function limparLogExecucao() {
  const contexto = inicializarAmbiente();
  limparAbaMantendoCabecalho(contexto.abaLog);

  registrarLog(
    contexto,
    "WARN",
    "MANUTENCAO",
    "🧹 Log de execução foi limpo manualmente."
  );
}


// Use esta função apenas para testar o escopo das pastas sem chamar IA.
// Ela registra quais pastas seriam processadas, ignoradas ou puladas.
function testarEscopoDasPastas() {
  const contexto = inicializarAmbiente();
  const pastaMae = DriveApp.getFolderById(PASTA_MAE_ID);
  const subpastas = pastaMae.getFolders();

  registrarLog(contexto, "INFO", "TESTE_ESCOPO", "🧪 Iniciando teste de escopo de pastas.");

  while (subpastas.hasNext()) {
    const subpasta = subpastas.next();
    const nome = subpasta.getName();

    if (ehPastaIgnorada(nome)) {
      registrarLog(contexto, "INFO", "TESTE_ESCOPO", `⛔ Seria ignorada: ${nome}`);
    } else if (ehPastaRaizPermitida(nome)) {
      registrarLog(contexto, "INFO", "TESTE_ESCOPO", `✅ Seria processada: ${nome}`);
    } else {
      registrarLog(contexto, "INFO", "TESTE_ESCOPO", `⏭️ Fora do escopo: ${nome}`);
    }
  }

  registrarLog(contexto, "INFO", "TESTE_ESCOPO", "🏁 Teste de escopo concluído.");
}

// Execute uma vez e informe sua chave para salvá-la com segurança nas propriedades do script.
function configurarChaveGemini() {
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.prompt(
    "Configurar chave Gemini",
    "Cole a chave da API Gemini:",
    ui.ButtonSet.OK_CANCEL
  );

  if (resposta.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const chave = resposta.getResponseText().trim();

  if (!chave) {
    throw new Error("A chave não pode ficar vazia.");
  }

  PropertiesService.getScriptProperties().setProperty("GEMINI_API_KEY", chave);
  ui.alert("Chave Gemini salva nas propriedades do script.");
}


function removerChaveGemini() {
  PropertiesService.getScriptProperties().deleteProperty("GEMINI_API_KEY");
}


// Recalcula todos os CRLVs existentes usando os calendários cadastrados,
// sem chamar a API Gemini e sem consumir tokens.
function recalcularTodosOsCRLVsPeloCalendario() {
  const contexto = inicializarAmbiente();
  const total = reclassificarRegistrosAntigosCRLVeRNC(contexto);
  const pendencias = atualizarPendenciasRevisao(contexto);

  registrarLog(
    contexto,
    "INFO",
    "CALENDARIO_CRLV",
    `📅 Recálculo concluído. Registros alterados: ${total}. Pendências atuais: ${pendencias}.`
  );
}
