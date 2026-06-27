// ======================================================
// CONFIG
// ======================================================

// ---------- Nomes das abas ----------
const ABA_PL = "Cód de Barras (PL Financ)";
const ABA_G = "Guepardo";
const ABA_OUT = "Base Conciliacao ISS";
const ABA_LN = "Base Analise LN";
const ABA_DIAG = "Diagnostico ISS";
const ABA_DIM_MUNICIPIO = "DIM_MUNICIPIO";

// =====================================================
// LEITURA DINÂMICA DE CABEÇALHOS
// =====================================================

function criarSchema(cabecalho) {

  const schema = {};

  cabecalho.forEach((nome, index) => {

    const chave = String(nome || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g,"")
      .toUpperCase()
      .trim()
      .replace(/\s+/g,"_");

    schema[chave] = index;

  });


  return schema;
}



function localizarColuna(schema, nomesPossiveis) {


  for (let nome of nomesPossiveis) {

    const chave = nome
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g,"")
      .toUpperCase()
      .trim()
      .replace(/\s+/g,"_");


    if (schema[chave] !== undefined) {

      return schema[chave];

    }

  }


  return null;

}

// CACHE MUNICÍPIOS NORMALIZADOS
let MAPA_MUNICIPIOS = {};

// ---------- Controle ----------
const ATIVAR_GUEPARDO = true;

// ---------- Colunas que serão localizadas dinamicamente ----------
const HEADERS = {

  // PL Financeira
  PL: {
    TIPO: null,
    EMPRESA: null,
    MUNICIPIO: null,
    SISTEMA: null,
    LN: null,
    COMPETENCIA: null,
    DOCUMENTO_SAP: null,
    CHAMADO_FINANCEIRO: null,
    ISS: null,
    MULTA: null,
    JUROS: null
  },

  // Guepardo
  G: {
    EMPRESA: null,
    MUNICIPIO: null,
    SISTEMA: null,
    LN: null,
    COMPETENCIA: null,
    DOCUMENTO: null,
    VALOR: null,
    TIPO_GUIA: null
  }

};

// ---------- Objetos Globais ----------
const docGlobalMap = {};

// ======================================================
// CONTROLE DE ABAS
// ======================================================

function mostrarDimMunicipio(ss) {

  const aba = ss.getSheetByName(ABA_DIM_MUNICIPIO);

  if (aba && aba.isSheetHidden()) {
    aba.showSheet();
  }

}

function ocultarDimMunicipio(ss) {

  const aba = ss.getSheetByName(ABA_DIM_MUNICIPIO);

  if (aba) {
    aba.hideSheet();
  }

}


// ======================================================
// NORMALIZAÇÃO
// ======================================================

function getValorPL(r, PL) {

  if (!r || !PL) return 0;


  const iss = parseMoney(
    r[PL.VALOR_PRINCIPAL] || 0
  );


  const multa = parseMoney(
    r[PL.MULTA] || 0
  );


  const juros = parseMoney(
    r[PL.JUROS] || 0
  );


  return iss + multa + juros;

}

function normDoc(doc) {
  return String(doc || "")
    .replace(/\D/g, "")
    .slice(-10);
}

function norm(x) {
  return String(x || "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toUpperCase().trim();
}

function safe(x) {
  const v = norm(x);
  return v ? v : "NA";
}

function normLN(ln) {
  if (!ln) return "0000";
  let cleanLN = String(ln).replace(/\D/g, "");
  return cleanLN.padStart(4, "0");
}

function parseMoney(v) {
  if (typeof v === "number") return v;
  if (!v) return 0;
  return Number(
    String(v).replace(/[^\d,-]/g, "")
      .replace(/\./g, "")
      .replace(",", ".")
  ) || 0;
}

function normMunicipio(x) {
  return String(x || "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\u00A0/g, " ")     // espaço invisível
    .replace(/\t/g, " ")         // tab escondido
    .replace(/\s+/g, " ")         // espaços duplicados
    .toUpperCase()
    .trim();
}


function normalizarSistema(s) {
  s = norm(s);

  if (["RETAIL", "RET"].includes(s)) return "RETAIL";
  if (["R3", "ECC"].includes(s)) return "ECC";
  if (["S4", "TDF"].includes(s)) return "S4";

  return s;
}

function normalizarCompetencia(c) {
  let s = String(c || "").trim();

  // transforma 1/2026 em 01/2026
  let [mes, ano] = s.split("/");
  if (!mes || !ano) return s;

  mes = mes.padStart(2, "0");

  return `${mes}/${ano}`;
}

function mmYYYY(d) {
  if (!d) return normalizarCompetenciaFallback("");

  // Date real
  if (Object.prototype.toString.call(d) === "[object Date]" && !isNaN(d)) {
    return ("0" + (d.getMonth() + 1)).slice(-2) + "/" + d.getFullYear();
  }

  let s = String(d).trim();

  // dd.mm.aaaa (PL)
  if (/^\d{2}\.\d{2}\.\d{4}$/.test(s)) {
    const [dd, mm, yyyy] = s.split(".");
    return `${mm}/${yyyy}`;
  }

  // mm/dd/yyyy (Guepardo)
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) {
    const [mm, dd, yyyy] = s.split("/");
    return ("0" + mm).slice(-2) + "/" + yyyy;
  }

  return normalizarCompetenciaFallback(s);
}


function gerarCompetenciasExpandida(comp) {
  if (!comp || comp === "NA") return ["SEM_COMP"];

  const [m, a] = comp.split("/").map(Number);
  const lista = [];

  for (let i = -12; i <= 12; i++) {
    let mes = m + i;
    let ano = a;

    while (mes <= 0) {
      mes += 12;
      ano--;
    }

    while (mes > 12) {
      mes -= 12;
      ano++;
    }

    lista.push(("0" + mes).slice(-2) + "/" + ano);
  }

  return lista;
}

function normalizarMunicipioComparacao(m) {

  let mun = normMunicipioFull(m);


  // remove palavras comuns adicionais
  mun = mun
    .replace(/\bDE\b/g, "")
    .replace(/\bDA\b/g, "")
    .replace(/\bDO\b/g, "")
    .replace(/\bDAS\b/g, "")
    .replace(/\bDOS\b/g, "")
    .replace(/\bCIDADE\b/g, "")
    .replace(/\bMUNICIPAL\b/g, "")
    .replace(/\bADM\b/g, "")
    .replace(/\bTRIBUTOS?\b/g, "")
    .replace(/\bFAZENDA\b/g, "")
    .replace(/\bSECRETARIA\b/g, "")
    .replace(/\bPREFEITURA\b/g, "")
    .replace(/\bMUNICIPIO\b/g, "")
    .replace(/\s+/g, " ")
    .trim();


  // usa mapa carregado da aba Dim Município
  if (MAPA_MUNICIPIOS[mun]) {

    return MAPA_MUNICIPIOS[mun];

  }


  return mun;

}


function normMunicipioFull(x) {
  return String(x || "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\u00A0/g, " ")
    .replace(/\t/g, " ")
    .replace(/\s+/g, " ")
    .replace(/[-/().]/g, " ")
    .replace(/\bSP\b/g, "")
    .replace(/\bRJ\b/g, "")
    .replace(/\bMG\b/g, "")
    .replace(/\bPR\b/g, "")
    .replace(/\bSC\b/g, "")
    .replace(/\bRS\b/g, "")
    .replace(/\bISS\b/g, "")
    .replace(/\bPM\b/g, "")
    .replace(/\bPREF\b/g, "")
    .replace(/\bSECRETARIA\b/g, "")
    .replace(/\bPREFEITURA\b/g, "")
    .replace(/\bMUNICIPIO\b/g, "")
    .toUpperCase()
    .trim();
}

function carregarMapaMunicipios() {

  const mapa = {};

  try {

    const ss = SpreadsheetApp.getActive();

    const aba = ss.getSheetByName(ABA_DIM_MUNICIPIO);


    if (!aba) {

      Logger.log("⚠️ Aba DIM_MUNICIPIO não encontrada");

      return mapa;

    }


    const dados = aba.getDataRange().getValues();


    for (let i = 1; i < dados.length; i++) {


      const origem = normMunicipioFull(dados[i][0]);

      const destino = normMunicipioFull(dados[i][1]);


      if (origem && destino) {

        mapa[origem] = destino;

      }

    }


    Logger.log(
      "✅ Municípios carregados: " + Object.keys(mapa).length
    );


    return mapa;


  } catch(e) {


    Logger.log("🚨 ERRO carregarMapaMunicipios");

    Logger.log(e);


    return mapa;

  }

}

// ======================================================
// CORE
// ======================================================

function localizarColunas(sheet, linhaCabecalho = 2) {

  const headers = sheet
    .getRange(linhaCabecalho, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const mapa = {};

  headers.forEach((h, i) => {
    mapa[norm(h)] = i;
  });

  return mapa;
}

function idx(mapa, nome) {

  const coluna = mapa[norm(nome)];

  if (coluna === undefined) {
    throw new Error("Cabeçalho não encontrado: " + nome);
  }

  return coluna;
}


function cruzarISS() {

  const stats = {
    plRaw: 0,
    plValid: 0,
    totalPL: 0,
    indexCount: 0,
    totalIndex: 0,
    totalSistema: 0,
    totalPago: 0,
    naoConciliadoQtd: 0,
    naoConciliadoValor: 0,
    match: {
      perfeito: 0,
      semSistema: 0,
      semComp: 0,
      lnComp: 0,
      ln: 0,
      global: 0,
      semMatch: 0
    }
  };

  Logger.log("=================================");
  Logger.log("🚀 INÍCIO CONCILIAÇÃO ISS");
  Logger.log("=================================");
  MAPA_MUNICIPIOS = carregarMapaMunicipios();

  const diagnosticoMap = {};
  const usadosGlobal = new Set(); // 🔥 controle único de consumo
  const reverseMapPL = {};
  const ss = SpreadsheetApp.getActive();
  mostrarDimMunicipio(ss);
  const pl = ss.getSheetByName(ABA_PL);
  const g = ss.getSheetByName(ABA_G);
  const errosPL = {};

  const out = ss.getSheetByName(ABA_OUT) || ss.insertSheet(ABA_OUT);
  out.clear();

  const lastRow = pl.getLastRow();
  const lastCol = pl.getLastColumn();

  const plData = pl
    .getRange(3, 1, lastRow - 2, lastCol)
    .getValues();
  const gData = g.getDataRange().getValues().slice(2);

  Logger.log("📥 Linhas PL: " + plData.length);
  // ========================
  // STATS - PL INICIAL
  // ========================
  stats.plRaw = plData.length;

  let linhasValidas = 0;
  plData.forEach(r => {
    if (linhaValidaPL(r)) linhasValidas++;
  });

  stats.plValid = linhasValidas;
  Logger.log("📥 Linhas Guepardo: " + gData.length);

  const totalPL = plData.reduce((s, r) => s + getValorPL(r), 0);
  stats.totalPL = totalPL;

  Logger.log("💰 TOTAL PL (original): " + totalPL);

  // ========================
  // PL
  // ========================
  const HPL = localizarColunas(pl);

const PL = {

  EMP: localizarColuna(headersPL,
  [
    "EMPRESA",
    "EMP"
  ]),


  LN: localizarColuna(headersPL,
  [
    "LN",
    "LOCAL NEGOCIO",
    "LOCAL DE NEGOCIO"
  ]),


  MUN: localizarColuna(headersPL,
  [
    "MUNICIPIO",
    "MUNICÍPIO"
  ]),


  COMP: localizarColuna(headersPL,
  [
    "COMPETENCIA",
    "PERIO APUR",
    "PERIO_APUR"
  ]),


  SIST: localizarColuna(headersPL,
  [
    "SISTEMA"
  ]),


  VALOR_PRINCIPAL: localizarColuna(headersPL,
  [
    "VALOR PRINCIPAL"
  ]),


  MULTA: localizarColuna(headersPL,
  [
    "MULTA - 337006",
    "MULTA"
  ]),


  JUROS: localizarColuna(headersPL,
  [
    "JUROS - 361003",
    "JUROS"
  ])

};

  // 🔥 DEBUG - TOTAL PL USADO NO CÓDIGO
  Logger.log("TOTAL PL (código): " + plData.reduce((s, r) => s + getValorPL(r), 0));

  const indexPL = {};
  const indexPL_lnValor = {};
  const listaPLUnica = [];
  const indexPLPorChave = {};

  plData.forEach((r, i) => {

    if (!linhaValidaPL(r)) return;

    const keyCompleta = buildKeyCompleta(r, PL);
    if (!keyCompleta) return;

    const chaves = gerarChavesVariantes(keyCompleta);
    const valor = getValorPL(r);
    const doc = String(r[PL.DOC] || "").trim();

    // ✅ DECLARA PRIMEIRO
    const emp = safe(r[PL.EMP]);
    const ln = normLN(r[PL.LN]);
    const comp = normalizarCompetencia(mmYYYY(r[PL.COMP]));

    // ✅ DEPOIS USA
    const item = {
      id: i,
      v: valor,
      doc,
      rowIndex: i,
      grupo: doc || keyCompleta,
      ln,
      comp,
      emp,
      duplicado: false,
      temLNMasOutraCompetencia: false,
      valorParcialDetectado: false
    };

    listaPLUnica.push(item);

    // index normal (já está certo)
    const chavePrincipal = chaves[0];

    if (!indexPL[chavePrincipal]) {
      indexPL[chavePrincipal] = { valores: [], docs: [], valor: 0 };
    }

    indexPL[chavePrincipal].valores.push(item);
    indexPL[chavePrincipal].valor += valor;

    if (doc) indexPL[chavePrincipal].docs.push(doc);

    for (let j = 1; j < chaves.length; j++) {
      const k = chaves[j];

      if (!indexPL[k]) {
        indexPL[k] = { valores: [], docs: [], valor: 0 };
      }

      if (!indexPL[k].valores.some(v => v.id === item.id)) {
        indexPL[k].valores.push(item);
      }

      if (j === 0) {
        indexPL[k].valor += valor;
      }

      if (doc && !indexPL[k].docs.includes(doc)) {
        indexPL[k].docs.push(doc);
      }
    }
    validarDuplicidadePL(listaPLUnica);

    // 🔥 AGORA SIM: LN + VALOR (DENTRO DO LOOP)
    // (usa as variáveis já declaradas acima)
    const keyLNValor = emp + "|" + ln + "|" + comp;


    if (!indexPL_lnValor[keyLNValor]) {
      indexPL_lnValor[keyLNValor] = {
        lista: [],
        grupos: {}
      };
    }

    const grupo = doc || keyCompleta;

    indexPL_lnValor[keyLNValor].lista.push(item);

    if (!indexPL_lnValor[keyLNValor].grupos[grupo]) {
      indexPL_lnValor[keyLNValor].grupos[grupo] = [];
    }

    indexPL_lnValor[keyLNValor].grupos[grupo].push(item);

  });

  // grupo (doc ou chave)

  const totalIndexPL = listaPLUnica.reduce((s, v) => s + v.v, 0);
  const countIndexPL = listaPLUnica.length;

  // ======================================================
  // 🔥 DETECÇÃO AUTOMÁTICA DE ERROS (PL)
  // ======================================================
  listaPLUnica.forEach(item => {

    const keyLN = item.emp + "|" + item.ln;

    // 🔥 EXISTE LN EM OUTRA COMPETÊNCIA?
    const existeOutraComp = Object.keys(indexPL_lnValor)
      .some(k => k.startsWith(keyLN) && !k.includes(item.comp));

    if (existeOutraComp) {
      item.temLNMasOutraCompetencia = true;
    }

    // 🔥 VALOR PARCIAL
    const pool = Object.entries(indexPL_lnValor)
      .filter(([k]) => k.startsWith(keyLN))
      .flatMap(([, v]) => v.lista);

    const soma = pool.reduce((s, x) => s + x.v, 0);

    if (soma > 0 && Math.abs(soma - item.v) > 0.01) {
      item.valorParcialDetectado = true;
    }

  });
  // ========================
  // STATS - INDEX
  // ========================
  stats.totalIndex = totalIndexPL;
  stats.indexCount = countIndexPL;

  Logger.log("📊 TOTAL INDEX PL (somado): " + totalIndexPL);
  Logger.log("📊 QTD ITENS INDEXADOS: " + countIndexPL);

  // ========================
  // GUEPARDO
  // ========================
  const HG = localizarColunas(g);

  const G = {

    SIST: idx(HG, "Sistema"),

    EMP_COD: idx(HG, "Cod Emp"),

    EMP: idx(HG, "Empresa"),

    LN: idx(HG, "LN"),

    MUN: idx(HG, "Munic. Recolhimento"),

    COMP: idx(HG, "Competência NF"),

    VAL: idx(HG, "Valor ISS"),

    DOC: idx(HG, "Doc Guia"),
 
    TIPO_GUIA: idx(HG, "PA"),

    DOC_PL: idx(HG, "Doc Guia PL")

  };

  const map = {};

  gData.forEach((r, i) => {

    const key = buildKeyCompleta(r, G); // ✅ PRIMEIRO DECLARA

    if (!key) {
      Logger.log("🚨 Linha inválida GUEPARDO: " + JSON.stringify(r));
      return;
    }

    if (!map[key]) {
      const [empresa, ln, mun, comp, sistema] = key.split("|");

      map[key] = {
        empresa,
        ln,
        comp,
        mun,
        sistema,
        valor: 0,
        docs: new Set(),
        tipoGuia: r[G.TIPO_GUIA],
        rowIndex: i,
        doc: r[G.DOC]
      };
    }

    // ✅ AGORA SIM pode usar
    map[key].rowIndex = i;
    map[key].doc = r[G.DOC];

    map[key].valor += parseMoney(r[G.VAL]);

    const doc = String(r[G.DOC] || "").trim();
    if (doc) map[key].docs.add(doc);
  });

  // ======================================================
  // MATCH ENGINE (VERSÃO PROFISSIONAL)
  // ======================================================

  function executarMatchEngine(gRow, indexPL, indexPL_lnValor) {

    const tipoGuia = String(gRow.tipoGuia || "").toUpperCase();

    const emp = gRow.empresa;
    const ln = gRow.ln;
    const comp = gRow.comp;
    const mun = normalizarMunicipioComparacao(gRow.mun);
    const sist = gRow.sistema;
    const valor = gRow.valor;

    const keyLN = emp + "|" + ln;
    const tolerancia = calcularTolerancia(valor);

    const usadosTemp = new Set(usadosGlobal);

    // ======================================================
    // 🔍 VALIDAÇÃO LN
    // ======================================================
    const existeLN = Object.keys(indexPL_lnValor)
      .some(k => k.startsWith(keyLN));

    if (!existeLN && tipoGuia !== "DAM-BACK") {
      return {
        valor: 0,
        docs: [],
        tipo: "SEM MATCH",
        chave: null,
        itens: [],
        motivo: "LN NAO EXISTE NO PL"
      };
    }

    // ======================================================
    // 🔥 DAM-BACK (BACKLOG PURO)
    // ======================================================
    if (tipoGuia === "DAM-BACK") {

      const poolLN = Object.entries(indexPL_lnValor)
        .filter(([k]) => k.startsWith(keyLN))
        .flatMap(([, v]) => v.lista)
        .filter(v => !usadosGlobal.has(v.id));

      if (poolLN.length) {

        const matchComb = encontrarCombinacao(
          poolLN,
          valor,
          tolerancia + 0.05,
          usadosTemp
        );

        if (matchComb) {
          return consumirEMontar(
            { soma: matchComb.soma, itens: matchComb.itens },
            "BACKLOG",
            keyLN + "|BACKLOG",
            usadosGlobal,
            valor
          );
        }

        const matchGreedy = simularConsumo(
          poolLN,
          valor,
          usadosTemp,
          tolerancia
        );

        if (matchGreedy) {
          return consumirEMontar(
            { soma: matchGreedy.soma, itens: matchGreedy.itens },
            "BACKLOG PARCIAL",
            keyLN + "|BACKLOG",
            usadosGlobal,
            valor
          );
        }
      }
      // 👉 se falhar, continua fluxo normal
    }

    let melhorMatch = null;

    function avaliar(match, tipo, key, prioridade) {
      if (!match) return;

      const diff = Math.abs(match.soma - valor);
      const score = diff + (match.itens.length * 0.01);

      if (
        !melhorMatch ||
        prioridade < melhorMatch.prioridade ||
        (prioridade === melhorMatch.prioridade && score < melhorMatch.score)
      ) {
        melhorMatch = { match, tipo, key, prioridade, score };
      }
    }

    // ======================================================
    // 1. PERFEITO
    // ======================================================
    let key = [emp, ln, mun, comp, sist].join("|");

    let match = tentarMatch(key, valor, indexPL, tolerancia, usadosTemp);

    if (match) {
      match.rowGuepardo = gRow.rowIndex; return consumirEMontar(match, "PERFEITO", key, usadosGlobal, valor);
    }

    // ======================================================
    // 2. SEM SISTEMA (LN + COMP)
    // ======================================================
    const keyLNComp = keyLN + "|" + comp;

    match = tentarMatch(keyLNComp, valor, indexPL_lnValor, tolerancia, usadosTemp);

    if (match) {
      return consumirEMontar(match, "SEM SISTEMA", keyLNComp, usadosGlobal, valor);
    }

    // =========================
    // 🚨 REGRA DAM-BACK + NF BACKLOG
    // =========================
    const flagDamBack = String(gRow.colunaR || "")
      .toUpperCase()
      .includes("DAM-BACK");

    if (flagDamBack) {
      for (let g of listaGuepardo) {
        if (usadosGlobal.has(g.id)) continue;

        const flagBacklog = String(g.coluna1 || "")
          .toUpperCase()
          .includes("NF BACKLOG");

        if (
          flagBacklog &&
          g.valor === valorPL &&
          g.ln === itemPL.ln &&
          g.empresa === itemPL.empresa
        ) {
          usadosGlobal.add(g.id);

          return {
            tipoMatch: "DAM_BACK_MATCH",
            status: "OK",
            diff: "-",
            diagnostico: "DAM-BACK + NF BACKLOG: competência ignorada"
          };
        }
      }
    }

    // ======================================================
    // 3. COMPETÊNCIA EXPANDIDA
    // ======================================================
    const comps = gerarCompetenciasExpandida(comp);

    for (let c of comps) {

      const keyAlt = keyLN + "|" + c;

      match = tentarMatch(keyAlt, valor, indexPL_lnValor, tolerancia, usadosTemp);

      if (match) {
        return consumirEMontar(
          match,
          c === comp ? "COMP NORMAL" : "COMP AJUSTADA",
          keyAlt,
          usadosGlobal,
          valor
        );
      }
    }

    // ======================================================
    // 4. SEM COMPETÊNCIA
    // ======================================================
    key = [emp, ln, mun, "", ""].join("|");

    match = tentarMatch(key, valor, indexPL, tolerancia, usadosTemp);

    if (match) {
      return consumirEMontar(match, "SEM COMPETENCIA", key, usadosGlobal, valor);
    }

    // ======================================================
    // 5. LN + COMP (fallback direto)
    // ======================================================
    match = tentarMatch(keyLNComp, valor, indexPL_lnValor, tolerancia, usadosTemp);

    avaliar(match, "LN + COMP", keyLNComp, 4);

    // ======================================================
    // 6. POOL LN (único e reutilizado)
    // ======================================================
    const poolLN = Object.entries(indexPL_lnValor)
      .filter(([k]) => k.startsWith(keyLN))
      .flatMap(([, v]) => v.lista)
      .filter(v => !usadosTemp.has(v.id));

    if (poolLN.length) {

      // combinação
      const matchComb = encontrarCombinacao(
        poolLN,
        valor,
        tolerancia + 0.05,
        usadosTemp
      );

      if (matchComb) {
        return consumirEMontar(
          { soma: matchComb.soma, itens: matchComb.itens },
          "LN COMBINADO",
          keyLN,
          usadosGlobal,
          valor
        );
      }

      // greedy
      const matchPool = simularConsumo(
        poolLN,
        valor,
        usadosTemp,
        tolerancia
      );

      avaliar(matchPool, "POOL", keyLN, 5);

      // valor próximo (último fallback seguro)
      const candidato = poolLN.find(v =>
        Math.abs(v.v - valor) <= tolerancia
      );

      if (candidato) {
        return consumirEMontar(
          { soma: candidato.v, itens: [candidato] },
          "LN_VALOR_PROXIMO",
          keyLN,
          usadosGlobal,
          valor
        );
      }
    }

    // ======================================================
    // MELHOR MATCH (se houver)
    // ======================================================
    if (melhorMatch && melhorMatch.match) {
      melhorMatch.match.rowGuepardo = gRow.rowIndex;

      return consumirEMontar(
        melhorMatch.match,
        melhorMatch.tipo,
        melhorMatch.key,
        usadosGlobal,
        valor
      );
    }

    // ======================================================
    // ❌ SEM MATCH
    // ======================================================
    return {
      valor: 0,
      docs: [],
      tipo: "SEM MATCH",
      chave: null
    };
  }



  // ========================
  // OUTPUT
  // ========================
  const res = [[
    "Empresa", "LN", "Competência", "Município", "Sistema",
    "Valor Sistema", "Docs Guepardo",
    "Valor Pago", "Docs PL",
    "Diferença",
    "Status", "Compliance Fiscal (%)", "Compliance Processo (%)",
    "Tipo Match", "Tipo Erro",
    "Nível Match", "Quebra", "Diagnóstico",
    "Sugestão"
  ]];

  Object.keys(map).forEach(key => {

    const gRow = map[key];

    // ================= DEBUG INVESTIGAÇÃO =================

    const keyLN = gRow.empresa + "|" + gRow.ln;

    const existeLN = Object.keys(indexPL_lnValor)
      .some(k => k.startsWith(keyLN));

    const munNorm = normalizarMunicipioComparacao(gRow.mun);

    const existeMunicipio = Object.keys(indexPL)
      .some(k => k.includes(munNorm));

    const valorExiste = listaPLUnica.some(v =>
      Math.abs(v.v - gRow.valor) <= 0.01
    );

    Logger.log("🧪 DEBUG MATCH:");
    Logger.log({
      empresa: gRow.empresa,
      ln: gRow.ln,
      comp: gRow.comp,
      mun: gRow.mun,
      valor: gRow.valor,

      existeLN,
      existeMunicipio,
      valorExiste
    });
    // ================= DEBUG INVESTIGAÇÃO =================

    let resultadoMatch = executarMatchEngine(gRow, indexPL, indexPL_lnValor);

    // 🔥 NOVO BLOCO
    if (resultadoMatch.valor === 0) {
      const matchErro = executarMatchComErro(gRow, indexPL_lnValor, usadosGlobal);

      if (matchErro) {
        resultadoMatch = matchErro;
        resultadoMatch.forcado = true; // 🔥 flag crítica
      }
    }

    // =========================
    // 🔥 MAPA REVERSO REAL
    // =========================
    const itensMatch = resultadoMatch?.itens || [];

    if (itensMatch.length > 0) {

      itensMatch.forEach(itemPL => {

        const keyPL = itemPL.id;

        // 🔥 agora sim pode usar itemPL
        const keyBaseLN = itemPL.emp + "|" + itemPL.ln;

        if (!reverseMapPL[keyBaseLN]) {
          reverseMapPL[keyBaseLN] = new Set();
        }

        if (!reverseMapPL[keyPL]) {
          reverseMapPL[keyPL] = new Set();
        }

        if (gRow.doc) {
          const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/edit#gid=${g.getSheetId()}&range=A${gRow.rowIndex + 3}`;
          const keyDoc = `${gRow.doc}|${url}`;

          reverseMapPL[keyPL].add(keyDoc);
          reverseMapPL[keyBaseLN].add(keyDoc);
        }

      });
    }
    // 🔥 DEBUG REAL - centavos ignorados
    const candidatoProximo = listaPLUnica.find(v =>
      !usadosGlobal.has(v.id) &&
      Math.abs(v.v - gRow.valor) <= 0.05
    );

    if (resultadoMatch.valor === 0 && candidatoProximo) {
      Logger.log("🔥 MATCH PERDIDO POR CENTAVOS (REAL):");
      Logger.log({
        G: gRow.valor,
        PL: candidatoProximo.v,
        diff: Math.abs(candidatoProximo.v - gRow.valor),
        empresa: gRow.empresa,
        ln: gRow.ln,
        comp: gRow.comp
      });
    }

    let match = {
      valor: resultadoMatch.valor,
      docs: resultadoMatch.docs,
      keyPL: resultadoMatch.chave
    };

    let tipoMatch = resultadoMatch.tipo;

    if (resultadoMatch.valor > 0 && tipoMatch === "SEM MATCH") {
      tipoMatch = "LN + VALOR"; // ou "MATCH_REAL"
    }


    // 🔥 CONTADOR DE MATCH (CORRIGIDO)
    if (resultadoMatch.valor > 0 && resultadoMatch.itens?.length) {

      switch (tipoMatch) {
        case "PERFEITO":
          stats.match.perfeito++;
          break;

        case "SEM SISTEMA":
          stats.match.semSistema++;
          break;

        case "SEM COMPETENCIA":
          stats.match.semComp++;
          break;

        case "LN + COMP":
          stats.match.lnComp++;
          break;

        case "LN + VALOR":
          stats.match.ln++;
          break;

        case "GLOBAL":
          stats.match.global++;
          break;

        default:
          // 🔥 TINHA VALOR MAS NÃO CLASSIFICOU DIREITO
          stats.match.ln++; // joga como LN (match fraco)
          break;
      }

    } else {
      // 🔥 AQUI SIM é SEM MATCH real
      stats.match.semMatch++;
    }

    if (tipoMatch === "LN + COMP" || tipoMatch === "SEM MATCH") {
      const munPL = match.keyPL ? match.keyPL.split("|")[2] : "NA";

      if (munPL !== normalizarMunicipioComparacao(gRow.mun)) {
        Logger.log("🌍 MUNICÍPIO DIVERGENTE:");
        Logger.log({
          G: gRow.mun,
          PL: munPL
        });
      }
    }


    // 🔥 DEBUG CERTO (AQUI)
    if (tipoMatch === "PERFEITO") {
      Logger.log("🔥 MATCH PERFEITO REAL:");
      Logger.log("KEY G: " + key);
      Logger.log("KEY PL: " + resultadoMatch.chave);
    }

    // 🔥 DEBUG REAL
    if (tipoMatch !== "SEM MATCH") {
      Logger.log("----- MATCH DEBUG -----");
      Logger.log("Tipo: " + tipoMatch);
      Logger.log("Valor G: " + gRow.valor);
      Logger.log("Valor PL: " + resultadoMatch.valor);
      Logger.log("Chave G: " + key);
      Logger.log("Chave PL: " + resultadoMatch.chave);
    }

    // ======================================================
    // 🔥 MATCH FINAL AGRESSIVO (POOL LN)
    // ======================================================
    if (match.valor === 0) {
      const tolerancia = calcularTolerancia(gRow.valor); // ✅ FIX

      const keyLNComp = gRow.empresa + "|" + gRow.ln + "|" + gRow.comp;
      const objLNComp = indexPL_lnValor[keyLNComp];

      if (objLNComp && objLNComp.lista.length) {
        const listaLNComp = objLNComp.lista;
        const soma = listaLNComp.reduce((s, x) => s + x.v, 0);
        const diffLocal = Math.abs(soma - gRow.valor);

        if (diffLocal <= tolerancia) {
          if (diffLocal <= tolerancia) {

            const novoMatch = consumirEMontar(
              { soma, itens: listaLNComp },
              diffLocal <= 0.01 ? "LN + COMP" : "LN + COMP (AJUSTE VALOR)",
              key,
              usadosGlobal,
              gRow.valor
            );

            match = {
              valor: novoMatch.valor,
              docs: novoMatch.docs,
              keyPL: novoMatch.chave
            };

            tipoMatch = novoMatch.tipo;
          }

        }

      }
    }
    let sugestao = "";

    if (tipoMatch === "SEM MATCH") {
      sugestao = sugerirCorrecao(gRow, indexPL, indexPL_lnValor);
    }

    // ======================================================
    // RESULTADO FINAL
    // ======================================================
    const pl = match.valor;
    const docsPL = match.docs; // agora é array de objetos
    const docsPorPL = {};
    const keyBase = [
      gRow.empresa,
      gRow.ln,
      "",
      gRow.comp,
      ""
    ].join("|");

    let docsGObjs = [];

    if (reverseMapPL[keyBase]) {
      docsGObjs = [...reverseMapPL[keyBase]].map(d => {
        const [text, url] = d.split("|");
        return { text, url };
      });
    } else if (gRow.docs && gRow.docs.size) {
      docsGObjs = [...gRow.docs].map(d => ({ text: d, url: "" }));
    }

    const docsG = montarRichText(docsGObjs);
    const diff = +(gRow.valor - pl).toFixed(2);
    const status = Math.abs(diff) <= 0.01 ? "OK" : pl === 0 ? "SEM PAGAMENTO" : "DIVERGENTE";

    let tipoErro = "";

    if (resultadoMatch.valor === 0) {

      if (tipoMatch === "SEM MATCH") {
        tipoErro = "NAO_PAGO_REAL";

      } else if (tipoMatch.includes("COMP")) {
        tipoErro = "COMPETENCIA";

      } else if (tipoMatch.includes("LN")) {
        tipoErro = "ERRO_CHAVE_LN";

      } else {
        tipoErro = "ERRO_MATCH";
      }

    } else {

      if (Math.abs(diff) <= 0.01) {
        tipoErro = "OK";

      } else if (tipoMatch.includes("COMP")) {
        tipoErro = "COMPETENCIA";

      } else {
        tipoErro = "VALOR";
      }

    }

    const fiscal = gRow.valor > 0 ? (pl / gRow.valor) * 100 : 100;
    const proc = gRow.valor > 0 ? Math.max(0, 100 - (Math.abs(diff) / gRow.valor * 100)) : 100;
    const keyParts = key.split("|"); // chave original

    let analise;

    if (match.keyPL) {
      const keyPartsReal = match.keyPL.split("|");
      analise = analisarChaves(gRow, keyPartsReal);
    } else {
      analise = {
        nivel: 0,
        quebra: ["SEM MATCH"]
      };
    }

    let diagnostico;

    if (tipoMatch === "SEM MATCH") {
      diagnostico = diagnosticoDetalhado(gRow, indexPL, indexPL_lnValor);
    } else {
      diagnostico = diagnosticoDetalhado(gRow, match, diff, tipoMatch);
    }
    const quebraStr = analise.quebra.join(", ") || "OK";

    const matchExecutivo = classificarMatchExecutivo(tipoMatch);

    Logger.log("--------------------------------------------------");
    Logger.log(`🔎 MATCH - ${gRow.empresa} | LN ${gRow.ln} | ${gRow.comp}`);
    Logger.log("💰 Valor Guepardo: " + gRow.valor);
    Logger.log("💰 Valor Match PL: " + match.valor);
    Logger.log("📌 Tipo Match: " + tipoMatch);

    if (match.valor > gRow.valor + 0.01) {
      Logger.log("🚨 OVERMATCH DETECTADO");
      Logger.log("Valor G: " + gRow.valor);
      Logger.log("Valor PL: " + match.valor);
    }
    if (tipoMatch === "SEM MATCH") {
      Logger.log("🧠 DEBUG NÃO MATCH:");
      Logger.log(diagnostico);
    }
    if (!match) {
      match = { valor: 0, docs: [], keyPL: null };
    }

    res.push([
      gRow.empresa,
      "'" + gRow.ln,
      gRow.comp,
      gRow.mun,
      gRow.sistema,
      gRow.valor,
      docsGObjs.length
        ? docsGObjs.map(d => d.text).join(" | ")
        : "SEM DOC",
      pl,
      docsPL && docsPL.length
        ? docsPL.map(d => d.text || d).join(" | ")
        : "SEM DOC",
      diff,
      status,
      fiscal,
      proc,
      matchExecutivo,
      tipoErro,
      analise.nivel + "/5",
      quebraStr,
      diagnostico,
      sugestao
    ]);
    diagnosticoMap[key] = {
      docs: match.docs,
      matchExecutivo,
      tipoMatch,
      status,
      diff,
      analise,
      quebraStr,
      diagnostico,
      sugestao
    };
  });


  // ======================================================
  // 🔥 DIAGNÓSTICO - PL NÃO UTILIZADO
  // ======================================================

  const naoConciliadoPL = listaPLUnica
    .filter(v => !usadosGlobal.has(v.id) && v.v > 0)
    .map(v => v.v);

  // 🔥 SOMA TOTAL NÃO CONCILIADA
  const totalNaoConciliado = naoConciliadoPL.reduce((a, b) => a + b, 0);
  // ========================
  // STATS - NÃO CONCILIADO
  // ========================
  stats.naoConciliadoQtd = naoConciliadoPL.length;
  stats.naoConciliadoValor = totalNaoConciliado;

  Logger.log("💥 TOTAL PL NÃO CONCILIADO: " + totalNaoConciliado);
  Logger.log("📊 QTD ITENS NÃO CONCILIADOS: " + naoConciliadoPL.length);

  const diag = ss.getSheetByName("Diagnostico ISS") || ss.insertSheet("Diagnostico ISS");
  diag.clear();

  diag.getRange(1, 1).setValue("Valores PL não conciliados");

  const linhas = naoConciliadoPL.map(v => [v]);

  if (linhas.length) {
    diag.getRange(2, 1, linhas.length, 1).setValues(linhas);
  }

  // ========================
  // 🎨 HIGHLIGHT MATCH SEMÂNTICO
  // ========================
  const rangeOut = out.getDataRange();
  const values = rangeOut.getValues();

  for (let i = 1; i < values.length; i++) {
    if (values[i][13].includes("COMP")) {
      out.getRange(i + 1, 1, 1, values[0].length)
        .setBackground("#FFF3CD");
    }
  }

  // ========================
  // WRITEBACK INTELIGENTE
  // ========================

  // pega header
  const header = g.getRange(2, 1, 1, g.getLastColumn()).getValues()[0];
  const headerPL = pl.getRange(2, 1, 1, pl.getLastColumn()).getValues()[0];

  // encontra coluna existente
  let colDocPL = header.findIndex(h => norm(h) === norm("Doc Guia PL")) + 1;
  let colDocG = headerPL.findIndex(h => norm(h) === norm("Doc Guepardo")) + 1;

  // se não existir, cria
  if (colDocPL === 0) {
    colDocPL = g.getLastColumn() + 1;
    g.getRange(2, colDocPL).setValue("Doc Guia PL");
  }
  if (colDocG === 0) {
    colDocG = pl.getLastColumn() + 1;
    pl.getRange(2, colDocG).setValue("Doc Guepardo");
  }

  const write = [];
  const notes = [];

  // ========================
  // MONTA OS DADOS PRIMEIRO
  // ========================
  for (let i = 0; i < gData.length; i++) {

    const r = gData[i];
    const key = buildKeyCompleta(r, G);

    if (!key) {
      write.push([""]);
      notes.push(["Linha inválida"]);
      continue;
    }

    const info = diagnosticoMap[key];

    // DOCS
    const docsObjs = info && info.docs ? info.docs : [];
    write.push([docsObjs.map(d => d.text).join(" | ")]);

    // NOTE
    // NOTE
    let texto = "❌ Não conciliado";

    if (info) {

      let explicacao = "";
      const tipo = info.tipoMatch || "";

      // 🔥 PRIORIDADE: MATCH COM ERRO
      if (tipo.includes("ERRO")) {

        explicacao = "Match forçado ignorando divergências (município, competência ou valor)";

        texto =
          `⚠️ MATCH COM ERRO

    Tipo: ${info.tipoMatch}
    Status: ${info.status}
    Diferença: ${info.diff}

    Diagnóstico:
    ${info.diagnostico}

    Explicação:
    ${explicacao}`;

      } else {

        // 🔽 fluxo normal
        switch (tipo) {

          case "TOTAL":
            explicacao = "Valor total encontrado com mesma chave (match perfeito)";
            break;

          case "MATCH_PERFEITO":
            explicacao = "Match exato por todas as dimensões (Empresa, LN, Competência, Município e Sistema)";
            break;

          case "MATCH_SEM_SISTEMA":
            explicacao = "Match encontrado sem correspondência de sistema";
            break;

          case "MATCH_SEM_COMPETENCIA":
            explicacao = "Match encontrado com divergência de competência";
            break;

          case "MATCH_LN_COMP":
            explicacao = "Match baseado em LN e competência";
            break;

          case "MATCH_LN_VALOR":
            explicacao = "Match apenas por LN e valor (baixa confiabilidade)";
            break;

          case "COMP AJUSTADA (TOTAL)":
          case "COMP AJUSTADA (COMB)":
            explicacao = "Pagamento encontrado em outra competência próxima";
            break;

          case "LN + COMP":
            explicacao = "Match por LN e competência sem município";
            break;

          default:
            explicacao = "Regra de fallback aplicada";
        }

        texto =
          `Match: ${info.tipoMatch}
Status: ${info.status}
Diferença: ${info.diff}

Diagnóstico:
${info.diagnostico}

Explicação:
${explicacao}`;
      }
    }

    notes.push([texto]);
  }

  // ========================
  // 🔥 AGORA SIM CRIA O RANGE
  // ========================
  if (write.length > 0) {

    const rangeWrite = g.getRange(3, colDocPL, write.length, 1);

    rangeWrite.setValues(write);
    rangeWrite.setNotes(notes);

  } else {
    Logger.log("⚠️ Nenhuma linha para escrever no writeback");
  }

  // ========================
  // 🔗 APLICA LINKS (RichText)
  // ========================
  for (let i = 0; i < gData.length; i++) {
    const r = gData[i];
    const key = buildKeyCompleta(r, G);
    const info = diagnosticoMap[key];

    if (!info || !info.docs || info.docs.length === 0) continue;

    const rich = montarRichText(info.docs); // ✅ JÁ ESTÁ NO FORMATO CERTO

    g.getRange(i + 3, colDocPL).setRichTextValue(rich);
  }


  // ========================
  // 🎨 HIGHLIGHT PL CONCILIADA (SÓ DOC)
  // ========================
  const startRow = 3; // por causa do slice(2)
  const docCol = PL.DOC + 1;

  // pega fundo atual
  const rangePL = pl.getRange(startRow, 1, plData.length, pl.getLastColumn());
  const bg = rangePL.getBackgrounds();

  // marca em memória (rápido)
  Object.values(indexPL).forEach(obj => {
    obj.valores.forEach(v => {
      if (usadosGlobal.has(v.id) && bg[v.rowIndex]) {
        bg[v.rowIndex][PL.DOC] = "#FFF3CD";
      }
    });
  });

  // aplica tudo de uma vez (rápido)
  rangePL.setBackgrounds(bg);

  // ========================
  // 🔗 WRITEBACK GUEPARDO → PL (NOVO)
  // ========================
  const colDocsPL = 9;

  for (let i = 0; i < listaPLUnica.length; i++) {

    const item = listaPLUnica[i];
    const keyBaseLN = item.emp + "|" + item.ln + "|" + item.comp;

    const docs = reverseMapPL[keyBaseLN]
      ? Array.from(reverseMapPL[keyBaseLN]).map(d => {
        try {
          return JSON.parse(d);
        } catch (e) {
          const [text, url] = String(d).split("|");
          return { text, url };
        }
      })
      : [];

    if (docs.length === 0) continue;

    const rich = montarRichText(docs);

    const rowPL = item.rowIndex + 3;

    pl.getRange(rowPL, colDocsPL)
      .setRichTextValue(rich);
  }


  // ======================================================
  // BASE LN
  // ======================================================
  function gerarBaseLN(ss) {

    const base = ss.getSheetByName(ABA_OUT);
    const s = ss.getSheetByName(ABA_LN) || ss.insertSheet(ABA_LN);
    s.clear();

    const data = base.getDataRange().getValues();
    const header = data.shift();

    const idx = {};
    header.forEach((h, i) => idx[norm(h)] = i);

    const map = {};

    data.forEach(r => {

      const ln = normLN(r[idx["LN"]]);
      if (!map[ln]) map[ln] = { v: 0, p: 0, d: 0 };

      map[ln].v += parseMoney(r[idx["VALOR SISTEMA"]]);
      map[ln].p += parseMoney(r[idx["VALOR PAGO"]]);
      map[ln].d += Math.abs(parseMoney(r[idx["DIFERENCA"]]));
    });

    const out = [[
      "LN", "Valor Sistema", "Valor Pago", "Diferença",
      "Compliance Fiscal (%)", "Compliance Processo (%)"
    ]];

    Object.entries(map).forEach(([ln, v]) => {

      const fiscal = v.v > 0 ? (v.p / v.v) * 100 : 100;
      const proc = v.v > 0 ? Math.max(0, 100 - (v.d / v.v * 100)) : 100;

      out.push([
        "'" + ln,
        v.v,
        v.p,
        v.d,
        fiscal,
        proc
      ]);
    });

    s.getRange(1, 1, out.length, out[0].length).setValues(out);
  }

  function analisarChaves(gRow, plKeyParts) {

    const resultado = {
      nivel: 0,
      quebra: []
    };

    const empPL = plKeyParts[0];
    const lnPL = plKeyParts[1];
    const munPL = plKeyParts[2]; // ✅ índice correto
    const compPL = plKeyParts[3]; // ✅ índice correto
    const sistPL = plKeyParts[4]; // ✅ índice correto

    const munG = normalizarMunicipioComparacao(gRow.mun);
    const sistG = normalizarSistema(gRow.sistema);

    const checks = [
      { nome: "EMPRESA", ok: gRow.empresa === empPL },
      { nome: "LN", ok: gRow.ln === lnPL },
      { nome: "MUNICIPIO", ok: municipioSimilar(munG, munPL) },
      { nome: "COMPETENCIA", ok: gRow.comp === compPL },
      { nome: "SISTEMA", ok: sistG === sistPL }
    ];

    checks.forEach(c => {
      if (c.ok) {
        resultado.nivel++;
      } else {
        resultado.quebra.push(c.nome);
      }
    });

    return resultado;
  }

  const totalPago = res.slice(1).reduce((s, r) => s + parseMoney(r[7]), 0);
  const totalSistema = res.slice(1).reduce((s, r) => s + parseMoney(r[5]), 0);

  // ========================
  // STATS - RESULTADO FINAL
  // ========================
  stats.totalPago = totalPago;
  stats.totalSistema = totalSistema;

  const statsPL = calcularStatsPL(listaPLUnica, usadosGlobal);

  // ========================
  // 🔥 PL NÃO CONCILIADO (BASE REAL)
  // ========================
  const naoConciliadosPL = new Set(
    listaPLUnica
      .filter(v => !usadosGlobal.has(v.id))
      .map(v => v.v.toFixed(2))
  );


  // ========================
  // 🔥 BASE DE ERROS DO PL (AUDITORIA)
  // ========================

  const naoConciliados = listaPLUnica.filter(v => !usadosGlobal.has(v.id));

  naoConciliados.forEach(item => {

    let tipo = "NAO_CLASSIFICADO";

    // 🔥 1. ERRO LN
    if (!validarLNReal(item.ln)) {
      tipo = "ERRO_LN";
    }

    // 🔥 2. COMPETÊNCIA
    else if (item.temLNMasOutraCompetencia) {
      tipo = "ERRO_COMPETENCIA";
    }

    // 🔥 3. VALOR PARCIAL
    else if (item.valorParcialDetectado) {
      tipo = "VALOR_PARCIAL";
    }

    // 🔥 4. DUPLICIDADE
    else if (item.duplicado) {
      tipo = "DUPLICIDADE";
    }

    if (!errosPL[tipo]) {
      errosPL[tipo] = { qtd: 0, valor: 0 };
    }

    errosPL[tipo].qtd++;
    errosPL[tipo].valor += item.v;

  });

  // DEBUG (opcional, mas recomendo)
  Logger.log("📊 ERROS PL:");
  Logger.log(errosPL);

  // ========================
  // ✍️ WRITE ERROS NA PL (COLUNA V)
  // ========================

  const colErroPL = 22; // coluna V

  const outputPL = [];

  // cria mapa por linha real
  const erroPorLinha = {};

  // classifica cada item
  listaPLUnica.forEach(item => {

    let tipo = "";
    let descricao = "";

    const foiUsado = usadosGlobal.has(item.id);

    if (foiUsado) {
      tipo = "CONCILIADO";
      descricao = "Guia conciliada com sucesso";
    } else {

      if (!validarLNReal(item.ln)) {
        tipo = "ERRO_LN";
        descricao = "LN inválida ou não encontrada no sistema";
      }

      else if (item.temLNMasOutraCompetencia) {
        tipo = "ERRO_COMPETENCIA";
        descricao = "Pagamento encontrado em outra competência";
      }

      else if (item.valorParcialDetectado) {
        tipo = "VALOR_PARCIAL";
        descricao = "Valor não bate com o total da LN";
      }

      else if (item.duplicado) {
        tipo = "DUPLICIDADE";
        descricao = "Documento duplicado na base PL";
      }

      else {
        tipo = "NAO_CONCILIADO";
        descricao = "Não houve match com o sistema";
      }
    }

    const textoFinal = `${tipo} | ${descricao}`;

    const rowReal = item.rowIndex + 3;

    erroPorLinha[rowReal] = textoFinal;
  });

  // monta matriz completa (evita múltiplos setValue)
  for (let i = 0; i < plData.length; i++) {
    const rowReal = i + 3;

    outputPL.push([
      erroPorLinha[rowReal] || ""
    ]);
  }

  // escreve tudo de uma vez (performance alta)
  pl.getRange(3, colErroPL, outputPL.length, 1).setValues(outputPL);


  // ========================
  // 🔥 FILTRO FINAL DA BASE
  // ========================
  const resFiltrado = [res[0]];

  for (let i = 1; i < res.length; i++) {
    const row = res[i];
    const valorPL = parseMoney(row[7]);
    const tipoErro = row[14];

    const ehConciliado = valorPL > 0;
    const ehErroPL = tipoErro !== "OK";

    if (ehConciliado || ehErroPL) {
      resFiltrado.push(row);
    }
  }

  // escreve base limpa
  out.clear();

  out.getRange(1, 1, resFiltrado.length, resFiltrado[0].length)
    .setValues(resFiltrado);

  /*// ========================
  // 🔗 WRITE DOCS GUEPARDO (CORRETO)
  // ========================
  const colDocsG = 7; // coluna "Docs Guepardo"

  for (let i = 1; i < resFiltrado.length; i++) {

    const row = resFiltrado[i];

    // 🔥 CHAVE CORRETA (MESMA DO MAPA)
    const keyBase = row[0] + "|" + row[1].replace("'", "");

    const docsSet = reverseMapPL[keyBase]; // ✅ CORRETO

    if (!docsSet || docsSet.size === 0) continue;

    // 🔥 CONVERSÃO PARA RICHTEXT
    const docsFormatados = [...docsSet].map(d => {
      const [text, url] = d.split("|");
      return { text, url };
    });

    out.getRange(i + 1, colDocsG)
      .setRichTextValue(montarRichText(docsFormatados));
  }
 */
  // ========================
  // 🔗 APLICAÇÃO DE HYPERLINKS (RICHTEXT) — VERSÃO CORRETA
  // ========================

  for (let i = 1; i < resFiltrado.length; i++) {
    const row = resFiltrado[i];

    const key = [
      row[0], // empresa
      row[1].replace("'", ""), // LN
      normalizarMunicipioComparacao(row[3]),
      row[2], // competência
      normalizarSistema(row[4])
    ].join("|");

    const info = diagnosticoMap[key];

    if (!info || !info.docs || info.docs.length === 0) continue;

    const rich = montarRichText(info.docs);

    out.getRange(i + 1, colDocsPL).setRichTextValue(rich);
  }


 function buildKeyCompleta(r, schema) {

  let emp, ln, mun, comp, sist;

  try {
    emp = safe(r[schema.EMP]);
  } catch(e){
    Logger.log("ERRO EMP");
    Logger.log(e);
    Logger.log(r);
    return null;
  }

  try {
    ln = normLN(r[schema.LN]);
  } catch(e){
    Logger.log("ERRO LN");
    Logger.log(e);
    Logger.log(r);
    return null;
  }

  try {
    mun = normalizarMunicipioComparacao(r[schema.MUN]);
  } catch(e){
    Logger.log("ERRO MUNICIPIO");
    Logger.log(e);
    Logger.log(r);
    return null;
  }

  try {
    comp = normalizarCompetencia(mmYYYY(r[schema.COMP]));
  } catch(e){
    Logger.log("ERRO COMPETENCIA");
    Logger.log(e);
    Logger.log(r);
    return null;
  }

  try {
    sist = normalizarSistema(r[schema.SIST]);
  } catch(e){
    Logger.log("ERRO SISTEMA");
    Logger.log(e);
    Logger.log(r);
    return null;
  }

  // Sem empresa ou LN não existe possibilidade de cruzamento
if (!emp || ln === "0000") {
    return null;
}


// Competência é obrigatória
if (comp === "SEM_COMP") {
    return null;
}


// Monta chave completa quando possível
return [
  emp,
  ln,
  mun || "",
  comp || "",
  sist || ""

].join("|");
}

  function gerarChavesPriorizadas(key) {
    const [emp, ln, mun, comp, sist] = key.split("|");

    return [
      { tipo: "PERFEITO", key: [emp, ln, mun, comp, sist].join("|") },
      { tipo: "SEM_SIST", key: [emp, ln, mun, comp, ""].join("|") },
      { tipo: "SEM_COMP", key: [emp, ln, mun, "", ""].join("|") },
      { tipo: "LN_COMP", key: [emp, ln, "", comp, ""].join("|") },
      { tipo: "LN", key: [emp, ln, "", "", ""].join("|") }
    ];
  }




  function diagnosticoDetalhado(gRow, indexPL, indexPL_lnValor) {

    const logs = [];

    const emp = gRow.empresa;
    const ln = gRow.ln;
    const comp = gRow.comp;
    const mun = normalizarMunicipioComparacao(gRow.mun);

    // 🔥 VALIDAÇÃO DE LN
    if (!validarLNReal(ln)) {
      logs.push("🚨 LN inválida ou genérica");
    }

    // ========================
    // 1. EXISTE LN NO PL?
    // ========================
    const temLN = Object.keys(indexPL_lnValor)
      .some(k => k.startsWith(emp + "|" + ln));

    if (!temLN) {
      logs.push("❌ LN não encontrada no PL");
      return logs.join(" | ");
    }

    logs.push("✔ LN encontrada no PL");

    // ========================
    // 2. EXISTE LN + COMP?
    // ========================
    const keyLNComp = emp + "|" + ln + "|" + comp;
    const objLNComp = indexPL_lnValor[keyLNComp];
    const listaLNComp = objLNComp ? objLNComp.lista : null;

    if (!listaLNComp || listaLNComp.length === 0) {
      logs.push("❌ LN encontrada, mas não na mesma competência");

      // tenta ver se existe em outra competência
      const comps = gerarCompetenciasExpandida(comp);

      const existeOutraComp = comps.some(c => {
        const key = emp + "|" + ln + "|" + c;
        return indexPL_lnValor[key];
      });

      if (existeOutraComp) {
        logs.push("⚠️ Existe pagamento em outra competência");
      }

    } else {
      logs.push("✔ LN + competência encontrada");

      // ========================
      // 3. VALOR BATE?
      // ========================
      const soma = listaLNComp.reduce((s, x) => s + x.v, 0);
      const diff = Math.abs(soma - gRow.valor);

      if (diff <= 0.01) {
        logs.push("✔ Valor bate (mas não casou por outra chave)");
      } else if (diff < 10) {
        logs.push(`⚠️ Diferença pequena (${diff.toFixed(2)}) → possível multa/juros`);
      } else {
        logs.push(`❌ Valor não bate (PL: ${soma.toFixed(2)} | Diff: ${diff.toFixed(2)})`);
      }

      // ========================
      // 4. MUNICÍPIO
      // ========================
      const existeMunicipio = Object.keys(indexPL).some(k => {
        return k.includes(emp + "|" + ln) && k.includes(mun);
      });

      if (!existeMunicipio) {
        logs.push("❌ Município divergente");
      } else {
        logs.push("✔ Município encontrado");
      }

      // ========================
      // 5. VALOR ISOLADO
      // ========================
      const valorExiste = Object.values(indexPL)
        .some(obj => Math.abs(obj.valor - gRow.valor) <= 0.01);

      if (!valorExiste) {
        logs.push("❌ Valor não existe no PL");
      } else {
        logs.push("⚠️ Valor existe no PL (mas com outra chave)");
      }

      return logs.join(" | ");
    }
  }
  ocultarDimMunicipio(ss);

  gerarDashboardISS(ss, statsPL, errosPL);
  // 🔥 apagar aba diagnóstico
  const diagSheet = ss.getSheetByName(ABA_DIAG);

  if (diagSheet) {
    ss.deleteSheet(diagSheet);
  }
}

function sugerirCorrecao(gRow, indexPL, indexPL_lnValor) {

  const sugestoes = [];

  const emp = gRow.empresa;
  const ln = gRow.ln;
  const comp = gRow.comp;
  const valor = gRow.valor;
  const mun = normalizarMunicipioComparacao(gRow.mun);

  // ========================
  // 1. COMPETÊNCIA
  // ========================
  const comps = gerarCompetenciasExpandida(comp);

  for (let c of comps) {
    const key = emp + "|" + ln + "|" + c;
    const obj = indexPL_lnValor[key];

    if (obj && obj.lista) {
      const soma = obj.lista.reduce((s, x) => s + x.v, 0);
      const diff = Math.abs(soma - gRow.valor);

      if (Math.abs(soma - valor) <= 0.01) {
        sugestoes.push(`💡 Ajustar competência para ${c}`);
        break;
      }

      if (soma > 0) {
        sugestoes.push(`💡 Possível pagamento em ${c} (R$ ${soma.toFixed(2)})`);
      }
    }
  }

  // ========================
  // 2. MUNICÍPIO
  // ========================
  const possiveisMunicipios = [];

  Object.keys(indexPL).forEach(k => {
    const [e, l, m] = k.split("|");

    if (e === emp && l === ln && m !== mun) {
      possiveisMunicipios.push(m);
    }
  });

  if (possiveisMunicipios.length) {
    const sugestaoMun = [...new Set(possiveisMunicipios)].slice(0, 2).join(", ");
    sugestoes.push(`💡 Verificar município (possível: ${sugestaoMun})`);
  }

  // ========================
  // 3. VALOR PARCIAL
  // ========================
  const keyLN = emp + "|" + ln;

  const listaLN = Object.entries(indexPL_lnValor)
    .filter(([k]) => k.startsWith(keyLN))
    .flatMap(([, v]) => v.lista);

  if (listaLN.length) {
    const somaTotal = listaLN.reduce((s, x) => s + x.v, 0);

    if (somaTotal > 0 && somaTotal < valor) {
      sugestoes.push(`💡 Pagamento parcial encontrado (R$ ${somaTotal.toFixed(2)} de R$ ${valor.toFixed(2)})`);
    }
  }

  // ========================
  // 4. VALOR EM OUTRA LN
  // ========================
  const existeValor = Object.entries(indexPL_lnValor).find(([k, obj]) => {

    if (!obj || !obj.lista) return false;

    const soma = obj.lista.reduce((s, x) => s + x.v, 0);
    const diff = Math.abs(soma - gRow.valor);

    return Math.abs(soma - valor) <= 0.01 && !k.startsWith(keyLN);
  });

  if (existeValor) {
    const outraLN = existeValor[0].split("|")[1];
    sugestoes.push(`💡 Valor encontrado em outra LN (${outraLN})`);
  }

  // ========================
  // 5. NADA ENCONTRADO
  // ========================
  if (sugestoes.length === 0) {
    sugestoes.push("💡 Nenhum indício encontrado → possível não pagamento");
  }

  return sugestoes.join(" | ");
}

function gerarDashboardISS(ss, statsPL, errosPL) {

  const base = ss.getSheetByName("Base Conciliacao ISS");
  const plSheet = ss.getSheetByName("Cód de Barras (PL Financ)");
  let dash = ss.getSheetByName("ISS Overview");

  if (dash) ss.deleteSheet(dash);
  dash = ss.insertSheet("ISS Overview");

  dash.clear();
  dash.setHiddenGridlines(true);

  // =========================
  // 🎨 PALETA
  // =========================
  const AZUL = "#1F3B5C";        // principal
  const AZUL_CLARO = "#4A90E2";  // destaque
  const VERDE = "#27AE60";       // sucesso
  const ROXO = "#6C5CE7";        // analytics
  const CINZA = "#F5F7FA";
  const BORDA = "#E0E0E0";

  //controlar largura das colunas

  dash.setColumnWidth(1, 260); // 🔥 coluna A maior (município)
  dash.setColumnWidths(2, 16, 120); // resto padrão
  dash.setColumnWidth(6, 30); // F
  dash.setColumnWidth(7, 30); // G
  dash.setColumnWidth(3, 30);
  dash.setColumnWidth(4, 140); // Coluna C 
  //dash.setRowHeight(15, 12);
  //dash.setRowHeight(16, 12);


  for (let i = 1; i <= 40; i++) dash.setRowHeight(i, 28);

  // =========================
  // HEADER
  // =========================
  dash.getRange("A1:Q2").merge()
    .setValue("📊 ISS OVERVIEW EXECUTIVO")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground(AZUL)
    .setFontColor("#FFFFFF");

  // =========================
  // BASE
  // =========================
  const data = base.getDataRange().getValues();

  // 🔥 KPI REAL DE CONCILIAÇÃO (BASE PL)
  const listaPLUnica = []; // você precisa passar isso como parâmetro idealmente
  const usadosGlobal = new Set(); // idem

  const totalGuias = statsPL.totalQtd;
  const guiasConciliadas = statsPL.conciliadoQtd;
  const guiasAbertas = statsPL.naoConciliadoQtd;
  const percConciliadoQtd = totalGuias > 0 ? guiasConciliadas / totalGuias : 0;

  const header = data.shift();

  const idx = {};
  header.forEach((h, i) => idx[h] = i);


  const startRow = 3;

  const lastRow = plSheet.getLastRow();
  const lastCol = plSheet.getLastColumn();

  const raw = plSheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol)
    .getValues();

  let linhasValidas = 0;

  const plData = raw.filter(r => {
    const ok = linhaValidaPL(r);
    if (ok) linhasValidas++;
    return ok;
  });

  //Logger.log("📥 Linhas PL (RAW range): " + raw.length);
  Logger.log("📥 Linhas PL (válidas): " + linhasValidas);

  const totalPL = plData.reduce((s, r) => s + getValorPL(r), 0);

  let totalSistema = 0;
  let totalPago = 0;

  const municipioMap = {};

  data.forEach(r => {
    const sistema = parseFloat(r[idx["Valor Sistema"]]) || 0;
    const pago = parseFloat(r[idx["Valor Pago"]]) || 0;
    const municipio = r[idx["Município"]] || "NA";

    totalSistema += sistema;
    totalPago += pago;

    if (!municipioMap[municipio]) {
      municipioMap[municipio] = { qtd: 0, valor: 0 };
    }

    const aberto = sistema - pago;
    const status = r[idx["Status"]];

    if (
      aberto > 0.01 &&
      (status === "SEM PAGAMENTO" || status === "DIVERGENTE")
    ) {
      if (!municipioMap[municipio]) {
        municipioMap[municipio] = { qtd: 0, valor: 0 };
      }

      municipioMap[municipio].qtd++;
      municipioMap[municipio].valor += aberto;
    }
  });

  const emAberto = totalSistema - totalPago;
  const pendenteManual = totalPL - totalPago;

  const percSistema = totalSistema > 0 ? totalPago / totalSistema : 0;
  const percPL = totalPL > 0 ? totalPago / totalPL : 0;

  // =========================
  // 🔝 KPI EXECUTIVO (CARDS)
  // =========================
  const kpis = [
    ["🚀 Análise Base de Dados", ""],
    ["Total Sistema (Guepardo)", totalSistema],
    ["Total Provisionado (PL)", totalPL],
    ["Total Conciliado", totalPago],
    ["Em Aberto", emAberto],
  ];

  dash.getRange(4, 1, kpis.length, 2).setValues(kpis)
    .setBorder(true, true, true, true, false, false, BORDA, SpreadsheetApp.BorderStyle.SOLID);

  // 🔥 HEADER (linha do foguete)
  dash.getRange(4, 1, 1, 2).merge()
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#F5F7FA")
    .setFontSize(12)
    .setFontColor("#1F3B5C")
    .setBorder(true, true, true, true, false, false, BORDA, SpreadsheetApp.BorderStyle.SOLID);

  // =========================
  // 📊 KPI OPERACIONAL (GUIAS)
  // =========================

  const kpisOperacionais = [
    ["📊 Impacto Operacional", ""],
    ["Qtd Guias (PL)", statsPL.totalQtd],
    ["Guias Conciliadas", statsPL.conciliadoQtd],
    ["Guias em Aberto", statsPL.naoConciliadoQtd],
    ["% Conciliação (Qtd)", statsPL.totalQtd > 0 ? statsPL.conciliadoQtd / statsPL.totalQtd : 0],
    ["", ""],
    ["Última Atualização", new Date()]
  ];

  dash.getRange(10, 4, kpisOperacionais.length, 2)
    .setValues(kpisOperacionais)
    .setBorder(true, true, true, true, false, false, BORDA, SpreadsheetApp.BorderStyle.SOLID);

  // 🔥 HEADER
  dash.getRange(10, 4, 1, 2).merge()
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#F5F7FA")
    .setFontSize(12)
    .setFontColor("#1F3B5C")
    .setBorder(true, true, true, true, false, false, BORDA, SpreadsheetApp.BorderStyle.SOLID);

  // 🎨 formatação IGUAL ao bloco de qualidade
  const rangeKPIs = dash.getRange(10, 4, kpisOperacionais.length, 2);

  rangeKPIs
    .setBackground(CINZA)
    .setBorder(true, true, true, true, false, false, BORDA, SpreadsheetApp.BorderStyle.SOLID);

  dash.getRange("D10:D16").setFontWeight("bold");
  dash.getRange("E10:E13")
    .setHorizontalAlignment("right");

  // formatos
  dash.getRange("E11:E13").setNumberFormat("0");
  dash.getRange("E14").setNumberFormat("0.00%");
  dash.getRange("E16").setNumberFormat("dd/MM/yyyy HH:mm");

  const kpiRange = dash.getRange(4, 1, kpis.length, 2);

  dash.getRange("B4:B7").setHorizontalAlignment("right");
  dash.getRange("E4:E7").setHorizontalAlignment("right");

  kpiRange
    .setBackground(CINZA)
    .setBorder(true, true, true, true, false, false, "#E0E0E0", SpreadsheetApp.BorderStyle.SOLID);

  dash.getRange("A4:A8").setFontWeight("bold");

  // FORMATAÇÃO
  dash.getRange(4, 2, kpis.length - 1, 1)
    .setNumberFormat("R$ #,##0.00");

  dash.getRange(4 + kpis.length - 1, 2)
    .setNumberFormat("R$ #,##0.00");

  dash.getRange(4, 8, kpisOperacionais.length, 1)
    .setNumberFormat("0");

  dash.getRange(7, 8)
    .setNumberFormat("0.00%");

  // CORES INTELIGENTES
  dash.getRange("B6").setFontColor(VERDE);       // conciliado
  dash.getRange("B7").setFontColor("#6C5CE7");   // aberto (roxo elegante)

  // =========================
  // 📊 FLUXO (COLUNA)
  // =========================
  const chartFunil = dash.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dash.getRange("A5:B7"))
    .setOption("title", "Fluxo de Conciliação")
    .setOption("colors", [AZUL_CLARO])
    .setPosition(4, 7, 0, 0)
    .setOption("width", 600)
    .setOption("height", 350)
    .build();

  dash.insertChart(chartFunil);

  // =========================
  // 📊 QUALIDADE
  // =========================
  const rangeQualidade = dash.getRange("D4:E8");

  // mantém estrutura 4 linhas (igual KPI)
  const qualidade = [
    ["", ""],
    ["Cobertura Fiscal", percSistema],
    ["Eficiência Processo", percPL],
    ["", ""],
  ];

  dash.getRange(4, 4, 1, 2).merge()
    .setValue("💰 Impacto Financeiro")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(CINZA)
    .setFontSize(12)
    .setFontColor("#1F3B5C")
    .setBorder(true, true, true, true, false, false, BORDA, SpreadsheetApp.BorderStyle.SOLID);

  dash.getRange(5, 4, qualidade.length, 2).setValues(qualidade)
    .setBorder(true, true, true, true, false, false, BORDA, SpreadsheetApp.BorderStyle.SOLID);

  // 🎯 formato %
  // 🎯 FORMATAÇÃO %
  dash.getRange("E4:E7")
    .setNumberFormat("0.00%")
    .setFontColor(AZUL)
    .setFontWeight("bold");

  // 💬 TOOLTIP EXECUTIVO (mais curto e claro)
  dash.getRange("E6")
    .setNote("Conciliado / Sistema (Guepardo)");

  dash.getRange("E7")
    .setNote("Conciliado / Provisionado (PL Financeiro)");

  // 🎨 mesmo visual dos KPIs
  rangeQualidade
    .setBackground(CINZA)
    .setBorder(true, true, true, true, false, false, BORDA, SpreadsheetApp.BorderStyle.SOLID);


  dash.getRange("D4:D7")
    .setFontWeight("bold");

  dash.getRange("E4:E7")
    .setHorizontalAlignment("right");




  // =========================
  // 🥧 DISTRIBUIÇÃO
  // =========================
  const hidden = 18;

  dash.getRange(51, 8, 3, 2).setValues([
    ["Status", "Valor"],
    ["Conciliado", totalPago],
    ["Em Aberto", emAberto]
  ]);

  dash.insertChart(
    dash.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dash.getRange(51, 8, 3, 2))
      .setPosition(10, 8, 0, 0)
      .setOption("title", "Total Sistema (Guepardo)")
      .setPosition(4, 13, 0, 0)
      .setOption("colors", ["#4A90E2", "#D6E4FF"])
      .setOption("width", 600)
      .setOption("height", 350)
      .build()
  );

  dash.getRange(55, 8, 3, 2).setValues([
    ["Status", "Valor"],
    ["Conciliado", totalPago],
    ["Não Conciliado", pendenteManual]
  ]);

  dash.insertChart(
    dash.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dash.getRange(51 + 4, 8, 3, 2))
      .setPosition(16, 8, 0, 0)
      .setOption("title", "Total Provisionado (Fin.)")
      .setPosition(18, 13, 0, 0)
      .setOption("colors", ["#27AE60", "#EAF7EF"])
      .setOption("width", 600)
      .setOption("height", 350)
      .build()

  );

  const startRowDashboard = 51; // 🔥 começa abaixo do insight

  // título
  dash.getRange(startRowDashboard, 1, 1, 3).merge()
    .setValue("Top Municípios (Origem: Guepardo)")
    .setFontWeight("bold");

  // dados
  const munArr = Object.entries(municipioMap)
    .map(([m, v]) => [m, v.qtd, v.valor])
    .sort((a, b) => b[2] - a[2])
    .slice(0, 10);

  if (munArr.length === 0) {
    dash.getRange(startRowDashboard + 1, 1)
      .setValue("Sem dados disponíveis");
    return;
  }

  const munTable = [
    ["Município", "Qtd Guias", "Valor (R$)"],
    ...munArr
  ];

  // 🔥 escreve corretamente abaixo do título
  dash.getRange(startRowDashboard + 1, 1, munTable.length, 3)
    .setValues(munTable);

  // 🔥 formatação correta (evita % bugado)
  dash.getRange(startRowDashboard + 2, 2, munArr.length, 1)
    .setNumberFormat("0"); // qtd

  dash.getRange(startRowDashboard + 2, 3, munArr.length, 1)
    .setNumberFormat("R$ #,##0.00"); // valor

  const munChartRange = dash.getRange(startRow + 2, 1, munArr.length, 2);
  const munChartData = munArr.map(r => [r[0], r[1], r[2]]);


  // =========================
  // 🧠 INSIGHT EXECUTIVO
  // =========================
  // 🔥 TEXTO (sem o título dentro)
  const insightTexto =
    "• " + (percPL * 100).toFixed(1) + "% conciliado automaticamente do que foi provisionado\n" +
    "• " + (percSistema * 100).toFixed(1) + "% cobertura do sistema\n" +
    "• R$ " + (emAberto / 1000).toFixed(0) + "K em aberto no sistema\n\n" +
    "⚠️ Alerta:\n" +
    "Esta análise considera apenas o impacto financeiro das guias conciliadas.\n" +
    "Não reflete eficiência operacional — guias não conciliadas com alto valor podem distorcer a leitura.";

  // 🔥 HEADER (linha separada)
  dash.getRange("A10:B10").merge()
    .setValue("📌 Insight Executivo")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground(CINZA)
    .setFontSize(12)
    .setFontColor("#1F3B5C")
    .setBorder(true, true, true, true, false, false, BORDA, SpreadsheetApp.BorderStyle.SOLID);

  // 🔥 CONTEÚDO
  dash.getRange("A11:B16").merge()
    .setValue(insightTexto)
    .setWrap(true)
    .setBackground(CINZA)
    .setBorder(true, true, true, true, false, false, BORDA, SpreadsheetApp.BorderStyle.SOLID)
    .setVerticalAlignment("top");

  const startDiag = 18;

  dash.getRange(startDiag, 1, 1, 5).merge()
    .setValue("🚨 Divergências e Erros Mapeados")
    .setFontWeight("bold")
    .setBackground("#FDECEA")
    .setHorizontalAlignment("center")   // 🔥 centraliza horizontal
    .setVerticalAlignment("middle");    // 🔥 centraliza vertical;

  const divergencias = {};

  data.forEach(r => {
    const tipoErro = r[idx["Tipo Erro"]] || "OUTROS";
    const valor = parseFloat(r[idx["Diferença"]]) || 0;

    if (!divergencias[tipoErro]) {
      divergencias[tipoErro] = { qtd: 0, valor: 0 };
    }

    divergencias[tipoErro].qtd++;
    divergencias[tipoErro].valor += Math.abs(valor);
  });

  const tabela = [["Tipo Erro (PL)", "Qtd", "", "Impacto (R$)"]];

  Object.entries(errosPL).forEach(([k, v]) => {
    tabela.push([k, v.qtd, "", v.valor]);
  });

  // escreve tabela
  dash.getRange(startDiag + 1, 1, tabela.length, 4)
    .setValues(tabela);

  // formata header
  dash.getRange(startDiag + 1, 1, 1, 5)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#FDECEA")

  // formata números
  dash.getRange(startDiag + 2, 2, tabela.length - 1, 1)
    .setNumberFormat("0");

  dash.getRange(startDiag + 2, 4, tabela.length - 1, 1)
    .setNumberFormat("R$ #,##0.00");

  dash.insertChart(
    dash.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(dash.getRange(startDiag + 2, 1, tabela.length, 4))
      .setPosition(startDiag, 7, 0, 0)
      .setOption("title", "Distribuição de Erros")
      .setOption("width", 600)
      .setOption("height", 350)
      .build()
  );


}


function classificarMatchExecutivo(tipo) {

  if (tipo === "TOTAL") return "MATCH_PERFEITO";

  if ([
    "COMBINACAO",
    "COMP AJUSTADA (TOTAL)",
    "COMP AJUSTADA (COMB)"
  ].includes(tipo)) {
    return "MATCH_COM_AJUSTE";
  }

  if ([
    "LN + COMP"
  ].includes(tipo)) {
    return "MATCH_FRACO";
  }

  return "SEM_MATCH";
}

function gerarChavesVariantes(key) {
  const [emp, ln, mun, comp, sist] = key.split("|");

  return [
    [emp, ln, mun, comp, sist].join("|"),   // completo
    [emp, ln, mun, comp, ""].join("|"),     // sem sistema
    [emp, ln, mun, "", ""].join("|"),       // sem competência
    [emp, ln, "", comp, ""].join("|"),      // LN + COMP
    [emp, ln, "", "", ""].join("|")         // só LN
  ];
}

function encontrarCombinacao(lista, alvo, tolerancia, usadosGlobal) {
  const n = lista.length;

  if (n > 18) return simularConsumo(lista, alvo, usadosGlobal);
  for (let i = 0; i < (1 << n); i++) {
    if (lista.length > 18) {
      return simularConsumo(lista, alvo, usadosGlobal);
    }
    let soma = 0;
    let usados = [];

    for (let j = 0; j < n; j++) {
      if (i & (1 << j)) {
        soma += lista[j].v;
        usados.push(lista[j]);
      }
    }

    if (Math.abs(soma - alvo) <= tolerancia) {
      return { soma, itens: usados };
    }

  }

  return null;
}

function normalizeComp(c) {
  if (!c) return normalizarCompetenciaFallback(s);

  const s = String(c).trim();

  const [m, a] = s.split("/");

  return ("0" + Number(m)).slice(-2) + "/" + a;
}


function simularConsumo(lista, valorAlvo, usadosTemp, tolerancia = 0.01) {

  const disponiveis = lista.filter(v => !usadosTemp.has(v.id));

  if (disponiveis.length === 0) return null;

  if (!Array.isArray(lista) || lista.length === 0) return null;

  const listaDisponivel = lista.filter(v => !usadosTemp.has(v.id));

  if (listaDisponivel.length === 0) return null;

  const candidatos = listaDisponivel
    .sort((a, b) => Math.abs(b.v - valorAlvo) - Math.abs(a.v - valorAlvo))
    .slice(0, 15);

  let soma = 0;
  const usados = [];

  const grupos = new Set();

  for (let item of candidatos) {

    // 🔥 NÃO mistura grupos diferentes
    if (grupos.size > 0 && !grupos.has(item.grupo)) {
      continue;
    }

    grupos.add(item.grupo);

    if (soma + item.v <= valorAlvo + tolerancia) {
      soma += item.v;
      usados.push(item);
    }

    if (Math.abs(soma - valorAlvo) <= tolerancia + 0.05) {
      return {
        soma,
        itens: usados
      };
    }
  }

  return null;
}

function tentarMatch(key, valorAlvo, indexPL, tolerancia, usadosTemp) {

  const obj = indexPL[key];
  if (!obj) return null;

  let lista = Array.isArray(obj)
    ? obj
    : obj.lista || obj.valores;

  if (!Array.isArray(lista)) return null;

  const disponiveis = lista.filter(v => !usadosTemp.has(v.id));
  if (disponiveis.length === 0) return null;

  const soma = disponiveis.reduce((s, v) => s + v.v, 0);

  // MATCH DIRETO
  if (Math.abs(soma - valorAlvo) <= tolerancia) {
    return { soma, itens: disponiveis, chaveReal: key };
  }

  // COMBINAÇÃO
  const comb = simularConsumo(disponiveis, valorAlvo, usadosTemp);

  if (comb) {
    return {
      soma: comb.soma,
      itens: comb.itens,
      chaveReal: key
    };
  }

  return null;
}



function consumirEMontar(match, tipo, chaveReal, usadosGlobal, valorGuepardo) {
  const docs = [];
  let somaConsumida = 0;

  const ss = SpreadsheetApp.getActive();
  mostrarDimMunicipio(ss);
  const spreadsheetId = ss.getId();
  const gidPL = ss.getSheetByName(ABA_PL).getSheetId();

  // 🔥 FILTRA ANTES DE CONSUMIR
  const itensValidos = match.itens.filter(item => !usadosGlobal.has(item.id));

  if (itensValidos.length === 0) {
    return {
      valor: 0,
      docs: [],
      tipo: "SEM MATCH",
      chave: null,
      itens: []
    };
  }

  for (let item of itensValidos) {
    const rowReal = item.rowIndex + 3;

    const link = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${gidPL}&range=A${rowReal}`;

    usadosGlobal.add(item.id);

    docs.push({
      text: String(item.doc || ""),
      url: link,
      rowPL: item.rowIndex,
      rowG: match.rowGuepardo // 👈 NOVO
    });

    for (let item of itensValidos) {
      const key = item.emp + "|" + item.ln + "|" + item.comp;

      if (!docGlobalMap[key]) {
        docGlobalMap[key] = new Set();
      }

      docGlobalMap[key].add(item.doc);
    }

    somaConsumida += item.v;

  }

  return {
    valor: somaConsumida,
    docs,
    tipo,
    chave: chaveReal,
    itens: itensValidos,
    diff: +(valorGuepardo - somaConsumida).toFixed(2)
  };
}


function validarLNReal(ln) {
  return ln && ln !== "0000" && ln.length >= 4;
}

function normalizarCompetenciaFallback(s) {
  const numeros = s.replace(/\D/g, "");

  if (numeros.length === 6) {
    return numeros.slice(0, 2) + "/" + numeros.slice(2);
  }

  if (numeros.length === 8) {
    return numeros.slice(2, 4) + "/" + numeros.slice(4);
  }

  return "SEM_COMP";
}

function similarity(a, b) {
  const dist = levenshtein(a, b);
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  return 1 - dist / maxLen;
}

function levenshtein(a, b) {
  const matrix = [];

  for (let i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }

  for (let j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }

  return matrix[b.length][a.length];
}

function carregarMunicipios() {
  if (MUNICIPIOS) return MUNICIPIOS;

  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("DIM_MUNICIPIO");

  const data = sheet.getDataRange().getValues();
  data.shift(); // remove cabeçalho

  const lista = [];

  data.forEach(r => {
    const nome = normMunicipioFull(r[0]);
    if (nome) lista.push(nome);
  });

  MUNICIPIOS = lista;
  return lista;
}



function processarGuepardo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaG = ss.getSheetByName(ABA_G);
  const abaPL = ss.getSheetByName(ABA_PL);
  const abaOut = ss.getSheetByName(ABA_OUT);

  if (!abaG || !abaPL || !abaOut) {
    SpreadsheetApp.getUi().alert("Erro: Certifique-se que as abas Guepardo, PL e Base Conciliacao existem.");
    return;
  }

  // Limpa a aba de saída (mantendo o cabeçalho)
  const lastRowOut = abaOut.getLastRow();
  if (lastRowOut > 1) {
    abaOut.getRange(2, 1, lastRowOut - 1, abaOut.getLastColumn()).clearContent();
  }

  // Carrega dados do PL para indexação (melhora performance)
  const PL = criarSchemaDinamico(abaPL);

const dadosPL = PL.dados;
const schemaPL = PL.schema;
  const indexPL = {}; // Chave: LN + Competência + Valor

  const cabecalhoPL = dadosPL[0];

const schemaPLBase = criarSchema(cabecalhoPL);


const schemaPL = {

  EMP: localizarColuna(schemaPLBase,
  ["EMPRESA","EMP","COD_EMPRESA"]),

  LN: localizarColuna(schemaPLBase,
  ["LN","LOCAL_NEGOCIO","LOCAL DE NEGOCIO"]),

  COMP: localizarColuna(schemaPLBase,
  ["COMPETENCIA","PERIO_APUR","PERIO APUR"]),

  MUN: localizarColuna(schemaPLBase,
  ["MUNICIPIO","MUNICÍPIO"]),

  VALOR: localizarColuna(schemaPLBase,
  ["VALOR","VALOR_ISS","ISS"])

};


Logger.log("SCHEMA PL");
Logger.log(schemaPL);

  // Pula cabeçalho do PL (considerando linha 2 como início dos dados úteis)
  for (let i = 2; i < dadosPL.length; i++) {
    const linhaPL = dadosPL[i];
    const valorTotalPL = getValorPL(linhaPL, PL);
    const lnPL = normLN(
  linhaPL[schemaPL.LN]
);


const compPL = String(
  linhaPL[schemaPL.COMP]
).trim();

    const chave = `${lnPL}_${compPL}_${valorTotalPL.toFixed(2)}`;

Logger.log("EXEMPLO CHAVE PL");
Logger.log(chave);

indexPL[chave] = linhaPL;
  }

  // Processamento do Guepardo
  const G = criarSchemaDinamico(abaG);

const dadosG = G.dados;
const schemaG = G.schema;
  const resultados = [];

  const cabecalhoG = dadosG[0];

const schemaGBase = criarSchema(cabecalhoG);


const schemaG = {

  EMP: localizarColuna(schemaGBase,
  ["EMPRESA","EMP","COD_EMPRESA"]),


  LN: localizarColuna(schemaGBase,
  ["LN","LOCAL_NEGOCIO","LOCAL DE NEGOCIO"]),


  MUN: localizarColuna(schemaGBase,
  ["MUNICIPIO","MUNICÍPIO","Munic. Recolhimento"]),


  COMP: localizarColuna(schemaGBase,
  ["COMPETENCIA","PERIO_APUR"]),


  VALOR: localizarColuna(schemaGBase,
  ["VALOR","ISS","VALOR_ISS","Valor ISS"])

};


Logger.log("SCHEMA GUEPARDO");
Logger.log(schemaG);

  // Início do loop (pula cabeçalho)
  for (let i = 2; i < dadosG.length; i++) {
    const r = dadosG[i];

    // --- FILTRO DE STATUS ---
    // Supondo que 'r' é a linha atual e Column V (índice 21) é o Status
    const statusEscrituracao = String(r[21]).trim();

    if (statusEscrituracao !== "Escriturado") {
      continue; // Pula linhas que não estão prontas para pagamento
    }

    // Captura e normalização da LN
    const lnFormatada = normLN(r[5]); // Coluna LN (Índice 5)
    const competencia = String(r[1]).trim(); // Competência NF (Índice 1)
    const valorISS = parseMoney(r[18]); // Valor ISS (Índice 18)
    const municipio = String(r[6]).trim(); // Munic. Recolhimento (Índice 6)

    // Lógica de Match Simples para a Base de Conciliação
    const chaveBusca = `${lnFormatada}_${competencia}_${valorISS.toFixed(2)}`;
    const matchPL = indexPL[chaveBusca];

    let statusMatch = "SEM PAGAMENTO";
    let docPL = "";

    if (matchPL) {
      statusMatch = "CONCILIADO";
      docPL = matchPL[19]; // Doc. Guia (Índice 19 no PL)
    }

    // Monta linha para a aba "Base Conciliacao ISS"
    // Estrutura sugerida: Empresa, LN, Competência, Município, Valor, Doc Guepardo, Status, Doc PL
    resultados.push([
      r[3],           // Empresa
      lnFormatada,    // LN formatada
      competencia,    // Período
      municipio,      // Município
      valorISS,       // Valor no Sistema
      r[27],          // Doc Contábil (Guepardo)
      statusMatch,    // Resultado da conciliação
      docPL           // Documento no PL
    ]);
  }

  // Grava os resultados na aba de saída
  if (resultados.length > 0) {
    abaOut.getRange(2, 1, resultados.length, resultados[0].length).setValues(resultados);
  }

  Logger.log("Processamento concluído: " + resultados.length + " linhas escrituradas processadas.");
}


function encontrarMatch(itemGuepardo, indexPL) {
  const lnG = normLN(itemGuepardo.ln);
  const munG = itemGuepardo.municipio;
  const valorG = itemGuepardo.valor;

  // 1. Busca no index do PL usando a LN formatada como chave primária 
  const candidatos = indexPL[lnG] || [];

  if (candidatos.length === 0) {
    return {
      tipo: "SEM_MATCH",
      motivo: "❌ LN não encontrada no PL",
      valorMatch: 0
    };
  }

  // 2. Filtro e Priorização de Candidatos 
  // Tentamos encontrar um candidato que coincida em Valor e Município
  let matchPerfeito = candidatos.find(c =>
    Math.abs(c.total - valorG) < 0.05 && (c.municipio === munG)
  );

  if (matchPerfeito) {
    return {
      tipo: "MATCH_TOTAL",
      dados: matchPerfeito,
      valorMatch: matchPerfeito.total
    };
  }

  // 3. Tratamento de Município Secundário (Caso {G=TO, PL=NA}) 
  // Se não houver match perfeito, busca por valor ignorando o município se o PL for "NA"
  let matchValorApenas = candidatos.find(c =>
    Math.abs(c.total - valorG) < 0.05 && (c.municipio === "NA" || !c.municipio)
  );

  if (matchValorApenas) {
    return {
      tipo: "MATCH_VALOR_MUN_NA",
      dados: matchValorApenas,
      valorMatch: matchValorApenas.total,
      obs: "⚠️ Município no PL consta como NA, mas LN e Valor coincidem"
    };
  }

  // 4. Identificação de Divergência de Valor 
  // Se a LN existe mas os valores não batem
  return {
    tipo: "DIVERGENTE",
    motivo: "💰 Diferença de valor na LN encontrada",
    valorMatch: candidatos[0].total // Retorna o valor do primeiro card encontrado para análise
  };
}

function validarDuplicidadePL(listaPLUnica) {
  const docsMap = {};
  const duplicados = [];

  listaPLUnica.forEach(item => {
    if (!item.doc) return;

    if (!docsMap[item.doc]) {
      docsMap[item.doc] = [];
    }

    docsMap[item.doc].push(item);
  });

  Object.entries(docsMap).forEach(([doc, itens]) => {
    if (itens.length > 1) {
      itens.forEach(i => i.duplicado = true);
      duplicados.push({
        doc,
        qtd: itens.length,
        ids: itens.map(i => i.id)
      });

      Logger.log("🚨 DOC DUPLICADO NO PL:");
      Logger.log({ doc, qtd: itens.length });
    }
  });

  return duplicados;
}

function classificarRelacao(matchItens) {
  if (!matchItens || matchItens.length === 0) return "SEM_MATCH";

  if (matchItens.length === 1) return "1:1";

  if (matchItens.length > 1) return "N:1";

  return "ERRO";
}

function calcularStatsPL(listaPLUnica, usadosGlobal) {
  const total = listaPLUnica.length;

  const conciliados = listaPLUnica.filter(v =>
    usadosGlobal.has(v.id)
  );

  const naoConciliados = listaPLUnica.filter(v => !usadosGlobal.has(v.id));

  const totalConciliado = conciliados.reduce((s, v) => s + v.v, 0);
  const totalNaoConciliado = naoConciliados.reduce((s, v) => s + v.v, 0);

  return {
    totalQtd: total,
    conciliadoQtd: conciliados.length,
    naoConciliadoQtd: naoConciliados.length,
    totalConciliado,
    totalNaoConciliado
  };
}

function inicializarStatsEstruturado() {
  return {
    G: {
      total: 0,
      match: 0,
      semMatch: 0
    },
    PL: {
      total: 0,
      conciliado: 0,
      naoConciliado: 0
    }
  };
}


function calcularTolerancia(valor) {
  return Math.max(
    0.05,
    valor * 0.05,   // 🔥 sobe pra 5%
    20              // 🔥 piso realista pra multa
  );
}

function extrairCompetencia(data) {
  if (!data) return null;

  const d = new Date(data);
  if (isNaN(d)) return null;

  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM");
}

function municipioSimilar(a, b) {
  if (!a || !b) return false;

  a = normalizarMunicipioComparacao(a);
  b = normalizarMunicipioComparacao(b);

  if (a === b) return true;

  // contém (ex: "SAO PAULO" dentro de "PREF SAO PAULO")
  if (a.includes(b) || b.includes(a)) return true;

  const score = similarity(a, b);

  return score >= 0.65; // mais tolerante
}

function montarRichText(docs) {
  docs = docs.map(d => ({
    text: d.text || "DOC",
    url: d.url || ""
  }));

  const separador = " | ";

  let textoFinal = "";
  const ranges = [];

  let cursor = 0;

  docs.forEach((d, i) => {

    const text = String(d.text || "");

    if (i > 0) {
      textoFinal += separador;
      cursor += separador.length;
    }

    const start = cursor;
    const end = start + text.length;

    textoFinal += text;

    ranges.push({
      start,
      end,
      url: d.url
    });

    cursor = end;
  });

  const builder = SpreadsheetApp.newRichTextValue()
    .setText(textoFinal);

  ranges.forEach(r => {
    if (r.url) {
      builder.setLinkUrl(r.start, r.end, r.url);
    }
  });

  return builder.build();
}

function linhaValidaPL(r) {

  const valor = getValorPL(r);
  const ln = normLN(r[5]);

  if (valor <= 0) return false;
  if (ln === "0000") return false;

  return true;
}

function executarMatchComErro(gRow, indexPL_lnValor, usadosGlobal) {
  const keyLN = gRow.empresa + "|" + gRow.ln;
  const valor = Number(gRow.valor.toFixed(2));
  const tolerancia = calcularTolerancia(valor);

  // =========================
  // 🔎 Pool mais seguro
  // =========================
  const pool = (indexPL_lnValor[keyLN]?.lista || [])
    .filter(v => !usadosGlobal.has(v.id));

  if (!pool.length) return null;

  // =========================
  // 🧠 Ordena por proximidade (MELHORA MUITO o match)
  // =========================
  pool.sort((a, b) =>
    Math.abs(a.v - valor) - Math.abs(b.v - valor)
  );

  // =========================
  // 🔥 1. Combinação inteligente
  // =========================
  const matchComb = encontrarCombinacao(
    pool,
    valor,
    tolerancia + 0.05, // mais justo
    usadosGlobal
  );

  if (matchComb) {
    return consumirEMontar(
      { soma: matchComb.soma, itens: matchComb.itens },
      "MATCH_COM_ERRO_COMBINADO",
      keyLN,
      usadosGlobal,
      valor
    );
  }

  // =========================
  // 🎯 2. Match direto (valor quase igual)
  // =========================
  const candidatoProximo = pool.find(v =>
    Math.abs(v.v - valor) <= tolerancia
  );

  if (candidatoProximo) {
    return consumirEMontar(
      { soma: candidatoProximo.v, itens: [candidatoProximo] },
      "MATCH_COM_ERRO_DIRETO",
      keyLN,
      usadosGlobal,
      valor
    );
  }

  // =========================
  // ⚠️ 3. Fallback final CONTROLADO
  // =========================
  const fallback = pool.find(v =>
    Math.abs(v.v - valor) <= tolerancia * 2
  );

  if (fallback) {
    return consumirEMontar(
      { soma: fallback.v, itens: [fallback] },
      "MATCH_FORCADO",
      keyLN,
      usadosGlobal,
      valor
    );
  }

  return null;
}

function removerAcentos(texto) {

  if (!texto) return "";

  return texto
    .toString()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

}

function criarSchemaDinamico(aba){

  const dados = aba.getDataRange().getValues();

  const cab = dados[0];

  const schema = {};

  cab.forEach((titulo,index)=>{

    const nome = String(titulo || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g,"")
      .toUpperCase()
      .trim();

    schema[nome] = index;

  });


  return {
    dados:dados,
    schema:schema
  };

}




27.06 - 18:31
/*Teste27/06 1.0
*/
