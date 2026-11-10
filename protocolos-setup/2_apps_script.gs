// ============================================================
// PROTOCOLOS -> SUPABASE
//
// Como instalar (não precisa apagar nada do script atual):
//   1. No editor do Apps Script, clique no "+" ao lado de "Arquivos" > Script,
//      dê o nome "ProtocolosSupabase" e cole TODO este conteúdo.
//   2. No arquivo principal, dentro da função syncAll, logo abaixo da linha
//        conferirProtocolos().forEach(l => { resultados.push(l); Logger.log(l); });
//      acrescente esta linha:
//        enviarProtocolosSupabase().forEach(l => { resultados.push(l); Logger.log(l); });
//   3. Salve e execute syncAll uma vez.
//
// O que faz: envia ao Supabase a extração MAIS RECENTE da aba PROTOCOLOS e o conteúdo
// das abas SALDO PROTOCOLOS e MOVIMENTO PROTOCOLOS. O nome do agente vai como o nome
// curto da aba DE-PARA, mesmo que na aba PROTOCOLOS esteja o nome completo do sistema. Só reenvia quando algo mudou na
// planilha, para não regravar ~22 mil linhas a cada hora.
// Para forçar um reenvio, execute a função reenviarProtocolosSupabase.
// ============================================================

const PROTO_SUPABASE = {
  tabelaProtocolos: "protocolos",
  tabelaSaldo: "protocolos_saldo",
  tabelaMovimento: "protocolos_movimento",
  propriedade: "PROTO_SUPABASE_ASSINATURA",
  lote: 1000
};

function enviarProtocolosSupabase(forcar) {
  try {
    if (!SERVICE_ROLE_KEY) return ["ERRO em Protocolos (Supabase): chave SERVICE_ROLE_KEY não cadastrada"];
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const dia = v => { const d = _paraData(v); return d ? Utilities.formatDate(d, tz, "yyyy-MM-dd") : null; };
    const diaHora = v => { const d = _paraData(v); return d ? Utilities.formatDate(d, tz, "yyyy-MM-dd'T'HH:mm:ss") : null; };
    const texto = v => (v === "" || v === null || v === undefined) ? null : String(v).trim();
    const numero = v => {
      if (v === "" || v === null || v === undefined) return null;
      const t = String(v).trim();
      const n = typeof v === "number" ? v : parseFloat(t.indexOf(",") !== -1 ? t.replace(/\./g, "").replace(",", ".") : t);
      return isNaN(n) ? null : n;
    };
    const inteiro = v => { const n = numero(v); return n === null ? null : Math.round(n); };
    // Troca o nome que vem do sistema pelo nome curto da aba DE-PARA (coluna A -> coluna B).
    // Se a aba não existir ou o nome não estiver nela, mantém o nome como está.
    const curtos = _protoNomesCurtos();
    const agente = v => { const t = texto(v); return t === null ? null : (curtos[_protoChaveNome(t)] || t.replace(/\s+/g, " ")); };

    // ---- aba PROTOCOLOS: só a extração mais recente ----
    const abaDados = findSheetCaseInsensitive(PROTOCOLOS.abaDados);
    if (!abaDados) return ["ERRO em Protocolos (Supabase): aba não encontrada: " + PROTOCOLOS.abaDados];
    const dados = _protoLerAba(abaDados);
    if (!dados.linhas.length) return ["OK: Protocolos (Supabase): aba vazia, nada enviado"];
    ["data_extracao", "numero_protocolo", "status", "criado_em", "agente", "setor", "categoria"].forEach(c => {
      if (dados.idx[c] === undefined) throw new Error("coluna não encontrada na aba " + PROTOCOLOS.abaDados + ": " + c);
    });
    let ultima = null;
    dados.linhas.forEach(r => { const d = dia(r[dados.idx.data_extracao]); if (d && (!ultima || d > ultima)) ultima = d; });
    if (!ultima) return ["ERRO em Protocolos (Supabase): nenhuma data_extracao válida na aba " + PROTOCOLOS.abaDados];

    const vistos = {};
    const protocolos = [];
    dados.linhas.forEach(r => {
      if (dia(r[dados.idx.data_extracao]) !== ultima) return;
      const num = texto(r[dados.idx.numero_protocolo]);
      if (!num || vistos[num]) return;
      vistos[num] = true;
      protocolos.push({
        data_extracao: ultima,
        numero_protocolo: num,
        status: texto(r[dados.idx.status]),
        criado_em: diaHora(r[dados.idx.criado_em]),
        agente: agente(r[dados.idx.agente]),
        setor: texto(r[dados.idx.setor]) || "Sem setor",
        categoria: texto(r[dados.idx.categoria])
      });
    });

    // ---- aba SALDO PROTOCOLOS ----
    const saldo = [];
    const abaSaldo = findSheetCaseInsensitive(PROTOCOLOS.abaSaldo);
    if (abaSaldo) {
      const s = _protoLerAba(abaSaldo);
      const col = nome => s.idx[nome];
      const colPend = Object.keys(s.idx).filter(k => k.indexOf("pendentes_mais_") === 0)[0];
      s.linhas.forEach(r => {
        const d = dia(r[col("data_extracao")]);
        if (!d) return;
        saldo.push({
          data_extracao: d,
          tipo: texto(r[col("tipo")]),
          setor: texto(r[col("setor")]) || "Sem setor",
          saldo_anterior: inteiro(r[col("saldo_anterior")]),
          novos: inteiro(r[col("novos")]),
          transferidos_entrada: inteiro(r[col("transferidos_entrada")]),
          transferidos_saida: inteiro(r[col("transferidos_saida")]),
          fechados: inteiro(r[col("fechados")]),
          saldo_atual: inteiro(r[col("saldo_atual")]),
          pendentes_mais_30_dias: colPend ? inteiro(r[s.idx[colPend]]) : null,
          tempo_medio_em_aberto_dias: numero(r[col("tempo_medio_em_aberto_dias")]),
          taxa_fechamento: numero(r[col("taxa_fechamento")])
        });
      });
    }

    // ---- aba MOVIMENTO PROTOCOLOS ----
    const movimento = [];
    const abaMov = findSheetCaseInsensitive(PROTOCOLOS.abaMovimento);
    if (abaMov) {
      const m = _protoLerAba(abaMov);
      const col = nome => m.idx[nome];
      m.linhas.forEach(r => {
        const d = dia(r[col("data_extracao")]);
        if (!d) return;
        movimento.push({
          data_extracao: d,
          numero_protocolo: texto(r[col("numero_protocolo")]),
          movimento: texto(r[col("movimento")]),
          setor: texto(r[col("setor")]) || "Sem setor",
          setor_anterior: texto(r[col("setor_anterior")]),
          agente: agente(r[col("agente")]),
          categoria: texto(r[col("categoria")]),
          criado_em: diaHora(r[col("criado_em")]),
          dias_em_aberto: inteiro(r[col("dias_em_aberto")])
        });
      });
    }

    // ---- só reenvia se algo mudou desde o último envio ----
    const props = PropertiesService.getScriptProperties();
    const assinatura = _protoAssinatura([protocolos, saldo, movimento]);
    const resumo = "extração " + ultima.split("-").reverse().join("/") + ": " + protocolos.length +
      " protocolos, " + saldo.length + " linhas de saldo, " + movimento.length + " de movimento";
    if (!forcar && props.getProperty(PROTO_SUPABASE.propriedade) === assinatura) {
      return ["OK: Protocolos (Supabase): sem mudança, nada reenviado (" + resumo + ")"];
    }

    // Se falhar no meio, a assinatura não é gravada e a próxima execução tenta de novo.
    props.deleteProperty(PROTO_SUPABASE.propriedade);
    [[PROTO_SUPABASE.tabelaProtocolos, protocolos],
     [PROTO_SUPABASE.tabelaSaldo, saldo],
     [PROTO_SUPABASE.tabelaMovimento, movimento]].forEach(par => {
      deleteAllRows(par[0]);
      if (par[1].length) insertInBatches(par[0], par[1], PROTO_SUPABASE.lote);
    });
    props.setProperty(PROTO_SUPABASE.propriedade, assinatura);
    return ["OK: Protocolos (Supabase): enviado (" + resumo + ")"];
  } catch (e) {
    return ["ERRO em Protocolos (Supabase): " + e.message];
  }
}

// Força o reenvio mesmo sem mudança (útil depois de recriar as tabelas no Supabase).
function reenviarProtocolosSupabase() {
  const linhas = enviarProtocolosSupabase(true);
  linhas.forEach(l => Logger.log(l));
  escreverLog(linhas);
  return linhas;
}

function _protoLerAba(aba) {
  const values = aba.getDataRange().getValues();
  const idx = {};
  if (!values.length) return { idx: idx, linhas: [] };
  values[0].forEach((h, i) => { idx[String(h).trim().toLowerCase()] = i; });
  const linhas = values.slice(1).filter(r => r.some(v => v !== "" && v !== null));
  return { idx: idx, linhas: linhas };
}

function _protoChaveNome(nome) {
  return String(nome).replace(/\s+/g, " ").trim().toLowerCase();
}

// Lê a aba DE-PARA: coluna A = nome no sistema, coluna B = nome curto.
function _protoNomesCurtos() {
  const mapa = {};
  const aba = findSheetCaseInsensitive("DE-PARA");
  if (!aba) return mapa;
  aba.getDataRange().getValues().slice(1).forEach(r => {
    const sistema = String(r[0] === null ? "" : r[0]).trim();
    const curto = String(r[1] === null || r[1] === undefined ? "" : r[1]).trim();
    if (sistema && curto) mapa[_protoChaveNome(sistema)] = curto;
  });
  return mapa;
}

function _protoAssinatura(obj) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(obj), Utilities.Charset.UTF_8);
  return bytes.map(b => ("0" + ((b + 256) % 256).toString(16)).slice(-2)).join("");
}
