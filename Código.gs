function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Scanner") return;
  if (e.range.getA1Notation() !== "A2") return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const produtos = ss.getSheetByName("Produtos");

  let codigo = e.value;
  if (!codigo) return;

  codigo = codigo.toString().trim();
  const tipoEnvio = sheet.getRange("B2").getValue() || "Não definido";
  const agora = new Date();

  const dados = produtos.getDataRange().getValues();

  let encontrados = [];

  // 🔎 PROCURA TODAS OCORRÊNCIAS DO CÓDIGO
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0].toString().trim() === codigo) {
      encontrados.push({
        linha: i + 1,
        status: dados[i][2]
      });
    }
  }

  // ============================
  // 🚫 SE JÁ EXISTE MAIS DE UMA VEZ → ERRO
  // ============================
  if (encontrados.length > 1) {
    sheet.getRange("C2").setValue("❌ DUPLICADO NA BASE");
    sheet.getRange("A2").clearContent().activate();
    return;
  }

  // ============================
  // ✅ SE EXISTE 1 VEZ
  // ============================
  if (encontrados.length === 1) {
    const item = encontrados[0];

    if (item.status === "Coletado") {
      // 🔴 BLOQUEIA NOVO BIPE
      sheet.getRange("C2").setValue("⚠ JÁ COLETADO");
      sheet.getRange("A2").clearContent().activate();
      return;
    }

    if (item.status === "Pronto") {
      // ✅ MARCA COMO COLETADO
      produtos.getRange(item.linha, 3).setValue("Coletado");
      produtos.getRange(item.linha, 5).setValue(agora);

      sheet.getRange("C2").setValue("✔ COLETADO");
      sheet.getRange("A2").clearContent().activate();
      return;
    }
  }

  // ============================
  // 🆕 SE NÃO EXISTE → CRIA
  // ============================
  produtos.appendRow([codigo, tipoEnvio, "Pronto", agora, ""]);

  sheet.getRange("C2").setValue("📦 PRONTO PARA ENVIO");
  sheet.getRange("A2").clearContent().activate();
}
// ===============================
// 2️⃣ CONTADORES GERAIS (SITE)
// ===============================
function getContadoresGerais() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const produtos = ss.getSheetByName("Produtos");
  const dados = produtos.getDataRange().getValues();

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  let contadores = {
    pronto: { correios: 0, transportadora: 0 },
    coletadoHoje: { correios: 0, transportadora: 0 }
  };

  for (let i = 1; i < dados.length; i++) {
    const tipo = dados[i][1];
    const status = dados[i][2];
    const dataColeta = dados[i][4];

    if (status === "Pronto") {
      if (tipo === "Correios") contadores.pronto.correios++;
      if (tipo === "Transportadora") contadores.pronto.transportadora++;
    }

    if (status === "Coletado" && dataColeta instanceof Date) {
      const d = new Date(dataColeta);
      d.setHours(0, 0, 0, 0);
      if (d.getTime() === hoje.getTime()) {
        if (tipo === "Correios") contadores.coletadoHoje.correios++;
        if (tipo === "Transportadora") contadores.coletadoHoje.transportadora++;
      }
    }
  }

  return contadores;
}

// ===============================
// 3️⃣ API PROCESSAMENTO
// ===============================
function receberCodigo(codigo, tipoEnvio) {
  if (!codigo) return "Código vazio";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const produtos = ss.getSheetByName("Produtos");
  const dados = produtos.getDataRange().getValues();
  const agora = new Date();

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0].toString().trim() === codigo) {
      if (dados[i][2] === "Pronto") {
        produtos.getRange(i + 1, 3).setValue("Coletado");
        produtos.getRange(i + 1, 5).setValue(agora);
        return "✔ Coletado";
      }
      return "⚠ Já coletado";
    }
  }

  produtos.appendRow([codigo, tipoEnvio, "Pronto", agora, ""]);
  return "📦 Pronto para envio";
}

// ===============================
// 4️⃣ RELATORIOS
// ===============================
function getRelatorios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const produtos = ss.getSheetByName("Produtos");
  const dados = produtos.getDataRange().getValues();
  const hoje = new Date();
  let relatorios = [];

  for (let i = 1; i < dados.length; i++) {
    const dataPronto = dados[i][3];
    const dataColetado = dados[i][4];

    let status = dados[i][2];
    let icone = "🟡";

    if (status === "Pronto" && dataPronto instanceof Date) {
      const diff = (hoje - dataPronto) / 86400000;
      if (diff >= 1) {
        status = "Pendente";
        icone = "🔴";
      }
    }

    if (dados[i][2] === "Coletado") icone = "🟢";

    relatorios.push({
      codigo: dados[i][0],
      tipo: dados[i][1],
      status,
      icone,
      pronto: dataPronto instanceof Date
        ? Utilities.formatDate(dataPronto, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
        : "",
      coletado: dataColetado instanceof Date
        ? Utilities.formatDate(dataColetado, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
        : ""
    });
  }

  return relatorios.reverse();
}

// ===============================
// 5️⃣ DEVOLUCAO SHOPEE - FOTOS
// ===============================
const PASTA_DEVOLUCAO = "1M0119JX-qJ7XpePfpLDYsqs_B50BAJXY";

function salvarFoto(dataUrl, nomeArquivo) {
  const pasta = DriveApp.getFolderById(PASTA_DEVOLUCAO);

  const blob = Utilities.newBlob(
    Utilities.base64Decode(dataUrl.split(",")[1]),
    "image/png",
    nomeArquivo
  );

  const arquivo = pasta.createFile(blob);

  // mantém privado (uso interno)
  return arquivo.getUrl();
}

function getFotosDevolucao() {
  const pasta = DriveApp.getFolderById(PASTA_DEVOLUCAO);
  const arquivos = pasta.getFiles();
  const hoje = new Date();
  const tz = Session.getScriptTimeZone();

  hoje.setHours(0, 0, 0, 0);

  let lista = [];
  let contadorHoje = 0;

  while (arquivos.hasNext()) {
    const f = arquivos.next();
    const data = f.getDateCreated();
    const d = new Date(data);
    d.setHours(0, 0, 0, 0);

    // 🔹 SOMENTE HOJE
    if (d.getTime() !== hoje.getTime()) continue;

    contadorHoje++;

    lista.push({
      nome: f.getName(),
      url: f.getUrl(),
      data: Utilities.formatDate(data, tz, "dd/MM/yyyy"),
      hora: Utilities.formatDate(data, tz, "HH:mm:ss")
    });
  }

  // mais recentes primeiro
  lista.sort((a, b) => b.hora.localeCompare(a.hora));

  return {
    contadorHoje: contadorHoje,
    fotos: lista
  };
}


// ===============================
// 6️⃣ WEB APP
// ===============================
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("Conferência de Pacotes")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
