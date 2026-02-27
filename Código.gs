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
// 3️⃣ API PROCESSAMENTO (ATUALIZADO)
// ===============================
function receberCodigo(codigo, tipo){

  if (!codigo) return "Código vazio";

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName("Pedidos"); // Aba correta
  const dados = aba.getRange("A2:A").getValues().flat(); // Coluna A (Códigos)

  const linha = dados.indexOf(codigo);

  if(linha === -1){
    return "❌ Código não encontrado";
  }

  const linhaReal = linha + 2; // Porque começa em A2

  const status = aba.getRange(linhaReal, 2).getValue(); // Coluna B = Status

  if(status === "CONFERIDO"){
    return "⚠ Já conferido";
  }

  // Marca como conferido
  aba.getRange(linhaReal, 2).setValue("CONFERIDO"); // Coluna B
  aba.getRange(linhaReal, 3).setValue(new Date());  // Coluna C = Data/Hora

  return "✔ Código Conferido com sucesso";
}
}
function getRelatorios(){
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pedidos");
  const dados = aba.getRange("A2:C").getValues();

  return dados.map(l => ({
    codigo: l[0],
    status: l[1],
    horario: l[2]
  }));
}
