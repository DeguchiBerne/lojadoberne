function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var action = e.parameter.action;
  var output = {};
  
  try {
    if (action === 'getData') {
      output = getCatalogData();
    } else if (action === 'saveBatch') {
      var postData = "";
      if (e.postData && e.postData.contents) {
        postData = JSON.parse(e.postData.contents);
      } else if (e.parameter.data) {
        postData = JSON.parse(e.parameter.data);
      }
      output = saveBatchData(postData.password, postData.items);
    } else {
      output = { status: "API Ativa" };
    }
  } catch (err) {
    output = { success: false, message: "Erro interno no servidor." };
  }
  
  return ContentService.createTextOutput(JSON.stringify(output))
      .setMimeType(ContentService.MimeType.JSON);
}

function getCatalogData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  
  var colCat = headers.indexOf("Categoria");
  var colNome = headers.indexOf("Nome do Item");
  var colDesc = headers.indexOf("Descrição");
  var colPreco = headers.indexOf("Preço Sugerido");
  var colLink = headers.indexOf("Link da foto");
  if (colLink === -1) colLink = headers.indexOf("Link da Foto");
  var colArquivo = headers.indexOf("Nome do arquivo");
  var colStatus = headers.indexOf("STATUS");
  if (colStatus === -1) colStatus = headers.indexOf("Status");

  if (colCat === -1) colCat = 0;
  if (colNome === -1) colNome = 1;
  if (colDesc === -1) colDesc = 2;
  if (colPreco === -1) colPreco = 3;
  if (colLink === -1) colLink = 4;
  if (colArquivo === -1) colArquivo = 5;

  var items = [];
  
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var nome = row[colNome];
    if (!nome) continue;
    
    var desc = String(row[colDesc] || "");
    var statusVal = colStatus !== -1 ? String(row[colStatus] || "").trim() : "";
    var vendido = statusVal.toUpperCase() === "VENDIDO" || desc.includes("[VENDIDO]");
    
    // O backend agora envia apenas a string crua. O front-end processará as imagens.
    var rawLinks = String(row[colLink] || "");

    items.push({
      rowIndex: i + 1,
      categoria: row[colCat],
      nome: nome,
      descricao: desc.replace("[VENDIDO]", "").trim(),
      preco: row[colPreco],
      linkRaw: rawLinks,
      arquivo: row[colArquivo],
      status: vendido ? "VENDIDO" : ""
    });
  }
  
  return items;
}

function saveBatchData(password, items) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var adminPassword = scriptProperties.getProperty("senha_admin");
  
  // MENSAGEM EXATA SOLICITADA PARA ERRO DE SENHA
  if (!adminPassword || password !== adminPassword) {
    return { success: false, message: "Erro na senha, não tente novamente." };
  }
  
  if (!items || !items.length) {
    return { success: false, message: "Nenhum dado enviado." };
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var colCat = headers.indexOf("Categoria") + 1;
  var colNome = headers.indexOf("Nome do Item") + 1;
  var colDesc = headers.indexOf("Descrição") + 1;
  var colPreco = headers.indexOf("Preço Sugerido") + 1;
  var colLink = headers.indexOf("Link da foto") + 1;
  if (colLink === 0) colLink = headers.indexOf("Link da Foto") + 1;
  var colArquivo = headers.indexOf("Nome do arquivo") + 1;
  var colStatus = headers.indexOf("STATUS") + 1;
  if (colStatus === 0) colStatus = headers.indexOf("Status") + 1;

  for (var k = 0; k < items.length; k++) {
    var it = items[k];
    var r = Number(it.rowIndex);
    if (!r || r < 2) continue;
    
    var isVendido = String(it.status).toUpperCase().trim() === "VENDIDO";
    var descFinal = String(it.descricao || "").replace("[VENDIDO]", "").trim();
    if (isVendido) descFinal = "[VENDIDO] " + descFinal;
    
    if (colCat > 0) sheet.getRange(r, colCat).setValue(it.categoria);
    if (colNome > 0) sheet.getRange(r, colNome).setValue(it.nome);
    if (colDesc > 0) sheet.getRange(r, colDesc).setValue(descFinal);
    if (colPreco > 0) sheet.getRange(r, colPreco).setValue(it.preco);
    if (colLink > 0) sheet.getRange(r, colLink).setValue(it.linkRaw);
    if (colArquivo > 0) sheet.getRange(r, colArquivo).setValue(it.arquivo);
    if (colStatus > 0) sheet.getRange(r, colStatus).setValue(isVendido ? "VENDIDO" : "");
  }
  
  return { success: true, message: "OK" }; // O front-end vai mostrar o ERRO 404 ao ler success: true
}
