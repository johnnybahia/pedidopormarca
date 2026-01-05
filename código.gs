// --- ARQUIVO: Código.gs ---

// 1. O SITE (Para o ser humano ver)
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
      .setTitle('Pedidos por Marca - Marfim Bahia')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 2. A API (Para o Robô Python enviar dados)
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName("Dados"); // Certifique-se que o nome da aba é 'Dados'
    
    if (!sheet) {
      // Se não existir, cria e põe cabeçalho
      sheet = doc.insertSheet("Dados");
      sheet.appendRow(["Data", "Arquivo", "Cliente", "Marca", "Local", "Qtd", "Unidade", "Valor"]);
    }

    var json = JSON.parse(e.postData.contents);
    var lista = json.pedidos; // O Python manda { "pedidos": [...] }
    var novasLinhas = [];
    
    // Verificação simples de duplicidade (olhando ultimos 500 registros para ser rápido)
    var ultimaLinha = sheet.getLastRow();
    var arquivosExistentes = [];
    if (ultimaLinha > 1) {
      // Pega apenas a coluna B (Arquivo)
      var dadosB = sheet.getRange(Math.max(2, ultimaLinha - 500), 2, Math.min(500, ultimaLinha-1), 1).getValues();
      arquivosExistentes = dadosB.map(function(r){ return r[0]; });
    }

    for (var i = 0; i < lista.length; i++) {
      var p = lista[i];
      if (arquivosExistentes.indexOf(p.arquivo) === -1) {
        novasLinhas.push([
          p.data, p.arquivo, p.cliente, p.marca, p.local, p.qtd, p.unidade, p.valor
        ]);
      }
    }

    if (novasLinhas.length > 0) {
      sheet.getRange(ultimaLinha + 1, 1, novasLinhas.length, 8).setValues(novasLinhas);
      return ContentService.createTextOutput(JSON.stringify({"status":"Sucesso", "msg": novasLinhas.length + " novos."})).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({"status":"Neutro", "msg": "Sem novidades."})).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (erro) {
    return ContentService.createTextOutput(JSON.stringify({"status":"Erro", "msg": erro.toString()})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// 3. FUNÇÃO QUE O SITE CHAMA PARA PEGAR DADOS DA PLANILHA
function getDadosPlanilha() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados");
  if (!sheet) return [];
  
  // Pega tudo da linha 2 até o fim
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var dados = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  // Retorna array puro para o Javascript do navegador processar
  return dados; 
}
