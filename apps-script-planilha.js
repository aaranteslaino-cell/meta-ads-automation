/**
 * Recebe os dados do dashboard Meta Ads e preenche APENAS 3 colunas:
 * "Alunos PP", "Investimento" e "Faturamento".
 * As colunas com fórmula (Lucro, CPA, TKM, ROAS, CPO) NÃO são tocadas —
 * elas se recalculam sozinhas a partir dessas três.
 *
 * COMO INSTALAR (1 vez, ~3 min):
 * 1. Abra a planilha → Extensões → Apps Script
 * 2. Apague o conteúdo e cole este arquivo inteiro
 * 3. Implantar → Nova implantação → tipo "App da Web"
 *      Executar como: Eu
 *      Quem pode acessar: Qualquer pessoa
 * 4. Copie a URL gerada (termina em /exec) e cole no dashboard
 */

/**
 * ⚠️ COLE AQUI O ID DA SUA PLANILHA.
 * Está na URL dela, entre "/d/" e "/edit":
 * docs.google.com/spreadsheets/d/  ESTE_PEDAÇO_AQUI  /edit
 *
 * Se você abriu o Apps Script por dentro da planilha (Extensões → Apps Script),
 * pode deixar vazio que ele acha sozinho. Preenchendo, funciona nos dois casos.
 */
var PLANILHA_ID = '';

function doPost(e) {
  var out = function (o) {
    return ContentService.createTextOutput(JSON.stringify(o))
      .setMimeType(ContentService.MimeType.JSON);
  };

  try {
    var b = JSON.parse(e.postData.contents);
    var id = b.planilhaId || PLANILHA_ID;
    var ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return out({ ok: false, erro: 'Sem planilha. Cole o ID dela na variável PLANILHA_ID no topo do script.' });
    }

    var sh = b.aba ? ss.getSheetByName(b.aba) : ss.getSheets()[0];
    if (!sh) {
      return out({ ok: false, erro: 'Aba não encontrada: ' + b.aba });
    }

    var dados = sh.getDataRange().getValues();
    var norm = function (s) {
      return String(s == null ? '' : s)
        .normalize('NFD').replace(/[̀-ͯ]/g, '')
        .trim().toLowerCase();
    };

    // Acha a linha de cabeçalho procurando as 3 colunas na MESMA linha.
    var hRow = -1, cA = -1, cI = -1, cF = -1;
    for (var r = 0; r < Math.min(dados.length, 60) && hRow < 0; r++) {
      var a = -1, i = -1, f = -1;
      for (var c = 0; c < dados[r].length; c++) {
        var v = norm(dados[r][c]);
        if (v.indexOf('alunos') >= 0) a = c;
        else if (v.indexOf('investimento') >= 0) i = c;
        else if (v.indexOf('faturamento') >= 0) f = c;
      }
      if (a >= 0 && i >= 0 && f >= 0) { hRow = r; cA = a; cI = i; cF = f; }
    }
    if (hRow < 0) {
      return out({ ok: false, erro: 'Não achei o cabeçalho com Alunos PP / Investimento / Faturamento nesta aba.' });
    }

    // Linha alvo: a informada, ou a primeira vazia abaixo do cabeçalho.
    var linha = -1;
    if (b.linha && b.linha > 0) {
      linha = b.linha - 1;                       // 1-indexed vindo do dashboard
    } else {
      for (var r2 = hRow + 1; r2 < dados.length; r2++) {
        if (norm(dados[r2][cA]) === '') { linha = r2; break; }
      }
      if (linha < 0) linha = dados.length;       // acabou a tabela: acrescenta no fim
    }

    // Modo consulta: só devolve onde escreveria, sem gravar nada.
    if (b.simular) {
      return out({ ok: true, simulacao: true, linha: linha + 1, aba: sh.getName() });
    }

    var num = function (x) { var n = parseFloat(x); return isNaN(n) ? 0 : n; };
    sh.getRange(linha + 1, cA + 1).setValue(Math.round(num(b.alunos)));
    sh.getRange(linha + 1, cI + 1).setValue(num(b.investimento));
    sh.getRange(linha + 1, cF + 1).setValue(num(b.faturamento));

    return out({
      ok: true,
      aba: sh.getName(),
      linha: linha + 1,
      gravado: {
        alunos: Math.round(num(b.alunos)),
        investimento: num(b.investimento),
        faturamento: num(b.faturamento)
      }
    });

  } catch (err) {
    return out({ ok: false, erro: String(err) });
  }
}

// Permite abrir a URL no navegador para conferir que a implantação está no ar.
function doGet() {
  return ContentService
    .createTextOutput('Web App ativo. Use POST para gravar.')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * TESTE — rode esta função aqui no editor (botão "Executar") ANTES de implantar.
 * Ela NÃO grava nada: só confirma que acha a planilha, a aba e as 3 colunas,
 * e diz em qual linha o dashboard escreveria. Veja o resultado em "Registro de execução".
 *
 * Troque o nome da aba abaixo pelo da sua.
 */
function testarConexao() {
  var ABA = 'PERPÉTUO NATH 37';   // <-- ajuste se o nome for outro

  var r = doPost({ postData: { contents: JSON.stringify({ aba: ABA, simular: true }) } });
  var resp = JSON.parse(r.getContent());

  if (resp.ok) {
    Logger.log('✅ Tudo certo! Aba "%s" encontrada. O dashboard gravaria na linha %s.', resp.aba, resp.linha);
  } else {
    Logger.log('❌ Problema: %s', resp.erro);
  }
  return resp;
}
