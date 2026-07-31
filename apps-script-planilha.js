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
    var colLetra = function (c) {   // 0 -> A, 1 -> B ...
      var s = '';
      c = c + 1;
      while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; }
      return s;
    };

    // Acha TODOS os blocos "Alunos PP / Investimento / Faturamento" da aba.
    var blocos = [];
    for (var r = 0; r < dados.length; r++) {
      for (var c = 0; c < dados[r].length; c++) {
        if (norm(dados[r][c]).indexOf('alunos') < 0) continue;
        var i = -1, f = -1;
        for (var c2 = c + 1; c2 < Math.min(dados[r].length, c + 12); c2++) {
          var v2 = norm(dados[r][c2]);
          if (i < 0 && v2.indexOf('investimento') >= 0) i = c2;
          else if (f < 0 && v2.indexOf('faturamento') >= 0) f = c2;
        }
        if (i < 0 || f < 0) continue;

        // primeira linha vazia na coluna de Alunos, abaixo do cabeçalho
        var vazia = -1;
        for (var r2 = r + 1; r2 < dados.length; r2++) {
          if (norm(dados[r2][c]) === '') { vazia = r2; break; }
        }
        if (vazia < 0) vazia = dados.length;

        // título/banner: primeiro texto não vazio nas 3 linhas acima
        var titulo = '';
        for (var up = r - 1; up >= 0 && up >= r - 3 && !titulo; up--) {
          for (var cc = Math.max(0, c - 2); cc < Math.min(dados[up].length, f + 3); cc++) {
            var t = String(dados[up][cc] || '').trim();
            if (t) { titulo = t; break; }
          }
        }
        blocos.push({
          titulo: titulo, cabecalho: r + 1,
          colAlunos: colLetra(c), colInvestimento: colLetra(i), colFaturamento: colLetra(f),
          primeiraVazia: vazia + 1,
          _c: c, _i: i, _f: f, _h: r
        });
      }
    }
    if (!blocos.length) {
      return out({ ok: false, erro: 'Não achei nenhum bloco com Alunos PP / Investimento / Faturamento nesta aba.' });
    }

    // Escolhe o bloco: por coluna informada, por âncora (texto do título), ou o primeiro.
    var alvo = null;
    if (b.colAlunos) {
      var alvoLetra = String(b.colAlunos).trim().toUpperCase();
      alvo = blocos.filter(function (x) { return x.colAlunos === alvoLetra; })[0];
      if (!alvo) return out({ ok: false, erro: 'Não achei bloco com Alunos na coluna ' + alvoLetra, blocos: blocos });
    } else if (b.ancora) {
      var anc = norm(b.ancora);
      alvo = blocos.filter(function (x) { return norm(x.titulo).indexOf(anc) >= 0; })[0];
      if (!alvo) return out({ ok: false, erro: 'Não achei bloco com o título contendo "' + b.ancora + '"', blocos: blocos });
    } else {
      alvo = blocos[0];
    }

    // Lista os blocos para você conferir/escolher, sem gravar nada.
    if (b.listar) {
      return out({ ok: true, aba: sh.getName(), blocos: blocos });
    }

    var hRow = alvo._h, cA = alvo._c, cI = alvo._i, cF = alvo._f;

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
      return out({
        ok: true, simulacao: true, aba: sh.getName(), linha: linha + 1,
        bloco: { titulo: alvo.titulo, colAlunos: alvo.colAlunos, colInvestimento: alvo.colInvestimento, colFaturamento: alvo.colFaturamento },
        totalBlocos: blocos.length
      });
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
  var ABA = '📅 Registro Diário';   // <-- nome da guia (embaixo, na planilha)

  var r = doPost({ postData: { contents: JSON.stringify({ aba: ABA, listar: true }) } });
  var resp = JSON.parse(r.getContent());

  if (!resp.ok) { Logger.log('❌ Problema: %s', resp.erro); return resp; }

  Logger.log('Aba "%s" — %s bloco(s) encontrado(s):', resp.aba, resp.blocos.length);
  for (var i = 0; i < resp.blocos.length; i++) {
    var x = resp.blocos[i];
    Logger.log('  [%s] título: "%s" | Alunos=%s Investimento=%s Faturamento=%s | 1ª linha vazia: %s',
      i + 1, x.titulo, x.colAlunos, x.colInvestimento, x.colFaturamento, x.primeiraVazia);
  }
  Logger.log('');
  Logger.log('➡️  Escolha o bloco certo e use a LETRA da coluna Alunos no dashboard (campo "Coluna Alunos").');
  return resp;
}
