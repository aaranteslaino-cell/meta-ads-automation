/**
 * Recebe os dados do dashboard Meta Ads e preenche APENAS 3 colunas:
 * "Alunos PP", "Investimento" e "Faturamento".
 * As colunas com fórmula (Lucro, CPA, TKM, ROAS, CPO) NÃO são tocadas —
 * elas se recalculam sozinhas a partir dessas três.
 *
 * JÁ CONFIGURADO para a planilha do Arthur:
 *   Planilha .... 1JOsq_-n8_4Fn0AJRnr-KFZrIh1gVa2_m0qlZ6UJ72yo
 *   Aba ......... 📅 Registro Diário
 *   Bloco ....... coluna AX (PERPÉTUO NATH 37)
 *
 * COMO IMPLANTAR:
 *   Implantar → Nova implantação → tipo "App da Web"
 *     Executar como: Eu
 *     Quem pode acessar: Qualquer pessoa
 *   Copie a URL gerada (termina em /exec) e cole no dashboard.
 */

// ── Configuração ────────────────────────────────────────────────────────────
var PLANILHA_ID = '1JOsq_-n8_4Fn0AJRnr-KFZrIh1gVa2_m0qlZ6UJ72yo';
var ABA_PADRAO  = '📅 Registro Diário';
var COL_PADRAO  = 'AX';   // coluna da "Alunos PP" do bloco PERPÉTUO NATH 37

// ── Endpoint ────────────────────────────────────────────────────────────────
function doPost(e) {
  var out = function (o) {
    return ContentService.createTextOutput(JSON.stringify(o))
      .setMimeType(ContentService.MimeType.JSON);
  };

  try {
    var b = JSON.parse(e.postData.contents);

    var ss = SpreadsheetApp.openById(b.planilhaId || PLANILHA_ID);
    if (!ss) return out({ ok: false, erro: 'Planilha não encontrada.' });

    var nomeAba = b.aba || ABA_PADRAO;
    var sh = ss.getSheetByName(nomeAba);
    if (!sh) {
      return out({
        ok: false,
        erro: 'Aba não encontrada: "' + nomeAba + '"',
        abasDisponiveis: ss.getSheets().map(function (s) { return s.getName(); })
      });
    }

    var dados = sh.getDataRange().getValues();

    var norm = function (s) {
      return String(s == null ? '' : s)
        .normalize('NFD').replace(/[̀-ͯ]/g, '')
        .trim().toLowerCase();
    };
    var colLetra = function (c) {          // 0 -> A, 27 -> AB ...
      var s = '', n = c + 1;
      while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
      return s;
    };

    // Mapeia TODOS os blocos "Alunos / Investimento / Faturamento" da aba.
    var blocos = [];
    for (var r = 0; r < dados.length; r++) {
      for (var c = 0; c < dados[r].length; c++) {
        if (norm(dados[r][c]).indexOf('alunos') < 0) continue;

        var ci = -1, cf = -1;
        for (var c2 = c + 1; c2 < Math.min(dados[r].length, c + 12); c2++) {
          var v = norm(dados[r][c2]);
          if (ci < 0 && v.indexOf('investimento') >= 0) ci = c2;
          else if (cf < 0 && v.indexOf('faturamento') >= 0) cf = c2;
        }
        if (ci < 0 || cf < 0) continue;

        // primeira linha vazia na coluna de Alunos, abaixo do cabeçalho
        var vazia = -1;
        for (var r2 = r + 1; r2 < dados.length; r2++) {
          if (norm(dados[r2][c]) === '') { vazia = r2; break; }
        }
        if (vazia < 0) vazia = dados.length;

        // título do bloco: primeiro texto não vazio nas 3 linhas acima
        var titulo = '';
        for (var up = r - 1; up >= 0 && up >= r - 3 && !titulo; up--) {
          for (var cc = Math.max(0, c - 2); cc < Math.min(dados[up].length, cf + 3); cc++) {
            var t = String(dados[up][cc] || '').trim();
            if (t) { titulo = t; break; }
          }
        }

        blocos.push({
          titulo: titulo,
          cabecalho: r + 1,
          colAlunos: colLetra(c),
          colInvestimento: colLetra(ci),
          colFaturamento: colLetra(cf),
          primeiraVazia: vazia + 1,
          _h: r, _c: c, _i: ci, _f: cf
        });
      }
    }
    if (!blocos.length) {
      return out({ ok: false, erro: 'Nenhum bloco com Alunos/Investimento/Faturamento nesta aba.' });
    }

    // Só listar os blocos (não grava nada)
    if (b.listar) return out({ ok: true, aba: sh.getName(), blocos: blocos });

    // Escolhe o bloco pela letra da coluna
    var letra = String(b.colAlunos || COL_PADRAO).trim().toUpperCase();
    var alvo = null;
    for (var i = 0; i < blocos.length; i++) {
      if (blocos[i].colAlunos === letra) { alvo = blocos[i]; break; }
    }
    if (!alvo) {
      return out({ ok: false, erro: 'Não achei bloco com Alunos na coluna ' + letra, blocos: blocos });
    }

    // Linha alvo: a informada, ou a primeira vazia do bloco
    var linha = (b.linha && b.linha > 0) ? (b.linha - 1) : (alvo.primeiraVazia - 1);

    if (b.simular) {
      return out({
        ok: true, simulacao: true, aba: sh.getName(), linha: linha + 1,
        bloco: {
          titulo: alvo.titulo, colAlunos: alvo.colAlunos,
          colInvestimento: alvo.colInvestimento, colFaturamento: alvo.colFaturamento
        },
        totalBlocos: blocos.length
      });
    }

    var num = function (x) { var n = parseFloat(x); return isNaN(n) ? 0 : n; };
    sh.getRange(linha + 1, alvo._c + 1).setValue(Math.round(num(b.alunos)));
    sh.getRange(linha + 1, alvo._i + 1).setValue(num(b.investimento));
    sh.getRange(linha + 1, alvo._f + 1).setValue(num(b.faturamento));

    return out({
      ok: true,
      aba: sh.getName(),
      linha: linha + 1,
      bloco: alvo.titulo,
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

// Abrir a URL no navegador confirma que a implantação está no ar.
function doGet() {
  return ContentService
    .createTextOutput('Web App ativo. Use POST para gravar.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── Testes (rode aqui no editor, não gravam nada) ────────────────────────────

/**
 * Confirma que acha a planilha, a aba e o bloco da coluna AX,
 * e diz em qual linha o dashboard escreveria.
 */
function testarConexao() {
  var r = doPost({ postData: { contents: JSON.stringify({
    aba: ABA_PADRAO, colAlunos: COL_PADRAO, simular: true
  }) } });
  var resp = JSON.parse(r.getContent());

  if (resp.ok) {
    Logger.log('✅ Aba "%s" · bloco "%s" (colunas %s/%s/%s)',
      resp.aba, resp.bloco.titulo, resp.bloco.colAlunos,
      resp.bloco.colInvestimento, resp.bloco.colFaturamento);
    Logger.log('   O dashboard gravaria na LINHA %s.', resp.linha);
  } else {
    Logger.log('❌ %s', resp.erro);
    if (resp.blocos) {
      Logger.log('Blocos disponíveis:');
      resp.blocos.forEach(function (x, i) {
        Logger.log('  [%s] "%s" | Alunos=%s | 1ª vazia: %s', i + 1, x.titulo, x.colAlunos, x.primeiraVazia);
      });
    }
    if (resp.abasDisponiveis) Logger.log('Abas: %s', resp.abasDisponiveis.join(' | '));
  }
  return resp;
}

/** Lista todos os blocos da aba, caso precise reconferir as colunas. */
function listarBlocos() {
  var r = doPost({ postData: { contents: JSON.stringify({ aba: ABA_PADRAO, listar: true }) } });
  var resp = JSON.parse(r.getContent());

  if (!resp.ok) { Logger.log('❌ %s', resp.erro); return resp; }

  Logger.log('Aba "%s" — %s bloco(s):', resp.aba, resp.blocos.length);
  resp.blocos.forEach(function (x, i) {
    Logger.log('  [%s] "%s" | Alunos=%s Investimento=%s Faturamento=%s | 1ª vazia: %s',
      i + 1, x.titulo, x.colAlunos, x.colInvestimento, x.colFaturamento, x.primeiraVazia);
  });
  return resp;
}
