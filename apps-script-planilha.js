/**
 * Preenche o "Relatório Diário MKT" a partir do dashboard Meta Ads.
 *
 * Grava APENAS 3 colunas por bloco: Alunos PP, Investimento e Faturamento.
 * As colunas com 🔒 (Lucro, CPA, TKM, ROAS, CPO) são fórmulas e NÃO são tocadas.
 *
 * Escreve na LINHA DA DATA (coluna "Data"), não na primeira linha vazia —
 * assim cada dia cai no lugar certo, mesmo se você preencher fora de ordem.
 *
 * Blocos existentes na aba (detectados pelo título na linha 4):
 *   PERPÉTUO TOMÉ ...... AP / AQ / AR
 *   PERPÉTUO NATH 37 ... AX / AY / AZ
 *   PERPÉTUO GERAL ..... BF / BG / BH
 *
 * IMPLANTAR: Implantar → Nova implantação → "App da Web"
 *            Executar como: Eu | Quem pode acessar: Qualquer pessoa
 *            Copie a URL (/exec) e cole no dashboard.
 */

// ── Configuração ────────────────────────────────────────────────────────────
var PLANILHA_ID = '1Xk3jipROXslXYMdp6x994NzbEMt_puoacSKe-MlYFJM';
var ABA_PADRAO  = '📅 Registro Diário';

// ── Helpers ─────────────────────────────────────────────────────────────────
function _norm(s) {
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .trim().toLowerCase();
}
function _colLetra(c) {              // 0 -> A, 27 -> AB
  var s = '', n = c + 1;
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}
function _normData(v) {              // qualquer formato -> "DD/MM"
  if (v instanceof Date) {
    return ('0' + v.getDate()).slice(-2) + '/' + ('0' + (v.getMonth() + 1)).slice(-2);
  }
  var m = String(v == null ? '' : v).trim().match(/^(\d{1,2})[\/\-.](\d{1,2})/);
  return m ? ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) : '';
}

/** Lê a aba e devolve o mapa: cabeçalho, coluna de data, blocos e linhas por data. */
function _mapear(sh) {
  var dados = sh.getDataRange().getValues();

  // 1) linha de cabeçalho = a que tem "Alunos PP"
  var hRow = -1;
  for (var r = 0; r < Math.min(dados.length, 30) && hRow < 0; r++) {
    for (var c = 0; c < dados[r].length; c++) {
      if (_norm(dados[r][c]).indexOf('alunos') >= 0) { hRow = r; break; }
    }
  }
  if (hRow < 0) return { erro: 'Não achei a linha de cabeçalho (nenhum "Alunos PP").' };

  // 2) coluna de data
  var cData = -1;
  for (var c2 = 0; c2 < dados[hRow].length; c2++) {
    if (_norm(dados[hRow][c2]) === 'data') { cData = c2; break; }
  }
  if (cData < 0) return { erro: 'Não achei a coluna "Data" no cabeçalho.' };

  // 3) blocos Alunos/Investimento/Faturamento + título (procurado nas 3 linhas acima)
  var blocos = [];
  for (var c3 = 0; c3 < dados[hRow].length; c3++) {
    if (_norm(dados[hRow][c3]).indexOf('alunos') < 0) continue;
    var ci = -1, cf = -1;
    for (var c4 = c3 + 1; c4 < Math.min(dados[hRow].length, c3 + 10); c4++) {
      var v = _norm(dados[hRow][c4]);
      if (ci < 0 && v.indexOf('investimento') >= 0) ci = c4;
      else if (cf < 0 && v.indexOf('faturamento') >= 0) cf = c4;
    }
    if (ci < 0 || cf < 0) continue;

    var titulo = '';
    for (var up = hRow - 1; up >= 0 && up >= hRow - 4 && !titulo; up--) {
      for (var cc = c3; cc <= cf && cc < dados[up].length; cc++) {
        var t = String(dados[up][cc] || '').trim();
        if (t && t !== '🔒') { titulo = t; break; }
      }
    }
    blocos.push({
      titulo: titulo,
      colAlunos: _colLetra(c3), colInvestimento: _colLetra(ci), colFaturamento: _colLetra(cf),
      _c: c3, _i: ci, _f: cf
    });
  }
  if (!blocos.length) return { erro: 'Nenhum bloco Alunos/Investimento/Faturamento encontrado.' };

  // 4) datas -> linha da planilha
  var linhaPorData = {};
  for (var r2 = hRow + 1; r2 < dados.length; r2++) {
    var d = _normData(dados[r2][cData]);
    if (d && !linhaPorData[d]) linhaPorData[d] = r2 + 1;
  }

  return { hRow: hRow, cData: cData, blocos: blocos, linhaPorData: linhaPorData, totalLinhas: dados.length };
}

// ── Endpoint ────────────────────────────────────────────────────────────────
function doPost(e) {
  var out = function (o) {
    return ContentService.createTextOutput(JSON.stringify(o))
      .setMimeType(ContentService.MimeType.JSON);
  };

  try {
    var b = JSON.parse(e.postData.contents);

    var ss = SpreadsheetApp.openById(b.planilhaId || PLANILHA_ID);
    var nomeAba = b.aba || ABA_PADRAO;
    var sh = ss.getSheetByName(nomeAba);
    if (!sh) {
      return out({ ok: false, erro: 'Aba não encontrada: "' + nomeAba + '"',
                   abasDisponiveis: ss.getSheets().map(function (s) { return s.getName(); }) });
    }

    var M = _mapear(sh);
    if (M.erro) return out({ ok: false, erro: M.erro });

    if (b.listar) {
      return out({ ok: true, aba: sh.getName(), blocos: M.blocos,
                   colunaData: _colLetra(M.cData), cabecalho: M.hRow + 1,
                   datas: Object.keys(M.linhaPorData).length,
                   primeiraData: Object.keys(M.linhaPorData)[0],
                   ultimaData: Object.keys(M.linhaPorData).slice(-1)[0] });
    }

    // Bloco: por título (ex: "NATH 37") ou pela letra da coluna
    var alvo = null;
    if (b.colAlunos) {
      var letra = String(b.colAlunos).trim().toUpperCase();
      for (var i = 0; i < M.blocos.length; i++) if (M.blocos[i].colAlunos === letra) alvo = M.blocos[i];
      if (!alvo) return out({ ok: false, erro: 'Sem bloco na coluna ' + letra, blocos: M.blocos });
    } else if (b.bloco) {
      var alvoN = _norm(b.bloco);
      for (var j = 0; j < M.blocos.length; j++) {
        if (_norm(M.blocos[j].titulo).indexOf(alvoN) >= 0) { alvo = M.blocos[j]; break; }
      }
      if (!alvo) return out({ ok: false, erro: 'Sem bloco com título contendo "' + b.bloco + '"', blocos: M.blocos });
    } else {
      return out({ ok: false, erro: 'Informe "bloco" (ex: NATH 37) ou "colAlunos" (ex: AX).', blocos: M.blocos });
    }

    // Linha: pela data
    var dataAlvo = _normData(b.data);
    if (!dataAlvo) return out({ ok: false, erro: 'Data inválida: "' + b.data + '". Use DD/MM.' });
    var linha = M.linhaPorData[dataAlvo];
    if (!linha) {
      return out({ ok: false, erro: 'A data ' + dataAlvo + ' não existe na coluna Data desta aba.',
                   datasDisponiveis: Object.keys(M.linhaPorData).slice(0, 5).join(', ') + '…' });
    }

    if (b.simular) {
      return out({ ok: true, simulacao: true, aba: sh.getName(), linha: linha,
                   data: dataAlvo, bloco: alvo.titulo,
                   colunas: alvo.colAlunos + '/' + alvo.colInvestimento + '/' + alvo.colFaturamento });
    }

    var num = function (x) { var n = parseFloat(x); return isNaN(n) ? 0 : n; };
    sh.getRange(linha, alvo._c + 1).setValue(Math.round(num(b.alunos)));
    sh.getRange(linha, alvo._i + 1).setValue(num(b.investimento));
    sh.getRange(linha, alvo._f + 1).setValue(num(b.faturamento));

    return out({
      ok: true, aba: sh.getName(), linha: linha, data: dataAlvo, bloco: alvo.titulo,
      gravado: { alunos: Math.round(num(b.alunos)), investimento: num(b.investimento), faturamento: num(b.faturamento) }
    });

  } catch (err) {
    return out({ ok: false, erro: String(err) });
  }
}

function doGet() {
  return ContentService.createTextOutput('Web App ativo. Use POST para gravar.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── Testes (rode no editor; NÃO gravam nada) ────────────────────────────────

/** Mostra os blocos, a coluna de data e o intervalo de datas da aba. */
function listarBlocos() {
  var r = doPost({ postData: { contents: JSON.stringify({ aba: ABA_PADRAO, listar: true }) } });
  var resp = JSON.parse(r.getContent());
  if (!resp.ok) { Logger.log('❌ %s', resp.erro); return resp; }

  Logger.log('Aba "%s" · cabeçalho na linha %s · datas na coluna %s', resp.aba, resp.cabecalho, resp.colunaData);
  Logger.log('%s datas mapeadas (%s … %s)', resp.datas, resp.primeiraData, resp.ultimaData);
  Logger.log('Blocos:');
  resp.blocos.forEach(function (x, i) {
    Logger.log('  [%s] "%s" → %s / %s / %s', i + 1, x.titulo, x.colAlunos, x.colInvestimento, x.colFaturamento);
  });
  return resp;
}

/** Simula a gravação de um dia no bloco da Nath (não escreve nada). */
function testarConexao() {
  var hoje = new Date();
  var dataTeste = ('0' + hoje.getDate()).slice(-2) + '/' + ('0' + (hoje.getMonth() + 1)).slice(-2);

  var r = doPost({ postData: { contents: JSON.stringify({
    aba: ABA_PADRAO, bloco: 'NATH 37', data: dataTeste, simular: true
  }) } });
  var resp = JSON.parse(r.getContent());

  if (resp.ok) {
    Logger.log('✅ Bloco "%s" (colunas %s)', resp.bloco, resp.colunas);
    Logger.log('   Data %s → gravaria na LINHA %s da aba "%s".', resp.data, resp.linha, resp.aba);
  } else {
    Logger.log('❌ %s', resp.erro);
    if (resp.blocos) resp.blocos.forEach(function (x) { Logger.log('   bloco: "%s" (%s)', x.titulo, x.colAlunos); });
    if (resp.datasDisponiveis) Logger.log('   datas: %s', resp.datasDisponiveis);
  }
  return resp;
}
