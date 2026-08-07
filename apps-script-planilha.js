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

// ════════════════════════════════════════════════════════════════════════════
//  AUTOMAÇÃO DIÁRIA — preenche sozinho o dia anterior, sem abrir o dashboard
// ════════════════════════════════════════════════════════════════════════════
//
// COMO ATIVAR (uma vez):
//   1. ⚙️ Configurações do projeto → Propriedades do script → Adicionar
//        Propriedade: META_TOKEN     Valor: (o token gerado no Business Manager)
//   2. Rode  criarGatilhoDiario  uma vez
//   3. Todo dia às 5h ele preenche o dia anterior.
//
// Testes seguros (não gravam): simularOntem, simularMesAteOntem

var SB_URL = 'https://mnrnnmfaupvmjfgpzkou.supabase.co';
var SB_KEY = 'sb_publishable_KraV-pqYINerxc0nCgdHaA_xHj7ew87';
var USD_BRL = 5.40;          // mesma taxa padrão do dashboard
var HORA_GATILHO = 5;        // 5h da manhã (fuso da planilha)

// Blocos a preencher.
//   match → filtra as campanhas do Meta pelo nome e usa funnel_revenue daquele funil
//   soma  → não consulta nada: soma os blocos já calculados
//
// GERAL é soma de NATH + TOMÉ de propósito. Usar p_funil:'' traria também a
// receita de Consultorias & Mentorias (backend) e de orgânico, que não têm
// gasto de mídia — o ROAS do bloco sairia inflado.
var FUNIS = [
  { match: 'NATH', bloco: 'NATH 37' },
  { match: 'TOME', bloco: 'TOMÉ'    },
  { soma: ['NATH 37', 'TOMÉ'], bloco: 'GERAL' }
];

function _semAcento(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
}
function _metaToken() {
  var t = PropertiesService.getScriptProperties().getProperty('META_TOKEN');
  if (!t) throw new Error('Falta o META_TOKEN nas Propriedades do script (⚙️ Configurações do projeto).');
  return t;
}
function _isoDia(d) { return Utilities.formatDate(d, 'America/Sao_Paulo', 'yyyy-MM-dd'); }
function _ddmmDe(d) { return Utilities.formatDate(d, 'America/Sao_Paulo', 'dd/MM'); }

/** Gasto por campanha no Meta, no dia (converte USD→BRL como o dashboard). */
function _metaGastoPorCampanha(dia) {
  var tok = _metaToken(), API = 'https://graph.facebook.com/v21.0';
  var contas = [], url = API + '/me/adaccounts?fields=id,name,currency&limit=500&access_token=' + encodeURIComponent(tok);
  while (url) {
    var j = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
    if (j.error) throw new Error('Meta /me/adaccounts: ' + j.error.message);
    contas = contas.concat(j.data || []);
    url = (j.paging && j.paging.next) ? j.paging.next : null;
  }

  var camps = {}, tr = encodeURIComponent(JSON.stringify({ since: dia, until: dia }));
  for (var i = 0; i < contas.length; i++) {
    var acc = contas[i], usd = (acc.currency === 'USD');
    var u = API + '/' + acc.id + '/insights?level=campaign&fields=spend,campaign_id,campaign_name' +
            '&time_range=' + tr + '&limit=500&access_token=' + encodeURIComponent(tok);
    var guard = 0;
    while (u && guard++ < 20) {
      var res = JSON.parse(UrlFetchApp.fetch(u, { muteHttpExceptions: true }).getContentText());
      if (res.error) { Logger.log('  ⚠️ %s: %s', acc.name, res.error.message); break; }
      (res.data || []).forEach(function (r) {
        var v = parseFloat(r.spend) || 0; if (!v) return;
        var id = r.campaign_id;
        if (!camps[id]) camps[id] = { nome: r.campaign_name || '', gasto: 0 };
        camps[id].gasto += usd ? v * USD_BRL : v;
      });
      u = (res.paging && res.paging.next) ? res.paging.next : null;
    }
  }
  return camps;
}

function _sbRpc(nome, corpo) {
  var r = UrlFetchApp.fetch(SB_URL + '/rest/v1/rpc/' + nome, {
    method: 'post', contentType: 'application/json',
    headers: { apikey: SB_KEY, Authorization: 'Bearer ' + SB_KEY },
    payload: JSON.stringify(corpo), muteHttpExceptions: true
  });
  var t = r.getContentText();
  if (r.getResponseCode() >= 300) throw new Error('Supabase ' + nome + ': ' + t);
  return JSON.parse(t);
}

/**
 * Calcula os valores de todos os blocos para um dia.
 * Fonte única usada por preencher e simular, para os dois nunca divergirem.
 * @return {Object} { 'NATH 37': {alunos, investimento, faturamento}, ... }
 */
function _calcularDia(quando) {
  var dia = _isoDia(quando);
  var camps = _metaGastoPorCampanha(dia);
  var since = dia + 'T00:00:00-03:00';
  var until = Utilities.formatDate(new Date(quando.getTime() + 864e5), 'America/Sao_Paulo', 'yyyy-MM-dd') + 'T00:00:00-03:00';

  // vendas front por campanha (mesma fonte do dashboard)
  var frontPorCamp = {};
  _sbRpc('meta_revenue_by_ad', { p_since: since, p_until: until }).forEach(function (r) {
    var id = String(r.campaign_id || ''); if (!id) return;
    frontPorCamp[id] = (frontPorCamp[id] || 0) + (parseInt(r.vendas_front) || 0);
  });

  var out = {};
  // 1ª passada: blocos que consultam de fato
  FUNIS.forEach(function (f) {
    if (f.soma) return;
    var m = _semAcento(f.match), invest = 0, alunos = 0;
    Object.keys(camps).forEach(function (id) {
      if (m && _semAcento(camps[id].nome).indexOf(m) < 0) return;
      invest += camps[id].gasto; alunos += frontPorCamp[id] || 0;
    });
    var fr = _sbRpc('funnel_revenue', { p_since: since, p_until: until, p_funil: f.match })[0] || {};
    out[f.bloco] = { alunos: alunos, investimento: Math.round(invest * 100) / 100,
                     faturamento: parseFloat(fr.receita_bruta) || 0 };
  });
  // 2ª passada: blocos que são soma de outros
  FUNIS.forEach(function (f) {
    if (!f.soma) return;
    var t = { alunos: 0, investimento: 0, faturamento: 0 };
    f.soma.forEach(function (nome) {
      var v = out[nome]; if (!v) return;
      t.alunos += v.alunos; t.investimento += v.investimento; t.faturamento += v.faturamento;
    });
    t.investimento = Math.round(t.investimento * 100) / 100;
    out[f.bloco] = t;
  });
  return out;
}

/** Escreve os 3 valores na linha da data, no bloco indicado. */
function _escrever(sh, dataDDMM, blocoNome, vals) {
  var M = _mapear(sh);
  if (M.erro) throw new Error(M.erro);

  var alvo = null, bn = _semAcento(blocoNome);
  for (var i = 0; i < M.blocos.length; i++) {
    if (_semAcento(M.blocos[i].titulo).indexOf(bn) >= 0) { alvo = M.blocos[i]; break; }
  }
  if (!alvo) throw new Error('Bloco "' + blocoNome + '" não encontrado.');

  var linha = M.linhaPorData[dataDDMM];
  if (!linha) throw new Error('Data ' + dataDDMM + ' não existe na aba.');

  sh.getRange(linha, alvo._c + 1).setValue(Math.round(vals.alunos));
  sh.getRange(linha, alvo._i + 1).setValue(vals.investimento);
  sh.getRange(linha, alvo._f + 1).setValue(vals.faturamento);
  return { linha: linha, bloco: alvo.titulo };
}

/** Preenche (ou simula) um dia em todos os blocos. */
function _rodarDia(quando, simular) {
  var ddmm = _ddmmDe(quando);
  var vals = _calcularDia(quando);
  var sh = simular ? null : SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(ABA_PADRAO);
  if (!simular && !sh) throw new Error('Aba "' + ABA_PADRAO + '" não encontrada.');

  var linha = '';
  FUNIS.forEach(function (f) {
    var v = vals[f.bloco]; if (!v) return;
    if (!simular) _escrever(sh, ddmm, f.bloco, v);
    linha += Utilities.formatString('  %s: %s alunos / R$ %s / R$ %s',
      f.bloco, v.alunos, v.investimento.toFixed(2), v.faturamento.toFixed(2));
  });
  Logger.log('%s %s%s', simular ? '  ' : '✅', ddmm, linha);
  return vals;
}

function preencherDia(quando) { return _rodarDia(quando, false); }
function simularDia(quando)   { return _rodarDia(quando, true); }

/** Roda automaticamente pelo gatilho: preenche o dia anterior. */
function preencherOntem() {
  Logger.log('▶ Preenchendo ontem');
  return _rodarDia(new Date(Date.now() - 864e5), false);
}
/** Mostra o que gravaria ontem, sem escrever. */
function simularOntem() {
  Logger.log('🔍 Simulando ontem — nada será gravado');
  return _rodarDia(new Date(Date.now() - 864e5), true);
}

/** Cria (ou recria) o gatilho diário. Rode UMA vez. */
function criarGatilhoDiario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'preencherOntem') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('preencherOntem').timeBased().atHour(HORA_GATILHO).everyDays(1).create();
  Logger.log('✅ Gatilho criado: preencherOntem, todo dia por volta das %sh.', HORA_GATILHO);
}
/** Remove o gatilho, se quiser desligar a automação. */
function removerGatilhoDiario() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'preencherOntem') { ScriptApp.deleteTrigger(t); n++; }
  });
  Logger.log('%s gatilho(s) removido(s).', n);
}

// ── Preenchimento retroativo ────────────────────────────────────────────────
function _diasEntre(ini, fim) {
  var out = [], d = new Date(ini.getTime());
  while (d.getTime() <= fim.getTime()) { out.push(new Date(d.getTime())); d.setDate(d.getDate() + 1); }
  return out;
}
function _hojeSP() {
  var s = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy/MM/dd').split('/');
  return new Date(+s[0], +s[1] - 1, +s[2]);
}
function _rodarIntervalo(ini, fim, simular) {
  var dias = _diasEntre(ini, fim);
  Logger.log('%s %s dia(s): %s → %s', simular ? '🔍 SIMULANDO' : '▶ PREENCHENDO',
    dias.length, _ddmmDe(ini), _ddmmDe(fim));
  Logger.log('');
  var oks = 0, falhas = 0;
  for (var i = 0; i < dias.length; i++) {
    try { _rodarDia(dias[i], simular); oks++; }
    catch (e) { Logger.log('❌ %s — %s', _ddmmDe(dias[i]), e.message); falhas++; }
  }
  Logger.log('');
  Logger.log('%s %s dia(s) ok, %s falha(s).', simular ? '🔍 Simulação:' : '✅ Concluído:', oks, falhas);
}

/** SIMULA o mês atual, do dia 1 até ontem. Não grava nada. */
function simularMesAteOntem() {
  var hoje = _hojeSP(), ini = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  var ontem = new Date(hoje.getTime() - 864e5);
  if (ontem.getTime() < ini.getTime()) { Logger.log('Ainda é dia 1 — nada a preencher.'); return; }
  _rodarIntervalo(ini, ontem, true);
}
/** PREENCHE o mês atual, do dia 1 até ontem. Sobrescreve o que houver. */
function preencherMesAteOntem() {
  var hoje = _hojeSP(), ini = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  var ontem = new Date(hoje.getTime() - 864e5);
  if (ontem.getTime() < ini.getTime()) { Logger.log('Ainda é dia 1 — nada a preencher.'); return; }
  _rodarIntervalo(ini, ontem, false);
}
/** Intervalo livre — edite as datas (mês é 0-based: agosto = 7). */
function preencherIntervaloManual() {
  _rodarIntervalo(new Date(2026, 7, 1), new Date(2026, 7, 5), false);
}
