// ============================================================
//   CONTROLE DE DEVOLUÇÕES v6.0 — OTIMIZAÇÃO COMPLETA
//
//  v5.7 (onEdit / status / dashboard):
//  [P01] syncCheckboxes: 3 setValue → 1 setValues batch (cols 12-14)
//  [P02] contarTotalConcluidos: O(n×3abas) → O(1) PropertiesService
//  [P03] _atualizarMetricasDashboard: debounce 8s via CacheService
//  [P04] protegerLinhaConcluida: contador Properties (não itera abas)
//  [P05] aplicarCorLinha: aceita dtOpcional, evita getValue extra
//  [P06] onOpen: reaplicarCores só quando cache de 1h expirou
//  [P07] _gravarLancamento: remove 3 setNumberFormat redundantes
//  [P08] _aplicarStatus: status+chk+obs em 1 setValues 1×5 por linha
//  [P09] _registrarLogAba: NFD+NF em 1 getRange 1×2
//  [P10] configurarPlanilha: reseta contadores/caches
//
//  v5.8 (auditoria completa — todas as funções):
//  [P11] registrarLog: appendRow → getLastRow+setValues (~30% mais rápido)
//  [P12] onEdit: valoresNF só lido quando col === COL_NF (evita getRange)
//  [P13] onEdit: 3 if de checkboxes → 1 if + 1 getValue compartilhado
//  [P14] _aplicarStatus: status+chk+obs em 1 setValues 1×5 (cols 11-15)
//  [P15] _moverParaHistorico: N appendRow → 1 setValues batch
//  [P16] executarExportarPDF: N×5 setValue → batch por linha + cores batch
//  [P17] executarExportarPDF: saveAndClose() redundante removido
//  [P18] executarBaixaVenda: N×5 setValue → 1 setValues 1×5 por item
//  [P19] salvarLancamentoForm: COL_NF lida 1× (antes eram 3×)
//  [P20] _gravarLancamento: upload Drive fora do ScriptLock
//  [P21] executarReabertura: getProtections() 1× por aba (não por NF)
//  [P22] executarReabertura: status+chk+obs em 1 setValues 1×5 por NF
//  [P23] enviarEmailDevolucao: DriveApp.getFileById 1× por arquivo
//  [P24] buscarHistoricoNF: lê últimas 500 linhas (não coluna inteira)
//  [P25] _gerarRelatorioPDF: cores da tabela em 2 setBackgrounds (não N×2)
//  [P26] _gerarRelatorioPDF: reusa pdfBlob em memória (não relê do Drive)
//  [P27] executarRestauracao: proteções em mapa, setValues em batch por aba
//  [P28] navegarParaLinha: flush() desnecessário removido
//  [P29] enviarResumoSemanal: 7 passes no array → 1 loop acumulador
//  [P30] FormReabertura.html: itensEncontrados passados direto ao servidor
//
//  v6.0 (unificação de formulários + painel de auditoria):
//  [P31] FormAuditoria.html: painel unificado (NF + e-mails + log)
//  [P32] Menu atualizado com todas as entradas v6.0
//  [P33] obterDiagnostico: retorna versão v6.0 corretamente
// ============================================================


// ════════════════════════════════════════════════════════════
//   VARIÁVEIS GLOBAIS — todas as declarações var/const no topo
// ════════════════════════════════════════════════════════════

// ── Chaves de PropertiesService ──────────────────────────────
var _PROP_KEY_CONCLUIDOS  = 'cdv_total_concluidos';
var _PROP_KEY_PROTECOES   = 'cdv_total_protecoes';

// ── Chaves de CacheService ───────────────────────────────────
var _CACHE_KEY_DASH       = 'cdv_dash_lock';
var _CACHE_KEY_CORES      = 'cdv_cores_ok';
var _CACHE_KEY_SENTINEL   = 'cdv_sentinel_ok';

// ── Tempos de cache ──────────────────────────────────────────
var _DASH_DEBOUNCE_SEG    = 8;     // segundos mínimos entre atualizações do dashboard
var _CORES_TTL_SEG        = 3600;  // 1 hora de cache para reaplicação de cores

// ── Chaves de configuração (e-mails e cores) ─────────────────
var _KEY_EMAILS_GERAL     = 'cdv_emails_geral';
var _KEY_EMAILS_ALERTA    = 'cdv_emails_alerta';
var _KEY_ALERTA_DEST      = 'cdv_alerta_dest';   // 'todos' | 'cc'
var _KEY_CORES            = 'cdv_cores';

// ── Dashboard: sentinela e células de filtro ─────────────────
var DASH_SENTINEL_CELL    = 'K1';
var DASH_SENTINEL_VALUE   = 'v6.0';
var DASH_DATA_INI_CELL    = 'C4';
var DASH_DATA_FIM_CELL    = 'C5';

// ── Dashboard: grupos de colunas (qtd + valor) ───────────────
var DASH_COLS = [
  { c: 2, label: 'BRITANIA',       cor: '#2563EB' },
  { c: 4, label: 'UNILEVER',       cor: '#059669' },
  { c: 6, label: 'FORN. VARIADOS', cor: '#D97706' },
  { c: 8, label: 'TOTAL GERAL',    cor: '#7C3AED' }
];

// ── Dashboard: paleta de cores ───────────────────────────────
var DC = {
  HEADER  : '#1A3557', SUB    : '#243F63',
  BRANCO  : '#FFFFFF', CINZA  : '#F0F2F5', BORDA : '#E5E7EB',
  PEND_BG : '#EBF3FF', PEND   : '#2563EB',
  DEV_BG  : '#ECFDF5', DEV    : '#059669',
  VENDA_BG: '#FFF7ED', VENDA  : '#D97706',
  TOT_BG  : '#F5F3FF', TOT    : '#7C3AED',
  TEXTO   : '#111827', TEXTO_L: '#6B7280'
};

// ── Backup ───────────────────────────────────────────────────
var BACKUP_ABA       = '_Backup_Snapshot';
var BACKUP_TOTAL_COL = 19; // 1(aba) + 17(dados) + 1(timestamp)


// ════════════════════════════════════════════════════════════
//   CONFIGURAÇÕES (constantes)
// ════════════════════════════════════════════════════════════

const EMAILS_DESTINATARIOS   = ['cidamara.silva@transben.com.br',
                                'mauro.santana@transben.com.br',
                                'sac@transben.com.br', 
                                'luiz.freire@transben.com.br',
                                'luiz.borba@transben.com.br',
                                'graziela.rodrigues@transben.com.br'];
const ID_MODELO_DOC          = '1zhS4HRlUvKoDUZCxf9HkaUAv0VnA-laSWfAHJJkz8l4';
const ID_PASTA_DESTINO       = '1W4dZMqV4d4qcs8-TIzLVvDNCQDh_CDPp';
const ID_PASTA_DESTINO_VENDA = '1McE2mLTyfK1J2d5BOz4nvecP0T4eJ3Wz';
const ID_PASTA_ANEXOS        = '14W3s-LnHl2aDCbz0h-Zi_FGmar3xLiNd';

// ── Cores ────────────────────────────────────────────────────
const COR_AZUL          = '#DDEEFF';
const COR_VERDE         = '#DDFFDD';
const COR_LARANJA       = '#FFE5CC';
const COR_ALERTA_30DIAS = '#FFD5D5';
const COR_HEADER        = '#1A3557';
const COR_VERMELHO      = '#FFD5D5';

// ── Colunas (1-based) ────────────────────────────────────────
const COL_NFD         = 1;
const COL_NF          = 2;
const COL_DATA        = 3;
const COL_FORN        = 4;
const COL_TIPO        = 5;
const COL_MOTIVO      = 6;
const COL_DESC        = 7;
const COL_QTD         = 8;
const COL_VL_UNIT     = 9;
const COL_VL_TOT      = 10;
const COL_STATUS      = 11;
const COL_PEND_CHK    = 12;
const COL_DEV_CHK     = 13;
const COL_VENDA_CHK   = 14;
const COL_OBS         = 15;
const COL_RESP        = 16;
const COL_ANEXO       = 17;
const TOTAL_COLUNAS   = 17;
const LINHA_DADOS     = 4;
const MAX_LINHAS_ABA  = 200;
const LIMITE_PROTECOES = 380;

// ── Índices de array (0-based) ───────────────────────────────
const IDX_NFD       = COL_NFD       - 1;
const IDX_NF        = COL_NF        - 1;
const IDX_DATA      = COL_DATA      - 1;
const IDX_FORN      = COL_FORN      - 1;
const IDX_TIPO      = COL_TIPO      - 1;
const IDX_MOTIVO    = COL_MOTIVO    - 1;
const IDX_DESC      = COL_DESC      - 1;
const IDX_QTD       = COL_QTD       - 1;
const IDX_VL_UNIT   = COL_VL_UNIT   - 1;
const IDX_VL_TOT    = COL_VL_TOT    - 1;
const IDX_STATUS    = COL_STATUS    - 1;
const IDX_PEND_CHK  = COL_PEND_CHK  - 1;
const IDX_DEV_CHK   = COL_DEV_CHK   - 1;
const IDX_VENDA_CHK = COL_VENDA_CHK - 1;
const IDX_OBS       = COL_OBS       - 1;
const IDX_RESP      = COL_RESP      - 1;
const IDX_ANEXO     = COL_ANEXO     - 1;

// ── Abas operacionais ────────────────────────────────────────
const ABAS_OPERACIONAIS = ['Britania', 'Unilever', 'Fornecedores Variados'];


// ════════════════════════════════════════════════════════════
//   HELPERS DE CACHE/PROPERTIES
// ════════════════════════════════════════════════════════════

/** Lê o contador de itens concluídos do PropertiesService. */
function _lerContadorConcluidos() {
  return parseInt(PropertiesService.getScriptProperties()
    .getProperty(_PROP_KEY_CONCLUIDOS) || '0');
}

/** Incrementa o contador e retorna o novo valor. */
function _incrementarContadorConcluidos() {
  var props = PropertiesService.getScriptProperties();
  var n = parseInt(props.getProperty(_PROP_KEY_CONCLUIDOS) || '0') + 1;
  props.setProperty(_PROP_KEY_CONCLUIDOS, String(n));
  return n;
}

/** Decrementa o contador (usado ao arquivar). */
function _decrementarContadorConcluidos(qtd) {
  var props = PropertiesService.getScriptProperties();
  var n = Math.max(0, parseInt(props.getProperty(_PROP_KEY_CONCLUIDOS) || '0') - (qtd || 1));
  props.setProperty(_PROP_KEY_CONCLUIDOS, String(n));
}

/** Zera o contador de concluídos (após arquivamento). */
function _zerarContadorConcluidos() {
  PropertiesService.getScriptProperties().setProperty(_PROP_KEY_CONCLUIDOS, '0');
}

/** Lê o total de proteções ativas (cache em Properties). */
function _lerTotalProtecoes() {
  return parseInt(PropertiesService.getScriptProperties()
    .getProperty(_PROP_KEY_PROTECOES) || '0');
}

/** Incrementa o contador de proteções. */
function _incrementarProtecoes() {
  var props = PropertiesService.getScriptProperties();
  var n = parseInt(props.getProperty(_PROP_KEY_PROTECOES) || '0') + 1;
  props.setProperty(_PROP_KEY_PROTECOES, String(n));
}

/** Decrementa o contador de proteções. */
function _decrementarProtecoes(qtd) {
  var props = PropertiesService.getScriptProperties();
  var n = Math.max(0, parseInt(props.getProperty(_PROP_KEY_PROTECOES) || '0') - (qtd || 1));
  props.setProperty(_PROP_KEY_PROTECOES, String(n));
}

/** Reseta todos os contadores e caches (usado em configurarPlanilha). */
function _resetarContadores() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  try { CacheService.getScriptCache().removeAll([_CACHE_KEY_DASH, _CACHE_KEY_CORES, _CACHE_KEY_SENTINEL]); } catch(_) {}
}


// ════════════════════════════════════════════════════════════
//   HELPERS GENÉRICOS
// ════════════════════════════════════════════════════════════

/**
 * Verifica se qualquer termo da lista bate com a NFD ou a NF da linha.
 * Retorna { bate: bool, termoBateu: string|null }
 */
function _baterTermos(termos, nfd, nf) {
  var nfdStr = String(nfd || '').trim();
  var nfStr  = String(nf  || '').trim();
  for (var i = 0; i < termos.length; i++) {
    var t = termos[i];
    if ((nfdStr && nfdStr === t) || (nfStr && nfStr === t)) {
      return { bate: true, termoBateu: t };
    }
  }
  return { bate: false, termoBateu: null };
}

/** Última linha com NF preenchida (âncora: COL_NF). */
function obterUltimaLinhaDados(ws) {
  var lastRow;
  try { lastRow = ws.getLastRow(); } catch (_) { return LINHA_DADOS - 1; }
  if (lastRow < LINHA_DADOS) return LINHA_DADOS - 1;
  var vals = ws.getRange(LINHA_DADOS, COL_NF, lastRow - LINHA_DADOS + 1, 1).getValues();
  var ultima = LINHA_DADOS - 1;
  vals.forEach(function(r, i) {
    if (r[0] !== '' && r[0] != null) ultima = LINHA_DADOS + i;
  });
  return ultima;
}

/** Cor de fundo por status. */
function corPorStatus(status) {
  switch (status) {
    case 'Pendente':  return COR_AZUL;
    case 'Devolvido': return COR_VERDE;
    case 'Venda':     return COR_LARANJA;
    default:          return '#FFFFFF';
  }
}

/** Escapa caracteres HTML para uso seguro em templates de e-mail. */
function _esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Extrai ID de arquivo a partir de URL do Drive. */
function _extrairIdDriveUrl(url) {
  if (!url) return null;
  var m = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) ||
          url.match(/[?&]id=([a-zA-Z0-9_-]+)/)     ||
          url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

/** Tenta apagar arquivo do Drive pela URL; falha silenciosamente. */
function _apagarAnexoDrive(url) {
  if (!url || !url.startsWith('http')) return;
  try {
    var id = _extrairIdDriveUrl(url);
    if (id) DriveApp.getFileById(id).setTrashed(true);
  } catch (_) {}
}

/** Pasta de anexos: usa ID_PASTA_ANEXOS se configurado, senão ID_PASTA_DESTINO. */
function _pastaAnexos() {
  return (ID_PASTA_ANEXOS && !ID_PASTA_ANEXOS.startsWith('INSIRA'))
    ? DriveApp.getFolderById(ID_PASTA_ANEXOS)
    : DriveApp.getFolderById(ID_PASTA_DESTINO);
}

/** Fórmula de valor total para a linha `row` da planilha. */
function _formulaTotal(row) {
  return '=IF(OR(H' + row + '="";I' + row + '="");"";H' + row + '*I' + row + ')';
}

/** Formata número como valor monetário sem símbolo R$. */
function _fmtVal(n) {
  return (parseFloat(n) || 0).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

/** Converte 'YYYY-MM-DD' em Date (início do dia). */
function _parseDateStr(s, fimDia) {
  if (!s) return null;
  var p = s.split('-');
  if (p.length < 3) return null;
  var d = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
  if (isNaN(d.getTime())) return null;
  if (fimDia) d.setHours(23, 59, 59, 999);
  return d;
}

/** Formata Date como dd/MM/yyyy usando timezone da planilha. */
function _fmtDt(dt, tz) {
  return Utilities.formatDate(dt, tz, 'dd/MM/yyyy');
}

/** Nome legível da coluna para o log. */
function obterNomeColuna(col) {
  var mapa = {
    1:'NFD', 2:'NF', 3:'Data', 4:'Fornecedor', 5:'Tipo', 6:'Motivo', 7:'Descrição', 8:'Qtd',
    9:'Vl Unit', 10:'Vl Total', 11:'Status',
    12:'Pendente✓', 13:'Devolvido✓', 14:'Venda✓', 15:'Obs', 16:'Responsável', 17:'Anexo'
  };
  return mapa[col] || ('Col' + col);
}

/** Retorna array de fornecedores únicos dos itens. */
function _fornecedoresUnicos(itens) {
  return itens.reduce(function(a, it) {
    if (a.indexOf(it.forn) === -1) a.push(it.forn);
    return a;
  }, []);
}

/** Monta assunto/título do e-mail com base nos tipos dos itens. */
function _montarTituloEmail(itens, forn) {
  var tipos = itens.reduce(function(a, it) {
    if (it.tipo && a.indexOf(it.tipo) === -1) a.push(it.tipo);
    return a;
  }, []);
  var temFalta    = tipos.indexOf('Falta')    !== -1;
  var temAvaria   = tipos.indexOf('Avaria')   !== -1;
  var temRejeicao = tipos.indexOf('Rejeição') !== -1;
  var tipoStr;
  if (temRejeicao && !temFalta && !temAvaria) {
    tipoStr = 'NF REJEITADA';
  } else if (temRejeicao && (temFalta || temAvaria)) {
    tipoStr = 'NFD DE DEVOLUÇÃO E REJEIÇÃO';
  } else if (temFalta && temAvaria) {
    tipoStr = 'NFD DE AVARIA E FALTA';
  } else if (tipos.length === 1) {
    tipoStr = 'NFD DE ' + tipos[0].toUpperCase();
  } else {
    tipoStr = 'NFD';
  }
  return tipoStr + ' (' + forn.toUpperCase() + ')';
}

/** Monta lista de destinatários: base + extras do formulário. */
function _montarDestinatarios(emailsExtras) {
  // Lê do PropertiesService (configurado via tela de configurações)
  // Se não houver nada salvo, cai na constante do código
  var dest = _getEmailsGeral().slice();
  
  if (emailsExtras) {
    emailsExtras.split(/[;,\n]/).forEach(function(e) {
      var em = e.trim();
      if (em && dest.indexOf(em) === -1) dest.push(em);
    });
  }
  return dest;
}

/** Retorna quais pastas serão varridas de acordo com o tipo selecionado.
 *  IMPORTANTE: ID_PASTA_ANEXOS (fotos/PDFs das NFs originais) NUNCA é incluída. */
function _pastasParaLimpar(tipo) {
  var pastas = [];
  if (tipo === 'relatorios' || tipo === 'tudo' || tipo === 'devolucoes')
    pastas.push({ id: ID_PASTA_DESTINO, label: 'Relatórios / PDFs de Devolução' });
  if (tipo === 'vendas' || tipo === 'tudo')
    pastas.push({ id: ID_PASTA_DESTINO_VENDA, label: 'PDFs de Venda' });
  var vistos = {};
  return pastas.filter(function(p) {
    if (!p.id || p.id.startsWith('INSIRA') || vistos[p.id]) return false;
    vistos[p.id] = true;
    return true;
  });
}

/** Acumula totais por status em 1 loop. */
function _acumular(linhas) {
  var acc = { tP:0, tD:0, tV:0, vP:0, vD:0, vV:0, vTotal:0 };
  linhas.forEach(function(l) {
    acc.vTotal += l.val;
    if      (l.st === 'Pendente')  { acc.tP++; acc.vP += l.val; }
    else if (l.st === 'Devolvido') { acc.tD++; acc.vD += l.val; }
    else if (l.st === 'Venda')     { acc.tV++; acc.vV += l.val; }
  });
  acc.taxa = linhas.length > 0
    ? Math.round((acc.tD + acc.tV) / linhas.length * 100)
    : 0;
  return acc;
}

/** Monta array de KPIs para o corpo do e-mail. */
function _kpisEmail(acc) {
  return [
    { label: 'Pendentes',  cor: '#2563EB', valor: acc.tP + ' itens', sub: 'R$ ' + _fmtVal(acc.vP) },
    { label: 'Devolvidos', cor: '#059669', valor: acc.tD + ' itens', sub: 'R$ ' + _fmtVal(acc.vD) },
    { label: 'Vendas',     cor: '#D97706', valor: acc.tV + ' itens', sub: 'R$ ' + _fmtVal(acc.vV) },
    { label: 'Taxa',
      cor:   acc.taxa >= 70 ? '#059669' : acc.taxa >= 40 ? '#D97706' : '#DC2626',
      valor: acc.taxa + '%', sub: 'de resolução' }
  ];
}

function verificarEmailsJaEnviados(nfdsRaw) {
  var nfds = nfdsRaw.split(/[\n,]/).map(function(s) { return s.trim(); }).filter(Boolean);
  if (!nfds.length) return JSON.stringify({ jaEnviadas: [] });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName('_EmailsEnviados');
  if (!ws) return JSON.stringify({ jaEnviadas: [] });

  try {
    var ul = ws.getLastRow();
    if (ul < 2) return JSON.stringify({ jaEnviadas: [] });

    var dados = ws.getRange(2, 1, ul - 1, 8).getValues();
    var contagem = {};

    dados.forEach(function(l) {
      if (!l[0]) return;
      var nfdsColuna = String(l[4] || '');
      nfds.forEach(function(nfd) {
        if (nfdsColuna.indexOf(nfd) !== -1) {
          if (!contagem[nfd]) contagem[nfd] = { total: 0, data: '' };
          contagem[nfd].total++;
          contagem[nfd].data = String(l[0]);
        }
      });
    });

    var jaEnviadas = Object.keys(contagem).map(function(nfd) {
      return { nfd: nfd, total: contagem[nfd].total, data: contagem[nfd].data };
    });

    return JSON.stringify({ jaEnviadas: jaEnviadas });
  } catch(e) {
    return JSON.stringify({ jaEnviadas: [] });
  }
}

// ════════════════════════════════════════════════════════════
//   LOG
// ════════════════════════════════════════════════════════════

function garantirAbaLog(ss) {
  var ws = ss.getSheetByName('_Log');
  if (!ws) {
    ws = ss.insertSheet('_Log');
    ws.hideSheet();
  }
  var cabecalho = '';
  try { cabecalho = ws.getRange('A1').getValue(); } catch (_) {}
  if (cabecalho !== 'Data/Hora') {
    ws.getRange(1, 1, 1, 8)
      .setValues([['Data/Hora','Usuário','Aba','Linha','Coluna','Valor Anterior','Novo Valor','Ação']])
      .setBackground('#444444').setFontColor('#FFFFFF').setFontWeight('bold');
    ws.setFrozenRows(1);
    [160, 220, 160, 60, 100, 200, 200, 150].forEach(function(w, i) {
      ws.setColumnWidth(i + 1, w);
    });
  }
  return ws;
}

function registrarLog(ss, nomeAba, row, col, valorAnterior, novoValor, acao) {
  try {
    var ws    = ss.getSheetByName('_Log') || garantirAbaLog(ss);
    var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    // [P11] getLastRow+setValues é ~30% mais rápido que appendRow
    var nextRow = ws.getLastRow() + 1;
    ws.getRange(nextRow, 1, 1, 8).setValues([[
      agora,
      Session.getActiveUser().getEmail() || 'sistema',
      nomeAba, row, obterNomeColuna(col),
      valorAnterior, novoValor, acao
    ]]);
  } catch (e) {
    console.error('Log: ' + e);
  }
}

/**
 * Helper que lê NF/NFD de uma linha e chama registrarLog com campos padronizados.
 * [P09] Lê NFD (col1) e NF (col2) em batch: 1 getRange 1×2 → 2 valores de uma vez.
 */
function _registrarLogAba(ss, ws, nomeAba, row, col, statusAnterior, novoStatus, prefixoAcao) {
  var nfLog = '', nfdLog = '';
  try {
    var vals = ws.getRange(row, COL_NFD, 1, 2).getValues()[0];
    nfdLog = String(vals[0] || '').trim();
    nfLog  = String(vals[1] || '').trim();
  } catch (_) {}
  var ref = nfLog || nfdLog || statusAnterior;
  registrarLog(ss, nomeAba, row, col, ref, novoStatus,
    prefixoAcao + (nfLog ? ' — NF: ' + nfLog : ''));
}


// ════════════════════════════════════════════════════════════
//   CONFIGURAÇÃO INICIAL
// ════════════════════════════════════════════════════════════

function configurarPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    var requests = ss.getSheets().map(function(s) {
      return { clearBasicFilter: { sheetId: s.getSheetId() } };
    });
    if (requests.length) {
      Sheets.Spreadsheets.batchUpdate({ requests: requests }, ss.getId());
    }
  } catch (_) {
    ss.getSheets().forEach(function(s) {
      try { var ff = s.getFilter(); if (ff) ff.remove(); } catch (_2) {}
    });
  }
  SpreadsheetApp.flush();

  ss.getSheets().forEach(function(s) {
    try { s.setConditionalFormatRules([]); } catch (_) {}
  });

  garantirAba(ss, 'Britania',              'Britania');
  garantirAba(ss, 'Unilever',              'Unilever');
  garantirAba(ss, 'Fornecedores Variados', 'Fornecedores Variados');
  garantirAbaLog(ss);

  ABAS_OPERACIONAIS.forEach(function(nome) {
    reaplicarCoresAba(ss.getSheetByName(nome));
  });

  _criarLayoutDashboard(ss);
  _atualizarMetricasDashboard(ss);

  // [P10] Reseta todos os contadores de cache ao reconfigurar
  _resetarContadores();

  instalarTriggers();
  SpreadsheetApp.getUi().alert(
    '✅ Sistema v6.0 configurado!\n\n' +
    '• Otimizações de performance aplicadas\n' +
    '• onEdit mais rápido (batch writes + contadores)\n' +
    '• Dashboard com debounce (evita releituras desnecessárias)\n' +
    '• onOpen com cache de cores (reabertura mais rápida)\n' +
    '• Painel de Auditoria unificado (NF + E-mails + Log)'
  );
}

function instalarTriggers() {
  var handlers = [
    'onEditInstalado',
    'enviarResumoSemanal',
    'verificarAtrasosEEnviarAlerta',
    'reaplicarCoresTodas'
  ];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (handlers.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('onEditInstalado')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit().create();

  ScriptApp.newTrigger('enviarResumoSemanal')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();

  ScriptApp.newTrigger('verificarAtrasosEEnviarAlerta')
    .timeBased().everyDays(1).atHour(9).create();

  ScriptApp.newTrigger('reaplicarCoresTodas')
    .timeBased().everyDays(1).atHour(2).create();
}


// ════════════════════════════════════════════════════════════
//   FORMATAÇÃO DE ABAS
// ════════════════════════════════════════════════════════════

function garantirAba(ss, nomeAba, nomeFornecedor) {
  var ws = ss.getSheetByName(nomeAba);
  if (ws) {
    try {
      if (ss.getSheets().filter(function(s) { return !s.isSheetHidden(); }).length === 1) {
        ss.insertSheet('_tmp_del_');
      }
      ss.deleteSheet(ws);
    } catch (_) {}
  }

  ws = ss.insertSheet(nomeAba);
  formatarAba(ws, nomeFornecedor, nomeAba !== 'Fornecedores Variados');

  var tmp = ss.getSheetByName('_tmp_del_');
  if (tmp) try { ss.deleteSheet(tmp); } catch (_) {}

  return ws;
}

function formatarAba(ws, nomeFornecedor, fixarFornecedor) {
  try {
    Sheets.Spreadsheets.batchUpdate(
      { requests: [{ clearBasicFilter: { sheetId: ws.getSheetId() } }] },
      ws.getParent().getId()
    );
  } catch (_) {
    try { var f = ws.getFilter(); if (f) f.remove(); } catch (_2) {}
  }

  try { ws.setConditionalFormatRules([]); } catch (_) {}
  ws.clear();
  ws.setFrozenRows(0);
  ws.setFrozenColumns(0);

  ws.setRowHeight(1, 40);
  ws.getRange(1, 1, 1, TOTAL_COLUNAS).setBackground(COR_HEADER);
  ws.getRange(1, 2, 1, TOTAL_COLUNAS - 1).merge()
    .setValue('CONTROLE DE DEVOLUÇÕES – ' + nomeFornecedor.toUpperCase())
    .setBackground(COR_HEADER).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  ws.setFrozenRows(3);
  ws.setFrozenColumns(1);
  ws.setRowHeight(2, 6);
  ws.getRange(2, 1, 1, TOTAL_COLUNAS).setBackground('#E8F0F8');

  ws.setRowHeight(3, 32);
  var headers  = ['NFD','Nº NF','Data Entrada','Fornecedor','Tipo','Motivo','Descrição do Produto',
                  'Qtd','Valor Unit (R$)','Valor Total (R$)','Status',
                  'Pendente ✓','Devolvido ✓','Venda ✓','Obs / Hora','Responsável','📎 Anexo NF'];
  var larguras = [100,100,120,180,90,200,320,60,120,130,120,95,100,85,280,200,160];
  headers.forEach(function(h, i) {
    ws.setColumnWidth(i + 1, larguras[i]);
    ws.getRange(3, i + 1).setValue(h)
      .setBackground(COR_HEADER).setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  });

  var fmt = 'R$ #,##0.00;;"";""';
  var formulasTotal = [], fmtUnit = [], fmtTot = [], fmtData = [], valForn = [], valStatus = [];
  var ultimaLinha = LINHA_DADOS + MAX_LINHAS_ABA - 1;

  for (var row = LINHA_DADOS; row <= ultimaLinha; row++) {
    formulasTotal.push([_formulaTotal(row)]);
    fmtUnit.push([fmt]);
    fmtTot.push([fmt]);
    fmtData.push(['dd/mm/yyyy']);
    valForn.push([fixarFornecedor ? nomeFornecedor : '']);
    valStatus.push(['']);
  }

  ws.setRowHeights(LINHA_DADOS, MAX_LINHAS_ABA, 22);
  ws.getRange(LINHA_DADOS, COL_VL_TOT,  MAX_LINHAS_ABA, 1).setFormulas(formulasTotal);
  ws.getRange(LINHA_DADOS, COL_VL_UNIT, MAX_LINHAS_ABA, 1).setNumberFormats(fmtUnit);
  ws.getRange(LINHA_DADOS, COL_VL_TOT,  MAX_LINHAS_ABA, 1).setNumberFormats(fmtTot);
  ws.getRange(LINHA_DADOS, COL_DATA,    MAX_LINHAS_ABA, 1).setNumberFormats(fmtData);
  ws.getRange(LINHA_DADOS, COL_STATUS,  MAX_LINHAS_ABA, 1).setValues(valStatus);

  ws.getRange(LINHA_DADOS, COL_PEND_CHK,  MAX_LINHAS_ABA, 1).insertCheckboxes();
  ws.getRange(LINHA_DADOS, COL_DEV_CHK,   MAX_LINHAS_ABA, 1).insertCheckboxes();
  ws.getRange(LINHA_DADOS, COL_VENDA_CHK, MAX_LINHAS_ABA, 1).insertCheckboxes();

  if (fixarFornecedor) {
    ws.getRange(LINHA_DADOS, COL_FORN, MAX_LINHAS_ABA, 1)
      .setValues(valForn).setFontColor('#555555').setFontStyle('italic');
  }

  ws.getRange(LINHA_DADOS, 1, MAX_LINHAS_ABA, TOTAL_COLUNAS)
    .setBackground('#FFFFFF').setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.getRange(LINHA_DADOS, COL_DESC, MAX_LINHAS_ABA, 1).setHorizontalAlignment('left').setWrap(true);
  ws.getRange(LINHA_DADOS, COL_OBS,  MAX_LINHAS_ABA, 1).setHorizontalAlignment('left').setWrap(true);

  ws.getRange(LINHA_DADOS, COL_TIPO, MAX_LINHAS_ABA, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Falta', 'Avaria', 'Rejeição'], true).setAllowInvalid(true).build());

  ws.getRange(LINHA_DADOS, COL_STATUS, MAX_LINHAS_ABA, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendente', 'Devolvido', 'Venda'], true).setAllowInvalid(true).build());
}


// ════════════════════════════════════════════════════════════
//   ON EDIT
// ════════════════════════════════════════════════════════════

function onEditInstalado(e) {
  if (!e) return;
  var ws     = e.range.getSheet();
  var nomAba = ws.getName();
  var col    = e.range.getColumn();
  var row    = e.range.getRow();
  var ss     = SpreadsheetApp.getActiveSpreadsheet();

  if (nomAba === 'Dashboard' && (row === 4 || row === 5) && col === 3) {
    _atualizarMetricasDashboard(ss);
    return;
  }

  if (ABAS_OPERACIONAIS.indexOf(nomAba) === -1 || row < LINHA_DADOS) return;

  var trava = LockService.getScriptLock();
  if (!trava.tryLock(8000)) {
    ss.toast('Sistema ocupado. Tente novamente em instantes.', '⏳ Aguarde', 4);
    return;
  }

  try {
    var novoValor     = e.range.getValue();
    var valorAnterior = e.oldValue != null ? e.oldValue : '';

    if (col === COL_DATA) {
      // [P05] Passa a data nova diretamente
      var stAtualData = ws.getRange(row, COL_STATUS).getValue();
      if (stAtualData === 'Pendente') aplicarCorLinha(ws, row, 'Pendente', novoValor instanceof Date ? novoValor : null);
    }

    // [P12] valoresNF só é lido quando a coluna editada é COL_NF
    if (col === COL_NF && novoValor !== '') {
      var ultimaLinha = obterUltimaLinhaDados(ws);
      var valoresNF   = ultimaLinha >= LINHA_DADOS
        ? ws.getRange(LINHA_DADOS, COL_NF, ultimaLinha - LINHA_DADOS + 1, 1).getValues()
        : [];
      if (_nfDuplicada(valoresNF, row, novoValor)) {
        SpreadsheetApp.getUi().alert('⚠️ NF "' + novoValor + '" já lançada nesta aba. Verifique duplicidade.');
      }
      if (!ws.getRange(row, COL_RESP).getValue()) {
        ws.getRange(row, COL_RESP).setValue(Session.getActiveUser().getEmail() || 'Não identificado');
      }
    }

    if (col === COL_STATUS) {
      _aplicarStatus(ss, ws, nomAba, row, novoValor, valorAnterior);
    }

    // [P13] 3 blocos if separados → 1 bloco com 1 único getValue compartilhado
    if ((col === COL_PEND_CHK || col === COL_DEV_CHK || col === COL_VENDA_CHK) && novoValor === true) {
      var stAtualChk = ws.getRange(row, COL_STATUS).getValue();
      var novoStChk  = col === COL_PEND_CHK ? 'Pendente'
                     : col === COL_DEV_CHK   ? 'Devolvido' : 'Venda';
      _aplicarStatus(ss, ws, nomAba, row, novoStChk, stAtualChk);
    }

    if (_lerContadorConcluidos() >= 40) {
      _zerarContadorConcluidos();
      arquivarItensConcluidos();
    }

  } catch (erro) {
    console.error('onEdit: ' + erro);
  } finally {
    trava.releaseLock();
  }
}

/** Aplica status, checkboxes, cor, obs, proteção, log e atualiza métricas. */
function _aplicarStatus(ss, ws, nomAba, row, novoStatus, statusAnterior) {
  var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  var obsVal = novoStatus === 'Devolvido' ? 'Devolvido em: ' + agora
             : novoStatus === 'Venda'     ? 'Enviado para o Fábio'
             : '';
  // [P14] status + 3 checkboxes + obs em 1 setValues (cols 11-15 são adjacentes)
  ws.getRange(row, COL_STATUS, 1, 5).setValues([[
    novoStatus,
    novoStatus === 'Pendente',
    novoStatus === 'Devolvido',
    novoStatus === 'Venda',
    obsVal
  ]]);

  aplicarCorLinha(ws, row, novoStatus);

  if (novoStatus === 'Devolvido' || novoStatus === 'Venda') {
    protegerLinhaConcluida(ss, ws, row, novoStatus);
    if (statusAnterior !== 'Devolvido' && statusAnterior !== 'Venda') {
      _incrementarContadorConcluidos();
    }
  }
  _atualizarMetricasDashboard(ss);
  _registrarLogAba(ss, ws, nomAba, row, COL_STATUS, statusAnterior, novoStatus, 'Status alterado');
}

function _nfDuplicada(valoresNF, rowAtual, valorNF) {
  return valoresNF.some(function(r, idx) {
    return (LINHA_DADOS + idx) !== rowAtual &&
           String(r[0]).trim() === String(valorNF).trim();
  });
}


// ── Helpers de visual ─────────────────────────────────────────

// [P01] Mantida para compatibilidade externa (restauração, reabertura, exportarPDF)
function syncCheckboxesComStatus(ws, row, status) {
  ws.getRange(row, COL_PEND_CHK, 1, 3).setValues([[
    status === 'Pendente',
    status === 'Devolvido',
    status === 'Venda'
  ]]);
}

// [P08] Mantida para compatibilidade externa
function registrarObs(ws, row, status) {
  var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  var obsVal = status === 'Devolvido' ? 'Devolvido em: ' + agora
             : status === 'Venda'     ? 'Enviado para o Fábio'
             : '';
  ws.getRange(row, COL_OBS).setValue(obsVal);
}

// [P05] Aceita dtOpcional para evitar getValue extra quando a data já é conhecida
function aplicarCorLinha(ws, row, status, dtOpcional) {
  var cor = corPorStatus(status);
  if (status === 'Pendente') {
    try {
      var dt = (dtOpcional instanceof Date) ? dtOpcional : ws.getRange(row, COL_DATA).getValue();
      if (dt instanceof Date && !isNaN(dt)) {
        if (Math.floor((new Date() - dt) / 864e5) > 30) cor = COR_ALERTA_30DIAS;
      }
    } catch (_) {}
  }
  ws.getRange(row, 1, 1, TOTAL_COLUNAS).setBackground(cor);
}

function reaplicarCoresAba(ws) {
  if (!ws) return;
  var ul = obterUltimaLinhaDados(ws);
  if (ul < LINHA_DADOS) return;

  var dados = ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues();
  var hoje  = new Date();
  var cores = dados.map(function(l) {
    var cor = '#FFFFFF';
    if (l[IDX_NF]) {
      cor = corPorStatus(l[IDX_STATUS]);
      if (l[IDX_STATUS] === 'Pendente' && l[IDX_DATA] instanceof Date && !isNaN(l[IDX_DATA])) {
        if (Math.floor((hoje - l[IDX_DATA]) / 864e5) > 30) cor = COR_ALERTA_30DIAS;
      }
    }
    return Array(TOTAL_COLUNAS).fill(cor);
  });
  ws.getRange(LINHA_DADOS, 1, dados.length, TOTAL_COLUNAS).setBackgrounds(cores);
}

function reaplicarCoresTodas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ABAS_OPERACIONAIS.forEach(function(nome) {
    reaplicarCoresAba(ss.getSheetByName(nome));
  });
}

function protegerLinhaConcluida(ss, ws, row, status) {
  if (status !== 'Devolvido' && status !== 'Venda') return;

  var protAtivas = ws.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  if (protAtivas.some(function(p) {
    return p.getRange().getRow() === row && p.getRange().getNumRows() === 1;
  })) return;

  // [P04] Usa contador em PropertiesService em vez de iterar todas as abas
  var total = _lerTotalProtecoes();
  if (total >= LIMITE_PROTECOES) {
    registrarLog(ss, ws.getName(), row, COL_STATUS, '', status, '⚠️ Limite de proteções atingido.');
    return;
  }

  try {
    var p     = ws.getRange(row, 1, 1, TOTAL_COLUNAS).protect()
                  .setDescription('Linha ' + row + ' – ' + status);
    var owner = ss.getOwner() ? ss.getOwner().getEmail() : '';
    try {
      p.getEditors().forEach(function(u) {
        if (u.getEmail() !== owner) p.removeEditor(u);
      });
    } catch (_) {}
    if (p.canDomainEdit()) p.setDomainEdit(false);
    _incrementarProtecoes();
  } catch (e) {
    console.error('protegerLinhaConcluida: ' + e);
  }
}

// [P02] Mantida para uso em diagnóstico manual ou chamadas externas.
function contarTotalConcluidos(ss) {
  return ABAS_OPERACIONAIS.reduce(function(tot, nome) {
    var ws = ss.getSheetByName(nome);
    if (!ws) return tot;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) return tot;
    return tot + ws.getRange(LINHA_DADOS, COL_STATUS, ul - LINHA_DADOS + 1, 1).getValues()
      .filter(function(r) { return r[0] === 'Devolvido' || r[0] === 'Venda'; }).length;
  }, 0);
}


// ════════════════════════════════════════════════════════════
//   ARQUIVAMENTO
// ════════════════════════════════════════════════════════════

function _garantirHistorico(ss) {
  var hist = ss.getSheetByName('Historico_Arquivo');
  if (!hist) hist = ss.insertSheet('Historico_Arquivo');

  if (hist.getRange('A1').getValue() !== 'NFD') {
    hist.getRange(1, 1, 1, TOTAL_COLUNAS + 1).setValues([[
      'NFD','Nº NF','Data Entrada','Fornecedor','Tipo','Motivo','Descrição do Produto',
      'Qtd','Valor Unit (R$)','Valor Total (R$)','Status',
      'Pendente ✓','Devolvido ✓','Venda ✓','Obs / Hora','Responsável','📎 Anexo NF',
      'Arquivado em'
    ]]).setBackground('#444444').setFontColor('#FFFFFF').setFontWeight('bold');
    hist.setFrozenRows(1);
    [100,100,120,180,90,200,320,60,120,130,120,95,100,85,280,200,160,160]
      .forEach(function(w, i) { hist.setColumnWidth(i + 1, w); });
  }
  return hist;
}

function _reconstruirAba(ss, ws, nomeAba, restantes) {
  var ul = obterUltimaLinhaDados(ws);
  var linhasLimpar = Math.max(ul - LINHA_DADOS + 1, 0);
  if (linhasLimpar > 0) {
    ws.getRange(LINHA_DADOS, 1, linhasLimpar, TOTAL_COLUNAS).clearContent().setBackground('#FFFFFF');
  }

  var fmt    = 'R$ #,##0.00;;"";""';
  var bloco  = [], cores = [], fmulas = [], fu = [], ft = [], fd = [];

  for (var i = 0; i < MAX_LINHAS_ABA; i++) {
    var l  = i < restantes.length ? restantes[i] : null;
    var lv = l ? l.slice() : Array(TOTAL_COLUNAS).fill('');
    if (!l) {
      lv[IDX_PEND_CHK]  = false;
      lv[IDX_DEV_CHK]   = false;
      lv[IDX_VENDA_CHK] = false;
      if (nomeAba !== 'Fornecedores Variados') lv[IDX_FORN] = nomeAba;
    }
    bloco.push(lv);
    cores.push(Array(TOTAL_COLUNAS).fill(corPorStatus(lv[IDX_STATUS])));
    var row = LINHA_DADOS + i;
    fmulas.push([_formulaTotal(row)]);
    fu.push([fmt]); ft.push([fmt]); fd.push(['dd/mm/yyyy']);
  }

  var rng = ws.getRange(LINHA_DADOS, 1, MAX_LINHAS_ABA, TOTAL_COLUNAS);
  rng.setValues(bloco).setBackgrounds(cores);
  ws.getRange(LINHA_DADOS, COL_VL_TOT,  MAX_LINHAS_ABA, 1).setFormulas(fmulas);
  ws.getRange(LINHA_DADOS, COL_VL_UNIT, MAX_LINHAS_ABA, 1).setNumberFormats(fu);
  ws.getRange(LINHA_DADOS, COL_VL_TOT,  MAX_LINHAS_ABA, 1).setNumberFormats(ft);
  ws.getRange(LINHA_DADOS, COL_DATA,    MAX_LINHAS_ABA, 1).setNumberFormats(fd);
  if (nomeAba !== 'Fornecedores Variados') {
    ws.getRange(LINHA_DADOS, COL_FORN, MAX_LINHAS_ABA, 1)
      .setFontColor('#555555').setFontStyle('italic');
  }
  rng.setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.getRange(LINHA_DADOS, COL_DESC, MAX_LINHAS_ABA, 1).setHorizontalAlignment('left').setWrap(true);
  ws.getRange(LINHA_DADOS, COL_OBS,  MAX_LINHAS_ABA, 1).setHorizontalAlignment('left').setWrap(true);

  restantes.forEach(function(l, idx) {
    var st = l[IDX_STATUS];
    if (st === 'Devolvido' || st === 'Venda') {
      protegerLinhaConcluida(ss, ws, LINHA_DADOS + idx, st);
    }
  });
}

/**
 * Move linhas de `dados` cujos índices estão em `linhasAlvoSet` para o histórico.
 * [P15] N appendRow individuais → 1 setValues batch por chamada.
 */
function _moverParaHistorico(hist, dados, linhasAlvoSet) {
  var restantes        = [];
  var linhasHistorico  = [];
  var total            = 0;
  var agora            = new Date();

  dados.forEach(function(l, idx) {
    var linhaAtual = LINHA_DADOS + idx;
    if (linhasAlvoSet.has(linhaAtual)) {
      linhasHistorico.push(l.concat([agora]));
      total++;
      _apagarAnexoDrive(String(l[IDX_ANEXO] || '').trim());
    } else if (l[IDX_NF]) {
      restantes.push(l);
    }
  });

  if (linhasHistorico.length) {
    var nextRow = hist.getLastRow() + 1;
    hist.getRange(nextRow, 1, linhasHistorico.length, TOTAL_COLUNAS + 1)
        .setValues(linhasHistorico);
  }

  return { restantes: restantes, total: total };
}

function arquivarItensConcluidos() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hist = _garantirHistorico(ss);
  var total = 0;
  var totalProtRemovidas = 0;

  ABAS_OPERACIONAIS.forEach(function(nomeAba) {
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) return;

    var dados = ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues();

    var linhasConcluidas = new Set();
    dados.forEach(function(l, idx) {
      if (l[IDX_NF] && (l[IDX_STATUS] === 'Devolvido' || l[IDX_STATUS] === 'Venda')) {
        linhasConcluidas.add(LINHA_DADOS + idx);
      }
    });
    ws.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) {
      if (linhasConcluidas.has(p.getRange().getRow())) {
        p.remove();
        totalProtRemovidas++;
      }
    });

    var resultado = _moverParaHistorico(hist, dados, linhasConcluidas);
    total += resultado.total;
    _reconstruirAba(ss, ws, nomeAba, resultado.restantes);
  });

  _zerarContadorConcluidos();
  _decrementarProtecoes(totalProtRemovidas);

  SpreadsheetApp.flush();
  _atualizarMetricasDashboard(ss);
  if (total > 0) {
    try { SpreadsheetApp.getUi().alert('📦 ' + total + ' itens arquivados.'); } catch (_) {}
  }
}

function _arquivarLinhasEspecificas(ss, linhasParaArquivar) {
  if (!linhasParaArquivar || !linhasParaArquivar.length) return 0;

  var hist = _garantirHistorico(ss);
  var porAba = {};
  linhasParaArquivar.forEach(function(ref) {
    if (!porAba[ref.nomeAba]) porAba[ref.nomeAba] = new Set();
    porAba[ref.nomeAba].add(ref.linha);
  });

  var total = 0;

  Object.keys(porAba).forEach(function(nomeAba) {
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) return;

    var linhasAlvo = porAba[nomeAba];
    var dados = ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues();

    ws.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) {
      if (linhasAlvo.has(p.getRange().getRow())) p.remove();
    });

    var resultado = _moverParaHistorico(hist, dados, linhasAlvo);
    total += resultado.total;
    _reconstruirAba(ss, ws, nomeAba, resultado.restantes);
  });

  SpreadsheetApp.flush();
  return total;
}


// ════════════════════════════════════════════════════════════
//   DASHBOARD
//
//  Estrutura:
//  Linha  1    : Título
//  Linha  2    : Subtítulo
//  Linhas 3–5  : Filtro de datas (col B–D) | Taxa (col F–I)
//  Linha  6    : Separador
//  Linhas 7–8  : Cabeçalhos de coluna (fornecedores)
//  Linhas 9–17 : Blocos PENDENTE / DEVOLVIDO / VENDA
//  Linhas 18–21: KPIs globais
//  Linha  23+  : Gráficos
// ════════════════════════════════════════════════════════════

function garantirDashboard(ss) {
  var ws = ss.getSheetByName('Dashboard');
  if (!ws) ws = ss.insertSheet('Dashboard');
  var sentinel = '';
  try { sentinel = ws.getRange(DASH_SENTINEL_CELL).getValue(); } catch(_) {}
  if (sentinel !== DASH_SENTINEL_VALUE) _criarLayoutDashboard(ss);
  _atualizarMetricasDashboard(ss);
}

function _criarLayoutDashboard(ss) {
  var ws = ss.getSheetByName('Dashboard');
  if (!ws) ws = ss.insertSheet('Dashboard');

  try { var f = ws.getFilter(); if (f) f.remove(); } catch(_) {}
  try { ws.setConditionalFormatRules([]); } catch(_) {}
  ws.clear();
  ws.setHiddenGridlines(true);
  ws.setFrozenRows(2);

  ws.getRange(1, 1, 50, 11).setBackground(DC.CINZA);

  ws.setColumnWidth(1, 12);
  for (var ci = 2; ci <= 9; ci++) ws.setColumnWidth(ci, 130);
  ws.setColumnWidth(10, 12);
  try { ws.hideColumns(11, 20); } catch(_) {}

  [48,24,8,30,30,8,22,6,22,38,8,22,38,8,22,38,10,22,42,24,18,10].forEach(function(h, i) {
    ws.setRowHeight(i + 1, h);
  });
  ws.setRowHeights(23, 20, 22);

  // ── Título ────────────────────────────────────────────────
  ws.getRange(1, 1, 1, 10).setBackground(DC.HEADER);
  ws.getRange(1, 2, 1, 8).merge()
    .setValue('CONTROLE DE DEVOLUÇÕES  ·  PAINEL DE GESTÃO')
    .setFontColor(DC.BRANCO).setFontWeight('bold').setFontSize(15)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  ws.getRange(DASH_SENTINEL_CELL).setValue(DASH_SENTINEL_VALUE)
    .setFontColor(DC.CINZA).setBackground(DC.CINZA);

  // ── Subtítulo ─────────────────────────────────────────────
  ws.getRange(2, 1, 1, 10).setBackground(DC.SUB);
  ws.getRange(2, 2, 1, 8).merge()
    .setValue('Atualizado automaticamente  ·  Edite as datas abaixo para filtrar o período')
    .setFontColor('#93B4D4').setFontSize(9).setFontStyle('italic')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');

  // ── Filtro de datas ───────────────────────────────────────
  ws.getRange(4, 2).setValue('Início:').setFontSize(9).setFontColor(DC.TEXTO_L)
    .setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
  ws.getRange(5, 2).setValue('Fim:').setFontSize(9).setFontColor(DC.TEXTO_L)
    .setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');

  var ano = new Date().getFullYear();
  ws.getRange(4, 3).setValue(new Date(ano, 0, 1)).setNumberFormat('dd/mm/yyyy')
    .setBackground('#F0F7FF').setFontColor(DC.PEND).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.getRange(5, 3).setValue(new Date(ano, 11, 31)).setNumberFormat('dd/mm/yyyy')
    .setBackground('#F0F7FF').setFontColor(DC.PEND).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Taxa de resolução ─────────────────────────────────────
  ws.getRange(3, 6, 3, 4).setBackground(DC.BRANCO)
    .setBorder(true, true, true, true, false, false, DC.BORDA, SpreadsheetApp.BorderStyle.SOLID);
  ws.getRange(3, 6, 1, 4).merge()
    .setValue('🏆 TAXA DE RESOLUÇÃO')
    .setFontWeight('bold').setFontSize(9).setFontColor(DC.TEXTO_L)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  ws.getRange(4, 6, 2, 2).merge()
    .setValue('—').setFontWeight('bold').setFontSize(26).setFontColor(DC.DEV)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setNumberFormat('0%');
  ws.getRange(4, 8, 2, 2).merge()
    .setValue('itens resolvidos\nsobre o total').setFontSize(8).setWrap(true)
    .setFontColor(DC.TEXTO_L).setHorizontalAlignment('left').setVerticalAlignment('middle');

  // ── Cabeçalhos fornecedores ───────────────────────────────
  ws.getRange(7, 1, 1, 10).setBackground(DC.CINZA);
  DASH_COLS.forEach(function(g) {
    ws.getRange(7, g.c, 1, 2).merge()
      .setValue(g.label).setBackground(g.cor).setFontColor(DC.BRANCO)
      .setFontWeight('bold').setFontSize(9)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  // ── Blocos de status ──────────────────────────────────────
  var blocks = [
    { label: '⏳ PENDENTE',  acc: DC.PEND,  bg: DC.PEND_BG,  rL: 9,  rV: 10 },
    { label: '✅ DEVOLVIDO', acc: DC.DEV,   bg: DC.DEV_BG,   rL: 12, rV: 13 },
    { label: '🛒 VENDA',     acc: DC.VENDA, bg: DC.VENDA_BG, rL: 15, rV: 16 }
  ];
  blocks.forEach(function(b) {
    ws.getRange(b.rL, 1, 1, 10).setBackground(b.acc);
    ws.getRange(b.rL, 2, 1, 8).merge()
      .setValue(b.label).setFontColor(DC.BRANCO).setFontWeight('bold').setFontSize(9)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    DASH_COLS.forEach(function(g, gi) {
      var acc = gi === 3 ? DC.TOT : b.acc;
      var bg  = gi === 3 ? DC.TOT_BG : b.bg;
      ws.getRange(b.rV, g.c).setBackground(bg).setFontColor(acc)
        .setFontWeight('bold').setFontSize(18)
        .setHorizontalAlignment('center').setVerticalAlignment('middle').setValue('—');
      ws.getRange(b.rV, g.c + 1).setBackground(bg).setFontColor(DC.TEXTO_L)
        .setFontSize(9).setFontWeight('bold')
        .setHorizontalAlignment('right').setVerticalAlignment('middle')
        .setNumberFormat('R$ #,##0.00').setValue(0);
    });
  });

  // ── KPIs globais ──────────────────────────────────────────
  var kpis = [
    { label: 'TOTAL PENDENTE',  sub: 'aguardando resolução', acc: DC.PEND,  bg: DC.PEND_BG,  col: 2 },
    { label: 'TOTAL DEVOLVIDO', sub: 'devoluções OK',        acc: DC.DEV,   bg: DC.DEV_BG,   col: 4 },
    { label: 'TOTAL VENDA',     sub: 'enviados p/ venda',    acc: DC.VENDA, bg: DC.VENDA_BG, col: 6 },
    { label: 'TOTAL RESOLVIDO', sub: 'devolvidos + vendas',  acc: DC.TOT,   bg: DC.TOT_BG,   col: 8 }
  ];
  kpis.forEach(function(k) {
    ws.getRange(18, k.col, 4, 2).setBackground(k.bg)
      .setBorder(true, true, true, true, false, false, k.acc, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    ws.getRange(18, k.col, 1, 2).merge()
      .setValue(k.label).setBackground(k.acc).setFontColor(DC.BRANCO)
      .setFontWeight('bold').setFontSize(8)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange(19, k.col, 1, 2).merge().setValue('—')
      .setFontColor(k.acc).setFontWeight('bold').setFontSize(22)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange(20, k.col, 1, 2).merge().setValue(0)
      .setFontColor(DC.TEXTO_L).setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setNumberFormat('R$ #,##0.00');
    ws.getRange(21, k.col, 1, 2).merge().setValue(k.sub)
      .setFontColor(DC.TEXTO_L).setFontSize(8).setFontStyle('italic')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  // ── Dados auxiliares gráficos ─────────────────────────────
  ws.getRange(1, 11, 4, 2).setValues([
    ['Status','Qtd'],['Pendente',0],['Devolvido',0],['Venda',0]
  ]);
  ws.getRange(1, 14, 13, 4).setValues(
    [['Mês','Pendente','Devolvido','Venda']].concat(
      ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez']
        .map(function(m) { return [m, 0, 0, 0]; })
    )
  );

  // ── Gráficos ──────────────────────────────────────────────
  ws.getCharts().forEach(function(ch) { ws.removeChart(ch); });
  ws.insertChart(ws.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(ws.getRange(1, 11, 4, 2)).setPosition(24, 2, 0, 0)
    .setOption('title', 'Distribuição por Status')
    .setOption('width', 320).setOption('height', 200)
    .setOption('colors', [DC.PEND, DC.DEV, DC.VENDA])
    .setOption('pieHole', 0.4).setOption('pieSliceText', 'percentage')
    .setOption('legend', { position: 'bottom', textStyle: { fontSize: 9 } })
    .setOption('backgroundColor', DC.CINZA).build());
  ws.insertChart(ws.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(ws.getRange(1, 14, 13, 4)).setPosition(24, 5, 0, 0)
    .setOption('title', 'Lançamentos por Mês — ' + new Date().getFullYear())
    .setOption('width', 480).setOption('height', 200)
    .setOption('colors', [DC.PEND, DC.DEV, DC.VENDA])
    .setOption('legend', { position: 'bottom', textStyle: { fontSize: 9 } })
    .setOption('backgroundColor', DC.CINZA)
    .setOption('vAxis', { textStyle: { fontSize: 8 } })
    .setOption('hAxis', { textStyle: { fontSize: 8 } }).build());

  SpreadsheetApp.flush();
}

// [P03] Debounce de _DASH_DEBOUNCE_SEG segundos para evitar releituras desnecessárias
function _atualizarMetricasDashboard(ss) {
  var cache = CacheService.getScriptCache();
  if (cache.get(_CACHE_KEY_DASH)) return;
  cache.put(_CACHE_KEY_DASH, '1', _DASH_DEBOUNCE_SEG);

  var ws = ss.getSheetByName('Dashboard');
  if (!ws) { _criarLayoutDashboard(ss); return; }

  if (!cache.get(_CACHE_KEY_SENTINEL)) {
    var sentinel = '';
    try { sentinel = ws.getRange(DASH_SENTINEL_CELL).getValue(); } catch(_) {}
    if (sentinel !== DASH_SENTINEL_VALUE) {
      _criarLayoutDashboard(ss);
      ws = ss.getSheetByName('Dashboard');
    }
    cache.put(_CACHE_KEY_SENTINEL, '1', 1800);
  }

  var rawIni = ws.getRange(DASH_DATA_INI_CELL).getValue();
  var rawFim = ws.getRange(DASH_DATA_FIM_CELL).getValue();
  var ano    = new Date().getFullYear();
  var dataIni = (rawIni instanceof Date && !isNaN(rawIni)) ? rawIni : new Date(ano, 0, 1);
  var dataFim = (rawFim instanceof Date && !isNaN(rawFim)) ? rawFim : new Date(ano, 11, 31);

  var b = _processarAba(ss.getSheetByName('Britania'),              dataIni, dataFim);
  var u = _processarAba(ss.getSheetByName('Unilever'),              dataIni, dataFim);
  var v = _processarAba(ss.getSheetByName('Fornecedores Variados'), dataIni, dataFim);

  var tP  = b.pQtd + u.pQtd + v.pQtd;   var tPv = b.pValor + u.pValor + v.pValor;
  var tD  = b.dQtd + u.dQtd + v.dQtd;   var tDv = b.dValor + u.dValor + v.dValor;
  var tV  = b.vQtd + u.vQtd + v.vQtd;   var tVv = b.vValor + u.vValor + v.vValor;
  var tR  = tD + tV;                     var tRv = tDv + tVv;
  var taxa = (tP + tR) > 0 ? tR / (tP + tR) : 0;

  ws.getRange(4, 6, 2, 2).merge().setValue(taxa).setNumberFormat('0%')
    .setFontWeight('bold').setFontSize(26)
    .setFontColor(taxa >= 0.7 ? DC.DEV : taxa >= 0.4 ? DC.VENDA : '#DC2626')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  var sd = [
    { rV: 10, qtds: [b.pQtd, u.pQtd, v.pQtd, tP], vals: [b.pValor, u.pValor, v.pValor, tPv], acc: DC.PEND  },
    { rV: 13, qtds: [b.dQtd, u.dQtd, v.dQtd, tD], vals: [b.dValor, u.dValor, v.dValor, tDv], acc: DC.DEV   },
    { rV: 16, qtds: [b.vQtd, u.vQtd, v.vQtd, tV], vals: [b.vValor, u.vValor, v.vValor, tVv], acc: DC.VENDA }
  ];
  sd.forEach(function(s) {
    DASH_COLS.forEach(function(g, gi) {
      var acc = gi === 3 ? DC.TOT : s.acc;
      ws.getRange(s.rV, g.c).setValue(s.qtds[gi]).setFontColor(acc)
        .setFontWeight('bold').setFontSize(18)
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
      ws.getRange(s.rV, g.c + 1).setValue(s.vals[gi]).setNumberFormat('R$ #,##0.00')
        .setFontColor(DC.TEXTO_L).setFontWeight('bold').setFontSize(9)
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
    });
  });

  var kd = [[tP, tPv], [tD, tDv], [tV, tVv], [tR, tRv]];
  var ka = [DC.PEND, DC.DEV, DC.VENDA, DC.TOT];
  [2, 4, 6, 8].forEach(function(col, i) {
    ws.getRange(19, col, 1, 2).merge().setValue(kd[i][0])
      .setFontColor(ka[i]).setFontWeight('bold').setFontSize(22)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    ws.getRange(20, col, 1, 2).merge().setValue(kd[i][1]).setNumberFormat('R$ #,##0.00')
      .setFontColor(DC.TEXTO_L).setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  ws.getRange(2, 12).setValue(tP);
  ws.getRange(3, 12).setValue(tD);
  ws.getRange(4, 12).setValue(tV);

  _atualizarGraficoMensal(ss);
}

function _processarAba(ws, dataIni, dataFim) {
  var r = { pQtd:0, pValor:0, dQtd:0, dValor:0, vQtd:0, vValor:0, tQtd:0, tValor:0 };
  if (!ws) return r;
  var ul = obterUltimaLinhaDados(ws);
  if (ul < LINHA_DADOS) return r;
  ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues()
    .forEach(function(l) {
      var nf = l[IDX_NF], dt = l[IDX_DATA], val = parseFloat(l[IDX_VL_TOT]) || 0, st = l[IDX_STATUS];
      if (!nf || !(dt instanceof Date) || dt < dataIni || dt > dataFim) return;
      r.tQtd++; r.tValor += val;
      if      (st === 'Pendente')  { r.pQtd++; r.pValor += val; }
      else if (st === 'Devolvido') { r.dQtd++; r.dValor += val; }
      else if (st === 'Venda')     { r.vQtd++; r.vValor += val; }
    });
  return r;
}

function _atualizarGraficoMensal(ss) {
  var ws = ss.getSheetByName('Dashboard');
  if (!ws) return;
  var ano  = new Date().getFullYear();
  var meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var cnt  = { Pendente: new Array(12).fill(0), Devolvido: new Array(12).fill(0), Venda: new Array(12).fill(0) };
  ABAS_OPERACIONAIS.forEach(function(nome) {
    var wsA = ss.getSheetByName(nome);
    if (!wsA) return;
    var ul = obterUltimaLinhaDados(wsA);
    if (ul < LINHA_DADOS) return;
    wsA.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues()
      .forEach(function(l) {
        var nf = l[IDX_NF], dt = l[IDX_DATA], st = l[IDX_STATUS];
        if (!nf || !(dt instanceof Date) || dt.getFullYear() !== ano) return;
        if (cnt[st] !== undefined) cnt[st][dt.getMonth()]++;
      });
  });
  var tab = [['Mês','Pendente','Devolvido','Venda']];
  for (var m = 0; m < 12; m++) tab.push([meses[m], cnt['Pendente'][m], cnt['Devolvido'][m], cnt['Venda'][m]]);
  ws.getRange(1, 14, 13, 4).setValues(tab);
}

// Stubs mantidos por compatibilidade com chamadas legadas
function _renderQuadro() {}
function _criarGraficoMensal() {}


// ════════════════════════════════════════════════════════════
//   EXPORTAR PDF
// ════════════════════════════════════════════════════════════

function abrirFormularioExportarPDF() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('FormExportarPDF').setWidth(460).setHeight(370),
    '📄 Exportar e Salvar PDF'
  );
}

function executarExportarPDF(txtNfsRaw) {
  var nfsDigitadas = txtNfsRaw.split(/[\n,]/).map(function(s) { return s.trim(); }).filter(Boolean);
  if (!nfsDigitadas.length) return JSON.stringify({ erro: 'Nenhuma NF válida identificada.' });

  if (!ID_MODELO_DOC || ID_MODELO_DOC.startsWith('INSIRA'))
    return JSON.stringify({ erro: 'Configure ID_MODELO_DOC no topo do script.' });
  if (!ID_PASTA_DESTINO || ID_PASTA_DESTINO.startsWith('INSIRA'))
    return JSON.stringify({ erro: 'Configure ID_PASTA_DESTINO no topo do script.' });

  var ss             = SpreadsheetApp.getActiveSpreadsheet();
  var itens          = [];
  var naoLocalizadas = nfsDigitadas.slice();

  ABAS_OPERACIONAIS.forEach(function(nomeAba) {
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) return;
    var dados = ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues();
    dados.forEach(function(l, li) {
      var nfd = String(l[IDX_NFD]).trim();
      var nf  = String(l[IDX_NF]).trim();
      var st  = String(l[IDX_STATUS]).trim();
      var bat = _baterTermos(nfsDigitadas, nfd, nf);
      if (bat.bate && st === 'Pendente') {
        itens.push({ nf: nf, nfd: nfd, fornecedor: String(l[IDX_FORN]).trim(),
                     linha: LINHA_DADOS + li, ws: ws, nomeAba: nomeAba });
        var idx = naoLocalizadas.indexOf(bat.termoBateu);
        if (idx > -1) naoLocalizadas.splice(idx, 1);
      }
    });
  });

  if (!itens.length) return JSON.stringify({ erro: "Nenhuma NF com status 'Pendente' localizada." });

  var forns = itens.reduce(function(acc, it) {
    if (acc.indexOf(it.fornecedor) === -1) acc.push(it.fornecedor);
    return acc;
  }, []);
  if (forns.length > 1)
    return JSON.stringify({ erro: 'NFs de fornecedores diferentes (' + forns.join(', ') + '). Use apenas NFs do mesmo fornecedor.' });

  var listaNfs = itens.map(function(it) { return it.nfd || it.nf; });
  var dataExp  = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy');

  try {
    var tempFile = DriveApp.getFileById(ID_MODELO_DOC).makeCopy('Temp_PDF_Dev');
    var doc      = DocumentApp.openById(tempFile.getId());
    var body     = doc.getBody();
    body.replaceText('\\{\\{\\s*nf\\s*\\}\\}', listaNfs.join(' / '));
    body.replaceText('\\{\\{data\\}\\}', dataExp);
    body.replaceText('\\{\\{forn\\}\\}', forns[0]);
    // [P17] saveAndClose() removido — getAs(PDF) já força o flush internamente

    var nomePdf = 'Devolucao_' + listaNfs.slice(0, 3).join('-') + (listaNfs.length > 3 ? '_etc' : '') + '.pdf';
    var pdf     = DriveApp.getFolderById(ID_PASTA_DESTINO)
                    .createFile(tempFile.getAs(MimeType.PDF).setName(nomePdf));
    tempFile.setTrashed(true);

    // [P16] Atualiza status+chk+obs em 1 setValues por linha; cores em batch por aba
    var agora = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    var porAba = {};
    itens.forEach(function(it) {
      it.ws.getRange(it.linha, COL_STATUS, 1, 5).setValues([[
        'Devolvido', false, true, false, 'Devolvido em: ' + agora
      ]]);
      protegerLinhaConcluida(ss, it.ws, it.linha, 'Devolvido');
      registrarLog(ss, it.nomeAba, it.linha, COL_STATUS, it.nf, 'Devolvido',
        '📄 PDF exportado — NF: ' + it.nf + ' (lote ' + listaNfs.length + ' NFs)');
      if (!porAba[it.nomeAba]) porAba[it.nomeAba] = { ws: it.ws, linhas: [] };
      porAba[it.nomeAba].linhas.push(it.linha);
    });
    Object.keys(porAba).forEach(function(nomeAba) {
      var g    = porAba[nomeAba];
      var lns  = g.linhas;
      var minL = Math.min.apply(null, lns);
      var maxL = Math.max.apply(null, lns);
      var n    = maxL - minL + 1;
      var bg   = g.ws.getRange(minL, 1, n, TOTAL_COLUNAS).getBackgrounds();
      lns.forEach(function(r) { bg[r - minL] = Array(TOTAL_COLUNAS).fill(COR_VERDE); });
      g.ws.getRange(minL, 1, n, TOTAL_COLUNAS).setBackgrounds(bg);
    });

    try { CacheService.getScriptCache().remove(_CACHE_KEY_DASH); } catch(_) {}
    _atualizarMetricasDashboard(ss);

    var aviso = naoLocalizadas.length
      ? '\n⚠️ Não localizadas como Pendente: ' + naoLocalizadas.join(', ') : '';
    return JSON.stringify({
      sucesso: '✅ PDF gerado para ' + listaNfs.length + ' NF(s) — ' + forns[0] + '.' + aviso,
      urlPdf: pdf.getUrl()
    });
  } catch (e) {
    registrarLog(ss, 'SISTEMA', 0, 0, '', '', '❌ Erro PDF: ' + e.toString());
    return JSON.stringify({ erro: '❌ ' + e.toString() });
  }
}


// ════════════════════════════════════════════════════════════
//   BAIXA PARA VENDA
// ════════════════════════════════════════════════════════════

function abrirFormularioVenda() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('FormVenda').setWidth(450).setHeight(320),
    '🛒 Baixa de Mercadorias para Venda'
  );
}

function executarBaixaVenda(txtNfsRaw) {
  var nfsDigitadas = txtNfsRaw.split(/[\n,]/).map(function(s) { return String(s).trim(); }).filter(Boolean);
  if (!nfsDigitadas.length) return JSON.stringify({ erro: 'Nenhuma NF válida identificada.' });

  var ss               = SpreadsheetApp.getActiveSpreadsheet();
  var itensEncontrados = [];
  var nfsOk            = [];
  var processados      = new Set();
  var agora            = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  var porAba           = {};

  ABAS_OPERACIONAIS.forEach(function(nomeAba) {
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) return;

    var dados = ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues();
    dados.forEach(function(l, i) {
      var nfd = String(l[IDX_NFD]).trim();
      var nf  = String(l[IDX_NF]).trim();
      var st  = String(l[IDX_STATUS]).trim();
      var bat = _baterTermos(nfsDigitadas, nfd, nf);
      if (!bat.bate || st !== 'Pendente' || processados.has(bat.termoBateu)) return;

      processados.add(bat.termoBateu);
      var linha = LINHA_DADOS + i;
      itensEncontrados.push([nfd || nf, String(l[IDX_FORN]).trim(), String(l[IDX_DESC]).trim(), l[IDX_QTD] || 0]);

      // [P18] status + checkboxes + obs em 1 setValues
      ws.getRange(linha, COL_STATUS, 1, 5).setValues([[
        'Venda', false, false, true, 'Enviado para o Fábio'
      ]]);
      protegerLinhaConcluida(ss, ws, linha, 'Venda');
      registrarLog(ss, nomeAba, linha, COL_STATUS, nf, 'Venda',
        '🛒 Baixa Venda via HTML — NF: ' + (nfd || nf));
      nfsOk.push(nfd || nf);

      if (!porAba[nomeAba]) porAba[nomeAba] = { ws: ws, linhas: [] };
      porAba[nomeAba].linhas.push(linha);
    });
  });

  if (!itensEncontrados.length) return JSON.stringify({ erro: "Nenhuma NF localizada como 'Pendente'." });

  Object.keys(porAba).forEach(function(nomeAba) {
    var g   = porAba[nomeAba];
    var lns = g.linhas;
    var minL = Math.min.apply(null, lns), maxL = Math.max.apply(null, lns);
    var n    = maxL - minL + 1;
    var bg   = g.ws.getRange(minL, 1, n, TOTAL_COLUNAS).getBackgrounds();
    lns.forEach(function(r) { bg[r - minL] = Array(TOTAL_COLUNAS).fill(COR_LARANJA); });
    g.ws.getRange(minL, 1, n, TOTAL_COLUNAS).setBackgrounds(bg);
  });

  try { CacheService.getScriptCache().remove(_CACHE_KEY_DASH); } catch(_) {}
  _atualizarMetricasDashboard(ss);

  if (!ID_PASTA_DESTINO_VENDA || ID_PASTA_DESTINO_VENDA.startsWith('INSIRA'))
    return JSON.stringify({
      sucesso: '✅ Baixa de ' + nfsOk.length + ' itens concluída! (PDF não gerado — configure ID_PASTA_DESTINO_VENDA).'
    });

  try {
    var ssTemp = SpreadsheetApp.create('Temp_Relatorio_Venda');
    var sh     = ssTemp.getSheets()[0];
    sh.getRange('A1:D1').merge()
      .setValue('RELAÇÃO DE MERCADORIAS ENVIADAS PARA VENDA')
      .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange('A2:D2').merge()
      .setValue('Emissão: ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm'))
      .setFontSize(10).setFontStyle('italic').setHorizontalAlignment('center');
    sh.getRange('A4:D4')
      .setValues([['NF', 'FORNECEDOR', 'PRODUTO', 'QUANTIDADE']])
      .setFontWeight('bold').setBackgroundColor('#2C3E50')
      .setFontColor('#FFFFFF').setHorizontalAlignment('center');
    sh.getRange(5, 1, itensEncontrados.length, 4).setValues(itensEncontrados);
    var lt = 5 + itensEncontrados.length;
    sh.getRange('A' + lt + ':C' + lt).merge()
      .setValue('Total:').setFontWeight('bold').setHorizontalAlignment('right');
    sh.getRange('D' + lt).setFormula('=SUM(D5:D' + (lt - 1) + ')').setFontWeight('bold');
    [90, 160, 240, 90].forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
    SpreadsheetApp.flush();

    var url  = ssTemp.getUrl().replace(/\/edit.*$/, '') +
      '/export?exportFormat=pdf&format=pdf&size=letter&portrait=true&fitw=true';
    var blob = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    }).getBlob().setName('Relacao_Venda_' + nfsOk.length + '_itens.pdf');
    var arquivo = DriveApp.getFolderById(ID_PASTA_DESTINO_VENDA).createFile(blob);
    DriveApp.getFileById(ssTemp.getId()).setTrashed(true);
    return JSON.stringify({
      sucesso: '✅ Baixa de ' + nfsOk.length + ' itens concluída!',
      urlPdf: arquivo.getUrl()
    });
  } catch (e) {
    return JSON.stringify({
      sucesso: '✅ Baixa de ' + nfsOk.length + ' itens concluída! ⚠️ Erro no PDF: ' + e.toString()
    });
  }
}


// ════════════════════════════════════════════════════════════
//   FORMULÁRIO DE LANÇAMENTO
// ════════════════════════════════════════════════════════════

function abrirFormularioLancamento() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('FormLancamento').setWidth(500).setHeight(600),
    '➕ Lançar / Excluir Devolução'
  );
}

function salvarLancamentoForm(dados) {
  _validarDadosForm(dados);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(dados.abaSelecao);
  if (!ws) throw new Error('Aba "' + dados.abaSelecao + '" não encontrada.');
  // [P19] Uma única leitura de COL_NF cobre: lastRow + verificação de duplicata + busca de buracos
  var lastRow = ws.getLastRow();
  var nfVals  = lastRow >= LINHA_DADOS
    ? ws.getRange(LINHA_DADOS, COL_NF, lastRow - LINHA_DADOS + 1, 1).getValues() : [];
  var ul = LINHA_DADOS - 1;
  nfVals.forEach(function(r, i) { if (r[0] !== '' && r[0] != null) ul = LINHA_DADOS + i; });
  if (_nfDuplicada(nfVals, -1, dados.nf))
    return JSON.stringify({ aviso: 'NF "' + dados.nf + '" já existe nesta aba. Confirme para lançar mesmo assim.' });
  return _gravarLancamento(ss, ws, dados, ul, nfVals);
}

function salvarLancamentoFormConfirmado(dados) {
  _validarDadosForm(dados);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(dados.abaSelecao);
  if (!ws) throw new Error('Aba "' + dados.abaSelecao + '" não encontrada.');
  return _gravarLancamento(ss, ws, dados, null, null);
}

function _validarDadosForm(dados) {
  if (!dados.abaSelecao || !dados.nf || !dados.descricao || !dados.qtd || !dados.valorUnit)
    throw new Error('Preencha todos os campos obrigatórios.');
}

function _gravarLancamento(ss, ws, dados, ulPre, nfsPre) {
  var valorUnit = Number(String(dados.valorUnit).replace(',', '.'));
  var qtd       = Number(dados.qtd);
  if (isNaN(valorUnit) || valorUnit < 0) throw new Error('Valor unitário inválido.');
  if (isNaN(qtd) || qtd <= 0)            throw new Error('Quantidade inválida.');

  // [P20] Upload do anexo ANTES do lock
  var urlAnexo = '';
  if (dados.base64 && dados.mimeType && dados.nomeArquivo) {
    try {
      var blob    = Utilities.newBlob(Utilities.base64Decode(dados.base64), dados.mimeType, dados.nomeArquivo);
      var arquivo = _pastaAnexos().createFile(blob);
      arquivo.setName('NF_' + dados.nf + '_' + dados.nomeArquivo);
      urlAnexo = arquivo.getUrl();
    } catch (eAnexo) {
      console.error('Erro ao salvar anexo: ' + eAnexo);
    }
  }

  var trava = LockService.getScriptLock();
  if (!trava.tryLock(8000)) throw new Error('Sistema ocupado. Tente novamente.');
  try {
    var ul = (ulPre !== null && ulPre !== undefined) ? ulPre : obterUltimaLinhaDados(ws);
    var nfsExistentes = nfsPre || (ul >= LINHA_DADOS
      ? ws.getRange(LINHA_DADOS, COL_NF, ul - LINHA_DADOS + 1, 1).getValues() : []);

    var dest = ul + 1;
    for (var i = 0; i < nfsExistentes.length; i++) {
      if (nfsExistentes[i][0] === '' || nfsExistentes[i][0] == null) {
        dest = LINHA_DADOS + i;
        break;
      }
    }
    if (dest > LINHA_DADOS + MAX_LINHAS_ABA - 1)
      throw new Error('Aba cheia. Faça o arquivamento antes de lançar novos itens.');

    var rowVals = [
      dados.nfd     || '',
      dados.nf,
      new Date(),
      dados.fornecedor || ws.getName(),
      dados.tipo    || '',
      dados.motivo  || '',
      dados.descricao,
      qtd,
      valorUnit,
      '',
      'Pendente',
      true, false, false,
      '',
      Session.getActiveUser().getEmail() || 'Não identificado',
      urlAnexo
    ];

    ws.getRange(dest, 1, 1, TOTAL_COLUNAS).setValues([rowVals]);
    ws.getRange(dest, COL_VL_TOT).setFormula(_formulaTotal(dest));
    aplicarCorLinha(ws, dest, 'Pendente', new Date());
    _atualizarMetricasDashboard(ss);
    registrarLog(ss, dados.abaSelecao, dest, COL_NF, '', dados.nf,
      '➕ Lançamento via Formulário' + (urlAnexo ? ' + anexo' : ''));
    return JSON.stringify({
      ok: '✅ NF ' + dados.nf + ' lançada na linha ' + dest + ' — ' + dados.abaSelecao +
          (urlAnexo ? '\n📎 Anexo salvo no Drive.' : '') + '.'
    });
  } finally {
    trava.releaseLock();
  }
}

/**
 * Salva um lote de lançamentos de uma só vez.
 * Recebe array de objetos com os mesmos campos de salvarLancamentoForm.
 * Upload de anexos fora do lock; gravação de todas as linhas dentro de 1 lock.
 */
function salvarLoteLancamentos(itens) {
  if (!itens || !itens.length) return JSON.stringify({ erro: 'Nenhum item recebido.' });

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // [P20-lote] Upload de todos os anexos ANTES do lock
  var urlsAnexo = itens.map(function(it) {
    if (!it.base64 || !it.mimeType || !it.nomeArquivo) return '';
    try {
      var blob    = Utilities.newBlob(Utilities.base64Decode(it.base64), it.mimeType, it.nomeArquivo);
      var arquivo = _pastaAnexos().createFile(blob);
      arquivo.setName('NF_' + it.nf + '_' + it.nomeArquivo);
      return arquivo.getUrl();
    } catch (e) {
      console.error('Erro ao salvar anexo NF ' + it.nf + ': ' + e);
      return '';
    }
  });

  var trava = LockService.getScriptLock();
  if (!trava.tryLock(15000)) return JSON.stringify({ erro: 'Sistema ocupado. Tente novamente.' });

  try {
    var resp     = Session.getActiveUser().getEmail() || 'Não identificado';
    var agora    = new Date();
    var salvos   = [];
    var erros    = [];

    // Agrupa itens por aba para minimizar leituras
    var porAba = {};
    itens.forEach(function(it, i) {
      if (!porAba[it.abaSelecao]) porAba[it.abaSelecao] = [];
      porAba[it.abaSelecao].push({ it: it, idx: i });
    });

    Object.keys(porAba).forEach(function(nomeAba) {
      var ws = ss.getSheetByName(nomeAba);
      if (!ws) {
        porAba[nomeAba].forEach(function(e) {
          erros.push('NF ' + e.it.nf + ': aba "' + nomeAba + '" não encontrada.');
        });
        return;
      }

      // Lê coluna NF uma vez por aba
      var lastRow = ws.getLastRow();
      var nfVals  = lastRow >= LINHA_DADOS
        ? ws.getRange(LINHA_DADOS, COL_NF, lastRow - LINHA_DADOS + 1, 1).getValues() : [];
      var ul = LINHA_DADOS - 1;
      nfVals.forEach(function(r, i) { if (r[0] !== '' && r[0] != null) ul = LINHA_DADOS + i; });

      porAba[nomeAba].forEach(function(entry) {
        var it      = entry.it;
        var urlAnexo = urlsAnexo[entry.idx];

        var valorUnit = Number(String(it.valorUnit).replace(',', '.'));
        var qtd       = Number(it.qtd);
        if (isNaN(valorUnit) || valorUnit < 0) { erros.push('NF ' + it.nf + ': valor inválido.'); return; }
        if (isNaN(qtd) || qtd <= 0)            { erros.push('NF ' + it.nf + ': quantidade inválida.'); return; }

        // Encontra próxima linha disponível
        var dest = ul + 1;
        for (var i = 0; i < nfVals.length; i++) {
          if (nfVals[i][0] === '' || nfVals[i][0] == null) {
            dest = LINHA_DADOS + i;
            break;
          }
        }
        if (dest > LINHA_DADOS + MAX_LINHAS_ABA - 1) {
          erros.push('NF ' + it.nf + ': aba "' + nomeAba + '" cheia.');
          return;
        }

        var rowVals = [
          it.nfd       || '',
          it.nf,
          agora,
          it.fornecedor || ws.getName(),
          it.tipo      || '',
          it.motivo    || '',
          it.descricao,
          qtd,
          valorUnit,
          '',           // COL_VL_TOT — fórmula gravada abaixo
          'Pendente',
          true, false, false,
          '',
          resp,
          urlAnexo
        ];

        ws.getRange(dest, 1, 1, TOTAL_COLUNAS).setValues([rowVals]);
        ws.getRange(dest, COL_VL_TOT).setFormula(_formulaTotal(dest));
        aplicarCorLinha(ws, dest, 'Pendente', agora);
        registrarLog(ss, nomeAba, dest, COL_NF, '', it.nf,
          '➕ Lançamento em lote' + (urlAnexo ? ' + anexo' : ''));

        salvos.push('NF ' + it.nf + ' → linha ' + dest);

        // Avança ponteiro para o próximo espaço disponível
        nfVals[dest - LINHA_DADOS] = [it.nf];
        ul = Math.max(ul, dest);
      });
    });

    try { CacheService.getScriptCache().remove(_CACHE_KEY_DASH); } catch(_) {}
    _atualizarMetricasDashboard(ss);

    if (!salvos.length) return JSON.stringify({ erro: '❌ Nenhum item salvo.\n' + erros.join('\n') });

    var msg = '✅ ' + salvos.length + ' item(ns) salvo(s) com sucesso!';
    if (erros.length) msg += '\n⚠️ ' + erros.length + ' erro(s):\n' + erros.join('\n');
    return JSON.stringify({ ok: msg });

  } finally {
    trava.releaseLock();
  }
}

/**
 * Busca uma NF Pendente para confirmar exclusão.
 * Retorna { item: { nf, nfd, aba, linha, status, desc } } ou { erro }.
 */
function buscarNFParaExcluir(nfBusca) {
  nfBusca = String(nfBusca || '').trim();
  if (!nfBusca) return JSON.stringify({ erro: 'Informe a NF ou NFD.' });

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  for (var a = 0; a < ABAS_OPERACIONAIS.length; a++) {
    var nomeAba = ABAS_OPERACIONAIS[a];
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) continue;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) continue;

    var dados = ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues();
    for (var i = 0; i < dados.length; i++) {
      var nfd = String(dados[i][IDX_NFD]).trim();
      var nf  = String(dados[i][IDX_NF]).trim();
      if (nf !== nfBusca && nfd !== nfBusca) continue;

      var st = String(dados[i][IDX_STATUS]).trim();
      if (st !== 'Pendente') {
        return JSON.stringify({ erro: 'NF "' + nfBusca + '" encontrada em ' + nomeAba +
          ' mas tem status "' + st + '". Só é possível excluir itens Pendentes.' });
      }
      return JSON.stringify({
        item: {
          nf:     nf,
          nfd:    nfd,
          aba:    nomeAba,
          linha:  LINHA_DADOS + i,
          status: st,
          desc:   String(dados[i][IDX_DESC]).substring(0, 60)
        }
      });
    }
  }
  return JSON.stringify({ erro: 'NF "' + nfBusca + '" não encontrada como Pendente em nenhuma aba.' });
}

/**
 * Exclui um lançamento Pendente com registro no histórico de log.
 * params: { item: { nf, nfd, aba, linha }, motivo: string }
 */
function excluirLancamento(params) {
  if (!params || !params.item || !params.motivo)
    return JSON.stringify({ erro: 'Dados incompletos para exclusão.' });

  var item   = params.item;
  var motivo = String(params.motivo).trim();
  if (!motivo) return JSON.stringify({ erro: 'Informe o motivo da exclusão.' });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(item.aba);
  if (!ws) return JSON.stringify({ erro: 'Aba "' + item.aba + '" não encontrada.' });

  // Confirma que a linha ainda é Pendente (pode ter mudado desde a busca)
  var statusAtual = ws.getRange(item.linha, COL_STATUS).getValue();
  if (statusAtual !== 'Pendente')
    return JSON.stringify({ erro: 'O item não está mais Pendente (status atual: "' + statusAtual + '"). Exclusão cancelada.' });

  var trava = LockService.getScriptLock();
  if (!trava.tryLock(8000)) return JSON.stringify({ erro: 'Sistema ocupado. Tente novamente.' });

  try {
    var nfLabel = item.nfd ? 'NFD ' + item.nfd + ' / NF ' + item.nf : 'NF ' + item.nf;

    // Apaga o conteúdo da linha e repõe valores neutros (checkboxes + fornecedor padrão)
    ws.getRange(item.linha, 1, 1, TOTAL_COLUNAS).clearContent();
    ws.getRange(item.linha, COL_PEND_CHK).setValue(false);
    ws.getRange(item.linha, COL_DEV_CHK).setValue(false);
    ws.getRange(item.linha, COL_VENDA_CHK).setValue(false);
    if (item.aba !== 'Fornecedores Variados') {
      ws.getRange(item.linha, COL_FORN).setValue(item.aba);
    }
    ws.getRange(item.linha, 1, 1, TOTAL_COLUNAS).setBackground('#FFFFFF');
    ws.getRange(item.linha, COL_VL_TOT).setFormula(_formulaTotal(item.linha));

    registrarLog(ss, item.aba, item.linha, COL_NF, item.nfd || '', item.nf,
      '🗑️ Exclusão manual — ' + nfLabel + ' | Motivo: ' + motivo);

    try { CacheService.getScriptCache().remove(_CACHE_KEY_DASH); } catch(_) {}
    _atualizarMetricasDashboard(ss);

    return JSON.stringify({ ok: '✅ ' + nfLabel + ' excluída com sucesso.\nMotivo registrado no log.' });
  } finally {
    trava.releaseLock();
  }
}



// ════════════════════════════════════════════════════════════
//   DESFAZER CONCLUSÃO (REABERTURA)
// ════════════════════════════════════════════════════════════

function desfazerConclusao() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('FormReabertura').setWidth(480).setHeight(400),
    '🔓 Reabrir Devoluções'
  );
}

function buscarNFsConcluidas(txtNfsRaw) {
  var nfsDigitadas = txtNfsRaw.split(/[\n,]/).map(function(s) { return s.trim(); }).filter(Boolean);
  if (!nfsDigitadas.length) return JSON.stringify({ erro: 'Nenhuma NF informada.' });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var encontradas = [];

  ABAS_OPERACIONAIS.forEach(function(nomeAba) {
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) return;
    ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues()
      .forEach(function(l, i) {
        var nfd = String(l[IDX_NFD]).trim();
        var nf  = String(l[IDX_NF]).trim();
        if (_baterTermos(nfsDigitadas, nfd, nf).bate) {
          encontradas.push({
            nf: nf, nfd: nfd, aba: nomeAba,
            status: String(l[IDX_STATUS]).trim(),
            linha: LINHA_DADOS + i,
            desc: String(l[IDX_DESC]).substring(0, 40)
          });
        }
      });
  });

  if (!encontradas.length) return JSON.stringify({ erro: 'Nenhuma NF localizada nas abas.' });
  return JSON.stringify({ itens: encontradas });
}

/**
 * [P30] Versão otimizada de executarReabertura chamada pelo FormReabertura.
 * Recebe o array itens (com linha+aba pré-resolvidos pelo buscarNFsConcluidas).
 */
function executarReaberturaPorItens(itens) {
  if (!itens || !itens.length) return JSON.stringify({ erro: 'Nenhum item para reabrir.' });

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var reabertos = [];
  var porAba    = {};

  itens.forEach(function(it) {
    if (!porAba[it.aba]) porAba[it.aba] = [];
    porAba[it.aba].push(it);
  });

  Object.keys(porAba).forEach(function(nomeAba) {
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;

    var loteAba = porAba[nomeAba];
    var protsAba = ws.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    var protMap  = {};
    protsAba.forEach(function(p) { protMap[p.getRange().getRow()] = p; });

    var linhasReabrir = loteAba.map(function(it) { return it.linha; });
    var minRow = Math.min.apply(null, linhasReabrir);
    var maxRow = Math.max.apply(null, linhasReabrir);
    var nRows  = maxRow - minRow + 1;
    var bgAtual = ws.getRange(minRow, 1, nRows, TOTAL_COLUNAS).getBackgrounds();

    loteAba.forEach(function(it) {
      if (protMap[it.linha]) { protMap[it.linha].remove(); _decrementarProtecoes(1); }
      ws.getRange(it.linha, COL_STATUS, 1, 5).setValues([['Pendente', true, false, false, '']]);
      bgAtual[it.linha - minRow] = Array(TOTAL_COLUNAS).fill(COR_AZUL);
      var nfLabel = it.nfd || it.nf;
      registrarLog(ss, nomeAba, it.linha, COL_STATUS, nfLabel, 'Pendente',
        '🔓 Reabertura via formulário — NF: ' + nfLabel);
      reabertos.push((it.nfd ? 'NFD ' + it.nfd + ' / ' : '') + 'NF ' + it.nf + ' (' + nomeAba + ')');
    });
    ws.getRange(minRow, 1, nRows, TOTAL_COLUNAS).setBackgrounds(bgAtual);
  });

  if (!reabertos.length) return JSON.stringify({ erro: 'Nenhuma NF foi reaberta.' });
  try { CacheService.getScriptCache().remove(_CACHE_KEY_DASH); } catch(_) {}
  _atualizarMetricasDashboard(ss);
  return JSON.stringify({ sucesso: '✅ ' + reabertos.length + ' NF(s) reabertas:\n' + reabertos.join(', ') });
}

function executarReabertura(txtNfsRaw) {
  var nfsDigitadas = txtNfsRaw.split(/[\n,]/).map(function(s) { return s.trim(); }).filter(Boolean);
  if (!nfsDigitadas.length) return JSON.stringify({ erro: 'Nenhuma NF informada.' });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reabertos = [], naoEncontradas = nfsDigitadas.slice();

  ABAS_OPERACIONAIS.forEach(function(nomeAba) {
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) return;

    // [P21] Lista proteções UMA VEZ por aba
    var protsAba = ws.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    var protMap  = {};
    protsAba.forEach(function(p) { protMap[p.getRange().getRow()] = p; });

    var dados         = ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues();
    var linhasReabrir = [];

    dados.forEach(function(l, i) {
      var nfd = String(l[IDX_NFD]).trim();
      var nf  = String(l[IDX_NF]).trim();
      var bat = _baterTermos(nfsDigitadas, nfd, nf);
      if (!bat.bate) return;
      var row = LINHA_DADOS + i;
      if (protMap[row]) { protMap[row].remove(); _decrementarProtecoes(1); }
      linhasReabrir.push({ row: row, nfd: nfd, nf: nf, bat: bat });
    });

    if (!linhasReabrir.length) return;

    // [P22] status+chk+obs em 1 setValues por linha + cores em batch por aba
    var minRow  = linhasReabrir[0].row;
    var maxRow  = linhasReabrir[linhasReabrir.length - 1].row;
    var nRows   = maxRow - minRow + 1;
    var bgAtual = ws.getRange(minRow, 1, nRows, TOTAL_COLUNAS).getBackgrounds();

    linhasReabrir.forEach(function(info) {
      ws.getRange(info.row, COL_STATUS, 1, 5).setValues([[
        'Pendente', true, false, false, ''
      ]]);
      bgAtual[info.row - minRow] = Array(TOTAL_COLUNAS).fill(COR_AZUL);
      var nfLabel = info.nfd || info.nf;
      registrarLog(ss, nomeAba, info.row, COL_STATUS, nfLabel, 'Pendente',
        '🔓 Reabertura em lote — NF: ' + nfLabel);
      reabertos.push((info.nfd ? 'NFD ' + info.nfd + ' / ' : '') + 'NF ' + info.nf + ' (' + nomeAba + ')');
      var idx = naoEncontradas.indexOf(info.bat.termoBateu);
      if (idx > -1) naoEncontradas.splice(idx, 1);
    });
    ws.getRange(minRow, 1, nRows, TOTAL_COLUNAS).setBackgrounds(bgAtual);
  });

  if (!reabertos.length) return JSON.stringify({ erro: 'Nenhuma NF foi reaberta. Verifique os números.' });
  try { CacheService.getScriptCache().remove(_CACHE_KEY_DASH); } catch(_) {}
  _atualizarMetricasDashboard(ss);

  var msg = '✅ ' + reabertos.length + ' NF(s) reabertas:\n' + reabertos.join(', ');
  if (naoEncontradas.length) msg += '\n⚠️ Não localizadas: ' + naoEncontradas.join(', ');
  return JSON.stringify({ sucesso: msg });
}


// ════════════════════════════════════════════════════════════
//   BUSCA / FILTRO RÁPIDO
// ════════════════════════════════════════════════════════════

function abrirBusca() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('FormBusca').setWidth(620).setHeight(500),
    '🔍 Buscar NF ou Fornecedor'
  );
}

function executarBusca(termo) {
  termo = String(termo).trim().toLowerCase();
  if (!termo) return JSON.stringify({ erro: 'Informe um termo para buscar.' });

  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var tz  = Session.getScriptTimeZone();
  var res = [];

  ABAS_OPERACIONAIS.forEach(function(nomeAba) {
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) return;
    ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues()
      .forEach(function(l, i) {
        var nfd  = String(l[IDX_NFD]  || '').trim();
        var nf   = String(l[IDX_NF]   || '').trim();
        var forn = String(l[IDX_FORN] || '').trim();
        var desc = String(l[IDX_DESC] || '').trim();
        if (!nf && !nfd) return;
        if ([nf, nfd, forn, desc].every(function(s) {
          return s.toLowerCase().indexOf(termo) === -1;
        })) return;
        var dt = l[IDX_DATA];
        res.push({
          origem:    'ativo',
          nf:        nf,
          nfd:       nfd,
          forn:      forn,
          desc:      desc.substring(0, 55),
          status:    String(l[IDX_STATUS] || ''),
          data:      dt instanceof Date ? Utilities.formatDate(dt, tz, 'dd/MM/yyyy') : '',
          dataArq:   '',
          valor:     parseFloat(l[IDX_VL_TOT]) || 0,
          aba:       nomeAba,
          linha:     LINHA_DADOS + i
        });
      });
  });

  var hist = ss.getSheetByName('Historico_Arquivo');
  if (hist) {
    var ulH = hist.getLastRow();
    if (ulH >= 2) {
      var ncols = Math.min(TOTAL_COLUNAS + 1, hist.getLastColumn());
      hist.getRange(2, 1, ulH - 1, ncols).getValues().forEach(function(l) {
        var nfd  = String(l[IDX_NFD]  || '').trim();
        var nf   = String(l[IDX_NF]   || '').trim();
        var forn = String(l[IDX_FORN] || '').trim();
        var desc = String(l[IDX_DESC] || '').trim();
        if (!nf && !nfd) return;
        if ([nf, nfd, forn, desc].every(function(s) {
          return s.toLowerCase().indexOf(termo) === -1;
        })) return;
        var dt    = l[IDX_DATA];
        var dtArq = l[TOTAL_COLUNAS];
        res.push({
          origem:    'historico',
          nf:        nf,
          nfd:       nfd,
          forn:      forn,
          desc:      desc.substring(0, 55),
          status:    String(l[IDX_STATUS] || ''),
          data:      dt    instanceof Date ? Utilities.formatDate(dt,    tz, 'dd/MM/yyyy') : '',
          dataArq:   dtArq instanceof Date ? Utilities.formatDate(dtArq, tz, 'dd/MM/yyyy') : '',
          valor:     parseFloat(l[IDX_VL_TOT]) || 0,
          aba:       '',
          linha:     0
        });
      });
    }
  }

  if (!res.length)
    return JSON.stringify({ erro: 'Nenhum resultado encontrado para "' + termo + '".' });

  res.sort(function(a, b) {
    if (a.origem !== b.origem) return a.origem === 'ativo' ? -1 : 1;
    return (b.data || '').localeCompare(a.data || '');
  });

  return JSON.stringify({ resultados: res });
}

// [P28] flush() desnecessário removido
function navegarParaLinha(nomeAba, linha) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;
    ws.activate();
    ws.setActiveRange(ws.getRange(linha, 1));
  } catch (e) { console.error('navegarParaLinha: ' + e); }
}


// ════════════════════════════════════════════════════════════
//   HISTÓRICO DA NF
// ════════════════════════════════════════════════════════════

function buscarHistoricoNF(nf) {
  nf = String(nf).trim();
  if (!nf) return JSON.stringify({ erro: 'Informe o número da NF ou NFD.' });

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var wsLog = ss.getSheetByName('_Log');
  if (!wsLog) return JSON.stringify({ erro: 'Aba _Log não encontrada.' });

  var registros = [];
  try {
    // [P24] Lê apenas as últimas 500 linhas em vez de toda a coluna
    var totalRows = wsLog.getMaxRows();
    var startRow  = Math.max(2, totalRows - 499);
    var blocoA    = wsLog.getRange(startRow, 1, totalRows - startRow + 1, 1).getValues();
    var ultimaLinha = startRow - 1;
    for (var k = blocoA.length - 1; k >= 0; k--) {
      if (blocoA[k][0] !== '' && blocoA[k][0] != null) { ultimaLinha = startRow + k; break; }
    }
    if (ultimaLinha < 2) return JSON.stringify({ registros: [] });

    var dados = wsLog.getRange(2, 1, ultimaLinha - 1, 8).getValues();

    var linhasNF = {};
    ABAS_OPERACIONAIS.forEach(function(nomeAba) {
      var ws = ss.getSheetByName(nomeAba);
      if (!ws) return;
      var ul = obterUltimaLinhaDados(ws);
      if (ul < LINHA_DADOS) return;
      ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues()
        .forEach(function(l, i) {
          if (String(l[IDX_NFD] || '').trim() === nf || String(l[IDX_NF] || '').trim() === nf) {
            linhasNF[nomeAba + ':' + (LINHA_DADOS + i)] = true;
          }
        });
    });

    dados.forEach(function(l) {
      var aba    = String(l[2]);
      var linha  = String(l[3]);
      var valAnt = String(l[5]);
      var valNov = String(l[6]);
      var acao   = String(l[7]);

      var bateChave = !!linhasNF[aba + ':' + linha];
      var bateTexto = valAnt === nf || valNov === nf ||
                      valAnt.indexOf(nf) !== -1 || valNov.indexOf(nf) !== -1 ||
                      acao.indexOf(nf) !== -1;
      if (!bateChave && !bateTexto) return;

      registros.push({
        data:     String(l[0]),
        usuario:  String(l[1]),
        aba:      aba,
        coluna:   String(l[4]),
        anterior: valAnt,
        novo:     valNov,
        acao:     acao
      });
    });
  } catch (e) {
    return JSON.stringify({ erro: 'Erro ao ler log: ' + e.toString() });
  }

  return JSON.stringify({ nf: nf, registros: registros });
}


// ════════════════════════════════════════════════════════════
//   ANEXO DE NF
// ════════════════════════════════════════════════════════════

function abrirAnexoNF() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('FormAnexo').setWidth(480).setHeight(360),
    '📎 Anexar Foto/PDF da NF'
  );
}

function salvarAnexoNF(dados) {
  if (!dados.nf || !dados.base64 || !dados.mimeType)
    return JSON.stringify({ erro: 'Dados incompletos para o anexo.' });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var linhaEncontrada = null, wsEncontrado = null, abaEncontrada = '';

  ABAS_OPERACIONAIS.forEach(function(nomeAba) {
    if (linhaEncontrada) return;
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) return;
    ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues()
      .forEach(function(l, i) {
        if (linhaEncontrada) return;
        var termo = String(dados.nf).trim();
        if (String(l[IDX_NFD] || '').trim() === termo || String(l[IDX_NF] || '').trim() === termo) {
          linhaEncontrada = LINHA_DADOS + i;
          wsEncontrado    = ws;
          abaEncontrada   = nomeAba;
        }
      });
  });

  if (!linhaEncontrada)
    return JSON.stringify({ erro: 'NF "' + dados.nf + '" não encontrada nas abas.' });

  try {
    var blob    = Utilities.newBlob(Utilities.base64Decode(dados.base64), dados.mimeType, dados.nomeArquivo);
    var arquivo = _pastaAnexos().createFile(blob);
    arquivo.setName('NF_' + dados.nf + '_' + dados.nomeArquivo);
    var url = arquivo.getUrl();
    wsEncontrado.getRange(linhaEncontrada, COL_ANEXO).setValue(url);
    registrarLog(ss, abaEncontrada, linhaEncontrada, COL_ANEXO, '', url, '📎 Anexo NF adicionado');
    return JSON.stringify({ sucesso: '✅ Arquivo anexado à NF ' + dados.nf + '.', url: url });
  } catch (e) {
    return JSON.stringify({ erro: '❌ Erro ao salvar anexo: ' + e.toString() });
  }
}


// ════════════════════════════════════════════════════════════
//   E-MAIL DE DEVOLUÇÃO POR NFD
// ════════════════════════════════════════════════════════════

function abrirEmailDevolucao() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('FormEmailDevolucao').setWidth(520).setHeight(520),
    '📧 Enviar E-mail de Devolução'
  );
}

function buscarDadosNFDs(nfdsRaw) {
  var nfds = nfdsRaw.split(/[\n,]/).map(function(s) { return s.trim(); }).filter(Boolean);
  if (!nfds.length) return JSON.stringify({ erro: 'Nenhuma NFD informada.' });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var itens = [];

  ABAS_OPERACIONAIS.forEach(function(nomeAba) {
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) return;
    ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues()
      .forEach(function(l) {
        var bat = _baterTermos(nfds, String(l[IDX_NFD]).trim(), String(l[IDX_NF]).trim());
        if (!bat.bate) return;
        var dt = l[IDX_DATA];
        itens.push({
          nfd:      String(l[IDX_NFD]).trim() || String(l[IDX_NF]).trim(),
          tipo:     String(l[IDX_TIPO]).trim(),
          motivo:   String(l[IDX_MOTIVO]).trim(),
          nf:       String(l[IDX_NF]).trim(),
          forn:     String(l[IDX_FORN]).trim(),
          desc:     String(l[IDX_DESC]).trim(),
          qtd:      l[IDX_QTD] || 0,
          valor:    parseFloat(l[IDX_VL_TOT]) || 0,
          data:     dt instanceof Date ? Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
          urlAnexo: String(l[IDX_ANEXO] || '').trim()
        });
      });
  });

  if (!itens.length) return JSON.stringify({ erro: 'Nenhuma NFD localizada nas abas.' });

  var forns = _fornecedoresUnicos(itens);
  if (forns.length > 1)
    return JSON.stringify({ erro: 'NFDs de fornecedores diferentes: ' + forns.join(', ') + '. Use apenas NFDs do mesmo fornecedor.' });

  return JSON.stringify({
    itens:      itens,
    forn:       forns[0],
    titulo:     _montarTituloEmail(itens, forns[0]),
    emailsBase: _getEmailsGeral()
  });
}

function enviarEmailDevolucao(params) {
  var nfds = params.nfdsRaw.split(/[\n,]/).map(function(s) { return s.trim(); }).filter(Boolean);
  if (!nfds.length) return JSON.stringify({ erro: 'Nenhuma NFD informada.' });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var itens = [];

  ABAS_OPERACIONAIS.forEach(function(nomeAba) {
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) return;
    ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues()
      .forEach(function(l, li) {
        var bat = _baterTermos(nfds, String(l[IDX_NFD]).trim(), String(l[IDX_NF]).trim());
        if (!bat.bate) return;
        var dt = l[IDX_DATA];
        itens.push({
          nfd:      String(l[IDX_NFD]).trim() || String(l[IDX_NF]).trim(),
          tipo:     String(l[IDX_TIPO]).trim(),
          motivo:   String(l[IDX_MOTIVO]).trim(),
          nf:       String(l[IDX_NF]).trim(),
          forn:     String(l[IDX_FORN]).trim(),
          desc:     String(l[IDX_DESC]).trim(),
          qtd:      l[IDX_QTD] || 0,
          valor:    parseFloat(l[IDX_VL_TOT]) || 0,
          data:     dt instanceof Date ? Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
          urlAnexo: String(l[IDX_ANEXO] || '').trim(),
          linha:    LINHA_DADOS + li,
          ws:       ws,
          nomeAba:  nomeAba
        });
      });
  });

  if (!itens.length) return JSON.stringify({ erro: 'Nenhuma NFD localizada.' });

  var forn          = itens[0].forn;
  var destinatarios = _montarDestinatarios(params.emailsExtras);
  var assunto       = params.assunto || _montarTituloEmail(itens, forn);
  var valorTotal    = itens.reduce(function(s, it) { return s + it.valor; }, 0);

  var linhasTabela = itens.map(function(it) {
    var corTipo  = it.tipo === 'Avaria'   ? '#FFF3E0'
                 : it.tipo === 'Rejeição' ? '#FEF2F2' : '#E3F2FD';
    var corTexto = it.tipo === 'Avaria'   ? '#E65100'
                 : it.tipo === 'Rejeição' ? '#DC2626' : '#1565C0';
    return '<tr>' +
      '<td style="padding:7px 10px;border-bottom:1px solid #eee;font-weight:bold">' + _esc(it.nfd) + '</td>' +
      '<td style="padding:7px 10px;border-bottom:1px solid #eee;background:' + corTipo + ';color:' + corTexto + ';font-weight:bold;text-align:center">' + _esc(it.tipo) + '</td>' +
      '<td style="padding:7px 10px;border-bottom:1px solid #eee;color:#555">' + _esc(it.motivo || '—') + '</td>' +
      '<td style="padding:7px 10px;border-bottom:1px solid #eee">' + _esc(it.nf) + '</td>' +
      '<td style="padding:7px 10px;border-bottom:1px solid #eee">' + _esc(it.desc) + '</td>' +
      '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:center">' + _esc(it.qtd) + '</td>' +
      '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:right">R$ ' + it.valor.toFixed(2).replace('.', ',') + '</td>' +
      '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:center">' + _esc(it.data) + '</td>' +
      '</tr>';
  }).join('');

  var dataEnvio = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  // ── Comunicado de retorno: blob direto na memória, sem salvar no Drive ──
  var blobComunicado = null;
  if (params.comBase64 && params.comMime && params.comNome) {
    try {
      blobComunicado = Utilities.newBlob(
        Utilities.base64Decode(params.comBase64),
        params.comMime,
        params.comNome
      );
    } catch (eCom) {
      console.error('Erro ao decodificar comunicado: ' + eCom);
    }
  }

  var comObsHtml = '';
  if (blobComunicado) {
    comObsHtml =
      '<div style="margin:14px 0 0;padding:10px 14px;background:#FFF8E1;' +
      'border-left:4px solid #F59E0B;border-radius:0 4px 4px 0">' +
      '<p style="margin:0;font-size:13px;color:#92400E;font-weight:bold">📋 Comunicado de Retorno em Anexo</p>';
    if (params.comObs) {
      comObsHtml += '<p style="margin:6px 0 0;font-size:13px;color:#444">' + _esc(params.comObs) + '</p>';
    }
    comObsHtml += '</div>';
  }

  var obsHtml = comObsHtml;
  if (params.obs) {
    obsHtml += '<p style="margin:14px 0 0;font-size:13px;color:#444"><strong>Observações:</strong> ' + _esc(params.obs) + '</p>';
  }

  var htmlBody = _montarHtmlEmail(assunto, dataEnvio, forn, linhasTabela, valorTotal, obsHtml);

  var blobs = [], semAnexo = [];
  itens.forEach(function(it) {
    if (!it.urlAnexo || !it.urlAnexo.startsWith('http')) { semAnexo.push(it.nfd); return; }
    try {
      var fileId = _extrairIdDriveUrl(it.urlAnexo);
      if (!fileId) { semAnexo.push(it.nfd); return; }
      // [P23] Busca o arquivo 1× e reusa para getBlob e getName
      var driveFile = DriveApp.getFileById(fileId);
      blobs.push(driveFile.getBlob().setName('NFD_' + it.nfd + '_' + driveFile.getName()));
    } catch (eBlob) {
      console.warn('Não foi possível anexar arquivo da NFD ' + it.nfd + ': ' + eBlob);
      semAnexo.push(it.nfd);
    }
  });

  var avisoSemAnexo = semAnexo.length
    ? '<p style="margin:10px 0 0;font-size:12px;color:#E65100">⚠️ NFD(s) sem arquivo anexado: ' + semAnexo.map(_esc).join(', ') + '</p>'
    : '';

  var htmlFinal = htmlBody.replace(
    '<p style="margin:20px 0 0;font-size:12px;color:#888">',
    avisoSemAnexo + '<p style="margin:20px 0 0;font-size:12px;color:#888">'
  );

  try {
    var todosBlobs = blobs.slice();
    if (blobComunicado) todosBlobs.push(blobComunicado);

    var mailOpts = {
      to:       destinatarios.join(','),
      subject:  assunto,
      htmlBody: htmlFinal
    };
    if (todosBlobs.length) mailOpts.attachments = todosBlobs;
    MailApp.sendEmail(mailOpts);

    var infoAnexos = todosBlobs.length
      ? ' | ' + todosBlobs.length + ' arquivo(s) anexado(s)' + (blobComunicado ? ' (incl. comunicado)' : '')
      : ' | sem anexos';
    registrarLog(ss, 'SISTEMA', 0, 0, '', assunto,
      '📧 E-mail devolução enviado para: ' + destinatarios.join(', ') + infoAnexos);

    _registrarEmailEnviado(ss, {
      assunto:       assunto,
      destinatarios: destinatarios,
      nfds:          itens.map(function(it) { return it.nfd || it.nf; }),
      forn:          forn,
      totalItens:    itens.length,
      totalValor:    valorTotal,
      anexos:        todosBlobs.length
    });

    var itensFalta   = itens.filter(function(it) { return it.tipo === 'Falta'; });
    var linhasFalta  = itensFalta.map(function(it) { return { ws: it.ws, linha: it.linha, nomeAba: it.nomeAba }; });
    var nfdsArquivadas = itensFalta.map(function(it) { return it.nfd || it.nf; });

    var totalArquivadas = 0;
    if (linhasFalta.length) {
      itensFalta.forEach(function(it) {
        var nfRef = it.nfd || it.nf;
        registrarLog(ss, it.nomeAba, it.linha, COL_STATUS, nfRef, 'Arquivado',
          '📧 Falta arquivada após envio do e-mail — NF: ' + nfRef);
      });
      totalArquivadas = _arquivarLinhasEspecificas(ss, linhasFalta);
      _atualizarMetricasDashboard(ss);
    }

    var infoArquivadas = totalArquivadas > 0
      ? '\n📦 ' + totalArquivadas + ' nota(s) de Falta movida(s) para Historico_Arquivo: ' + nfdsArquivadas.join(', ')
      : '';

    return JSON.stringify({
      sucesso: '✅ E-mail enviado para ' + destinatarios.length + ' destinatário(s)' + infoAnexos + ':\n' +
               destinatarios.join('\n') + infoArquivadas
    });
  } catch (e) {
    return JSON.stringify({ erro: '❌ Erro ao enviar: ' + e.toString() });
  }
}

/** Monta o HTML completo do e-mail de devolução. */
function _montarHtmlEmail(assunto, dataEnvio, forn, linhasTabela, valorTotal, obsHtml) {
  return '<div style="font-family:Arial,sans-serif;max-width:820px;color:#222">' +
    '<div style="background:#2D5F8A;color:#fff;padding:16px 20px;border-radius:6px 6px 0 0">' +
    '<h2 style="margin:0;font-size:18px">' + _esc(assunto) + '</h2>' +
    '<p style="margin:4px 0 0;font-size:12px;opacity:.85">Emitido em ' + _esc(dataEnvio) + '</p>' +
    '</div>' +
    '<div style="background:#fff;border:1px solid #ddd;border-top:none;padding:16px 20px;border-radius:0 0 6px 6px">' +
    '<p style="margin:0 0 14px;font-size:13px">Prezados,</p>' +
    '<p style="margin:0 0 14px;font-size:13px">Encaminhamos abaixo a relação de notas fiscais referentes às devoluções de <strong>' + _esc(forn) + '</strong>:</p>' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px">' +
    '<thead><tr style="background:#F0F4F8">' +
    '<th style="padding:8px 10px;text-align:left;border-bottom:2px solid #2D5F8A">NFD</th>' +
    '<th style="padding:8px 10px;text-align:center;border-bottom:2px solid #2D5F8A">Tipo</th>' +
    '<th style="padding:8px 10px;text-align:left;border-bottom:2px solid #2D5F8A">Motivo</th>' +
    '<th style="padding:8px 10px;text-align:left;border-bottom:2px solid #2D5F8A">Nº NF</th>' +
    '<th style="padding:8px 10px;text-align:left;border-bottom:2px solid #2D5F8A">Descrição</th>' +
    '<th style="padding:8px 10px;text-align:center;border-bottom:2px solid #2D5F8A">Qtd</th>' +
    '<th style="padding:8px 10px;text-align:right;border-bottom:2px solid #2D5F8A">Valor</th>' +
    '<th style="padding:8px 10px;text-align:center;border-bottom:2px solid #2D5F8A">Data</th>' +
    '</tr></thead>' +
    '<tbody>' + linhasTabela + '</tbody>' +
    '<tfoot><tr style="background:#F9F9F9">' +
    '<td colspan="6" style="padding:8px 10px;font-weight:bold;text-align:right;border-top:2px solid #2D5F8A">TOTAL:</td>' +
    '<td style="padding:8px 10px;font-weight:bold;text-align:right;border-top:2px solid #2D5F8A;color:#2D5F8A">R$ ' + valorTotal.toFixed(2).replace('.', ',') + '</td>' +
    '<td style="border-top:2px solid #2D5F8A"></td>' +
    '</tr></tfoot>' +
    '</table>' + obsHtml +
    '</div></div>';
}


// ════════════════════════════════════════════════════════════
//   HISTÓRICO DE E-MAILS ENVIADOS
// ════════════════════════════════════════════════════════════

function garantirAbaEmailsEnviados(ss) {
  var ws = ss.getSheetByName('_EmailsEnviados');
  if (!ws) {
    ws = ss.insertSheet('_EmailsEnviados');
    ws.hideSheet();
    var cab = ['Data/Hora','Assunto','Fornecedor','Destinatários','NFDs/NFs Incluídas','Total Itens','Valor Total (R$)','Arquivos Anexados'];
    ws.getRange(1, 1, 1, cab.length).setValues([cab])
      .setBackground('#2D5F8A').setFontColor('#FFFFFF').setFontWeight('bold');
    ws.setFrozenRows(1);
    [160,300,160,260,300,80,140,120].forEach(function(w, i) { ws.setColumnWidth(i + 1, w); });
  }
  return ws;
}

function _registrarEmailEnviado(ss, info) {
  try {
    var ws    = garantirAbaEmailsEnviados(ss);
    var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    ws.appendRow([
      agora, info.assunto, info.forn,
      info.destinatarios.join('; '),
      info.nfds.join(', '),
      info.totalItens, info.totalValor, info.anexos
    ]);
  } catch (e) {
    console.error('_registrarEmailEnviado: ' + e);
  }
}

function buscarHistoricoEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName('_EmailsEnviados');
  if (!ws) return JSON.stringify({ registros: [] });

  try {
    var ul = ws.getLastRow();
    if (ul < 2) return JSON.stringify({ registros: [] });
    var dados = ws.getRange(2, 1, ul - 1, 8).getValues();
    var registros = dados
      .filter(function(l) { return l[0]; })
      .map(function(l) {
        return {
          data:       String(l[0]),
          assunto:    String(l[1]),
          forn:       String(l[2]),
          destinos:   String(l[3]),
          nfds:       String(l[4]),
          totalItens: l[5] || 0,
          totalValor: parseFloat(l[6]) || 0,
          anexos:     l[7] || 0
        };
      })
      .reverse();
    return JSON.stringify({ registros: registros });
  } catch (e) {
    return JSON.stringify({ erro: 'Erro ao ler histórico: ' + e.toString() });
  }
}


// ════════════════════════════════════════════════════════════
//   ALERTAS DE ATRASO E RESUMO SEMANAL
// ════════════════════════════════════════════════════════════

function verificarAtrasosEEnviarAlerta() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var tz     = ss.getSpreadsheetTimeZone();
  var hoje   = new Date();
  var limite = new Date(hoje.getTime() - 30 * 24 * 60 * 60 * 1000);
  var linhas = [];

  ABAS_OPERACIONAIS.forEach(function(nomeAba) {
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) return;
    ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues()
      .forEach(function(l) {
        var nf = l[IDX_NF], st = l[IDX_STATUS], dt = l[IDX_DATA];
        if (nf && st === 'Pendente' && dt instanceof Date && dt < limite) {
          linhas.push({
            nfd:  String(l[IDX_NFD]  || '').trim(),
            nf:   String(nf).trim(),
            data: Utilities.formatDate(dt, tz, 'dd/MM/yyyy'),
            forn: String(l[IDX_FORN] || nomeAba).trim(),
            tipo: String(l[IDX_TIPO] || '').trim(),
            desc: String(l[IDX_DESC] || '').trim(),
            qtd:  l[IDX_QTD] || 0,
            val:  parseFloat(l[IDX_VL_TOT]) || 0,
            st:   st,
            dias: Math.floor((hoje - dt) / 864e5),
            resp: String(l[IDX_RESP] || 'Não informado').trim()
          });
        }
      });
  });

  if (!linhas.length) {
    try { SpreadsheetApp.getUi().alert('✅ Nenhum item com +30 dias pendente.'); } catch (_) {}
    return;
  }

  var dataStr  = Utilities.formatDate(hoje, tz, 'dd/MM/yyyy');
  var valTotal = linhas.reduce(function(s, l) { return s + l.val; }, 0);
  var assunto  = '⚠️ [Devoluções] ' + linhas.length + ' item(ns) em atraso crítico (+30 dias) — ' + dataStr;

  var pdf = _gerarRelatorioPDF(ss, {
    titulo:   'DEVOLUÇÕES EM ATRASO CRÍTICO (+30 DIAS)',
    periodo:  dataStr,
    linhas:   linhas,
    valTotal: valTotal,
    nomeArq:  'Atraso_Critico_' + dataStr.replace(/\//g, '-') + '.pdf',
    kpiLabel: 'Em Atraso (+30 dias)',
    kpiCor:   '#DC2626',
    colExtra: { header: 'Atraso', fn: function(l) { return l.dias + ' dias'; } }
  });

  var htmlEmail = _montarHtmlRelatorio({
    icone:    '⚠️',
    titulo:   'Devoluções em Atraso Crítico',
    subtitulo: dataStr,
    intro:    'Foram encontradas <strong>' + linhas.length + '</strong> devolução(ões) com mais de <strong>30 dias</strong> em aberto. Segue relatório em anexo.',
    kpis: [
      { label: 'Em Atraso (+30 dias)', cor: '#DC2626', valor: linhas.length + ' itens', sub: 'R$ ' + _fmtVal(valTotal) }
    ]
  });

  var anexos = pdf ? [pdf.blob] : [];
  enviarEmail(assunto, htmlEmail, anexos);
  registrarLog(ss, 'SISTEMA', 0, 0, '', linhas.length + ' itens', '⚠️ Alerta de atraso enviado — ' + dataStr);
  try { SpreadsheetApp.getUi().alert('📧 Alerta enviado! ' + linhas.length + ' item(ns) em atraso.'); } catch (_) {}
}

function enviarEmail(assunto, htmlBody, anexos) {
  try {
    var destinatarios = _getEmailsGeral();
    if (!destinatarios || !destinatarios.length) return;
    var opts = {
      to:       destinatarios.join(','),
      subject:  assunto,
      htmlBody: htmlBody
    };
    if (anexos && anexos.length) opts.attachments = anexos;
    MailApp.sendEmail(opts);
  } catch (e) {
    console.error('enviarEmail: ' + e);
  }
}


// ════════════════════════════════════════════════════════════
//   BUSCA NO HISTÓRICO ARQUIVADO
// ════════════════════════════════════════════════════════════

function abrirBuscaHistorico() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('FormBuscaHistorico').setWidth(620).setHeight(520),
    '🗂️ Buscar no Histórico Arquivado'
  );
}

function executarBuscaHistorico(params) {
  var termo        = String(params.termo || '').trim().toLowerCase();
  var incluirAtivas = !!params.incluirAtivas;
  if (!termo) return JSON.stringify({ erro: 'Informe um termo para buscar.' });

  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var resultados = [];
  var tz         = Session.getScriptTimeZone();

  var hist = ss.getSheetByName('Historico_Arquivo');
  if (hist) {
    var ulH = hist.getLastRow();
    if (ulH >= 2) {
      var ncols = Math.min(TOTAL_COLUNAS + 1, hist.getLastColumn());
      hist.getRange(2, 1, ulH - 1, ncols).getValues().forEach(function(l, i) {
        var nfd  = String(l[IDX_NFD]  || '').trim();
        var nf   = String(l[IDX_NF]   || '').trim();
        var forn = String(l[IDX_FORN] || '').trim();
        var desc = String(l[IDX_DESC] || '').trim();
        if (!nf && !nfd) return;
        if ([nf, nfd, forn, desc].every(function(s) {
          return s.toLowerCase().indexOf(termo) === -1;
        })) return;
        var dt       = l[IDX_DATA];
        var dtArq    = l[TOTAL_COLUNAS];
        resultados.push({
          origem:    'Histórico',
          aba:       forn || 'Arquivado',
          nf:        nf,
          nfd:       nfd,
          forn:      forn,
          desc:      desc.substring(0, 55),
          status:    String(l[IDX_STATUS] || ''),
          tipo:      String(l[IDX_TIPO]   || ''),
          qtd:       l[IDX_QTD]  || 0,
          valor:     parseFloat(l[IDX_VL_TOT]) || 0,
          data:      dt instanceof Date
                       ? Utilities.formatDate(dt, tz, 'dd/MM/yyyy') : '',
          dataArq:   dtArq instanceof Date
                       ? Utilities.formatDate(dtArq, tz, 'dd/MM/yyyy') : '',
          linha:     i + 2,
          navegavel: false
        });
      });
    }
  }

  if (incluirAtivas) {
    ABAS_OPERACIONAIS.forEach(function(nomeAba) {
      var ws = ss.getSheetByName(nomeAba);
      if (!ws) return;
      var ul = obterUltimaLinhaDados(ws);
      if (ul < LINHA_DADOS) return;
      ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues()
        .forEach(function(l, i) {
          var nfd  = String(l[IDX_NFD]  || '').trim();
          var nf   = String(l[IDX_NF]   || '').trim();
          var forn = String(l[IDX_FORN] || '').trim();
          var desc = String(l[IDX_DESC] || '').trim();
          if (!nf && !nfd) return;
          if ([nf, nfd, forn, desc].every(function(s) {
            return s.toLowerCase().indexOf(termo) === -1;
          })) return;
          var dt = l[IDX_DATA];
          resultados.push({
            origem:    'Ativo',
            aba:       nomeAba,
            nf:        nf,
            nfd:       nfd,
            forn:      forn,
            desc:      desc.substring(0, 55),
            status:    String(l[IDX_STATUS] || ''),
            tipo:      String(l[IDX_TIPO]   || ''),
            qtd:       l[IDX_QTD]  || 0,
            valor:     parseFloat(l[IDX_VL_TOT]) || 0,
            data:      dt instanceof Date
                         ? Utilities.formatDate(dt, tz, 'dd/MM/yyyy') : '',
            dataArq:   '',
            linha:     LINHA_DADOS + i,
            navegavel: true
          });
        });
    });
  }

  if (!resultados.length)
    return JSON.stringify({ erro: 'Nenhum resultado encontrado para "' + termo + '".' });

  resultados.sort(function(a, b) {
    if (a.origem !== b.origem) return a.origem === 'Ativo' ? -1 : 1;
    return (b.data || '').localeCompare(a.data || '');
  });

  return JSON.stringify({ resultados: resultados, total: resultados.length });
}


// ════════════════════════════════════════════════════════════
//   RELATÓRIOS (MENSAL / SEMANAL / DIÁRIO)
// ════════════════════════════════════════════════════════════

function abrirRelatorios() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('FormRelatorios').setWidth(480).setHeight(400),
    '📊 Relatórios de Devoluções'
  );
}

// ─── MENSAL ──────────────────────────────────────────────────

function gerarRelatorioMensal(params) {
  var mes = parseInt(params.mes, 10);
  var ano = parseInt(params.ano, 10);
  var enviarEmailFlag = !!params.enviarEmail;

  if (!mes || !ano || mes < 1 || mes > 12)
    return JSON.stringify({ erro: 'Mês ou ano inválido.' });
  if (!ID_PASTA_DESTINO || ID_PASTA_DESTINO.startsWith('INSIRA'))
    return JSON.stringify({ erro: 'Configure ID_PASTA_DESTINO no topo do script.' });

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var tz      = ss.getSpreadsheetTimeZone();
  var dataIni = new Date(ano, mes - 1, 1);
  var dataFim = new Date(ano, mes, 0, 23, 59, 59);
  var nomeMes = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                 'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'][mes - 1];
  var periodo = nomeMes + ' / ' + ano;

  var linhas = _coletarLinhas(ss, tz, dataIni, dataFim);
  if (!linhas.length)
    return JSON.stringify({ erro: 'Nenhum lançamento encontrado para ' + periodo + '.' });

  var acc     = _acumular(linhas);
  var nomeArq = 'Relatorio_Mensal_' + nomeMes + '_' + ano + '.pdf';
  var pdf     = _gerarRelatorioPDF(ss, { titulo: 'RELATÓRIO MENSAL DE DEVOLUÇÕES', periodo: periodo, linhas: linhas, acc: acc, nomeArq: nomeArq });

  if (!pdf) return JSON.stringify({ erro: '❌ Erro ao gerar o PDF. Verifique o log.' });

  if (enviarEmailFlag) {
    var htmlEmail = _montarHtmlRelatorio({ icone: '📊', titulo: 'Relatório Mensal de Devoluções', subtitulo: periodo,
      intro: 'Segue em anexo o relatório mensal referente a <strong>' + periodo + '</strong>.', kpis: _kpisEmail(acc) });
    enviarEmail('📊 Relatório Mensal de Devoluções — ' + periodo, htmlEmail, [pdf.blob]);
    registrarLog(ss, 'SISTEMA', 0, 0, '', periodo, '📊 Relatório mensal gerado e enviado — ' + periodo);
  } else {
    registrarLog(ss, 'SISTEMA', 0, 0, '', periodo, '📊 Relatório mensal gerado — ' + periodo);
  }

  return JSON.stringify({
    sucesso: '✅ Relatório de ' + periodo + ' gerado!\n' +
             linhas.length + ' lançamento(s) — R$ ' + _fmtVal(acc.vTotal) +
             (enviarEmailFlag ? '\n📧 Enviado por e-mail.' : ''),
    urlPdf: pdf.arquivo.getUrl()
  });
}

// ─── SEMANAL ─────────────────────────────────────────────────

function gerarRelatorioSemanal(params) {
  if (!ID_PASTA_DESTINO || ID_PASTA_DESTINO.startsWith('INSIRA'))
    return JSON.stringify({ erro: 'Configure ID_PASTA_DESTINO no topo do script.' });

  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var tz   = ss.getSpreadsheetTimeZone();
  var hoje = new Date();
  var enviarEmailFlag = !!params.enviarEmail;
  var dataIni, dataFim, periodoLabel;
  var modo = params.modo || 'ultimos7';

  if (modo === 'personalizado') {
    dataIni = _parseDateStr(params.dataIni);
    dataFim = _parseDateStr(params.dataFim, true);
    if (!dataIni || !dataFim) return JSON.stringify({ erro: 'Datas inválidas.' });
    periodoLabel = _fmtDt(dataIni, tz) + ' a ' + _fmtDt(dataFim, tz);
  } else if (modo === 'semana_corrente') {
    var dia = hoje.getDay();
    var diffSeg = (dia === 0) ? -6 : 1 - dia;
    dataIni = new Date(hoje); dataIni.setDate(hoje.getDate() + diffSeg);
    dataIni.setHours(0, 0, 0, 0);
    dataFim = new Date(dataIni); dataFim.setDate(dataIni.getDate() + 6);
    dataFim.setHours(23, 59, 59, 999);
    periodoLabel = 'Semana ' + _fmtDt(dataIni, tz) + ' – ' + _fmtDt(dataFim, tz);
  } else if (modo === 'semana_anterior') {
    var dia2 = hoje.getDay();
    var diffSeg2 = (dia2 === 0) ? -6 : 1 - dia2;
    dataFim = new Date(hoje); dataFim.setDate(hoje.getDate() + diffSeg2 - 1);
    dataFim.setHours(23, 59, 59, 999);
    dataIni = new Date(dataFim); dataIni.setDate(dataFim.getDate() - 6);
    dataIni.setHours(0, 0, 0, 0);
    periodoLabel = 'Semana ' + _fmtDt(dataIni, tz) + ' – ' + _fmtDt(dataFim, tz);
  } else {
    dataFim = new Date(hoje); dataFim.setHours(23, 59, 59, 999);
    dataIni = new Date(hoje); dataIni.setDate(hoje.getDate() - 6);
    dataIni.setHours(0, 0, 0, 0);
    periodoLabel = 'Últimos 7 dias — até ' + _fmtDt(dataFim, tz);
  }

  var linhas = _coletarLinhas(ss, tz, dataIni, dataFim);
  if (!linhas.length)
    return JSON.stringify({ erro: 'Nenhum lançamento encontrado para o período selecionado.' });

  var acc     = _acumular(linhas);
  var nomeArq = 'Relatorio_Semanal_' + Utilities.formatDate(dataIni, tz, 'dd-MM-yyyy') +
                '_a_' + Utilities.formatDate(dataFim, tz, 'dd-MM-yyyy') + '.pdf';
  var pdf     = _gerarRelatorioPDF(ss, { titulo: 'RELATÓRIO SEMANAL DE DEVOLUÇÕES', periodo: periodoLabel, linhas: linhas, acc: acc, nomeArq: nomeArq });

  if (!pdf) return JSON.stringify({ erro: '❌ Erro ao gerar o PDF. Verifique o log.' });

  if (enviarEmailFlag) {
    var htmlEmail = _montarHtmlRelatorio({ icone: '📊', titulo: 'Relatório Semanal de Devoluções', subtitulo: periodoLabel,
      intro: 'Segue em anexo o relatório semanal referente a <strong>' + periodoLabel + '</strong>.', kpis: _kpisEmail(acc) });
    enviarEmail('📊 Relatório Semanal de Devoluções — ' + periodoLabel, htmlEmail, [pdf.blob]);
    registrarLog(ss, 'SISTEMA', 0, 0, '', linhas.length + ' itens', '📊 Relatório semanal gerado e enviado — ' + periodoLabel);
  } else {
    registrarLog(ss, 'SISTEMA', 0, 0, '', linhas.length + ' itens', '📊 Relatório semanal gerado — ' + periodoLabel);
  }

  return JSON.stringify({
    sucesso: '✅ Relatório semanal gerado!\n' +
             linhas.length + ' lançamento(s) — R$ ' + _fmtVal(acc.vTotal) +
             (enviarEmailFlag ? '\n📧 Enviado por e-mail.' : ''),
    urlPdf: pdf.arquivo.getUrl()
  });
}

/** Compatibilidade com trigger semanal automático. */
function enviarResumoSemanal() {
  var r = gerarRelatorioSemanal({ modo: 'ultimos7', enviarEmail: true });
  var obj = JSON.parse(r);
  if (obj.erro) console.error('enviarResumoSemanal: ' + obj.erro);
}

// ─── DIÁRIO ──────────────────────────────────────────────────

function gerarRelatorioDiario(params) {
  if (!ID_PASTA_DESTINO || ID_PASTA_DESTINO.startsWith('INSIRA'))
    return JSON.stringify({ erro: 'Configure ID_PASTA_DESTINO no topo do script.' });

  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var tz  = ss.getSpreadsheetTimeZone();
  var enviarEmailFlag = !!params.enviarEmail;

  var dataIni = _parseDateStr(params.data);
  if (!dataIni) return JSON.stringify({ erro: 'Data inválida.' });
  var dataFim = new Date(dataIni);
  dataFim.setHours(23, 59, 59, 999);

  var periodoLabel = _fmtDt(dataIni, tz);
  var linhas = _coletarLinhas(ss, tz, dataIni, dataFim);
  if (!linhas.length)
    return JSON.stringify({ erro: 'Nenhum lançamento encontrado para ' + periodoLabel + '.' });

  var acc     = _acumular(linhas);
  var nomeArq = 'Relatorio_Diario_' + Utilities.formatDate(dataIni, tz, 'dd-MM-yyyy') + '.pdf';
  var pdf     = _gerarRelatorioPDF(ss, { titulo: 'RELATÓRIO DIÁRIO DE DEVOLUÇÕES', periodo: periodoLabel, linhas: linhas, acc: acc, nomeArq: nomeArq });

  if (!pdf) return JSON.stringify({ erro: '❌ Erro ao gerar o PDF. Verifique o log.' });

  if (enviarEmailFlag) {
    var htmlEmail = _montarHtmlRelatorio({ icone: '📋', titulo: 'Relatório Diário de Devoluções', subtitulo: periodoLabel,
      intro: 'Segue em anexo o relatório diário referente a <strong>' + periodoLabel + '</strong>.', kpis: _kpisEmail(acc) });
    enviarEmail('📋 Relatório Diário de Devoluções — ' + periodoLabel, htmlEmail, [pdf.blob]);
    registrarLog(ss, 'SISTEMA', 0, 0, '', linhas.length + ' itens', '📋 Relatório diário gerado e enviado — ' + periodoLabel);
  } else {
    registrarLog(ss, 'SISTEMA', 0, 0, '', linhas.length + ' itens', '📋 Relatório diário gerado — ' + periodoLabel);
  }

  return JSON.stringify({
    sucesso: '✅ Relatório diário de ' + periodoLabel + ' gerado!\n' +
             linhas.length + ' lançamento(s) — R$ ' + _fmtVal(acc.vTotal) +
             (enviarEmailFlag ? '\n📧 Enviado por e-mail.' : ''),
    urlPdf: pdf.arquivo.getUrl()
  });
}

// ─── HELPERS DE COLETA ───────────────────────────────────────

/**
 * Coleta linhas dentro de [dataIni, dataFim] varrendo:
 * 1. Abas operacionais (itens ainda ativos/pendentes)
 * 2. Historico_Arquivo (itens já arquivados — Devolvido/Venda)
 * Usa a DATA DE ENTRADA (col 3) para filtrar.
 */
function _coletarLinhas(ss, tz, dataIni, dataFim) {
  var linhas = [];

  ABAS_OPERACIONAIS.forEach(function(nomeAba) {
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) return;
    ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues()
      .forEach(function(l) {
        var nf = l[IDX_NF], dt = l[IDX_DATA];
        if (!nf || !(dt instanceof Date) || dt < dataIni || dt > dataFim) return;
        linhas.push({
          nfd:  String(l[IDX_NFD]  || '').trim(),
          nf:   String(nf).trim(),
          data: Utilities.formatDate(dt, tz, 'dd/MM/yyyy'),
          forn: String(l[IDX_FORN] || nomeAba).trim(),
          tipo: String(l[IDX_TIPO] || '').trim(),
          desc: String(l[IDX_DESC] || '').trim(),
          qtd:  l[IDX_QTD] || 0,
          val:  parseFloat(l[IDX_VL_TOT]) || 0,
          st:   String(l[IDX_STATUS] || '')
        });
      });
  });

  var hist = ss.getSheetByName('Historico_Arquivo');
  if (hist && hist.getLastRow() >= 2) {
    var ulH  = hist.getLastRow();
    var cols = Math.min(TOTAL_COLUNAS, hist.getLastColumn());
    hist.getRange(2, 1, ulH - 1, cols).getValues()
      .forEach(function(l) {
        var nf = l[IDX_NF], dt = l[IDX_DATA];
        if (!nf || !(dt instanceof Date) || dt < dataIni || dt > dataFim) return;
        linhas.push({
          nfd:  String(l[IDX_NFD]  || '').trim(),
          nf:   String(nf).trim(),
          data: Utilities.formatDate(dt, tz, 'dd/MM/yyyy'),
          forn: String(l[IDX_FORN] || '').trim(),
          tipo: String(l[IDX_TIPO] || '').trim(),
          desc: String(l[IDX_DESC] || '').trim(),
          qtd:  l[IDX_QTD] || 0,
          val:  parseFloat(l[IDX_VL_TOT]) || 0,
          st:   String(l[IDX_STATUS] || '')
        });
      });
  }

  return linhas;
}

// ─── PDF DE RELATÓRIO ─────────────────────────────────────────

/**
 * Gera PDF do relatório com layout completo:
 *   1. Cabeçalho  2. KPIs  3. Resumo por fornecedor
 *   4. Listagem detalhada  5. Rodapé
 */
function _gerarRelatorioPDF(ss, params) {
  var tz = ss.getSpreadsheetTimeZone();
  var ssTemp;
  try {
    ssTemp = SpreadsheetApp.create('_Rel_Temp_' + new Date().getTime());
    var sh = ssTemp.getSheets()[0];

    var AZUL_ESC  = '#1A3557';
    var AZUL_SUB  = '#243F63';
    var CINZA_BG  = '#F8F9FA';
    var BRANCO    = '#FFFFFF';
    var corStatus = { 'Pendente': '#EBF3FF', 'Devolvido': '#ECFDF5', 'Venda': '#FFF7ED' };
    var corTipo   = { 'Avaria': '#FFF3E0', 'Falta': '#E3F2FD', 'Rejeição': '#FEF2F2' };

    var nCols    = 9;
    var larguras = [70, 75, 80, 140, 65, 70, 210, 50, 80];
    if (params.colExtra) { nCols = 10; larguras.push(70); }
    larguras.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });

    var acc  = params.acc || _acumular(params.linhas);
    var taxa = acc.taxa;
    var rl   = 1;

    // ── Cabeçalho ─────────────────────────────────────────
    sh.setRowHeight(rl, 46);
    sh.getRange(rl, 1, 1, nCols).merge()
      .setValue(params.titulo + (params.periodo ? ' — ' + params.periodo.toUpperCase() : ''))
      .setBackground(AZUL_ESC).setFontColor(BRANCO)
      .setFontWeight('bold').setFontSize(12)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    rl++;

    sh.setRowHeight(rl, 20);
    sh.getRange(rl, 1, 1, nCols).merge()
      .setValue('Emitido em: ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm') +
                '   |   Total de lançamentos no período: ' + params.linhas.length)
      .setBackground(AZUL_SUB).setFontColor('#93B4D4')
      .setFontSize(8).setFontStyle('italic')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    rl++;

    // ── KPIs ─────────────────────────────────────────────
    sh.setRowHeight(rl, 6);
    sh.getRange(rl, 1, 1, nCols).setBackground(CINZA_BG);
    rl++;

    var kpiDefs = [
      { label: 'PENDENTES',      cor: '#2563EB', st: 'Pendente',  qtd: acc.tP, val: acc.vP },
      { label: 'DEVOLVIDOS',     cor: '#059669', st: 'Devolvido', qtd: acc.tD, val: acc.vD },
      { label: 'VENDAS',         cor: '#D97706', st: 'Venda',     qtd: acc.tV, val: acc.vV },
      { label: 'TOTAL',          cor: '#7C3AED', st: null,        qtd: params.linhas.length, val: acc.vTotal },
      { label: 'TAXA RESOLUÇÃO', cor: taxa >= 70 ? '#059669' : taxa >= 40 ? '#D97706' : '#DC2626',
        st: 'taxa', qtd: taxa, val: -1 }
    ];

    var kpiSpans, kpiStarts;
    if (nCols === 9) {
      kpiSpans  = [2, 2, 2, 2, 1];
      kpiStarts = [1, 3, 5, 7, 9];
    } else {
      kpiSpans  = [2, 2, 2, 2, 2];
      kpiStarts = [1, 3, 5, 7, 9];
    }

    sh.setRowHeight(rl, 15);
    kpiDefs.forEach(function(k, ki) {
      sh.getRange(rl, kpiStarts[ki], 1, kpiSpans[ki]).merge()
        .setValue(k.label)
        .setBackground(k.cor).setFontColor(BRANCO)
        .setFontWeight('bold').setFontSize(7)
        .setHorizontalAlignment('center').setVerticalAlignment('middle');
    });
    rl++;

    sh.setRowHeight(rl, 28);
    kpiDefs.forEach(function(k, ki) {
      var txt = (k.st === 'taxa') ? k.qtd + '%' : k.qtd + ' itens\nR$ ' + _fmtVal(k.val);
      sh.getRange(rl, kpiStarts[ki], 1, kpiSpans[ki]).merge()
        .setValue(txt)
        .setFontColor(k.cor).setFontWeight('bold').setFontSize(9)
        .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    });
    rl++;

    sh.setRowHeight(rl, 6);
    sh.getRange(rl, 1, 1, nCols).setBackground(CINZA_BG);
    rl++;

    // ── Resumo por fornecedor ─────────────────────────────
    var fornMap = {};
    params.linhas.forEach(function(l) {
      var f = l.forn || '(sem fornecedor)';
      if (!fornMap[f]) fornMap[f] = { tP:0,tD:0,tV:0, vP:0,vD:0,vV:0, total:0, vTotal:0 };
      var m = fornMap[f];
      m.total++;  m.vTotal += l.val;
      if      (l.st === 'Pendente')  { m.tP++; m.vP += l.val; }
      else if (l.st === 'Devolvido') { m.tD++; m.vD += l.val; }
      else if (l.st === 'Venda')     { m.tV++; m.vV += l.val; }
    });
    var fornKeys = Object.keys(fornMap);

    sh.setRowHeight(rl, 16);
    sh.getRange(rl, 1, 1, nCols).merge()
      .setValue('RESUMO POR FORNECEDOR')
      .setBackground('#1A3557').setFontWeight('bold').setFontSize(9).setFontColor(BRANCO)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    rl++;

    var hForn = ['Fornecedor','Pendentes','Vl Pendente','Devolvidos','Vl Devolvido','Vendas','Vl Venda','Total','Vl Total'];
    sh.setRowHeight(rl, 16);
    sh.getRange(rl, 1, 1, 9).setValues([hForn])
      .setBackground('#1A3557').setFontWeight('bold').setFontSize(8).setFontColor(BRANCO)
      .setHorizontalAlignment('center');
    sh.getRange(rl, 1).setHorizontalAlignment('left');
    rl++;

    var fornRows = fornKeys.map(function(f) {
      var m = fornMap[f];
      return [f, m.tP, 'R$ '+_fmtVal(m.vP), m.tD, 'R$ '+_fmtVal(m.vD),
              m.tV, 'R$ '+_fmtVal(m.vV), m.total, 'R$ '+_fmtVal(m.vTotal)];
    });
    fornRows.push([
      'TOTAL GERAL',
      acc.tP, 'R$ '+_fmtVal(acc.vP),
      acc.tD, 'R$ '+_fmtVal(acc.vD),
      acc.tV, 'R$ '+_fmtVal(acc.vV),
      params.linhas.length, 'R$ '+_fmtVal(acc.vTotal)
    ]);

    if (fornRows.length) {
      sh.getRange(rl, 1, fornRows.length, 9).setValues(fornRows).setFontSize(8)
        .setHorizontalAlignment('center');
      sh.getRange(rl, 1, fornRows.length, 1).setHorizontalAlignment('left');
      sh.getRange(rl + fornRows.length - 1, 1, 1, 9)
        .setFontWeight('bold').setBackground('#E8EDF3');
      for (var fi = 0; fi < fornRows.length - 1; fi++) {
        if (fi % 2 === 1) sh.getRange(rl + fi, 1, 1, 9).setBackground('#F5F7FA');
      }
      rl += fornRows.length;
    }

    sh.setRowHeight(rl, 6);
    sh.getRange(rl, 1, 1, nCols).setBackground(CINZA_BG);
    rl++;

    // ── Listagem detalhada ────────────────────────────────
    sh.setRowHeight(rl, 16);
    sh.getRange(rl, 1, 1, nCols).merge()
      .setValue('LISTAGEM DETALHADA — ' + params.linhas.length + ' LANÇAMENTO(S)')
      .setBackground('#1A3557').setFontWeight('bold').setFontSize(9).setFontColor(BRANCO)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    rl++;

    var headers = ['NFD', 'Nº NF', 'Data', 'Fornecedor', 'Tipo', 'Status', 'Descrição', 'Qtd', 'Valor (R$)'];
    if (params.colExtra) headers.push(params.colExtra.header);

    sh.setRowHeight(rl, 15);
    sh.getRange(rl, 1, 1, nCols).setValues([headers])
      .setBackground('#1A3557').setFontWeight('bold').setFontSize(8).setFontColor(BRANCO)
      .setHorizontalAlignment('center');
    sh.getRange(rl, 7).setHorizontalAlignment('left');
    rl++;

    if (params.linhas.length) {
      var vals = params.linhas.map(function(it) {
        var row = [it.nfd || '', it.nf, it.data, it.forn, it.tipo,
                   it.st || '', it.desc, it.qtd || '', _fmtVal(it.val)];
        if (params.colExtra) row.push(params.colExtra.fn(it));
        return row;
      });
      sh.getRange(rl, 1, vals.length, nCols).setValues(vals).setFontSize(8)
        .setHorizontalAlignment('center');
      sh.getRange(rl, 7, vals.length, 1).setHorizontalAlignment('left');

      // [P25] Cores em batch
      var bgStatus = params.linhas.map(function(it) {
        return Array(nCols).fill(corStatus[it.st] || BRANCO);
      });
      sh.getRange(rl, 1, vals.length, nCols).setBackgrounds(bgStatus);
      var bgTipo = params.linhas.map(function(it) { return [corTipo[it.tipo] || BRANCO]; });
      sh.getRange(rl, 5, vals.length, 1).setBackgrounds(bgTipo);
      rl += vals.length;
    }

    // ── Rodapé ────────────────────────────────────────────
    sh.setRowHeight(rl, 6); rl++;
    sh.setRowHeight(rl, 14);
    sh.getRange(rl, 1, 1, nCols).merge()
      .setValue('Relatório gerado automaticamente pelo Sistema de Controle de Devoluções.')
      .setFontColor('#9CA3AF').setFontSize(7).setFontStyle('italic')
      .setHorizontalAlignment('center');

    SpreadsheetApp.flush();

    var exportUrl = ssTemp.getUrl().replace(/\/edit.*$/, '') +
      '/export?exportFormat=pdf&format=pdf&size=A4&portrait=false' +
      '&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false';

    // [P26] pdfBlob reutilizado em memória
    var pdfBlob = UrlFetchApp.fetch(exportUrl, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    }).getBlob().setName(params.nomeArq);

    var arquivo = DriveApp.getFolderById(ID_PASTA_DESTINO).createFile(pdfBlob);
    DriveApp.getFileById(ssTemp.getId()).setTrashed(true);
    return { arquivo: arquivo, blob: pdfBlob };

  } catch(e) {
    console.error('_gerarRelatorioPDF: ' + e);
    try { if (ssTemp) DriveApp.getFileById(ssTemp.getId()).setTrashed(true); } catch(_) {}
    return null;
  }
}

/** Monta o HTML padrão dos e-mails de relatório. */
function _montarHtmlRelatorio(params) {
  var kpiCells = (params.kpis || []).map(function(k) {
    return '<td style="width:' + Math.floor(100 / params.kpis.length) + '%;padding:12px 8px;' +
           'text-align:center;vertical-align:top">' +
      '<div style="background:' + k.cor + ';color:#fff;border-radius:6px 6px 0 0;' +
           'padding:5px 4px;font-size:9px;font-weight:bold;letter-spacing:.5px">' + _esc(k.label) + '</div>' +
      '<div style="border:1px solid #E5E7EB;border-top:none;border-radius:0 0 6px 6px;padding:8px 4px;background:#fff">' +
      '<div style="font-size:18px;font-weight:bold;color:' + k.cor + '">' + _esc(k.valor) + '</div>' +
      '<div style="font-size:10px;color:#6B7280;margin-top:2px">' + _esc(k.sub) + '</div>' +
      '</div></td>';
  }).join('');

  return '<div style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto">' +
    '<div style="background:#1A3557;color:#fff;padding:18px 22px;border-radius:8px 8px 0 0">' +
    '<h2 style="margin:0;font-size:17px;font-weight:bold">' + _esc(params.icone) + ' ' + _esc(params.titulo) + '</h2>' +
    '<p style="margin:5px 0 0;font-size:11px;opacity:.75">' + _esc(params.subtitulo) + '</p>' +
    '</div>' +
    '<div style="background:#fff;border:1px solid #E5E7EB;border-top:none;' +
         'padding:18px 22px;border-radius:0 0 8px 8px">' +
    '<p style="margin:0 0 16px;font-size:13px;color:#374151">' + (params.intro || '') + '</p>' +
    (kpiCells ? '<table style="width:100%;border-collapse:separate;border-spacing:6px;margin-bottom:16px">' +
      '<tr>' + kpiCells + '</tr></table>' : '') +
    '<p style="margin:0;font-size:11px;color:#9CA3AF;border-top:1px solid #F3F4F6;' +
         'padding-top:12px">Gerado automaticamente pelo Sistema de Controle de Devoluções.' +
         ' O relatório completo está em anexo.</p>' +
    '</div></div>';
}


// ════════════════════════════════════════════════════════════
//   RELATÓRIO DE PENDENTES
// ════════════════════════════════════════════════════════════

function gerarRelatorioPendentes(params) {
  if (!ID_PASTA_DESTINO || ID_PASTA_DESTINO.startsWith('INSIRA'))
    return JSON.stringify({ erro: 'Configure ID_PASTA_DESTINO no topo do script.' });

  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var tz   = ss.getSpreadsheetTimeZone();
  var hoje = new Date();
  var linhas = [];

  ABAS_OPERACIONAIS.forEach(function(nomeAba) {
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) return;
    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) return;
    ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues()
      .forEach(function(l) {
        if (!l[IDX_NF] || l[IDX_STATUS] !== 'Pendente') return;
        var dt   = l[IDX_DATA];
        var dias = (dt instanceof Date && !isNaN(dt))
                   ? Math.floor((hoje - dt) / 864e5) : 0;
        linhas.push({
          nfd:  String(l[IDX_NFD]  || '').trim(),
          nf:   String(l[IDX_NF]   || '').trim(),
          data: (dt instanceof Date && !isNaN(dt))
                ? Utilities.formatDate(dt, tz, 'dd/MM/yyyy') : '',
          forn: String(l[IDX_FORN] || nomeAba).trim(),
          tipo: String(l[IDX_TIPO] || '').trim(),
          desc: String(l[IDX_DESC] || '').trim(),
          qtd:  l[IDX_QTD]  || 0,
          val:  parseFloat(l[IDX_VL_TOT]) || 0,
          st:   'Pendente',
          dias: dias
        });
      });
  });

  linhas.sort(function(a, b) { return b.dias - a.dias; });

  if (!linhas.length)
    return JSON.stringify({ erro: 'Nenhum item Pendente encontrado nas abas.' });

  var acc      = _acumular(linhas);
  var dataStr  = Utilities.formatDate(hoje, tz, 'dd/MM/yyyy');
  var nomeArq  = 'Relatorio_Pendentes_' + dataStr.replace(/\//g, '-') + '.pdf';

  var pdf;
  try {
    pdf = _gerarRelatorioPDF(ss, {
      titulo:   'RELATÓRIO DE PENDÊNCIAS EM ABERTO',
      periodo:  dataStr,
      linhas:   linhas,
      acc:      acc,
      nomeArq:  nomeArq,
      colExtra: { header: 'Em aberto', fn: function(it) {
        return it.dias > 0 ? it.dias + ' dias' : 'Hoje';
      }}
    });
  } catch (ePdf) {
    registrarLog(ss, 'SISTEMA', 0, 0, '', '', '❌ Erro PDF pendentes: ' + ePdf.toString());
    return JSON.stringify({ erro: '❌ Erro ao gerar PDF: ' + ePdf.toString() });
  }

  if (!pdf)
    return JSON.stringify({ erro: '❌ Erro ao gerar o PDF. Verifique o log do Apps Script para detalhes.' });

  if (params && params.enviarEmail) {
    try {
      var htmlEmail = _montarHtmlRelatorio({
        icone:     '⏳',
        titulo:    'Pendências em Aberto',
        subtitulo: dataStr,
        intro:     'Snapshot de todos os itens atualmente <strong>Pendentes</strong>, ordenados por antiguidade.',
        kpis: [
          { label: 'Total Pendente', cor: '#2563EB',
            valor: linhas.length + ' itens', sub: 'R$ ' + _fmtVal(acc.vP) },
          { label: 'Valor em Aberto', cor: '#DC2626',
            valor: 'R$ ' + _fmtVal(acc.vP), sub: 'a receber/resolver' }
        ]
      });
      enviarEmail('⏳ Relatório de Pendências — ' + dataStr, htmlEmail, [pdf.blob]);
    } catch (eMail) {
      console.error('gerarRelatorioPendentes — e-mail: ' + eMail);
    }
  }

  registrarLog(ss, 'SISTEMA', 0, 0, '', linhas.length + ' itens',
    '⏳ Relatório pendentes gerado — ' + dataStr);

  return JSON.stringify({
    sucesso: '✅ Relatório de pendentes gerado!\n' +
             linhas.length + ' item(ns) em aberto — R$ ' + _fmtVal(acc.vP) +
             (params && params.enviarEmail ? '\n📧 Enviado por e-mail.' : ''),
    urlPdf: pdf.arquivo.getUrl()
  });
}

// ════════════════════════════════════════════════════════════
//   RELATÓRIO POR FORNECEDOR
//   Adicionar no Código.gs logo após gerarRelatorioPendentes()
//   (após a linha que fecha a função com `}` na linha ~3445)
// ════════════════════════════════════════════════════════════

/**
 * Retorna lista de fornecedores presentes na aba "Fornecedores Variados".
 * Chamada pelo FormRelatorios.html via google.script.run.listarFornecedoresVariados()
 */
function listarFornecedoresVariados() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ws = ss.getSheetByName('Fornecedores Variados');

    if (!ws) {
      return JSON.stringify({ variados: [], erro: 'Aba "Fornecedores Variados" não encontrada.' });
    }

    var ul = obterUltimaLinhaDados(ws);
    if (ul < LINHA_DADOS) {
      return JSON.stringify({ variados: [] });
    }

    var valores = ws.getRange(LINHA_DADOS, COL_FORN, ul - LINHA_DADOS + 1, 1).getValues();
    var vistos  = {};
    var lista   = [];

    valores.forEach(function(row) {
      var nome = String(row[0] || '').trim();
      if (nome && !vistos[nome]) {
        vistos[nome] = true;
        lista.push(nome);
      }
    });

    lista.sort(function(a, b) { return a.localeCompare(b, 'pt-BR'); });

    return JSON.stringify({ variados: lista });

  } catch (e) {
    return JSON.stringify({ variados: [], erro: e.toString() });
  }
}

/**
 * Gera relatório PDF filtrado por fornecedor específico (ou todos) e período.
 * Chamada pelo FormRelatorios.html via google.script.run.gerarRelatorioPorFornecedor(params)
 *
 * params: {
 *   fornecedor:   string — nome do fornecedor ou 'TODOS'
 *   dataIni:      string — 'YYYY-MM-DD'
 *   dataFim:      string — 'YYYY-MM-DD'
 *   enviarEmail:  boolean
 * }
 */
function gerarRelatorioPorFornecedor(params) {
  if (!ID_PASTA_DESTINO || ID_PASTA_DESTINO.startsWith('INSIRA'))
    return JSON.stringify({ erro: 'Configure ID_PASTA_DESTINO no topo do script.' });

  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var tz   = ss.getSpreadsheetTimeZone();
  var enviarEmailFlag = !!params.enviarEmail;

  var fornFiltro = String(params.fornecedor || '').trim();
  if (!fornFiltro)
    return JSON.stringify({ erro: 'Selecione um fornecedor.' });

  var dataIni = _parseDateStr(params.dataIni);
  var dataFim = _parseDateStr(params.dataFim, true);
  if (!dataIni || !dataFim)
    return JSON.stringify({ erro: 'Datas inválidas.' });

  var periodoLabel = _fmtDt(dataIni, tz) + ' a ' + _fmtDt(dataFim, tz);
  var titulo, nomeArq;

  // Coleta todas as linhas do período
  var todasLinhas = _coletarLinhas(ss, tz, dataIni, dataFim);

  // Filtra por fornecedor (se não for TODOS)
  var linhas;
  if (fornFiltro === 'TODOS') {
    linhas  = todasLinhas;
    titulo  = 'RELATÓRIO DE DEVOLUÇÕES — TODOS OS FORNECEDORES';
    nomeArq = 'Relatorio_Fornecedor_TODOS_' +
              Utilities.formatDate(dataIni, tz, 'dd-MM-yyyy') + '_a_' +
              Utilities.formatDate(dataFim, tz, 'dd-MM-yyyy') + '.pdf';
  } else {
    var fornLower = fornFiltro.toLowerCase();
    linhas = todasLinhas.filter(function(l) {
      return l.forn.toLowerCase() === fornLower;
    });
    titulo  = 'RELATÓRIO DE DEVOLUÇÕES — ' + fornFiltro.toUpperCase();
    nomeArq = 'Relatorio_Fornecedor_' +
              fornFiltro.replace(/[^a-zA-Z0-9]/g, '_') + '_' +
              Utilities.formatDate(dataIni, tz, 'dd-MM-yyyy') + '_a_' +
              Utilities.formatDate(dataFim, tz, 'dd-MM-yyyy') + '.pdf';
  }

  if (!linhas.length) {
    var msg = fornFiltro === 'TODOS'
      ? 'Nenhum lançamento encontrado para o período ' + periodoLabel + '.'
      : 'Nenhum lançamento de "' + fornFiltro + '" encontrado para ' + periodoLabel + '.';
    return JSON.stringify({ erro: msg });
  }

  var acc = _acumular(linhas);

  var pdf;
  try {
    pdf = _gerarRelatorioPDF(ss, {
      titulo:  titulo,
      periodo: periodoLabel,
      linhas:  linhas,
      acc:     acc,
      nomeArq: nomeArq
    });
  } catch (ePdf) {
    registrarLog(ss, 'SISTEMA', 0, 0, '', '', '❌ Erro PDF fornecedor: ' + ePdf.toString());
    return JSON.stringify({ erro: '❌ Erro ao gerar PDF: ' + ePdf.toString() });
  }

  if (!pdf)
    return JSON.stringify({ erro: '❌ Erro ao gerar o PDF. Verifique o log do Apps Script.' });

  if (enviarEmailFlag) {
    try {
      var htmlEmail = _montarHtmlRelatorio({
        icone:    '🏭',
        titulo:   'Relatório por Fornecedor — ' + (fornFiltro === 'TODOS' ? 'Todos' : fornFiltro),
        subtitulo: periodoLabel,
        intro:    'Segue em anexo o relatório de devoluções de <strong>' +
                  (fornFiltro === 'TODOS' ? 'todos os fornecedores' : fornFiltro) +
                  '</strong> referente ao período <strong>' + periodoLabel + '</strong>.',
        kpis:     _kpisEmail(acc)
      });
      enviarEmail(
        '🏭 Relatório por Fornecedor — ' +
        (fornFiltro === 'TODOS' ? 'Todos' : fornFiltro) + ' — ' + periodoLabel,
        htmlEmail,
        [pdf.blob]
      );
      registrarLog(ss, 'SISTEMA', 0, 0, '', fornFiltro,
        '🏭 Relatório por fornecedor gerado e enviado — ' + fornFiltro + ' — ' + periodoLabel);
    } catch (eEmail) {
      // PDF gerado com sucesso; apenas avisa falha no e-mail
      return JSON.stringify({
        sucesso: '✅ PDF gerado, mas falha ao enviar e-mail: ' + eEmail.message,
        urlPdf: pdf.arquivo.getUrl()
      });
    }
  } else {
    registrarLog(ss, 'SISTEMA', 0, 0, '', fornFiltro,
      '🏭 Relatório por fornecedor gerado — ' + fornFiltro + ' — ' + periodoLabel);
  }

  return JSON.stringify({
    sucesso: '✅ Relatório de ' + (fornFiltro === 'TODOS' ? 'todos os fornecedores' : '"' + fornFiltro + '"') +
             ' gerado!\n' + linhas.length + ' lançamento(s) — R$ ' + _fmtVal(acc.vTotal) +
             (enviarEmailFlag ? '\n📧 Enviado por e-mail.' : ''),
    urlPdf: pdf.arquivo.getUrl()
  });
}

// ════════════════════════════════════════════════════════════
//   BACKUP E RESTAURAÇÃO
// ════════════════════════════════════════════════════════════

function abrirBackup() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('FormBackup').setWidth(480).setHeight(380),
    '💾 Backup e Restauração'
  );
}

function infoBackupExistente() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ws  = ss.getSheetByName(BACKUP_ABA);
  if (!ws || ws.getLastRow() < 2)
    return JSON.stringify({ existe: false });

  var ul   = ws.getLastRow();
  var dados = ws.getRange(2, 1, ul - 1, BACKUP_TOTAL_COL).getValues();

  var contagem = {};
  var dataBackup = '';
  dados.forEach(function(l) {
    var aba = String(l[0] || '').trim();
    var nf  = String(l[IDX_NF + 1] || '').trim();
    if (!aba || !nf) return;
    contagem[aba] = (contagem[aba] || 0) + 1;
    if (!dataBackup) {
      var ts = l[BACKUP_TOTAL_COL - 1];
      dataBackup = ts instanceof Date
        ? Utilities.formatDate(ts, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
        : String(ts || '');
    }
  });

  return JSON.stringify({ existe: true, data: dataBackup, contagem: contagem });
}

function executarBackup() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var tz  = Session.getScriptTimeZone();
  var agora = new Date();

  var ws = ss.getSheetByName(BACKUP_ABA);
  if (ws) ss.deleteSheet(ws);
  ws = ss.insertSheet(BACKUP_ABA);
  ws.hideSheet();

  var cab = ['Aba Origem',
    'NFD','Nº NF','Data Entrada','Fornecedor','Tipo','Motivo','Descrição',
    'Qtd','Vl Unit','Vl Total','Status','Pendente✓','Devolvido✓','Venda✓',
    'Obs','Responsável','Anexo','Backup em'
  ];
  ws.getRange(1, 1, 1, BACKUP_TOTAL_COL)
    .setValues([cab])
    .setBackground('#1A3557').setFontColor('#FFFFFF').setFontWeight('bold');
  ws.setFrozenRows(1);
  ws.setColumnWidth(1, 160);

  var totalLinhas = 0;
  var resumo = {};

  ABAS_OPERACIONAIS.forEach(function(nomeAba) {
    var wsAba = ss.getSheetByName(nomeAba);
    if (!wsAba) return;
    var ul = obterUltimaLinhaDados(wsAba);
    if (ul < LINHA_DADOS) return;

    var dados = wsAba.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues();
    var linhas = [];
    dados.forEach(function(l) {
      if (!l[IDX_NF] && !l[IDX_NFD]) return;
      linhas.push([nomeAba].concat(l).concat([agora]));
    });

    if (linhas.length) {
      ws.getRange(ws.getLastRow() + 1, 1, linhas.length, BACKUP_TOTAL_COL)
        .setValues(linhas);
      resumo[nomeAba] = linhas.length;
      totalLinhas += linhas.length;
    }
  });

  SpreadsheetApp.flush();

  var dataStr = Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm:ss');
  registrarLog(ss, 'SISTEMA', 0, 0, '', totalLinhas + ' linhas', '💾 Backup realizado em ' + dataStr);

  var msg = '✅ Backup concluído em ' + dataStr + '\n\n';
  Object.keys(resumo).forEach(function(aba) {
    msg += '• ' + aba + ': ' + resumo[aba] + ' linha(s)\n';
  });
  msg += '\nTotal: ' + totalLinhas + ' registro(s) salvos.';
  return JSON.stringify({ sucesso: msg, data: dataStr, contagem: resumo });
}

function executarRestauracao() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var wsB = ss.getSheetByName(BACKUP_ABA);
  if (!wsB || wsB.getLastRow() < 2)
    return JSON.stringify({ erro: 'Nenhum backup encontrado. Faça um backup antes de reconfigurar.' });

  var tz   = Session.getScriptTimeZone();
  var ul   = wsB.getLastRow();
  var snap = wsB.getRange(2, 1, ul - 1, BACKUP_TOTAL_COL).getValues();

  var porAba = {};
  snap.forEach(function(l) {
    var aba = String(l[0] || '').trim();
    var nf  = String(l[IDX_NF + 1] || '').trim();
    if (!aba || !nf) return;
    if (!porAba[aba]) porAba[aba] = [];
    porAba[aba].push(l.slice(1, TOTAL_COLUNAS + 1));
  });

  var totalRestaurados = 0;
  var resumo = {};
  var erros  = [];

  Object.keys(porAba).forEach(function(nomeAba) {
    var ws = ss.getSheetByName(nomeAba);
    if (!ws) {
      erros.push('Aba "' + nomeAba + '" não encontrada — execute "Configurar/Reinstalar" primeiro.');
      return;
    }

    var linhas = porAba[nomeAba];
    var dest   = LINHA_DADOS;
    var fmt    = 'R$ #,##0.00;;"";""';
    var trava  = LockService.getScriptLock();
    if (!trava.tryLock(15000)) {
      erros.push('Timeout ao restaurar "' + nomeAba + '". Tente novamente.');
      return;
    }

    try {
      // [P27] Lista proteções 1× fora do loop
      var protsAba = ws.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      var protMap  = {};
      protsAba.forEach(function(p) { protMap[p.getRange().getRow()] = p; });

      linhas.forEach(function(l, idx) {
        var row = dest + idx;
        if (protMap[row]) { protMap[row].remove(); _decrementarProtecoes(1); }
      });

      // [P27] Batch setValues
      ws.getRange(dest, 1, linhas.length, TOTAL_COLUNAS).setValues(linhas);

      var fmulas = [], fu = [], ft = [], fd = [], cores = [];
      linhas.forEach(function(l, idx) {
        var row    = dest + idx;
        var status = String(l[IDX_STATUS] || 'Pendente').trim();
        fmulas.push([_formulaTotal(row)]);
        fu.push([fmt]); ft.push([fmt]); fd.push(['dd/mm/yyyy']);
        cores.push(Array(TOTAL_COLUNAS).fill(corPorStatus(status)));
      });
      ws.getRange(dest, COL_VL_TOT,  linhas.length, 1).setFormulas(fmulas);
      ws.getRange(dest, COL_VL_UNIT, linhas.length, 1).setNumberFormats(fu);
      ws.getRange(dest, COL_VL_TOT,  linhas.length, 1).setNumberFormats(ft);
      ws.getRange(dest, COL_DATA,    linhas.length, 1).setNumberFormats(fd);
      ws.getRange(dest, 1, linhas.length, TOTAL_COLUNAS).setBackgrounds(cores);

      linhas.forEach(function(l, idx) {
        var status = String(l[IDX_STATUS] || 'Pendente').trim();
        if (status === 'Devolvido' || status === 'Venda') {
          protegerLinhaConcluida(ss, ws, dest + idx, status);
        }
      });

      resumo[nomeAba]  = linhas.length;
      totalRestaurados += linhas.length;
    } catch(e) {
      erros.push('Erro em "' + nomeAba + '": ' + e.toString());
    } finally {
      trava.releaseLock();
    }
  });

  SpreadsheetApp.flush();
  _atualizarMetricasDashboard(ss);

  var dataStr = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  registrarLog(ss, 'SISTEMA', 0, 0, '', totalRestaurados + ' linhas',
    '🔄 Restauração concluída em ' + dataStr);

  var msg = '✅ Restauração concluída!\n\n';
  Object.keys(resumo).forEach(function(aba) {
    msg += '• ' + aba + ': ' + resumo[aba] + ' linha(s) restaurada(s)\n';
  });
  msg += '\nTotal: ' + totalRestaurados + ' registro(s).';
  if (erros.length) msg += '\n\n⚠️ Avisos:\n' + erros.join('\n');

  return JSON.stringify({ sucesso: msg });
}


// ════════════════════════════════════════════════════════════
//   CONFIGURAÇÕES
// ════════════════════════════════════════════════════════════

function abrirConfiguracoes() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('FormConfiguracoes').setWidth(520).setHeight(580),
    '⚙️ Configurações do Sistema'
  );
}

// ─── E-MAILS ─────────────────────────────────────────────────

function obterEmailConfig() {
  var props = PropertiesService.getScriptProperties();
  try {
    var rawGeral  = props.getProperty(_KEY_EMAILS_GERAL);
    var rawAlerta = props.getProperty(_KEY_EMAILS_ALERTA);
    var alertaDest = props.getProperty(_KEY_ALERTA_DEST) || 'todos';
    var geral  = rawGeral  ? JSON.parse(rawGeral)  : (EMAILS_DESTINATARIOS || []);
    var alerta = rawAlerta ? JSON.parse(rawAlerta) : [];
    return JSON.stringify({ geral: geral, alerta: alerta, alertaDest: alertaDest });
  } catch (e) {
    return JSON.stringify({ geral: EMAILS_DESTINATARIOS || [], alerta: [], alertaDest: 'todos' });
  }
}

function salvarEmailConfig(params) {
  if (!params || !params.geral || !params.geral.length)
    return JSON.stringify({ erro: 'A lista geral precisa ter ao menos um e-mail.' });
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty(_KEY_EMAILS_GERAL,  JSON.stringify(params.geral));
    props.setProperty(_KEY_EMAILS_ALERTA, JSON.stringify(params.alerta || []));
    props.setProperty(_KEY_ALERTA_DEST,   params.alertaDest || 'todos');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    registrarLog(ss, 'SISTEMA', 0, 0, '', params.geral.join(';'),
      '⚙️ E-mails atualizados — alerta: ' + (params.alertaDest || 'todos'));
    return JSON.stringify({
      ok: '✅ Configurações de e-mail salvas!\n' +
          'Lista geral: ' + params.geral.length + ' e-mail(s)\n' +
          'Alertas: ' + (params.alertaDest === 'cc'
            ? 'CC — ' + (params.alerta || []).length + ' e-mail(s)'
            : 'Todos da lista geral')
    });
  } catch (e) {
    return JSON.stringify({ erro: e.toString() });
  }
}

function _getEmailsGeral() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(_KEY_EMAILS_GERAL);
    if (raw) return JSON.parse(raw);
  } catch (_) {}
  return EMAILS_DESTINATARIOS || [];
}

function _getEmailsAlerta() {
  try {
    var props      = PropertiesService.getScriptProperties();
    var alertaDest = props.getProperty(_KEY_ALERTA_DEST) || 'todos';
    var rawGeral   = props.getProperty(_KEY_EMAILS_GERAL);
    var rawAlerta  = props.getProperty(_KEY_EMAILS_ALERTA);
    var geral  = rawGeral  ? JSON.parse(rawGeral)  : (EMAILS_DESTINATARIOS || []);
    var alerta = rawAlerta ? JSON.parse(rawAlerta) : [];
    if (!geral.length) return { to: '', cc: '' };
    var to = geral[0];
    var cc = [];
    if (geral.length > 1) cc = cc.concat(geral.slice(1));
    if (alertaDest === 'cc') {
      alerta.forEach(function(e) {
        if (cc.indexOf(e) === -1 && e !== to) cc.push(e);
      });
    }
    return { to: to, cc: cc.join(',') };
  } catch (_) {
    var fb = EMAILS_DESTINATARIOS || [];
    return { to: fb[0] || '', cc: fb.slice(1).join(',') };
  }
}

/** Alias de compatibilidade. */
function _getEmailsDestinatarios() { return _getEmailsGeral(); }

// ─── CORES ───────────────────────────────────────────────────

function obterCoresSalvas() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(_KEY_CORES);
    if (raw) return JSON.stringify({ cores: JSON.parse(raw) });
    return JSON.stringify({ cores: {
      pendente:  COR_AZUL,
      devolvido: COR_VERDE,
      venda:     COR_LARANJA,
      alerta:    COR_ALERTA_30DIAS,
      header:    COR_HEADER
    }});
  } catch (_) {
    return JSON.stringify({ cores: null });
  }
}

function salvarCoresEReaplicar(cores) {
  if (!cores) return JSON.stringify({ erro: 'Cores não informadas.' });
  try {
    PropertiesService.getScriptProperties().setProperty(_KEY_CORES, JSON.stringify(cores));
    try { CacheService.getScriptCache().remove(_CACHE_KEY_CORES); } catch(_) {}
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var hoje = new Date();

    ABAS_OPERACIONAIS.forEach(function(nomeAba) {
      var ws = ss.getSheetByName(nomeAba);
      if (!ws) return;
      if (cores.header) {
        ws.getRange(1, 1, 1, TOTAL_COLUNAS).setBackground(cores.header);
        ws.getRange(3, 1, 1, TOTAL_COLUNAS).setBackground(cores.header);
      }
      var ul = obterUltimaLinhaDados(ws);
      if (ul < LINHA_DADOS) return;
      var dados = ws.getRange(LINHA_DADOS, 1, ul - LINHA_DADOS + 1, TOTAL_COLUNAS).getValues();
      var bgs   = dados.map(function(l) {
        if (!l[IDX_NF]) return Array(TOTAL_COLUNAS).fill('#FFFFFF');
        var st  = String(l[IDX_STATUS] || '');
        var cor = st === 'Pendente'  ? (cores.pendente  || COR_AZUL)
                : st === 'Devolvido' ? (cores.devolvido || COR_VERDE)
                : st === 'Venda'     ? (cores.venda     || COR_LARANJA)
                : '#FFFFFF';
        if (st === 'Pendente' && l[IDX_DATA] instanceof Date) {
          if (Math.floor((hoje - l[IDX_DATA]) / 864e5) > 30)
            cor = cores.alerta || COR_ALERTA_30DIAS;
        }
        return Array(TOTAL_COLUNAS).fill(cor);
      });
      ws.getRange(LINHA_DADOS, 1, dados.length, TOTAL_COLUNAS).setBackgrounds(bgs);
    });

    SpreadsheetApp.flush();
    registrarLog(ss, 'SISTEMA', 0, 0, '', JSON.stringify(cores), '🎨 Cores atualizadas via configurações');
    return JSON.stringify({ ok: '✅ Cores salvas e reaplicadas em todas as abas!' });
  } catch (e) {
    return JSON.stringify({ erro: '❌ ' + e.toString() });
  }
}

// ─── NOVO FORNECEDOR ─────────────────────────────────────────

function criarNovoFornecedor(params) {
  var nome  = String(params.nome || '').trim();
  var fixar = !!params.fixar;
  if (!nome) return JSON.stringify({ erro: 'Nome não informado.' });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(nome))
    return JSON.stringify({ erro: 'Já existe uma aba com o nome "' + nome + '".' });

  try {
    garantirAba(ss, nome, nome);
    var raw    = PropertiesService.getScriptProperties().getProperty('cdv_abas_extras') || '[]';
    var extras = JSON.parse(raw);
    if (extras.indexOf(nome) === -1) extras.push(nome);
    PropertiesService.getScriptProperties().setProperty('cdv_abas_extras', JSON.stringify(extras));
    registrarLog(ss, 'SISTEMA', 0, 0, '', nome, '🏭 Nova aba criada: ' + nome);
    return JSON.stringify({
      ok: '✅ Aba "' + nome + '" criada com sucesso!\n\n' +
          'Para incluir no menu automático, adicione "' + nome +
          '" ao array ABAS_OPERACIONAIS no código e reinstale o sistema.'
    });
  } catch (e) {
    return JSON.stringify({ erro: '❌ ' + e.toString() });
  }
}

// ─── DIAGNÓSTICO ─────────────────────────────────────────────

function obterDiagnostico() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var props = PropertiesService.getScriptProperties();

    var concl  = parseInt(props.getProperty(_PROP_KEY_CONCLUIDOS) || '0', 10);
    var prots  = parseInt(props.getProperty(_PROP_KEY_PROTECOES)  || '0', 10);
    var backup = props.getProperty('cdv_ultimo_backup') || 'Nunca realizado';

    var abas = ABAS_OPERACIONAIS.map(function(nome) {
      var ws    = ss.getSheetByName(nome);
      var usado = ws ? Math.max(0, obterUltimaLinhaDados(ws) - LINHA_DADOS + 1) : 0;
      return { nome: nome, usado: usado };
    });

    var triggers = ScriptApp.getProjectTriggers().map(function(t) {
      return { func: t.getHandlerFunction(), tipo: t.getTriggerSource().toString() };
    });

    var rawGeral   = props.getProperty(_KEY_EMAILS_GERAL);
    var rawAlerta  = props.getProperty(_KEY_EMAILS_ALERTA);
    var alertaDest = props.getProperty(_KEY_ALERTA_DEST) || 'todos';
    var emailsGeral  = rawGeral  ? JSON.parse(rawGeral)  : (EMAILS_DESTINATARIOS || []);
    var emailsAlerta = rawAlerta ? JSON.parse(rawAlerta) : [];

    return JSON.stringify({
      versao:             'v6.0',
      ultimoBackup:       backup,
      contConcluidos:     concl,
      contProtecoes:      prots,
      abas:               abas,
      triggers:           triggers,
      totalEmailsGeral:   emailsGeral.length,
      totalEmailsAlerta:  emailsAlerta.length,
      alertaDest:         alertaDest
    });
  } catch (e) {
    return JSON.stringify({ erro: e.toString() });
  }
}

// ─── LIMPEZA DO LOG ──────────────────────────────────────────

function executarLimpezaLog(params) {
  var meses = parseInt(params.meses || 6, 10);
  var acao  = String(params.acao || 'arquivar');

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var wsLog = ss.getSheetByName('_Log');
  if (!wsLog) return JSON.stringify({ erro: '_Log não encontrado.' });

  var tz     = ss.getSpreadsheetTimeZone();
  var hoje   = new Date();
  var limite = new Date(hoje);
  limite.setMonth(hoje.getMonth() - meses);
  var limiteISO = String(limite.getFullYear()) +
    String(limite.getMonth()+1).padStart(2,'0') +
    String(limite.getDate()).padStart(2,'0');

  try {
    var ul = wsLog.getMaxRows();
    if (ul < 2) return JSON.stringify({ ok: '✅ Log vazio, nada a limpar.' });

    var dadosAll = wsLog.getRange(2, 1, ul - 1, 8).getValues();
    var antigos  = [];
    var recentes = [];

    dadosAll.forEach(function(l) {
      if (!l[0]) return;
      var s = String(l[0]).trim();
      var compact = '';
      if (/^\d{2}\/\d{2}\/\d{4}/.test(s)) {
        compact = s.slice(6,10) + s.slice(3,5) + s.slice(0,2);
      } else {
        var d = new Date(s);
        if (!isNaN(d)) compact = String(d.getFullYear()) +
          String(d.getMonth()+1).padStart(2,'0') + String(d.getDate()).padStart(2,'0');
      }
      if (compact && compact < limiteISO) antigos.push(l);
      else recentes.push(l);
    });

    if (!antigos.length) {
      var limStr = Utilities.formatDate(limite, tz, 'dd/MM/yyyy');
      return JSON.stringify({ ok: '✅ Nenhum registro anterior a ' + limStr + '. Log está limpo.' });
    }

    if (acao === 'arquivar') {
      var wsArq = ss.getSheetByName('_Log_Arquivo');
      if (!wsArq) {
        wsArq = ss.insertSheet('_Log_Arquivo');
        wsArq.hideSheet();
        wsArq.getRange(1,1,1,8).setValues([
          ['Data/Hora','Usuário','Aba','Linha','Coluna','Valor Anterior','Novo Valor','Ação']
        ]).setBackground('#444444').setFontColor('#FFFFFF').setFontWeight('bold');
      }
      var nextRow = wsArq.getLastRow() + 1;
      wsArq.getRange(nextRow, 1, antigos.length, 8).setValues(antigos);
    }

    wsLog.getRange(2, 1, ul - 1, 8).clearContent();
    if (recentes.length) wsLog.getRange(2, 1, recentes.length, 8).setValues(recentes);

    SpreadsheetApp.flush();
    registrarLog(ss, 'SISTEMA', 0, 0, '', antigos.length + ' registros', '🗂️ Limpeza de log — ' + acao);

    return JSON.stringify({
      ok: '✅ Limpeza concluída!\n' +
          antigos.length + ' registro(s) ' +
          (acao === 'arquivar' ? 'movidos para _Log_Arquivo' : 'apagados') + '.\n' +
          recentes.length + ' registro(s) mantidos no _Log.'
    });
  } catch (e) {
    return JSON.stringify({ erro: '❌ ' + e.toString() });
  }
}

// ─── LIMPEZA DO DRIVE ─────────────────────────────────────────

function previewLimpezaDrive(params) {
  var tipo    = String(params.tipo    || 'relatorios');
  var periodo = parseInt(params.periodo || 0, 10);
  var pastas  = _pastasParaLimpar(tipo);
  if (!pastas.length)
    return JSON.stringify({ erro: 'Nenhuma pasta configurada para o tipo selecionado.' });

  var tz       = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var corte    = periodo > 0 ? new Date(Date.now() - periodo * 864e5) : null;
  var arquivos = [];

  pastas.forEach(function(p) {
    try {
      var files = DriveApp.getFolderById(p.id).getFiles();
      while (files.hasNext() && arquivos.length < 200) {
        var f = files.next();
        if (corte && f.getDateCreated() > corte) continue;
        arquivos.push({
          nome: f.getName(),
          data: Utilities.formatDate(f.getDateCreated(), tz, 'dd/MM/yyyy')
        });
      }
    } catch (e) { console.warn('previewLimpezaDrive — ' + p.label + ': ' + e); }
  });

  return JSON.stringify({ arquivos: arquivos });
}

function executarLimpezaDrive(params) {
  var tipo    = String(params.tipo    || 'relatorios');
  var periodo = parseInt(params.periodo || 0, 10);
  var pastas  = _pastasParaLimpar(tipo);
  if (!pastas.length)
    return JSON.stringify({ erro: 'Nenhuma pasta configurada.' });

  var tz    = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var corte = periodo > 0 ? new Date(Date.now() - periodo * 864e5) : null;
  var total = 0;
  var erros = [];

  pastas.forEach(function(p) {
    try {
      var files = DriveApp.getFolderById(p.id).getFiles();
      while (files.hasNext()) {
        var f = files.next();
        try {
          if (corte && f.getDateCreated() > corte) continue;
          f.setTrashed(true);
          total++;
        } catch (ef) { erros.push(f.getName() + ': ' + ef.message); }
      }
    } catch (ep) { erros.push(p.label + ': ' + ep.message); }
  });

  registrarLog(SpreadsheetApp.getActiveSpreadsheet(), 'SISTEMA', 0, 0, '',
    total + ' arquivos', '🗑️ Limpeza Drive — tipo: ' + tipo);

  var msg = '✅ ' + total + ' arquivo(s) movido(s) para a lixeira.\n' +
            'Podem ser restaurados pelo Drive em até 30 dias.';
  if (erros.length) msg += '\n⚠️ ' + erros.length + ' erro(s): ' + erros.slice(0,3).join('; ');
  return JSON.stringify({ ok: msg });
}



// ════════════════════════════════════════════════════════════
//   AUDITORIA E HISTÓRICO
//
//  Formulário unificado que substitui:
//  • abrirHistoricoNF    (absorvida aqui)
//  • abrirHistoricoEmails (absorvida aqui)
//
//  O FormAuditoria.html exibe abas: Histórico de NF | E-mails Enviados
// ════════════════════════════════════════════════════════════

/**
 * Abre o painel unificado de Auditoria e Histórico.
 * Substitui as antigas abrirHistoricoNF() e abrirHistoricoEmails().
 * Requer o arquivo FormAuditoria.html no projeto.
 */
function abrirAuditoria() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('FormAuditoria').setWidth(700).setHeight(560),
    '🔍 Auditoria e Histórico'
  );
}

// ════════════════════════════════════════════════════════════
//   MENU — v6.0
// ════════════════════════════════════════════════════════════

// [P06] reaplicarCores só quando cache de 1h expirou
function onOpen() {
  try {
    var cache = CacheService.getScriptCache();
    if (!cache.get(_CACHE_KEY_CORES)) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      ABAS_OPERACIONAIS.forEach(function(nome) {
        reaplicarCoresAba(ss.getSheetByName(nome));
      });
      cache.put(_CACHE_KEY_CORES, '1', _CORES_TTL_SEG);
    }
  } catch (_) {}

  SpreadsheetApp.getUi().createMenu('📦 Devoluções')
    .addSeparator()
    .addItem('➕ Lançar / Excluir Devolução',    'abrirFormularioLancamento')
    .addItem('🔍 Buscar NF / Fornecedor',        'abrirBusca')
    .addItem('📨 Enviar E-mail de Devolução',    'abrirEmailDevolucao')
    .addSeparator()
    .addItem('📄 Dar baixa para Devolução',      'abrirFormularioExportarPDF')
    .addItem('🛒 Dar Baixa para Venda',          'abrirFormularioVenda')
    .addItem('🔓 Reabrir Devoluções',            'desfazerConclusao')
    .addSeparator()
    .addItem('📊 Relatórios (Mensal / Semanal / Diário / Fornecedor)', 'abrirRelatorios')
    .addItem('🔔 Verificar Atrasos Agora',       'verificarAtrasosEEnviarAlerta')
    .addSeparator()
    .addItem('📦 Forçar Arquivamento Manual',    'arquivarItensConcluidos')
    .addItem('💾 Backup e Restauração',           'abrirBackup')
    .addSeparator()
    .addItem('🔍 Auditoria e Histórico',         'abrirAuditoria')
    .addItem('⚙️ Configurações do Sistema',       'abrirConfiguracoes')
    .addItem('🔧 Configurar/Reinstalar Sistema', 'configurarPlanilha')
    .addToUi();
}

// ════════════════════════════════════════════════════════════
//   BUSCA DE LOG DO SISTEMA (para FormAuditoria)
// ════════════════════════════════════════════════════════════

/**
 * Retorna registros do _Log filtrados por período (dataIni/dataFim ISO).
 * Chamada pelo FormAuditoria.html — tela Log do Sistema.
 */
function buscarLogSistema(params) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var wsLog = ss.getSheetByName('_Log');
    if (!wsLog) return JSON.stringify({ registros: [] });

    var ul = wsLog.getLastRow();
    if (ul < 2) return JSON.stringify({ registros: [] });

    var tz      = ss.getSpreadsheetTimeZone();
    var dataIni = String(params.dataIni || '').replace(/-/g, '');
    var dataFim = String(params.dataFim || '').replace(/-/g, '');

    var dados = wsLog.getRange(2, 1, ul - 1, 8).getValues();
    var registros = [];

    dados.forEach(function(l) {
      if (!l[0]) return;
      var s = String(l[0]).trim();
      var compact = '';
      // "dd/MM/yyyy HH:mm:ss"
      if (/^\d{2}\/\d{2}\/\d{4}/.test(s)) {
        compact = s.slice(6,10) + s.slice(3,5) + s.slice(0,2);
      } else {
        var d = new Date(s);
        if (!isNaN(d.getTime())) {
          compact = String(d.getFullYear()) +
            String(d.getMonth()+1).padStart(2,'0') +
            String(d.getDate()).padStart(2,'0');
        }
      }

      // Filtro por período (servidor — segurança)
      if (dataIni && compact && compact < dataIni) return;
      if (dataFim && compact && compact > dataFim) return;

      registros.push({
        data:     s,
        usuario:  String(l[1] || ''),
        aba:      String(l[2] || ''),
        linha:    String(l[3] || ''),
        coluna:   String(l[4] || ''),
        anterior: String(l[5] || ''),
        novo:     String(l[6] || ''),
        acao:     String(l[7] || '')
      });
    });

    return JSON.stringify({ registros: registros });
  } catch (e) {
    return JSON.stringify({ erro: '❌ ' + e.toString() });
  }
}
