/**********************************************************
 *  ESCALA DE PERMANÊNCIA HOSPITALAR
 *  Formulário web + controle de datas numa planilha do Google
 * ---------------------------------------------------------
 *  ESTE ARQUIVO VAI NO ARQUIVO CHAMADO "Código" (Código.gs)
 *  Arquivos do mesmo projeto (criar como HTML): Index, Styles, Script
 **********************************************************/

/* ===================== CONFIGURAÇÕES ===================== */

var ABA_CONFIG   = 'Config';
var ABA_DATAS    = 'Datas';
var ABA_RESERVAS = 'Reservas';

var COL_DATA   = 1;
var COL_STATUS = 2;
var COL_OBS    = 3;

var COL_RES_DATA   = 1;
var COL_RES_NOME   = 2;
var COL_RES_QUANDO = 3;

var FUSO = 'America/Sao_Paulo';

var MIN_NOME = 2;
var MAX_NOME = 80;
var ESPERA_LOCK_MS = 20000;
var MAX_DIAS_FRENTE = 365;   // até quantos dias à frente dá para reservar

var DIAS = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira',
            'quinta-feira', 'sexta-feira', 'sábado'];
var MESES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
             'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];

var CONFIG_PADRAO = {
  titulo: 'Escala de Permanência',
  subtitulo: 'Reserve a noite que você poderá ficar com seu tio no hospital.',
  paciente: 'Adriano Coutinho',
  hospital: '',
  responsavel: '',
  instrucoes: 'Combinem juntos o horário de entrada e tragam documento com foto.',
  aviso: ''
};


/* ===================== ENTRADA ============================ */

function doGet() {
  var config = configComDados();
  var t = HtmlService.createTemplateFromFile('Index');
  t.configJson = jsonSeguro(config);
  return t.evaluate()
    .setTitle(config.titulo || 'Escala de Permanência')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function jsonSeguro(obj) {
  return JSON.stringify(obj)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');
}


/* ===================== REGRAS DE NEGÓCIO =================== */
/*
 *  Qualquer data futura (até MAX_DIAS_FRENTE) está disponível, a menos que:
 *   - já tenha uma reserva na aba Reservas, ou
 *   - esteja bloqueada na aba Datas (Status "Não", "Bloqueada" etc.).
 */

/** Tudo o que a página precisa para responder na hora, sem ir ao servidor a cada data */
function obterDados() {
  var hoje = hojeISO();
  return {
    hoje: hoje,
    max: somarDias(hoje, MAX_DIAS_FRENTE),
    reservas: reservasPorData(),
    bloqueios: bloqueios()
  };
}

function configComDados() {
  var config = lerConfig();
  config.dados = obterDados();
  return config;
}

function reservar(payload) {
  payload = payload || {};

  var nome = normalizarNome(payload.nome);
  if (nome.length < MIN_NOME) {
    return { ok: false, mensagem: 'Escreva seu nome para concluir a reserva.' };
  }
  if (nome.length > MAX_NOME) {
    return { ok: false, mensagem: 'O nome está muito longo. Use apenas seu nome completo.' };
  }

  var iso = parsearData(payload.data);
  if (!iso) {
    return { ok: false, mensagem: 'Data inválida. Escolha a data pelo calendário.' };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(ESPERA_LOCK_MS);
  } catch (err) {
    return { ok: false, mensagem: 'Muitas pessoas acessando ao mesmo tempo. Tente de novo em instantes.' };
  }

  try {
    // Consulta de novo dentro da trava: a situação pode ter mudado desde a tela
    var atual = situacao(iso);
    if (atual.status !== 'DISPONIVEL') {
      return { ok: false, mensagem: atual.motivo, item: atual };
    }

    aba(ABA_RESERVAS).appendRow([
      iso,
      nome,
      Utilities.formatDate(new Date(), fuso(), 'dd/MM/yyyy HH:mm:ss')
    ]);
    // Grava de fato antes de liberar a trava: quem vier em seguida já enxerga esta reserva
    SpreadsheetApp.flush();

    return { ok: true, mensagem: 'Reserva confirmada', nome: nome,
             item: montarItem(iso, 'RESERVADA', '', nome) };
  } catch (err) {
    return { ok: false, mensagem: 'Não foi possível registrar agora. Tente novamente em instantes.' };
  } finally {
    lock.releaseLock();
  }
}

function situacao(iso) {
  var hoje = hojeISO();

  if (iso < hoje) {
    return montarItem(iso, 'PASSADA', 'Essa data já passou. Escolha uma noite a partir de hoje.', '');
  }
  if (iso > somarDias(hoje, MAX_DIAS_FRENTE)) {
    return montarItem(iso, 'FORA', 'Só dá para reservar com até ' + MAX_DIAS_FRENTE +
                      ' dias de antecedência. Confira o ano da data.', '');
  }

  var quem = reservasPorData()[iso];
  if (quem) {
    return montarItem(iso, 'RESERVADA', 'Esta noite já está com ' + quem + '.', quem);
  }

  var blq = bloqueios();
  if (blq.hasOwnProperty(iso)) {
    return montarItem(iso, 'BLOQUEADA',
      blq[iso] ? 'Indisponível: ' + blq[iso] : 'Esta noite não está disponível.', '');
  }

  return montarItem(iso, 'DISPONIVEL', '', '');
}


/* ===================== PLANILHA =========================== */

function planilha() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Abra a planilha antes de executar esta função.');
  return ss;
}

function aba(nome) {
  var ss = planilha();
  var a = ss.getSheetByName(nome);
  if (!a) a = ss.insertSheet(nome);
  return a;
}

// Consulta o fuso da planilha uma vez só por execução
var _fusoCache = null;
function fuso() {
  if (_fusoCache) return _fusoCache;
  try { _fusoCache = planilha().getSpreadsheetTimeZone() || FUSO; } catch (e) { _fusoCache = FUSO; }
  return _fusoCache;
}

/** { 'yyyy-MM-dd': 'observação' } das datas bloqueadas na aba Datas */
function bloqueios() {
  var valores = aba(ABA_DATAS).getDataRange().getValues();
  var mapa = {};
  for (var i = 1; i < valores.length; i++) {
    var iso = parsearData(valores[i][COL_DATA - 1]);
    if (!iso || ehLiberada(valores[i][COL_STATUS - 1])) continue;
    mapa[iso] = texto(valores[i][COL_OBS - 1]).trim();
  }
  return mapa;
}

/** { 'yyyy-MM-dd': 'nome de quem reservou' } da aba Reservas */
function reservasPorData() {
  var valores = aba(ABA_RESERVAS).getDataRange().getValues();
  var mapa = {};
  for (var i = 1; i < valores.length; i++) {
    var iso = parsearData(valores[i][COL_RES_DATA - 1]);
    var nome = normalizarNome(valores[i][COL_RES_NOME - 1]);
    if (!iso || mapa[iso]) continue;
    mapa[iso] = nome || 'outra pessoa';
  }
  return mapa;
}

function montarItem(iso, status, motivo, nome) {
  var d = dataDeISO(iso);
  return {
    data: iso,
    rotulo: DIAS[d.getDay()] + ', ' + d.getDate() + ' de ' +
            MESES[d.getMonth()] + ' de ' + d.getFullYear(),
    status: status,
    motivo: motivo,
    nome: nome
  };
}

function lerConfig() {
  var valores = aba(ABA_CONFIG).getDataRange().getValues();
  var config = {};
  for (var k in CONFIG_PADRAO) {
    if (CONFIG_PADRAO.hasOwnProperty(k)) config[k] = CONFIG_PADRAO[k];
  }
  for (var i = 1; i < valores.length; i++) {
    var chave = semAcento(texto(valores[i][0]));
    if (!chave) continue;
    config[chave] = texto(valores[i][1]);
  }
  config.maxDias = MAX_DIAS_FRENTE;
  return config;
}


/* ===================== UTILIDADES ========================= */

function parsearData(valor) {
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, fuso(), 'yyyy-MM-dd');
  }

  var s = texto(valor).trim();
  if (!s) return null;

  var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) {
    var ano = m[3].length === 2 ? '20' + m[3] : m[3];
    return montarISO(m[1], m[2], ano);
  }

  m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/);
  if (m) return montarISO(m[3], m[2], m[1]);

  return null;
}

function montarISO(dia, mes, ano) {
  var d = Number(dia), me = Number(mes), a = Number(ano);
  if (a < 1900 || me < 1 || me > 12 || d < 1 || d > 31) return null;
  var data = new Date(a, me - 1, d);
  if (data.getFullYear() !== a || data.getMonth() !== me - 1 || data.getDate() !== d) return null;
  return a + '-' + dois(me) + '-' + dois(d);
}

function dataDeISO(iso) {
  var p = String(iso).split('-');
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

function somarDias(iso, n) {
  var d = dataDeISO(iso);
  d.setDate(d.getDate() + n);
  return montarISO(d.getDate(), d.getMonth() + 1, d.getFullYear());
}

function hojeISO() {
  return Utilities.formatDate(new Date(), fuso(), 'yyyy-MM-dd');
}

function dois(n) { return (n < 10 ? '0' : '') + n; }

/** "Disponível " -> "disponivel" (minúsculo, sem acento, sem espaços nas pontas) */
function semAcento(s) {
  return String(s == null ? '' : s)
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function ehLiberada(valor) {
  var s = semAcento(texto(valor));
  if (!s) return true;
  return ['liberada', 'livre', 'liberado', 'disponivel', 'sim', 'ok',
          's', '1', 'x', 'y', 'true', 'yes', 'verdadeiro'].indexOf(s) >= 0;
}

function normalizarNome(valor) {
  return texto(valor).replace(/\s+/g, ' ').trim();
}

function texto(valor) {
  return (valor === null || valor === undefined) ? '' : String(valor);
}


/* ===================== MENU E UTILIDADES ================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Escala')
    .addItem('Abrir formulário', 'abrirFormulario')
    .addItem('Ver prévia do formulário', 'verPreviaDoFormulario')
    .addItem('Inicializar planilha', 'inicializarPlanilha')
    .addSeparator()
    .addItem('Remover bloqueios que já passaram', 'limparDatasPassadas')
    .addToUi();
}

function inicializarPlanilha() {
  criarAbas();
  planilha().toast('Pronto! Qualquer data futura já pode ser reservada. ' +
                   'Use a aba Datas só para bloquear noites.', 'Escala', 8);
}

function criarAbas() {
  comCabecalho(ABA_CONFIG,   ['Chave', 'Valor']);
  comCabecalho(ABA_DATAS,    ['Data', 'Status', 'Observação']);
  comCabecalho(ABA_RESERVAS, ['Data', 'Nome', 'Reservado em']);

  aba(ABA_DATAS).getRange('A1').setNote(
    'Use esta aba só para BLOQUEAR noites.\n' +
    'Data: a noite a bloquear\n' +
    'Status: escreva "Não" (ou "Bloqueada")\n' +
    'Observação: motivo, aparece para quem acessar o link.\n' +
    'Datas que não estão aqui ficam livres.');

  var cfg = aba(ABA_CONFIG);
  if (cfg.getLastRow() < 2) {
    var linhas = [];
    for (var chave in CONFIG_PADRAO) {
      if (CONFIG_PADRAO.hasOwnProperty(chave)) linhas.push([chave, CONFIG_PADRAO[chave]]);
    }
    cfg.getRange(2, 1, linhas.length, 2).setValues(linhas);
  }
}

function comCabecalho(nome, cabecalho) {
  var a = aba(nome);
  a.getRange(1, 1, 1, cabecalho.length)
    .setValues([cabecalho])
    .setFontWeight('bold')
    .setBackground('#0f3d3e')
    .setFontColor('#ffffff');
  a.setFrozenRows(1);
  return a;
}

function limparDatasPassadas() {
  var a = aba(ABA_DATAS);
  var hoje = hojeISO();
  var valores = a.getDataRange().getValues();
  for (var i = valores.length - 1; i >= 1; i--) {
    var iso = parsearData(valores[i][COL_DATA - 1]);
    if (iso && iso < hoje) a.deleteRow(i + 1);
  }
  planilha().toast('Bloqueios antigos removidos.', 'Escala', 4);
}

function abrirFormulario() {
  var url = ScriptApp.getService().getUrl();
  if (url) {
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput('<p><a href="' + url + '" target="_blank">Abrir formulário</a></p>')
        .setWidth(320).setHeight(80),
      'Escala');
  } else {
    planilha().toast('Falta fazer a implantação como aplicativo web.', 'Escala', 6);
  }
}

function verPreviaDoFormulario() {
  var url = ScriptApp.getService().getUrl();
  if (url) {
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput('<iframe style="width:100%;height:70vh;border:0" src="' +
        url + '"></iframe>').setWidth(900).setHeight(700),
      'Prévia do formulário');
    return;
  }
  var t = HtmlService.createTemplateFromFile('Index');
  t.configJson = jsonSeguro(configComDados());
  SpreadsheetApp.getUi().showModalDialog(
    t.evaluate().setWidth(900).setHeight(700),
    'Prévia do formulário');
}

// Lê o arquivo como texto bruto (sem validar como HTML), porque Script e Styles
// não são HTML: o "<" do JavaScript faria o Google rejeitar o arquivo.
function include(nomeArquivo) {
  return HtmlService.createTemplateFromFile(nomeArquivo).getRawContent();
}
