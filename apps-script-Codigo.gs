/**
 * Central 3D Ostras — banco de dados na Planilha Google
 *
 * 1. Troque o PIN abaixo pela senha que a equipe vai usar para entrar no app.
 * 2. Rode a função "configurar" uma vez (cria a planilha e carrega os produtos).
 * 3. Implante como App da Web: Executar como "Eu", Acesso "Qualquer pessoa".
 */
const PIN = '1234';

const ABAS = ['produtos', 'vendas', 'investimentos', 'testes', 'propostas', 'config'];
const NOME_PLANILHA = 'Central 3D Ostras — dados';

function configurar() {
  const ss = planilha_();
  ABAS.forEach(aba_);
  ss.getSheets().forEach(function (sh) {
    if (ABAS.indexOf(sh.getName()) < 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  if (aba_('produtos').getLastRow() < 2) {
    SEMENTE_PRODUTOS.forEach(function (p) { gravar_('produtos', p.id, p); });
  }
  if (aba_('config').getLastRow() < 2) {
    gravar_('config', 'params', SEMENTE_PARAMS);
  }
  Logger.log('Pronto. Planilha: ' + ss.getUrl());
}

function doGet() {
  return saida_({ ok: true, app: 'Central 3D Ostras' });
}

function doPost(e) {
  var b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return saida_({ ok: false, error: 'formato' }); }
  if (String(b.pin) !== String(PIN)) return saida_({ ok: false, error: 'pin' });
  try {
    if (b.action === 'all') return saida_({ ok: true, data: tudo_() });
    if (ABAS.indexOf(b.col) < 0 || !b.id) return saida_({ ok: false, error: 'colecao' });
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (b.action === 'set') gravar_(b.col, String(b.id), b.data || {});
      else if (b.action === 'delete') apagar_(b.col, String(b.id));
      else return saida_({ ok: false, error: 'acao' });
    } finally { lock.releaseLock(); }
    return saida_({ ok: true });
  } catch (err) {
    return saida_({ ok: false, error: String(err) });
  }
}

/* ---------- internos ---------- */
function planilha_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var ss = SpreadsheetApp.create(NOME_PLANILHA);
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}
function aba_(nome) {
  var ss = planilha_();
  var sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.appendRow(['id', 'atualizado em', 'dados (não editar à mão)']);
    sh.setFrozenRows(1);
    sh.setColumnWidth(3, 600);
  }
  return sh;
}
function tudo_() {
  var out = {};
  ABAS.forEach(function (nome) {
    var sh = aba_(nome), n = sh.getLastRow();
    out[nome] = [];
    if (n < 2) return;
    sh.getRange(2, 1, n - 1, 3).getValues().forEach(function (r) {
      if (!r[0]) return;
      try { var d = JSON.parse(r[2]); d.id = String(r[0]); out[nome].push(d); } catch (e) {}
    });
  });
  return out;
}
function linha_(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  var f = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}
function gravar_(col, id, data) {
  var sh = aba_(col), copia = JSON.parse(JSON.stringify(data));
  delete copia.id;
  var linha = [id, new Date(), JSON.stringify(copia)];
  var r = linha_(sh, id);
  if (r) sh.getRange(r, 1, 1, 3).setValues([linha]);
  else sh.appendRow(linha);
}
function apagar_(col, id) {
  var sh = aba_(col), r = linha_(sh, id);
  if (r) sh.deleteRow(r);
}
function saida_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- dados iniciais (da planilha Controle Negócio 3D) ---------- */
var OBS = 'Estimativa de pesquisa (set/2026), não medido em máquina';
var SEMENTE_PRODUTOS = [
  { id: 'pPIN', code: 'PIN', nome: 'Pin de lapela - laço de campanha', status: 'Ideia', peso: 1.5, tempo: 2.5, pos: 0.5, extras: 0.15, preco: 6, obs: OBS, ts: 1 },
  { id: 'pCHV', code: 'CHV', nome: 'Chaveiro laço + logo do cliente', status: 'Ideia', peso: 4, tempo: 6, pos: 1, extras: 0.75, preco: 9, obs: OBS, ts: 2 },
  { id: 'pCLP', code: 'CLP', nome: 'Clip / porta-crachá temático', status: 'Ideia', peso: 5, tempo: 7, pos: 1, extras: 0.5, preco: 11, obs: OBS, ts: 3 },
  { id: 'pNFC', code: 'NFC', nome: 'Chaveiro NFC - canal de escuta', status: 'Ideia', peso: 8, tempo: 12, pos: 2, extras: 4, preco: 32, obs: OBS, ts: 4 },
  { id: 'pDSP', code: 'DSP', nome: 'Display de mesa com QR e 188', status: 'Ideia', peso: 45, tempo: 55, pos: 3, extras: 1, preco: 59, obs: OBS, ts: 5 },
  { id: 'pCIPA', code: 'CIPA', nome: 'Placa de reconhecimento CIPA', status: 'Ideia', peso: 120, tempo: 150, pos: 10, extras: 3, preco: 140, obs: OBS, ts: 6 },
  { id: 'pGIR', code: 'GIR', nome: 'Girassol - chaveiro e peça de mesa', status: 'Ideia', peso: 5, tempo: 8, pos: 1, extras: 0.75, preco: 12, obs: OBS, ts: 7 },
  { id: 'pFDG', code: 'FDG', nome: 'Clicker fidget print-in-place', status: 'Ideia', peso: 12, tempo: 18, pos: 0.5, extras: 0.5, preco: 18, obs: OBS, ts: 8 }
];
var SEMENTE_PARAMS = {
  filamento: 95, kwh: 1, potencia: 100, maquina: 4000, vida: 5000, falha: 0.1, fixos: 0, pecasMes: 100,
  hora: 25, hHenrique: 6, hRenato: 6, margemMin: 50, markup: 3.5, descRevenda: 0.5,
  cartao: 0.0399, das: 0.05, mkt: 0.05, maquinas: 2, horasDia: 8, tetoInterno: 30000, tetoLegal: 81000, ano: 2026
};
