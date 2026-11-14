/**
 * Central 3D Ostras — banco de dados na Planilha Google (versão 2)
 * Novidades: estoque, movimentos de estoque, portfólio com fotos no Drive e página pública da bio.
 *
 * Depois de colar: rode "atualizar" uma vez (autoriza o Drive e carrega estoque e portfólio)
 * e publique uma nova versão em Implantar > Gerenciar implantações.
 */
const PIN = '1234';

const ABAS = ['produtos', 'vendas', 'investimentos', 'testes', 'propostas', 'estoque', 'movimentos', 'portfolio', 'config'];
const NOME_PLANILHA = 'Central 3D Ostras — dados';
const NOME_PASTA = 'Central 3D Ostras — fotos do portfólio';

function configurar() { atualizar(); }

function atualizar() {
  const ss = planilha_();
  ABAS.forEach(aba_);
  ss.getSheets().forEach(function (sh) {
    if (ABAS.indexOf(sh.getName()) < 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  if (aba_('produtos').getLastRow() < 2) SEMENTE_PRODUTOS.forEach(function (p) { gravar_('produtos', p.id, p); });
  if (!linha_(aba_('config'), 'params')) gravar_('config', 'params', SEMENTE_PARAMS);
  if (aba_('estoque').getLastRow() < 2) {
    var hoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
    SEMENTE_ESTOQUE.forEach(function (e) {
      gravar_('estoque', e.id, e);
      gravar_('movimentos', 'm-' + e.id, { item: e.id, tipo: 'Entrada', qtd: e.qtd, data: hoje, motivo: 'Saldo inicial', obs: '', ts: Date.now() });
    });
  }
  if (aba_('portfolio').getLastRow() < 2) SEMENTE_PORTFOLIO.forEach(function (p) { gravar_('portfolio', p.id, p); });
  if (!linha_(aba_('config'), 'vitrine')) gravar_('config', 'vitrine', SEMENTE_VITRINE);
  pasta_();
  Logger.log('Pronto. Planilha: ' + ss.getUrl());
}

function doGet() {
  return saida_({ ok: true, app: 'Central 3D Ostras', versao: 2 });
}

function doPost(e) {
  var b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return saida_({ ok: false, error: 'formato' }); }
  try {
    if (b.action === 'vitrine') return saida_({ ok: true, data: vitrine_() });
    if (String(b.pin) !== String(PIN)) return saida_({ ok: false, error: 'pin' });
    if (b.action === 'all') return saida_({ ok: true, data: tudo_() });
    if (b.action === 'upload') return saida_({ ok: true, file: upload_(b) });
    if (b.action === 'trash') { lixeira_(String(b.id || '')); return saida_({ ok: true }); }
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

/* ---------- página pública: só itens visíveis, sem dados internos ---------- */
function vitrine_() {
  var t = tudo_(['portfolio', 'config']);
  var cfg = (t.config || []).filter(function (c) { return c.id === 'vitrine'; })[0] || {};
  var itens = (t.portfolio || []).filter(function (p) { return p.visivel !== false; }).map(function (p) {
    var c = {}; ['id', 'categoria', 'icone', 'selo', 'nome', 'chamada', 'subtitulo', 'descricao', 'beneficios', 'personalizacao', 'specs', 'prazo', 'preco', 'fotos', 'ordem'].forEach(function (k) { c[k] = p[k]; });
    return c;
  });
  delete cfg.id;
  return { settings: cfg, itens: itens };
}

/* ---------- fotos no Google Drive ---------- */
function pasta_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder(NOME_PASTA);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}
function upload_(b) {
  var bytes = Utilities.base64Decode(String(b.data || ''));
  if (!bytes.length || bytes.length > 8 * 1024 * 1024) throw new Error('arquivo');
  var mime = /^image\/(jpeg|png|webp)$/.test(b.mime) ? b.mime : 'image/jpeg';
  var file = pasta_().createFile(Utilities.newBlob(bytes, mime, String(b.name || 'foto.jpg').slice(0, 80)));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { id: file.getId() };
}
function lixeira_(id) {
  if (!id) return;
  var file = DriveApp.getFileById(id), pasta = pasta_().getId(), it = file.getParents();
  while (it.hasNext()) { if (it.next().getId() === pasta) { file.setTrashed(true); return; } }
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
function tudo_(abas) {
  var out = {};
  (abas || ABAS).forEach(function (nome) {
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

/* ---------- dados iniciais ---------- */
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
var SEMENTE_ESTOQUE = [{"id": "e-ro-preto", "nome": "Filamento preto", "categoria": "Filamento", "material": "", "cor": "Preto", "unidade": "rolo", "qtd": 2, "minimo": 1, "custo": 95, "local": "Rio das Ostras", "fornecedor": "", "obs": "Estoque inicial (set/2026). Confira a contagem."}, {"id": "e-ro-branco", "nome": "Filamento branco", "categoria": "Filamento", "material": "", "cor": "Branco", "unidade": "rolo", "qtd": 2, "minimo": 1, "custo": 95, "local": "Rio das Ostras", "fornecedor": "", "obs": "Estoque inicial (set/2026). Confira a contagem."}, {"id": "e-ro-azul", "nome": "Filamento azul", "categoria": "Filamento", "material": "", "cor": "Azul", "unidade": "rolo", "qtd": 1, "minimo": 1, "custo": 95, "local": "Rio das Ostras", "fornecedor": "", "obs": "Estoque inicial (set/2026). Confira a contagem."}, {"id": "e-ro-vermelho", "nome": "Filamento vermelho", "categoria": "Filamento", "material": "", "cor": "Vermelho", "unidade": "rolo", "qtd": 1, "minimo": 1, "custo": 95, "local": "Rio das Ostras", "fornecedor": "", "obs": "Estoque inicial (set/2026). Confira a contagem."}, {"id": "e-ro-rosa", "nome": "Filamento rosa", "categoria": "Filamento", "material": "", "cor": "Rosa", "unidade": "rolo", "qtd": 1, "minimo": 1, "custo": 95, "local": "Rio das Ostras", "fornecedor": "", "obs": "Estoque inicial (set/2026). Confira a contagem."}, {"id": "e-ro-bege", "nome": "Filamento bege", "categoria": "Filamento", "material": "", "cor": "Bege", "unidade": "rolo", "qtd": 1, "minimo": 1, "custo": 95, "local": "Rio das Ostras", "fornecedor": "", "obs": "Estoque inicial (set/2026). Confira a contagem."}, {"id": "e-un-preto", "nome": "Filamento preto", "categoria": "Filamento", "material": "", "cor": "Preto", "unidade": "rolo", "qtd": 2, "minimo": 1, "custo": 95, "local": "Unamar", "fornecedor": "", "obs": "Estoque inicial (set/2026). Confira a contagem."}, {"id": "e-un-branco", "nome": "Filamento branco", "categoria": "Filamento", "material": "", "cor": "Branco", "unidade": "rolo", "qtd": 2, "minimo": 1, "custo": 95, "local": "Unamar", "fornecedor": "", "obs": "Estoque inicial (set/2026). Confira a contagem."}, {"id": "e-un-azul", "nome": "Filamento azul", "categoria": "Filamento", "material": "", "cor": "Azul", "unidade": "rolo", "qtd": 1, "minimo": 1, "custo": 95, "local": "Unamar", "fornecedor": "", "obs": "Estoque inicial (set/2026). Confira a contagem."}, {"id": "e-un-vermelho", "nome": "Filamento vermelho", "categoria": "Filamento", "material": "", "cor": "Vermelho", "unidade": "rolo", "qtd": 1, "minimo": 1, "custo": 95, "local": "Unamar", "fornecedor": "", "obs": "Estoque inicial (set/2026). Confira a contagem."}, {"id": "e-un-rosa", "nome": "Filamento rosa", "categoria": "Filamento", "material": "", "cor": "Rosa", "unidade": "rolo", "qtd": 1, "minimo": 1, "custo": 95, "local": "Unamar", "fornecedor": "", "obs": "Estoque inicial (set/2026). Confira a contagem."}, {"id": "e-un-bege", "nome": "Filamento bege", "categoria": "Filamento", "material": "", "cor": "Bege", "unidade": "rolo", "qtd": 1, "minimo": 1, "custo": 95, "local": "Unamar", "fornecedor": "", "obs": "Estoque inicial (set/2026). Confira a contagem."}];
var SEMENTE_PORTFOLIO = [{"id": "pf-combo-troca-de-faixa", "categoria": "artes-marciais", "icone": "kimono", "selo": "Mais pedido", "nome": "Combo Troca de Faixa", "chamada": "Cada graduação no lugar de honra.", "subtitulo": "Kimono 3D de 20 cm com base personalizada e suporte para todas as faixas.", "descricao": "Um kimono em miniatura sobre base com placa de nome, acompanhado de um suporte que guarda as faixas da sua jornada. Subiu de faixa? É só trocar a do kimono. A peça acompanha a trajetória do atleta no tatame, da branca à preta.", "beneficios": ["Faixas intercambiáveis que acompanham cada graduação", "Placa com nome, equipe e data de início", "Kimono branco, azul ou preto", "Presente marcante para aluno, atleta ou professor"], "personalizacao": ["Nome do atleta", "Logo ou bandeira da equipe", "Data de início", "Cor do kimono", "Jiu-jitsu, judô ou karatê"], "specs": ["Altura do kimono: 20 cm", "Suporte: faixas da branca à preta", "Material: PLA/PETG fosco", "Cores: direto do filamento, sem pintura"], "prazo": "7 a 10 dias úteis", "preco": "Sob consulta", "hashtags": ["#JiuJitsu", "#BJJ", "#Graduacao", "#PresenteParaAtleta"], "visivel": true, "fotos": [], "ordem": 1, "code": ""}, {"id": "pf-chaveiro-faixa-com-graus", "categoria": "artes-marciais", "icone": "chaveiro", "selo": "", "nome": "Chaveiro Faixa com Graus", "chamada": "Sua graduação no bolso.", "subtitulo": "Chaveiro em formato de faixa, com ponteira, graus e nome.", "descricao": "A faixa em miniatura, com a ponteira e os graus da graduação atual. Leve e firme, funciona como lembrança de graduação, item de venda na academia ou brinde de campeonato.", "beneficios": ["Cores de faixa da branca à preta", "De 0 a 4 graus na ponteira", "Nome ou equipe em relevo", "Sai bem em lote para graduações e eventos"], "personalizacao": ["Cor da faixa", "Quantidade de graus", "Nome ou equipe", "Argola ou mosquetão"], "specs": ["Comprimento: cerca de 8 cm", "Material: PLA", "Pedido: unidade ou lote"], "prazo": "3 a 5 dias úteis", "preco": "Sob consulta", "hashtags": ["#JiuJitsu", "#Graduacao", "#ChaveiroPersonalizado"], "visivel": true, "fotos": [], "ordem": 2, "code": ""}, {"id": "pf-trofeu-homenagem-ao-mestre", "categoria": "artes-marciais", "icone": "kimono", "selo": "Dia do Professor", "nome": "Troféu Homenagem ao Mestre", "chamada": "Para quem ensinou mais que técnica.", "subtitulo": "Kimono 3D com a faixa do professor e placa de homenagem da turma.", "descricao": "Peça comemorativa para Dia do Professor, aniversário da academia ou graduação do mestre. O kimono vem com a faixa e os graus do professor, e a base leva uma mensagem escrita pelos alunos.", "beneficios": ["Mensagem da turma gravada na base", "Faixa e graus fiéis à graduação do professor", "Presente coletivo: a academia inteira assina", "Embalagem pronta para entregar"], "personalizacao": ["Nome do mestre", "Mensagem da turma", "Logo da academia", "Data especial"], "specs": ["Altura: 20 cm", "Material: PLA/PETG fosco", "Placa: texto em relevo"], "prazo": "7 a 10 dias úteis", "preco": "Sob consulta", "hashtags": ["#JiuJitsu", "#DiaDoProfessor", "#Mestre"], "visivel": true, "fotos": [], "ordem": 3, "code": ""}, {"id": "pf-copa-lagos-beach-tennis", "categoria": "trofeus", "icone": "trofeu", "selo": "Modelo exemplo", "nome": "Troféu Copa Lagos de Beach Tennis", "chamada": "O pódio com a cara do seu torneio.", "subtitulo": "Troféu recortado com a arte do evento e base com categoria e colocação.", "descricao": "Modelo criado para a Copa Lagos de Beach Tennis, um evento fictício que mostra o que fazemos: a arte do torneio vira um troféu em relevo, nas cores da identidade, com base para categoria e colocação. Serve para esportes de areia, quadra ou pista.", "beneficios": ["Arte do evento em relevo e em cores, sem adesivo", "Base com categoria e colocação", "Mesmo modelo em alturas diferentes para o pódio", "Várias máquinas em paralelo para cumprir o prazo do lote"], "personalizacao": ["Logo e arte do evento", "Nome do torneio e ano", "Categorias", "Cores da marca"], "specs": ["Altura: 15 a 25 cm", "Material: PLA/PETG", "Base: com placa de colocação"], "prazo": "10 a 15 dias úteis, conforme o lote", "preco": "Sob consulta", "hashtags": ["#BeachTennis", "#Trofeu", "#Premiacao"], "visivel": true, "fotos": [], "ordem": 4, "code": ""}, {"id": "pf-medalha-corrida-costa-azul", "categoria": "trofeus", "icone": "medalha", "selo": "Modelo exemplo", "nome": "Medalha Corrida Costa Azul 10K", "chamada": "A medalha que o corredor faz questão de mostrar.", "subtitulo": "Medalha de finisher em camadas de cor, com relevo e passagem para fita.", "descricao": "Exemplo feito para a Corrida Costa Azul 10K, evento fictício. A medalha tem distância, ano e percurso em relevo, várias cores na mesma peça e bordas arredondadas. Pronta para receber a fita da sua prova.", "beneficios": ["Várias cores na mesma peça, direto da impressão", "Relevo com distância, ano e percurso", "Leve e sem cantos vivos", "Para corrida, ciclismo, natação e treino funcional"], "personalizacao": ["Arte do evento", "Distância e ano", "Cores", "Passagem para fita"], "specs": ["Diâmetro: 7 a 9 cm", "Espessura: cerca de 4 mm", "Material: PLA/PETG"], "prazo": "10 a 15 dias úteis, conforme o lote", "preco": "Sob consulta", "hashtags": ["#Corrida", "#MedalhaFinisher", "#Corredores"], "visivel": true, "fotos": [], "ordem": 5, "code": ""}, {"id": "pf-trofeu-desafio-mare-alta", "categoria": "trofeus", "icone": "trofeu", "selo": "Modelo exemplo", "nome": "Troféu Desafio Maré Alta", "chamada": "Esforço de competição merece troféu à altura.", "subtitulo": "Troféu para desafios de treino funcional, box e academias.", "descricao": "Modelo feito para o Desafio Maré Alta, evento fictício de treino funcional. Silhueta do atleta, nome da prova em relevo e base escura com placa de categoria. Pensado para box, academias e desafios internos entre alunos.", "beneficios": ["Silhueta e nome da prova em relevo", "Placas de categoria: RX, scaled, master", "Visual forte em foto de pódio", "Troféus e medalhas no mesmo padrão visual"], "personalizacao": ["Logo do box", "Nome do desafio", "Categorias", "Cores"], "specs": ["Altura: 18 a 25 cm", "Material: PLA/PETG", "Base: grafite com placa"], "prazo": "10 a 15 dias úteis, conforme o lote", "preco": "Sob consulta", "hashtags": ["#TreinoFuncional", "#Trofeu", "#Box"], "visivel": true, "fotos": [], "ordem": 6, "code": ""}, {"id": "pf-trofeu-destaque-do-ano", "categoria": "empresas", "icone": "trofeu", "selo": "B2B", "nome": "Troféu Destaque do Ano", "chamada": "Reconhecimento com a marca da sua empresa.", "subtitulo": "Premiação corporativa para colaboradores, metas e SIPAT.", "descricao": "Modelo feito para a Atlântica Energia, empresa fictícia, como exemplo de premiação interna: colaborador destaque, meta de segurança, tempo de casa ou encerramento de SIPAT. A marca ganha volume e a base leva nome, área e ano.", "beneficios": ["Logo em relevo, nas cores da marca", "Série com nomes diferentes no mesmo lote", "Para SIPAT, convenções e metas de segurança", "Prazo e quantidade combinados por escrito"], "personalizacao": ["Logo da empresa", "Nome, área e ano", "Categoria do prêmio", "Cores corporativas"], "specs": ["Altura: 15 a 25 cm", "Material: PETG", "Lote: sob medida"], "prazo": "Sob consulta", "preco": "Sob consulta", "hashtags": ["#SIPAT", "#Reconhecimento", "#BrindeCorporativo"], "visivel": true, "fotos": [], "ordem": 7, "code": ""}, {"id": "pf-plaquinha-de-mesa-profissao", "categoria": "profissoes", "icone": "placa", "selo": "Novo", "nome": "Plaquinha de Mesa por Profissão", "chamada": "Seu nome e sua profissão com presença na mesa.", "subtitulo": "Placa de mesa com nome, profissão, símbolo da área e registro.", "descricao": "Plaquinha de mesa em camadas de cor, com o símbolo da profissão, o nome e o registro profissional. Fica bem em consultório, escritório, balcão de atendimento e home office. Também é presente de formatura que a pessoa usa todo dia.", "beneficios": ["Símbolo da área: saúde, direito, engenharia, educação, TI e outras", "Nome e registro (CRM, OAB, CREA, CRO) em relevo", "Cores da profissão ou da sua marca", "Presente certeiro para formatura e emprego novo"], "personalizacao": ["Nome e título", "Profissão e símbolo", "Número de registro", "Cores", "Versão com NFC"], "specs": ["Largura: 15 a 20 cm", "Altura: cerca de 5 cm", "Material: PLA/PETG"], "prazo": "5 a 7 dias úteis", "preco": "Sob consulta", "hashtags": ["#PlacaDeMesa", "#PresenteDeFormatura", "#Profissao"], "visivel": true, "fotos": [], "ordem": 8, "code": ""}, {"id": "pf-display-nfc-siga-e-avalie", "categoria": "nfc", "icone": "nfc", "selo": "", "nome": "Display NFC Siga e Avalie", "chamada": "Um toque e o cliente já te segue.", "subtitulo": "Display de balcão com a sua marca, chip NFC e QR code.", "descricao": "Display de mesa com a sua marca e um chip NFC dentro da peça. O cliente aproxima o celular e abre seu Instagram, sua avaliação no Google, o WhatsApp ou a senha do Wi-Fi. O QR code atende quem tem celular sem NFC.", "beneficios": ["Abre Instagram, Google, WhatsApp ou cardápio em um toque", "QR code junto, para qualquer celular", "O destino do link pode mudar sem trocar a peça", "Para consultórios, cafés, salões, lojas e estandes"], "personalizacao": ["Logo e cores", "Destino do link", "Formato e tamanho", "De 1 a 3 ícones de ação"], "specs": ["Chip: NTAG213", "Tamanho: cerca de 10 x 12 cm", "Material: PLA/PETG"], "prazo": "5 a 7 dias úteis", "preco": "Sob consulta", "hashtags": ["#NFC", "#MarketingLocal", "#AvaliacaoGoogle"], "visivel": true, "fotos": [], "ordem": 9, "code": ""}, {"id": "pf-chaveiro-nfc-cartao-de-visita", "categoria": "nfc", "icone": "chaveiro", "selo": "", "nome": "Chaveiro NFC Cartão de Visita", "chamada": "Seu contato em um toque.", "subtitulo": "Chaveiro com a sua marca e chip NFC que abre seu perfil ou cartão digital.", "descricao": "Chaveiro com logo em relevo e chip NFC embutido. Encostou no celular, abre seu Instagram, WhatsApp ou cartão de visita digital. Para networking, feiras, corretores, vendedores e profissionais liberais.", "beneficios": ["Contato compartilhado sem digitar nada", "Link reprogramável", "Marca sempre à mão do cliente", "Em lote para equipes de venda"], "personalizacao": ["Logo e cores", "Destino do link", "Formato", "Quantidade"], "specs": ["Chip: NTAG213", "Tamanho: cerca de 5 cm", "Material: PLA/PETG"], "prazo": "5 a 7 dias úteis", "preco": "Sob consulta", "hashtags": ["#NFC", "#Networking", "#CartaoDigital"], "visivel": true, "fotos": [], "ordem": 10, "code": ""}, {"id": "pf-tag-nfc-inspecao-epi", "categoria": "empresas", "icone": "tag", "selo": "B2B", "nome": "Tag NFC de Inspeção de EPI", "chamada": "Inspeção registrada no celular, direto no equipamento.", "subtitulo": "Tag com chip NFC para EPI, cintas, ferramentas e instrumentos.", "descricao": "Tag presa ao equipamento, com chip NFC embutido durante a impressão. O técnico aproxima o celular e abre o formulário de inspeção ou o histórico daquele item. Também montamos os links e a planilha de controle da sua equipe.", "beneficios": ["Chip protegido dentro da peça", "Abre formulário, checklist ou histórico do item", "Cor por status ou por ciclo de inspeção", "Versão para superfície metálica"], "personalizacao": ["Código do item", "Cor por ciclo", "Destino do link", "Formato de fixação"], "specs": ["Material: PETG ou ASA para área externa", "Chip: NTAG213 ou anti-metal", "Implantação: links e planilha"], "prazo": "Sob consulta", "preco": "Sob consulta", "hashtags": ["#SegurancaDoTrabalho", "#SST", "#NR35", "#Inspecao"], "visivel": true, "fotos": [], "ordem": 11, "code": ""}, {"id": "pf-chaveiro-brinde-com-logo", "categoria": "empresas", "icone": "chaveiro", "selo": "", "nome": "Chaveiro Personalizado com Logo", "chamada": "O brinde que fica no bolso do cliente.", "subtitulo": "Chaveiro com a marca em relevo, feito em lote para eventos e campanhas.", "descricao": "Chaveiro com a logo em relevo e nas cores da empresa, produzido em lote para feiras, datas comemorativas e campanhas como Outubro Rosa e Novembro Azul. Pode ganhar clicker antiestresse ou chip NFC.", "beneficios": ["Logo em relevo, nas cores da marca", "Lotes pequenos ou grandes", "Opção com clicker ou NFC", "Embalagem individual com cartão"], "personalizacao": ["Logo e cores", "Formato", "Clicker ou NFC", "Embalagem"], "specs": ["Tamanho: 4 a 6 cm", "Material: PLA/PETG", "Pedido: lote"], "prazo": "7 a 12 dias úteis, conforme o lote", "preco": "Sob consulta", "hashtags": ["#BrindeCorporativo", "#OutubroRosa", "#NovembroAzul"], "visivel": true, "fotos": [], "ordem": 12, "code": ""}];
var SEMENTE_VITRINE = {"marca": "3D Ostras", "assinatura": "Soluções tridimensionais", "slogan": "Onde ideias ganham forma.", "intro": "Peças personalizadas, troféus, brindes e soluções em NFC, fabricados sob demanda em Rio das Ostras, RJ. Escolha um modelo e peça o seu orçamento.", "instagram": "3dostras", "cidade": "Rio das Ostras, RJ", "links": [{"rotulo": "Fale no WhatsApp", "sub": "Tire dúvidas e acompanhe seu pedido", "tipo": "whatsapp", "valor": "Olá! Vim pelo Instagram da 3D Ostras.", "destaque": false}, {"rotulo": "Ver catálogo de produtos", "sub": "Troféus, brindes, NFC e peças personalizadas", "tipo": "catalogo", "valor": "", "destaque": false}, {"rotulo": "Pedir orçamento", "sub": "Conte sua ideia e receba o valor", "tipo": "whatsapp", "valor": "Olá! Quero pedir um orçamento na 3D Ostras. Minha ideia é: ", "destaque": true}], "whatsapp": "22999047056", "categorias": [{"id": "artes-marciais", "nome": "Artes marciais"}, {"id": "trofeus", "nome": "Troféus e medalhas"}, {"id": "profissoes", "nome": "Profissões"}, {"id": "nfc", "nome": "NFC e contato"}, {"id": "empresas", "nome": "Empresas e brindes"}]};
