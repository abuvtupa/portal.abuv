/**
 * API de cadastros da ABUV (Google Apps Script + Planilha Google).
 *
 * Como publicar (feito uma vez, logado na conta Google que vai guardar os dados):
 *  1. Abra https://script.google.com > Novo projeto. Apague o conteúdo e cole este arquivo inteiro.
 *  2. Selecione a função "configurar" e clique em Executar. Autorize o acesso quando o Google pedir.
 *     Ela cria a planilha "ABUV - Cadastros (banco do site)" no seu Drive e grava os cadastros de exemplo.
 *  3. Implantar > Nova implantação > tipo "App da Web".
 *     Executar como: Eu  |  Quem pode acessar: Qualquer pessoa  > Implantar.
 *  4. Copie a URL que termina em /exec e coloque na constante API do index.html.
 *
 * Para mudar a senha de exclusão: Configurações do projeto > Propriedades do script > SENHA_ADMIN.
 * Ao alterar este código, publique em Implantar > Gerenciar implantações > Editar > Nova versão
 * (assim a URL /exec continua a mesma).
 */

const ABA = 'Cadastros';
const CAMPOS = ['id', 'nomeResponsavel', 'email', 'celular', 'logradouro', 'numero', 'bairro',
  'complemento', 'cidade', 'nomeAluno', 'idade', 'modalidade', 'observacoes', 'criadoEm', 'atualizadoEm'];
const EDITAVEIS = CAMPOS.filter(c => ['id', 'criadoEm', 'atualizadoEm'].indexOf(c) === -1);

const EXEMPLOS = [
  { nomeResponsavel: 'Maria Aparecida Souza', email: 'maria.souza@exemplo.com', celular: '(14) 99611-2034', logradouro: 'Rua das Acácias', numero: '215', bairro: 'Jardim Unesp', complemento: '', cidade: 'Tupã - SP', nomeAluno: 'Gabriel Souza', idade: 12, modalidade: 'Jiu Jitsu', observacoes: 'Treina às segundas e quartas.' },
  { nomeResponsavel: 'José Carlos Pereira', email: 'jose.pereira@exemplo.com', celular: '(14) 99702-4418', logradouro: 'Rua Ipê Amarelo', numero: '87', bairro: 'Jardim Carajás II', complemento: 'Casa 2', cidade: 'Tupã - SP', nomeAluno: 'Ana Clara Pereira', idade: 10, modalidade: 'Ginástica', observacoes: '' },
  { nomeResponsavel: 'Fernanda Lima', email: 'fernanda.lima@exemplo.com', celular: '(14) 99633-7120', logradouro: 'Avenida Tamoios', numero: '1302', bairro: 'Jardim Itaipu', complemento: '', cidade: 'Tupã - SP', nomeAluno: 'Lucas Lima', idade: 15, modalidade: 'Muay Thai', observacoes: 'Possui luvas próprias.' },
  { nomeResponsavel: 'Roberto Alves', email: 'roberto.alves@exemplo.com', celular: '(14) 99845-3301', logradouro: 'Rua Pioneiro Gallo Soares', numero: '98', bairro: 'José Feliciano', complemento: '', cidade: 'Tupã - SP', nomeAluno: 'Pedro Henrique Alves', idade: 14, modalidade: 'Krav Magá', observacoes: '' },
  { nomeResponsavel: 'Cláudia Ribeiro', email: 'claudia.ribeiro@exemplo.com', celular: '(14) 99780-5562', logradouro: 'Rua Guaranis', numero: '44', bairro: 'Jardim Aritana', complemento: 'Fundos', cidade: 'Tupã - SP', nomeAluno: 'Beatriz Ribeiro', idade: 9, modalidade: 'Ginástica', observacoes: 'Alergia a amendoim.' },
  { nomeResponsavel: 'Antônio Marcos Silva', email: 'antonio.silva@exemplo.com', celular: '(14) 99654-8890', logradouro: 'Rua Caingangues', numero: '310', bairro: 'Conj. Hab. José Maria Gonçalves Gameiro', complemento: '', cidade: 'Tupã - SP', nomeAluno: 'Matheus Silva', idade: 16, modalidade: 'Jiu Jitsu', observacoes: '' },
  { nomeResponsavel: 'Juliana Martins', email: 'juliana.martins@exemplo.com', celular: '(14) 99721-6643', logradouro: 'Rua Tupinambás', numero: '152', bairro: 'Jardim Unesp', complemento: '', cidade: 'Tupã - SP', nomeAluno: 'Juliana Martins', idade: 29, modalidade: 'Voluntário', observacoes: 'Professora voluntária de ginástica aos sábados.' },
  { nomeResponsavel: 'Sebastião Oliveira', email: 'sebastiao.oliveira@exemplo.com', celular: '(14) 99812-0457', logradouro: 'Rua Xavantes', numero: '601', bairro: 'Jardim Carajás II', complemento: '', cidade: 'Tupã - SP', nomeAluno: 'Sebastião Oliveira', idade: 41, modalidade: 'Voluntário', observacoes: 'Apoio na organização dos eventos comunitários.' }
];

/* ---------- configuração ---------- */

function configurar() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SENHA_ADMIN')) props.setProperty('SENHA_ADMIN', '1234');
  const sh = aba_();
  if (sh.getLastRow() <= 1) popularExemplos();
  Logger.log('Planilha: ' + sh.getParent().getUrl());
}

function popularExemplos() {
  EXEMPLOS.forEach(r => inserir_(r));
}

function aba_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty('PLANILHA_ID');
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = SpreadsheetApp.create('ABUV - Cadastros (banco do site)');
  props.setProperty('PLANILHA_ID', ss.getId());
  let sh = ss.getSheetByName(ABA);
  if (!sh) sh = ss.insertSheet(ABA);
  if (sh.getLastRow() === 0) {
    sh.appendRow(CAMPOS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CAMPOS.length).setFontWeight('bold');
  }
  return sh;
}

/* ---------- web app ---------- */

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  try {
    if (action === 'list') return json_({ ok: true, data: listar_() });
    return json_({ ok: false, error: 'Ação inválida: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { body = {}; }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (action === 'insert') return json_({ ok: true, data: inserir_(body) });
    if (action === 'update') return json_({ ok: true, data: atualizar_(body) });
    if (action === 'delete') { checarSenha_(body); excluir_(body.id); return json_({ ok: true, data: { id: Number(body.id) } }); }
    if (action === 'delete_all') { checarSenha_(body); excluirTodos_(); return json_({ ok: true, data: [] }); }
    return json_({ ok: false, error: 'Ação inválida: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- operações ---------- */

function listar_() {
  const sh = aba_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  const linhas = sh.getRange(2, 1, n, CAMPOS.length).getValues();
  return linhas.filter(l => l[0] !== '').map(registro_);
}

function inserir_(dados) {
  const r = limpar_(dados);
  validar_(r);
  const sh = aba_();
  const agora = new Date().toISOString();
  r.id = proximoId_(sh);
  r.criadoEm = agora;
  r.atualizadoEm = agora;
  sh.appendRow(CAMPOS.map(c => celula_(r[c])));
  return registro_(CAMPOS.map(c => r[c]));
}

function atualizar_(dados) {
  const sh = aba_();
  const linha = linhaDoId_(sh, dados.id);
  if (!linha) throw new Error('Registro não encontrado.');
  const atual = registro_(sh.getRange(linha, 1, 1, CAMPOS.length).getValues()[0]);
  const r = Object.assign(atual, limpar_(dados));
  validar_(r);
  r.atualizadoEm = new Date().toISOString();
  sh.getRange(linha, 1, 1, CAMPOS.length).setValues([CAMPOS.map(c => celula_(r[c]))]);
  return registro_(CAMPOS.map(c => r[c]));
}

function excluir_(id) {
  const sh = aba_();
  const linha = linhaDoId_(sh, id);
  if (!linha) throw new Error('Registro não encontrado.');
  sh.deleteRow(linha);
}

function excluirTodos_() {
  const sh = aba_();
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
}

/* ---------- apoio ---------- */

function checarSenha_(body) {
  const senha = PropertiesService.getScriptProperties().getProperty('SENHA_ADMIN') || '1234';
  if (String(body.senha || '') !== senha) throw new Error('Senha incorreta.');
}

function limpar_(dados) {
  const r = {};
  EDITAVEIS.forEach(c => {
    if (dados && dados[c] !== undefined && dados[c] !== null) r[c] = String(dados[c]).trim().slice(0, 500);
  });
  if (r.idade !== undefined) r.idade = r.idade === '' ? '' : Number(r.idade);
  return r;
}

function validar_(r) {
  ['nomeResponsavel', 'email', 'celular', 'cidade', 'nomeAluno', 'modalidade'].forEach(c => {
    if (!r[c]) throw new Error('Campo obrigatório: ' + c);
  });
  if (!(Number(r.idade) > 0)) throw new Error('Idade inválida.');
}

function proximoId_(sh) {
  const n = sh.getLastRow() - 1;
  if (n <= 0) return 1;
  const ids = sh.getRange(2, 1, n, 1).getValues().map(l => Number(l[0]) || 0);
  return Math.max.apply(null, ids) + 1;
}

function linhaDoId_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n <= 0) return 0;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (Number(ids[i][0]) === Number(id)) return i + 2;
  return 0;
}

// Evita que um texto começando com = + - @ vire fórmula na planilha.
function celula_(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v === undefined ? '' : v;
}

function registro_(linha) {
  const r = {};
  CAMPOS.forEach((c, i) => { r[c] = linha[i] instanceof Date ? linha[i].toISOString() : linha[i]; });
  r.id = Number(r.id);
  return r;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
