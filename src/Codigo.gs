/**
 * Violão – Acompanhamento de exercícios (Google Apps Script)
 * A aluna abre o link (sem login), marca os exercícios e envia.
 * Tudo é gravado na aba "Registros" desta planilha, com o nome dela.
 */

// Configuração opcional. Deixe em branco: o sistema descobre sozinho.
// 1) ID da planilha (o trecho entre /d/ e /edit). Em branco = usa a planilha onde o script está.
const PLANILHA_ID = '';
// 2) URL do web app (termina em /exec). Em branco = usa a do menu Violão > Definir URL do web app
//    ou a URL detectada automaticamente.
const URL_APP = '';

const ABA = { alunos: 'Alunos', aulas: 'Aulas', ex: 'Exercicios', envios: 'Envios', reg: 'Registros' };
const FUSO = 'America/Sao_Paulo';
// Quantos aparelhos cada aluno pode usar (ex.: celular + computador)
const LIMITE_APARELHOS = 2;

function onOpen() {
  try { // guarda o ID desta planilha para o web app encontrar
    const a = SpreadsheetApp.getActiveSpreadsheet();
    if (a) PropertiesService.getScriptProperties().setProperty('PLANILHA_ID', a.getId());
  } catch (e) {}
  SpreadsheetApp.getUi()
    .createMenu('Violão')
    .addItem('Abrir painel do professor', 'abrirPainel')
    .addItem('Link do painel no celular', 'linkPainelCelular')
    .addSeparator()
    .addItem('Gerar links pendentes', 'gerarLinks')
    .addSeparator()
    .addItem('Processar faltas agora', 'processarFaltas')
    .addItem('Ativar automação (AppSheet)', 'ativarAutomacao')
    .addSeparator()
    .addItem('Definir URL do web app', 'definirUrlApp')
    .addItem('Definir endereço do app (GitHub Pages)', 'definirEnderecoApp')
    .addItem('Trocar chave do painel', 'trocarChavePainel')
    .addToUi();
}

function ss_() {
  const props = PropertiesService.getScriptProperties();
  const id = PLANILHA_ID || props.getProperty('PLANILHA_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ativa = SpreadsheetApp.getActiveSpreadsheet();
  if (ativa) { props.setProperty('PLANILHA_ID', ativa.getId()); return ativa; }
  throw new Error('Planilha não encontrada. Abra a planilha e rode o Configurar uma vez.');
}

/** URL pública do web app (termina em /exec). */
function urlApp_() {
  const salva = URL_APP || PropertiesService.getScriptProperties().getProperty('URL_APP') || '';
  if (/\/exec$/.test(salva)) return salva;
  try {
    const auto = ScriptApp.getService().getUrl() || '';
    if (/\/exec$/.test(auto)) return auto;
  } catch (e) {}
  return '';
}

/** Endereço que vai nos links enviados aos alunos: o app instalável (GitHub Pages) ou, se não houver, o web app. */
function urlPublica_() {
  const app = PropertiesService.getScriptProperties().getProperty('URL_PUBLICA') || '';
  return /^https:\/\//.test(app) ? app.replace(/\/?$/, '/') : urlApp_();
}

function definirEnderecoApp() {
  soDono_();
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Endereço do app',
    'Cole o endereço do GitHub Pages (ex.: https://seuusuario.github.io/violao-acompanhamento/).\nDeixe em branco para voltar a usar o link direto do Google.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const url = r.getResponseText().trim();
  const props = PropertiesService.getScriptProperties();
  if (!url) { props.deleteProperty('URL_PUBLICA'); ui.alert('Os links voltam a usar o endereço do Google.'); return; }
  if (!/^https:\/\/[^\s]+$/.test(url)) { ui.alert('Endereço inválido: ele precisa começar com https://'); return; }
  props.setProperty('URL_PUBLICA', url);
  ui.alert('Pronto! Os próximos links já vão abrir pelo app.');
}

function definirUrlApp() {
  soDono_();
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('URL do web app', 'Cole a URL que termina em /exec (Implantar > Gerenciar implantações):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const url = r.getResponseText().trim();
  if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) { ui.alert('URL inválida: ela precisa começar com https://script.google.com e terminar em /exec.'); return; }
  PropertiesService.getScriptProperties().setProperty('URL_APP', url);
  ui.alert('URL salva.');
}

function tabela_(nome) {
  const valores = ss_().getSheetByName(nome).getDataRange().getValues();
  const cab = valores.shift();
  return valores
    .filter(l => l[0] !== '')
    .map(l => Object.fromEntries(cab.map((c, i) => [c, l[i]])));
}

function data_(d) {
  return d ? Utilities.formatDate(new Date(d), FUSO, 'dd/MM/yyyy') : '';
}

/** Menu Violão > Gerar links pendentes: cria token, link e mensagem de WhatsApp
 *  para cada linha da aba Envios que ainda não tem token. */
function gerarLinks() {
  soDono_();
  const URL = urlApp_();
  if (!URL) {
    SpreadsheetApp.getUi().alert('Implante o web app e salve a URL em Violão > Definir URL do web app.');
    return;
  }
  const alunos = tabela_(ABA.alunos);
  const aulas = tabela_(ABA.aulas);
  const sh = ss_().getSheetByName(ABA.envios);
  const v = sh.getDataRange().getValues();
  const cab = v[0];
  const c = n => cab.indexOf(n);
  let criados = 0;
  for (let i = 1; i < v.length; i++) {
    const l = v[i];
    if (!l[c('ID_Aluno')] || !l[c('ID_Aula')] || l[c('Token')]) continue;
    if (!l[c('Data_Envio')]) l[c('Data_Envio')] = new Date();
    const token = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    const link = urlPublica_() + '?t=' + token;
    const aluno = alunos.find(a => a.ID_Aluno === l[c('ID_Aluno')]) || {};
    const aula = aulas.find(a => a.ID_Aula === l[c('ID_Aula')]) || {};
    const primeiroNome = String(aluno.Nome || '').split(' ')[0];
    const msg = 'Olá, ' + primeiroNome + '! Seus exercícios da Aula ' + (aula.Numero || '') +
      ' (' + (aula.Titulo || '') + '). Marque cada dia que praticar. Link válido até ' +
      data_(l[c('Validade')]) + ': ' + link;
    const fone = String(aluno.Telefone || '').replace(/\D/g, '');
    l[c('Token')] = token;
    l[c('Link')] = link;
    l[c('WhatsApp')] = fone ? 'https://wa.me/55' + fone + '?text=' + encodeURIComponent(msg) : '';
    criados++;
  }
  sh.getRange(1, 1, v.length, cab.length).setValues(v);
  SpreadsheetApp.getUi().alert(criados + ' link(s) gerado(s). Clique na coluna WhatsApp para enviar.');
}

/** Monta tudo o que a página precisa a partir do token. */
function contexto_(token) {
  if (!token) return { erro: 'Link inválido. Peça um novo link ao professor.' };
  const env = tabela_(ABA.envios).find(r => String(r.Token) === token);
  if (!env) return { erro: 'Link inválido. Peça um novo link ao professor.' };

  const validade = new Date(env.Validade);
  validade.setHours(23, 59, 59);
  if (isNaN(validade) || new Date() > validade) {
    return { erro: 'Este link venceu em ' + data_(env.Validade) + '. Na próxima aula você recebe o link novo.' };
  }

  const aluno = tabela_(ABA.alunos).find(r => r.ID_Aluno === env.ID_Aluno) || {};
  const aula = tabela_(ABA.aulas).find(r => r.ID_Aula === env.ID_Aula) || {};
  const regs = tabela_(ABA.reg).filter(r => String(r.Token) === token);

  const vezes = {};
  regs.forEach(r => String(r.IDs_Exercicios).split(',').map(s => s.trim()).filter(Boolean)
    .forEach(id => { vezes[id] = (vezes[id] || 0) + 1; }));
  const dias = new Set(regs.map(r => data_(r.Data_Hora))).size;
  const minutos = regs.reduce((s, r) => s + (Number(r.Minutos) || 0), 0);

  const exercicios = tabela_(ABA.ex)
    .filter(r => r.ID_Aula === env.ID_Aula)
    .sort((a, b) => a.Ordem - b.Ordem)
    .map(r => ({ id: r.ID_Exercicio, titulo: r.Titulo, descricao: r.Descricao, tempo: r.Tempo_min, vezes: vezes[r.ID_Exercicio] || 0 }));

  return {
    token: token,
    idAluno: env.ID_Aluno,
    idAula: env.ID_Aula,
    nome: String(aluno.Nome || ''),
    aula: { numero: aula.Numero, titulo: aula.Titulo, tarefa: aula.Tarefa, apostila: aula.Link_Apostila },
    validade: data_(env.Validade),
    exercicios: exercicios,
    dias: dias,
    minutos: minutos,
    hoje: regs.some(r => data_(r.Data_Hora) === data_(new Date()))
  };
}

/* ---------- Acesso: cadastro no 1º link, depois só o nome ---------- */

function alunoCadastrado_(aluno) {
  return String(aluno.Cadastro_Completo || '').toLowerCase() === 'sim';
}

/** O que a página recebe ANTES de a aluna se identificar: nada de dados pessoais. */
function publico_(token) {
  const ctx = contexto_(token);
  if (ctx.erro) return { erro: ctx.erro };
  const aluno = tabela_(ABA.alunos).find(r => r.ID_Aluno === ctx.idAluno) || {};
  return {
    etapa: alunoCadastrado_(aluno) ? 'nome' : 'cadastro',
    token: token,
    aula: { numero: ctx.aula.numero, titulo: ctx.aula.titulo },
    validade: ctx.validade
  };
}

function doGet(e) {
  if (e && e.parameter && e.parameter.p === 'painel') return paginaPainel_(String(e.parameter.k || ''));
  const ctx = publico_(String((e && e.parameter && e.parameter.t) || '').trim());
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.dados = JSON.stringify(ctx).replace(/</g, '\\u003c');
  return tpl.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // permite abrir dentro do app (GitHub Pages)
    .setTitle('Violão – Meus exercícios')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function normaliza_(t) {
  return String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}

/** Aceita o nome completo ou o primeiro nome, sem diferenciar maiúsculas e acentos. */
function nomeConfere_(digitado, cadastrado) {
  const d = normaliza_(digitado), c = normaliza_(cadastrado);
  return d.length >= 2 && (d === c || d === c.split(' ')[0]);
}

/** Limita tentativas de nome por link (10 a cada 15 minutos). */
function tentativa_(token) {
  const cache = CacheService.getScriptCache();
  const k = 'tent_' + token;
  const n = Number(cache.get(k) || 0) + 1;
  cache.put(k, String(n), 900);
  if (n > 10) throw new Error('Muitas tentativas. Espere alguns minutos e tente de novo.');
}

/** Trava por aparelho: o link só abre nos aparelhos registrados do aluno.
 *  Para liberar um aparelho novo, apague a célula Aparelhos do aluno na aba Alunos. */
function verificarAparelho_(idAluno, aparelho, podeCadastrar) {
  aparelho = String(aparelho || '').replace(/[^a-z0-9]/gi, '').slice(0, 40);
  if (aparelho.length < 12) {
    throw new Error('Não foi possível reconhecer este aparelho. Abra o link no navegador do celular (Chrome ou Safari), fora da janela anônima.');
  }
  const sh = ss_().getSheetByName(ABA.alunos);
  const v = sh.getDataRange().getValues();
  const cab = v[0];
  const cA = cab.indexOf('Aparelhos');
  const linha = v.findIndex((l, i) => i > 0 && l[cab.indexOf('ID_Aluno')] === idAluno);
  if (cA < 0 || linha < 0) return;
  const lista = String(v[linha][cA] || '').split(',').map(x => x.trim()).filter(Boolean);
  if (lista.indexOf(aparelho) >= 0) return;
  if (podeCadastrar && lista.length < LIMITE_APARELHOS) {
    lista.push(aparelho);
    sh.getRange(linha + 1, cA + 1).setValue(lista.join(', '));
    return;
  }
  throw new Error('Este link só abre nos aparelhos cadastrados da aluna. Se você trocou de celular ou de navegador, peça ao professor para liberar.');
}

/** 1º acesso: grava o cadastro na aba Alunos e libera os exercícios. */
function cadastrar(d) {
  d = d || {};
  const token = String(d.token || '');
  const ctx = contexto_(token);
  if (ctx.erro) throw new Error(ctx.erro);
  const sh = ss_().getSheetByName(ABA.alunos);
  const v = sh.getDataRange().getValues();
  const cab = v[0];
  const linha = v.findIndex((l, i) => i > 0 && l[cab.indexOf('ID_Aluno')] === ctx.idAluno);
  if (linha < 0) throw new Error('Aluno não encontrado. Fale com o professor.');
  if (String(v[linha][cab.indexOf('Cadastro_Completo')]).toLowerCase() === 'sim') {
    throw new Error('Seu cadastro já foi feito. Atualize a página e entre com o seu nome.');
  }

  const limpa = (t, max) => { let x = String(t || '').trim().slice(0, max); if (/^[=+\-@]/.test(x)) x = "'" + x; return x; };
  const nome = limpa(d.nome, 80);
  const fone = String(d.telefone || '').replace(/\D/g, '');
  if (normaliza_(nome).length < 3) throw new Error('Digite seu nome completo.');
  if (fone.length < 10 || fone.length > 11) throw new Error('Digite o WhatsApp com DDD (ex.: 94 99123-4567).');
  verificarAparelho_(ctx.idAluno, d.aparelho, true);

  const campos = {
    Nome: nome, Telefone: fone, Email: limpa(d.email, 80),
    Tem_Violao: limpa(d.temViolao, 10), Tipo_Violao: limpa(d.tipoViolao, 20),
    Experiencia: limpa(d.experiencia, 300), Estilos: limpa(d.estilos, 300), Objetivo: limpa(d.objetivo, 300),
    Data_Cadastro: new Date(), Cadastro_Completo: 'Sim'
  };
  Object.keys(campos).forEach(k => {
    const col = cab.indexOf(k);
    if (col >= 0) sh.getRange(linha + 1, col + 1).setValue(campos[k]);
  });
  SpreadsheetApp.flush();
  const novo = contexto_(token);
  novo.nomeDigitado = nome;
  return novo;
}

/** Links seguintes: entra só com o nome. */
function entrar(token, nome, aparelho) {
  token = String(token || '');
  const ctx = contexto_(token);
  if (ctx.erro) throw new Error(ctx.erro);
  tentativa_(token);
  if (!nomeConfere_(nome, ctx.nome)) throw new Error('Esse nome não confere com o cadastro deste link. Confira e tente de novo.');
  verificarAparelho_(ctx.idAluno, aparelho, true);
  ctx.nomeDigitado = String(nome);
  return ctx;
}

/** Chamado pela página quando a aluna envia a prática do dia. */
function registrar(d) {
  d = d || {};
  const ctx = contexto_(String(d.token || ''));
  if (ctx.erro) throw new Error(ctx.erro);
  if (!nomeConfere_(d.nome, ctx.nome)) throw new Error('Entre de novo com o seu nome.');
  verificarAparelho_(ctx.idAluno, d.aparelho, false);
  const tipoAparelho = /Mobi|Android|iPhone|iPad/i.test(String(d.navegador || '')) ? 'Celular' : 'Computador';

  const validos = new Map(ctx.exercicios.map(x => [x.id, x.titulo]));
  const marcados = (d.exercicios || []).map(String).filter(id => validos.has(id));
  if (!marcados.length) throw new Error('Marque pelo menos um exercício.');

  const minutos = Math.max(0, Math.min(300, Math.round(Number(d.minutos) || 0)));
  const dificuldade = ['Fácil', 'Médio', 'Difícil'].indexOf(d.dificuldade) >= 0 ? d.dificuldade : '';
  let comentario = String(d.comentario || '').slice(0, 500);
  if (/^[=+\-@]/.test(comentario)) comentario = "'" + comentario; // evita fórmula na planilha

  const linha = [new Date(), ctx.token, ctx.idAluno, ctx.nome, ctx.idAula, marcados.join(', '),
    marcados.map(id => validos.get(id)).join(', '), minutos, dificuldade, comentario,
    tipoAparelho + ' ' + String(d.aparelho || '').slice(0, 6)];

  const trava = LockService.getScriptLock();
  trava.waitLock(10000);
  try {
    const sh = ss_().getSheetByName(ABA.reg);
    sh.getRange(sh.getLastRow() + 1, 1, 1, linha.length).setValues([linha]);
  } finally {
    trava.releaseLock();
  }
  const novo = contexto_(ctx.token);
  novo.nomeDigitado = String(d.nome);
  return novo;
}

/* =====================================================================
 * FREQUÊNCIA: reposição automática pela sua DISPONIBILIDADE
 * Ao marcar "Falta" ou "Falta justificada" numa aula da aba Frequencia:
 *  - procura o primeiro horário livre seu (aba Disponibilidade) depois da falta,
 *    que não coincida com outra aula já agendada, e agenda a reposição;
 *  - estende a validade do link (aba Envios) até a reposição;
 *  - se ela faltar também na reposição, o link passa a valer até a próxima aula agendada.
 * Funciona sozinho ao editar a planilha (onEdit). Para edições feitas pelo AppSheet,
 * use o menu Violão > Ativar automação (verifica a cada hora).
 * ===================================================================== */
const DIAS_BUSCA = 14;          // até quantos dias à frente procurar horário
const DURACAO_AULA_MIN = 60;    // cada aula ocupa 1 hora
const DIAS_SEMANA = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== 'Frequencia') return;
  const colPresenca = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf('Presenca') + 1;
  if (e.range.getColumn() > colPresenca || e.range.getLastColumn() < colPresenca) return;
  const msgs = processarFaltas_(e.source);
  if (msgs.length) e.source.toast(msgs.join('\n'), 'Reposição', 10);
}

/** Menu e gatilho por hora. */
function processarFaltas() {
  const ss = ss_();
  const msgs = processarFaltas_(ss);
  try { if (msgs.length) ss.toast(msgs.join('\n'), 'Reposição', 10); } catch (err) {}
}

function ativarAutomacao() {
  soDono_();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processarFaltas')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processarFaltas').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('Automação ativada: as faltas são verificadas a cada hora (inclusive as marcadas pelo AppSheet).');
}

// "18:00" ou um horário da planilha -> minutos desde 00:00
function minutos_(h) {
  if (h instanceof Date) return h.getHours() * 60 + h.getMinutes();
  const m = String(h).trim().match(/^(\d{1,2})[:h](\d{2})?/);
  return m ? Number(m[1]) * 60 + Number(m[2] || 0) : null;
}
function hhmm_(min) {
  return String(Math.floor(min / 60)).padStart(2, '0') + ':' + String(min % 60).padStart(2, '0');
}
function semAcento_(t) { return String(t).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim(); }

/** Primeiro horário livre depois de "depoisDe", ou null. Retorna {data, horario}. */
function acharHorario_(ss, depoisDe, ocupados) {
  const disp = ss.getSheetByName('Disponibilidade').getDataRange().getValues();
  const cab = disp[0];
  const c = n => cab.indexOf(n);
  const dia = d => Utilities.formatDate(new Date(d), FUSO, 'yyyy-MM-dd');

  // horários semanais (colunas A–E) e exceções por data (colunas G–J)
  const semanais = disp.slice(1).filter(l => l[c('Dia_Semana')] !== '' && semAcento_(l[c('Ativo')]) !== 'nao')
    .filter(l => ['reposicao', 'ambos'].indexOf(semAcento_(l[c('Uso')])) >= 0)
    .map(l => ({ dow: DIAS_SEMANA.map(semAcento_).indexOf(semAcento_(l[c('Dia_Semana')])),
                 ini: minutos_(l[c('Hora_Inicio')]), fim: minutos_(l[c('Hora_Fim')]) }));
  const excecoes = disp.slice(1).filter(l => l[c('Data')] instanceof Date)
    .map(l => ({ dia: dia(l[c('Data')]), tipo: semAcento_(l[c('Situacao')]),
                 ini: minutos_(l[c('Inicio_Extra')]), fim: minutos_(l[c('Fim_Extra')]) }));

  const d = new Date(depoisDe);
  for (let n = 1; n <= DIAS_BUSCA; n++) {
    const cand = new Date(d); cand.setDate(d.getDate() + n);
    const chave = dia(cand);
    const exDia = excecoes.filter(x => x.dia === chave);
    if (exDia.some(x => x.tipo === 'indisponivel')) continue;
    const blocos = exDia.filter(x => x.tipo === 'disponivel extra' && x.ini != null && x.fim != null)
      .concat(semanais.filter(s => s.dow === cand.getDay() && s.ini != null && s.fim != null));
    blocos.sort((a, b) => a.ini - b.ini);
    for (const b of blocos) {
      for (let t = b.ini; t + DURACAO_AULA_MIN <= b.fim; t += DURACAO_AULA_MIN) {
        const horario = hhmm_(t) + '–' + hhmm_(t + DURACAO_AULA_MIN);
        if (!ocupados.has(chave + ' ' + hhmm_(t))) return { data: cand, horario: horario, chave: chave + ' ' + hhmm_(t) };
      }
    }
  }
  return null;
}

function processarFaltas_(ss) {
  const fr = ss.getSheetByName('Frequencia');
  const env = ss.getSheetByName('Envios');
  const v = fr.getDataRange().getValues();
  const cab = v[0];
  const c = n => cab.indexOf(n);
  const [cData, cAluno, cAula, cPres, cRep, cObs, cHora, cTipo, cNome] =
    ['Data', 'ID_Aluno', 'ID_Aula', 'Presenca', 'Reposicao_Data', 'Obs', 'Horario', 'Tipo', 'Nome'].map(c);
  const dia = d => Utilities.formatDate(new Date(d), FUSO, 'yyyy-MM-dd');
  const br = d => Utilities.formatDate(new Date(d), FUSO, 'dd/MM');
  const msgs = [];

  // horários já ocupados por aulas agendadas (para não marcar duas no mesmo horário)
  const ocupados = new Set(v.slice(1)
    .filter(x => x[cData] instanceof Date && x[cPres] === 'Agendada')
    .map(x => dia(x[cData]) + ' ' + hhmm_(minutos_(String(x[cHora]).split(/[–-]/)[0]) || 0)));

  for (let i = 1; i < v.length; i++) {
    const l = v[i];
    const pres = l[cPres];
    if ((pres !== 'Falta' && pres !== 'Falta justificada') || l[cRep] !== '' || !(l[cData] instanceof Date)) continue;
    const aluno = l[cAluno], aula = l[cAula], data = l[cData];
    const nome = l[cNome] || aluno;
    const ehReposicao = l[cTipo] === 'Reposição';
    const proximaAula = () => v.slice(1)
      .filter(x => x[cAluno] === aluno && x[cPres] === 'Agendada' && x[cData] instanceof Date && x[cData] > data)
      .map(x => x[cData]).sort((a, b) => a - b)[0] || null;
    let novaValidade = null;

    const slot = ehReposicao ? null : acharHorario_(ss, data, ocupados);
    if (slot) {
      ocupados.add(slot.chave);
      novaValidade = slot.data;
      fr.getRange(i + 1, cRep + 1).setValue(slot.data).setNumberFormat('dd/MM/yyyy');
      let alvo = v.findIndex((x, k) => k > 0 && x[cData] === '' && x[cAluno] === '');
      if (alvo === -1) { alvo = v.length; v.push(new Array(cab.length).fill('')); }
      const r = alvo + 1;
      fr.getRange(r, cData + 1).setValue(slot.data).setNumberFormat('dd/MM/yyyy');
      fr.getRange(r, cAluno + 1).setValue(aluno);
      fr.getRange(r, cAula + 1).setValue(aula);
      fr.getRange(r, cPres + 1).setValue('Agendada');
      fr.getRange(r, cObs + 1).setValue('Reposição da aula de ' + br(data));
      fr.getRange(r, cHora + 1).setValue(slot.horario);
      fr.getRange(r, cTipo + 1).setValue('Reposição');
      fr.getRange(r, cNome + 1).setFormula('=IF(B' + r + '="","",IFERROR(INDEX(Alunos!B:B,MATCH(B' + r + ',Alunos!A:A,0)),""))');
      fr.getRange(r, c('Aula') + 1).setFormula('=IF(D' + r + '="","",IFERROR("Aula "&INDEX(Aulas!B:B,MATCH(D' + r + ',Aulas!A:A,0))&" – "&INDEX(Aulas!C:C,MATCH(D' + r + ',Aulas!A:A,0)),""))');
      v[alvo][cData] = slot.data; v[alvo][cAluno] = aluno;
      msgs.push(nome + ': reposição agendada para ' + DIAS_SEMANA[slot.data.getDay()].toLowerCase() + ' ' +
        br(slot.data) + ', ' + slot.horario + '.');
    } else {
      novaValidade = proximaAula();
      fr.getRange(i + 1, cRep + 1).setValue(ehReposicao ? 'Sem nova reposição' : 'A definir');
      msgs.push(nome + (ehReposicao ? ': faltou também na reposição. ' :
        ': sem horário livre nos próximos ' + DIAS_BUSCA + ' dias, combine a reposição. ') +
        (novaValidade ? 'O link vale até a próxima aula (' + br(novaValidade) + ').' : 'Não há próxima aula agendada.'));
    }

    // estende a validade do link que vencia na data da falta
    if (novaValidade) {
      const ev = env.getDataRange().getValues();
      const eA = ev[0].indexOf('ID_Aluno'), eV = ev[0].indexOf('Validade');
      for (let k = 1; k < ev.length; k++) {
        if (ev[k][eA] === aluno && ev[k][eV] instanceof Date && dia(ev[k][eV]) === dia(data)) {
          env.getRange(k + 1, eV + 1).setValue(novaValidade).setNumberFormat('dd/MM/yyyy');
        }
      }
    }
  }
  return msgs;
}
