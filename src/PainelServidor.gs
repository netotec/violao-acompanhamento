/**
 * PainelServidor.gs – Painel do professor (servidor)
 * Abra pelo menu Violão > Abrir painel do professor (na planilha)
 * ou pelo celular com o link secreto (Violão > Link do painel no celular).
 * Todas as funções do painel exigem a chave secreta, guardada nas propriedades do script.
 */

/* ---------- Segurança ---------- */

function chavePainel_() {
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('PAINEL_CHAVE');
  if (!k) {
    k = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    p.setProperty('PAINEL_CHAVE', k);
  }
  return k;
}

function checarChave_(k) {
  const certa = PropertiesService.getScriptProperties().getProperty('PAINEL_CHAVE');
  if (!k || !certa || String(k) !== certa) throw new Error('Acesso negado.');
}

/** Bloqueia funções administrativas quando chamadas pela página pública. */
function soDono_() {
  const ativo = Session.getActiveUser().getEmail();
  const dono = Session.getEffectiveUser().getEmail();
  if (!ativo || ativo !== dono) throw new Error('Acesso negado.');
}

/* ---------- Abrir o painel ---------- */

function abrirPainel() {
  soDono_();
  const t = HtmlService.createTemplateFromFile('Painel');
  t.chave = chavePainel_();
  t.modo = 'dialogo';
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(1180).setHeight(780), 'Painel do professor');
}

function linkPainelCelular() {
  soDono_();
  if (!urlApp_()) {
    SpreadsheetApp.getUi().alert('Salve a URL do web app em Violão > Definir URL do web app primeiro.');
    return;
  }
  const url = urlApp_() + '?p=painel&k=' + chavePainel_();
  const html = HtmlService.createHtmlOutput(
    '<div style="font:14px Arial;padding:8px">' +
    '<p><b>Link do seu painel no celular.</b> Guarde só para você: quem tiver este link consegue abrir o painel.</p>' +
    '<textarea style="width:100%;height:90px" onclick="this.select()">' + url + '</textarea>' +
    '<p>Se o link vazar, use <b>Violão &gt; Trocar chave do painel</b> e o link antigo para de funcionar.</p></div>'
  ).setWidth(520).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, 'Painel no celular');
}

function trocarChavePainel() {
  soDono_();
  PropertiesService.getScriptProperties().deleteProperty('PAINEL_CHAVE');
  chavePainel_();
  SpreadsheetApp.getUi().alert('Chave trocada. Os links antigos do painel pararam de funcionar. Gere o novo em Violão > Link do painel no celular.');
}

/** Chamado pelo doGet quando a URL tem ?p=painel&k=... */
function paginaPainel_(k) {
  try { checarChave_(k); } catch (e) {
    return HtmlService.createHtmlOutput('<p style="font:16px Arial;padding:24px">Acesso negado.</p>').setTitle('Painel');
  }
  const t = HtmlService.createTemplateFromFile('Painel');
  t.chave = k;
  t.modo = 'web';
  return t.evaluate().setTitle('Painel do professor – Violão').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- Ajudantes ---------- */

function iso_(d) {
  return d instanceof Date && !isNaN(d) ? Utilities.formatDate(d, FUSO, 'yyyy-MM-dd') : '';
}
function deIso_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('Data inválida.');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}
function hoje_() { const d = new Date(); d.setHours(0, 0, 0, 0); return d; }

/** Lê uma aba com o número de cada linha. */
function linhas_(nome) {
  const sh = ss_().getSheetByName(nome);
  const v = sh.getDataRange().getValues();
  const cab = v[0];
  const rows = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === '' || v[i][0] === null) continue;
    const o = { _r: i + 1 };
    cab.forEach((c, j) => { if (c) o[c] = v[i][j]; });
    rows.push(o);
  }
  return { sh: sh, cab: cab, rows: rows, total: v.length };
}

function primeiraLinhaVazia_(sh) {
  const col = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (let i = 1; i < col.length; i++) if (col[i][0] === '') return i + 1;
  return col.length + 1;
}

function gravar_(sh, cab, r, campos) {
  Object.keys(campos).forEach(k => {
    const c = cab.indexOf(k);
    if (c >= 0) sh.getRange(r, c + 1).setValue(campos[k]);
  });
}

function limpa_(t, max) {
  let x = String(t == null ? '' : t).trim().slice(0, max || 200);
  if (/^[=+\-@]/.test(x)) x = "'" + x;
  return x;
}

/* ---------- Dados do painel ---------- */

function painelDados(k) {
  checarChave_(k);
  const H = hoje_();
  const alunos = linhas_(ABA.alunos).rows;
  const aulas = linhas_(ABA.aulas).rows;
  const envios = linhas_(ABA.envios).rows;
  const freq = linhas_('Frequencia').rows;
  const regs = linhas_(ABA.reg).rows;
  const dispSh = ss_().getSheetByName('Disponibilidade').getRange(2, 1, 30, 10).getValues();
  const titulo = id => { const a = aulas.find(x => x.ID_Aula === id); return a ? 'Aula ' + a.Numero + ' – ' + a.Titulo : id; };
  const txt = v => (v instanceof Date ? Utilities.formatDate(v, FUSO, 'HH:mm') : String(v == null ? '' : v));

  const resumo = alunos.map(a => {
    const id = a.ID_Aluno;
    const meus = regs.filter(r => r.ID_Aluno === id);
    const ult7 = meus.filter(r => r.Data_Hora instanceof Date && (new Date() - r.Data_Hora) <= 7 * 864e5);
    const ultima = meus.map(r => r.Data_Hora).filter(d => d instanceof Date).sort((x, y) => y - x)[0];
    const minhas = freq.filter(f => f.ID_Aluno === id);
    const realizadas = minhas.filter(f => f.Tipo !== 'Reposição' && f.Presenca && f.Presenca !== 'Agendada').length;
    const presencas = minhas.filter(f => f.Presenca === 'Presente' || f.Presenca === 'Reposição').length;
    const prox = minhas.filter(f => f.Presenca === 'Agendada' && f.Data instanceof Date && f.Data >= H).sort((x, y) => x.Data - y.Data)[0];
    const env = envios.filter(e => e.ID_Aluno === id).sort((x, y) => (y.Data_Envio || 0) - (x.Data_Envio || 0))[0];
    return {
      id: id,
      praticas7: new Set(ult7.map(r => iso_(r.Data_Hora))).size,
      minutos7: ult7.reduce((s, r) => s + (Number(r.Minutos) || 0), 0),
      minutosTotal: meus.reduce((s, r) => s + (Number(r.Minutos) || 0), 0),
      ultima: ultima ? Utilities.formatDate(ultima, FUSO, 'dd/MM HH:mm') : '',
      diasSemPraticar: ultima ? Math.floor((new Date() - ultima) / 864e5) : null,
      frequencia: realizadas ? Math.min(1, presencas / realizadas) : null,
      realizadas: realizadas, presencas: presencas,
      proxima: prox ? { data: iso_(prox.Data), horario: txt(prox.Horario), aula: titulo(prox.ID_Aula) } : null,
      envio: env ? { aula: titulo(env.ID_Aula), validade: iso_(env.Validade), ativo: env.Validade instanceof Date && env.Validade >= H, link: env.Link || '', whatsapp: env.WhatsApp || '' } : null
    };
  });

  return {
    hoje: iso_(H),
    alunos: alunos.map(a => ({
      id: a.ID_Aluno, nome: String(a.Nome || ''), telefone: String(a.Telefone || ''), plano: String(a.Plano || ''),
      valor: a.Valor_Mensal === '' ? '' : Number(a.Valor_Mensal) || 0, status: String(a.Status || ''),
      cadastro: String(a.Cadastro_Completo || '').toLowerCase() === 'sim',
      aparelhos: String(a.Aparelhos || '').split(',').filter(x => x.trim()).length,
      estilos: String(a.Estilos || ''), objetivo: String(a.Objetivo || ''), violao: String(a.Tipo_Violao || '')
    })),
    aulas: aulas.map(a => ({ id: a.ID_Aula, numero: a.Numero, titulo: a.Titulo })),
    freq: freq.map(f => ({
      r: f._r, data: iso_(f.Data), aluno: f.ID_Aluno, aula: f.ID_Aula, aulaTitulo: titulo(f.ID_Aula),
      presenca: String(f.Presenca || ''), motivo: String(f.Motivo_Falta || ''),
      reposicao: f.Reposicao_Data instanceof Date ? iso_(f.Reposicao_Data) : String(f.Reposicao_Data || ''),
      obs: String(f.Obs || ''), horario: txt(f.Horario), tipo: String(f.Tipo || '')
    })).sort((x, y) => (x.data < y.data ? -1 : x.data > y.data ? 1 : 0)),
    registros: regs.filter(r => r.Data_Hora instanceof Date).sort((x, y) => y.Data_Hora - x.Data_Hora).slice(0, 80).map(r => ({
      quando: Utilities.formatDate(r.Data_Hora, FUSO, 'dd/MM HH:mm'), aluno: r.ID_Aluno, nome: String(r.Nome || ''),
      aula: titulo(r.ID_Aula), exercicios: String(r.Exercicios || ''), minutos: Number(r.Minutos) || 0,
      dificuldade: String(r.Dificuldade || ''), comentario: String(r.Comentario || ''), aparelho: String(r.Aparelho || '')
    })),
    disp: {
      semanais: dispSh.filter(l => l[0] !== '').map(l => ({ dia: String(l[0]), ini: txt(l[1]), fim: txt(l[2]), uso: String(l[3] || 'Ambos'), ativo: String(l[4] || 'Sim') })),
      excecoes: dispSh.filter(l => l[6] instanceof Date).map(l => ({ data: iso_(l[6]), situacao: String(l[7] || ''), ini: txt(l[8]), fim: txt(l[9]) }))
    },
    resumo: resumo,
    temUrl: !!urlApp_()
  };
}

/* ---------- Alunos ---------- */

function painelSalvarAluno(k, a) {
  checarChave_(k);
  a = a || {};
  const nome = limpa_(a.nome, 80);
  if (normaliza_(nome).length < 2) throw new Error('Digite o nome do aluno.');
  const fone = String(a.telefone || '').replace(/\D/g, '');
  if (fone && (fone.length < 10 || fone.length > 11)) throw new Error('Telefone com DDD, só números (10 ou 11 dígitos).');
  const L = linhas_(ABA.alunos);
  const campos = { Nome: nome, Telefone: fone, Plano: limpa_(a.plano, 40), Status: limpa_(a.status || 'Ativo', 20) };
  if (a.valor !== '' && a.valor != null) campos.Valor_Mensal = Number(a.valor) || 0;
  const trava = LockService.getScriptLock(); trava.waitLock(10000);
  try {
    let existente = a.id ? L.rows.find(r => r.ID_Aluno === a.id) : null;
    if (existente) {
      gravar_(L.sh, L.cab, existente._r, campos);
      return { id: a.id, novo: false };
    }
    const nums = L.rows.map(r => Number(String(r.ID_Aluno).replace(/^\D+/, ''))).filter(n => !isNaN(n));
    const id = 'A' + String((nums.length ? Math.max.apply(null, nums) : 0) + 1).padStart(2, '0');
    const r = primeiraLinhaVazia_(L.sh);
    L.sh.getRange(r, L.cab.indexOf('Telefone') + 1).setNumberFormat('@');
    campos.ID_Aluno = id;
    campos.Data_Inicio = hoje_();
    gravar_(L.sh, L.cab, r, campos);
    return { id: id, novo: true };
  } finally { trava.releaseLock(); }
}

/** oque = 'aparelhos' (libera aparelho novo) ou 'cadastro' (pede o cadastro de novo). */
function painelResetar(k, id, oque) {
  checarChave_(k);
  const L = linhas_(ABA.alunos);
  const a = L.rows.find(r => r.ID_Aluno === id);
  if (!a) throw new Error('Aluno não encontrado.');
  const campos = { Aparelhos: '' };
  if (oque === 'cadastro') campos.Cadastro_Completo = '';
  gravar_(L.sh, L.cab, a._r, campos);
  return true;
}

/* ---------- Enviar link ---------- */

function painelCriarEnvio(k, idAluno, idAula, validadeIso) {
  checarChave_(k);
  const URL = urlApp_();
  if (!URL) throw new Error('Salve a URL do web app em Violão > Definir URL do web app.');
  const aluno = tabela_(ABA.alunos).find(a => a.ID_Aluno === idAluno);
  const aula = tabela_(ABA.aulas).find(a => a.ID_Aula === idAula);
  if (!aluno) throw new Error('Escolha o aluno.');
  if (!aula) throw new Error('Escolha a aula.');
  const validade = deIso_(validadeIso);
  if (validade < hoje_()) throw new Error('A validade precisa ser hoje ou depois.');

  const token = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  const link = URL + '?t=' + token;
  const primeiro = String(aluno.Nome || '').split(' ')[0];
  const msg = 'Olá, ' + primeiro + '! Seus exercícios da Aula ' + aula.Numero + ' (' + aula.Titulo + '). ' +
    'Marque cada dia que praticar. Link válido até ' + data_(validade) + ': ' + link;
  const fone = String(aluno.Telefone || '').replace(/\D/g, '');
  const whats = fone ? 'https://wa.me/55' + fone + '?text=' + encodeURIComponent(msg) : '';

  const sh = ss_().getSheetByName(ABA.envios);
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const trava = LockService.getScriptLock(); trava.waitLock(10000);
  try {
    const r = primeiraLinhaVazia_(sh);
    gravar_(sh, cab, r, { ID_Aluno: idAluno, ID_Aula: idAula, Data_Envio: new Date(), Validade: validade, Token: token, Link: link, WhatsApp: whats });
  } finally { trava.releaseLock(); }
  return { link: link, whatsapp: whats, mensagem: msg };
}

/* ---------- Frequência ---------- */

function painelPresenca(k, linha, status, motivo) {
  checarChave_(k);
  const validos = ['Agendada', 'Presente', 'Falta', 'Falta justificada', 'Reposição'];
  if (validos.indexOf(status) < 0) throw new Error('Situação inválida.');
  const sh = ss_().getSheetByName('Frequencia');
  const cab = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  linha = Number(linha);
  if (!(linha >= 2) || linha > sh.getLastRow()) throw new Error('Aula não encontrada.');
  const ehFalta = st => st === 'Falta' || st === 'Falta justificada';
  const atual = sh.getRange(linha, 1, 1, cab.length).getValues()[0];
  const antes = atual[cab.indexOf('Presenca')];
  const msgs = [];

  // desfazendo uma falta: apaga a reposição automática que ainda não aconteceu
  if (ehFalta(antes) && !ehFalta(status)) {
    const dataFalta = atual[cab.indexOf('Data')];
    const obsRep = 'Reposição da aula de ' + Utilities.formatDate(new Date(dataFalta), FUSO, 'dd/MM');
    const v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      const l = v[i];
      if (l[cab.indexOf('ID_Aluno')] === atual[cab.indexOf('ID_Aluno')] && l[cab.indexOf('Tipo')] === 'Reposição' &&
          l[cab.indexOf('Obs')] === obsRep && l[cab.indexOf('Presenca')] === 'Agendada') {
        sh.getRange(i + 1, 1, 1, cab.length).clearContent();
        msgs.push('A reposição que tinha sido agendada foi cancelada.');
      }
    }
  }

  const campos = { Presenca: status };
  campos.Motivo_Falta = ehFalta(status) ? limpa_(motivo, 40) : '';
  if (!ehFalta(status) || !ehFalta(antes)) campos.Reposicao_Data = '';
  gravar_(sh, cab, linha, campos);
  SpreadsheetApp.flush();
  return msgs.concat(processarFaltas_(ss_()));
}

/** Agenda as próximas N aulas do aluno nos seus horários de "Aula" (ou "Ambos"). */
function painelAgendar(k, idAluno, n) {
  checarChave_(k);
  n = Math.max(1, Math.min(12, Number(n) || 4));
  const ss = ss_();
  const aulas = tabela_(ABA.aulas).sort((a, b) => a.Numero - b.Numero);
  const F = linhas_('Frequencia');
  const minhas = F.rows.filter(f => f.ID_Aluno === idAluno && f.Data instanceof Date);
  if (!tabela_(ABA.alunos).some(a => a.ID_Aluno === idAluno)) throw new Error('Escolha o aluno.');

  // próxima aula do programa
  const numDe = id => { const a = aulas.find(x => x.ID_Aula === id); return a ? Number(a.Numero) : 0; };
  let ultimoNum = minhas.filter(f => f.Tipo !== 'Reposição').reduce((m, f) => Math.max(m, numDe(f.ID_Aula)), 0);
  let inicio = minhas.reduce((m, f) => (f.Data > m ? f.Data : m), hoje_());

  // horários semanais de aula e exceções
  const disp = ss.getSheetByName('Disponibilidade').getRange(2, 1, 30, 10).getValues();
  const semanais = disp.filter(l => l[0] !== '' && semAcento_(l[4]) !== 'nao' && ['aula', 'ambos'].indexOf(semAcento_(l[3])) >= 0)
    .map(l => ({ dow: DIAS_SEMANA.map(semAcento_).indexOf(semAcento_(l[0])), ini: minutos_(l[1]), fim: minutos_(l[2]) }))
    .filter(s => s.dow >= 0 && s.ini != null && s.fim != null);
  if (!semanais.length) throw new Error('Cadastre pelo menos um horário com uso "Aula" ou "Ambos" na Disponibilidade.');
  const bloqueados = new Set(disp.filter(l => l[6] instanceof Date && semAcento_(l[7]) === 'indisponivel').map(l => iso_(l[6])));
  const ocupados = new Set(F.rows.filter(f => f.Data instanceof Date && f.Presenca === 'Agendada')
    .map(f => iso_(f.Data) + ' ' + String(f.Horario instanceof Date ? Utilities.formatDate(f.Horario, FUSO, 'HH:mm') : f.Horario).slice(0, 5)));

  const criadas = [];
  const d = new Date(inicio);
  for (let guarda = 0; criadas.length < n && guarda < 120; guarda++) {
    d.setDate(d.getDate() + 1);
    if (bloqueados.has(iso_(d))) continue;
    const doDia = semanais.filter(s => s.dow === d.getDay()).sort((a, b) => a.ini - b.ini);
    for (const s of doDia) {
      let achou = false;
      for (let t = s.ini; t + DURACAO_AULA_MIN <= s.fim; t += DURACAO_AULA_MIN) {
        const chave = iso_(d) + ' ' + hhmm_(t);
        if (ocupados.has(chave)) continue;
        const prox = aulas.find(a => Number(a.Numero) > ultimoNum);
        if (!prox) break;
        ultimoNum = Number(prox.Numero);
        ocupados.add(chave);
        criadas.push({ data: new Date(d), aula: prox.ID_Aula, horario: hhmm_(t) + '–' + hhmm_(t + DURACAO_AULA_MIN) });
        achou = true;
        break;
      }
      if (achou) break; // uma aula por dia
    }
    if (!aulas.some(a => Number(a.Numero) > ultimoNum)) break;
  }
  if (!criadas.length) throw new Error('Não encontrei horário livre, ou o aluno já está na última aula do programa.');

  const sh = F.sh, cab = F.cab;
  criadas.forEach(c => {
    const r = primeiraLinhaVazia_(sh);
    gravar_(sh, cab, r, { Data: c.data, ID_Aluno: idAluno, ID_Aula: c.aula, Presenca: 'Agendada', Horario: c.horario, Tipo: 'Regular' });
    sh.getRange(r, cab.indexOf('Data') + 1).setNumberFormat('dd/MM/yyyy');
    sh.getRange(r, cab.indexOf('Nome') + 1).setFormula('=IF(B' + r + '="","",IFERROR(INDEX(Alunos!B:B,MATCH(B' + r + ',Alunos!A:A,0)),""))');
    sh.getRange(r, cab.indexOf('Aula') + 1).setFormula('=IF(D' + r + '="","",IFERROR("Aula "&INDEX(Aulas!B:B,MATCH(D' + r + ',Aulas!A:A,0))&" – "&INDEX(Aulas!C:C,MATCH(D' + r + ',Aulas!A:A,0)),""))');
  });
  return criadas.map(c => ({ data: iso_(c.data), aula: c.aula, horario: c.horario }));
}

/* ---------- Disponibilidade ---------- */

function painelSalvarDisp(k, semanais, excecoes) {
  checarChave_(k);
  const dias = DIAS_SEMANA;
  const hora = h => { const m = minutos_(h); if (m == null) throw new Error('Horário inválido: ' + h); return hhmm_(m); };
  const S = (semanais || []).slice(0, 30).map(s => {
    const dia = dias.find(x => semAcento_(x) === semAcento_(s.dia));
    if (!dia) throw new Error('Dia inválido: ' + s.dia);
    const ini = hora(s.ini), fim = hora(s.fim);
    if (minutos_(fim) <= minutos_(ini)) throw new Error(dia + ': o fim precisa ser depois do início.');
    const uso = ['Aula', 'Reposição', 'Ambos'].indexOf(s.uso) >= 0 ? s.uso : 'Ambos';
    return [dia, ini, fim, uso, s.ativo === 'Não' ? 'Não' : 'Sim'];
  });
  const E = (excecoes || []).slice(0, 30).map(x => {
    const sit = x.situacao === 'Disponível extra' ? 'Disponível extra' : 'Indisponível';
    return [deIso_(x.data), sit, sit === 'Disponível extra' ? hora(x.ini) : '', sit === 'Disponível extra' ? hora(x.fim) : ''];
  });
  const sh = ss_().getSheetByName('Disponibilidade');
  sh.getRange(2, 1, 30, 5).clearContent();
  sh.getRange(2, 7, 30, 4).clearContent();
  sh.getRange(2, 2, 30, 2).setNumberFormat('@');
  sh.getRange(2, 9, 30, 2).setNumberFormat('@');
  if (S.length) sh.getRange(2, 1, S.length, 5).setValues(S);
  if (E.length) {
    sh.getRange(2, 7, E.length, 4).setValues(E);
    sh.getRange(2, 7, E.length, 1).setNumberFormat('dd/MM/yyyy');
  }
  return true;
}
