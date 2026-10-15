/**
 * Planner CSIRT — Kanban sobre Google Apps Script + Google Sheets
 * Carlos Quintas | SOC | CSIRT | SIEM Tech Leader
 *
 * Superfície exposta ao cliente (google.script.run): apiBootstrap, apiListTasks,
 * apiSaveTask, apiArchiveTask, apiRunRemindersNow, apiSetTheme,
 * apiListComments, apiAddComment, apiEditComment, apiDeleteComment,
 * apiListChecklist, apiAddChecklistItems, apiUpdateChecklistItem, apiDeleteChecklistItem, apiReorderChecklist.
 * Entradas de operação: setup() (editor), runReminders() (trigger diário).
 * Tudo com sufixo "_" é privado (não invocável via google.script.run).
 */

const APP = {
  NAME: 'Planner CSIRT',
  TZ: 'America/Sao_Paulo',
  SH_TASKS: 'Atividades',
  SH_TEAM: 'Equipe',
  SH_CFG: 'Config',
  SH_LOG: 'Auditoria',
  SH_COMMENTS: 'Comentarios',
  SH_REQ: 'Solicitacoes',
  SH_CHECK: 'Checklist',
  STATUSES: ['Backlog', 'A Fazer', 'Em Andamento', 'Bloqueado', 'Concluído'],
  DONE: 'Concluído',
  BLOCKED: 'Bloqueado',
  PRIORITIES: ['Crítica', 'Alta', 'Média', 'Baixa'],
  PROP_SSID: 'PLANNER_SSID'
};

const COLS = ['id', 'titulo', 'descricao', 'responsavel', 'categoria', 'prioridade', 'status',
  'inicio', 'prazo', 'lembreteDias', 'tags', 'criadoPor', 'criadoEm', 'atualizadoEm',
  'concluidoEm', 'ultimoLembrete', 'versao', 'arquivado', 'origemId'];
const C = Object.fromEntries(COLS.map((k, i) => [k, i]));
const DAY_COLS = ['inicio', 'prazo'];
const TS_COLS = ['criadoEm', 'atualizadoEm', 'concluidoEm', 'ultimoLembrete'];
const TEXT_COLS = ['id', 'titulo', 'descricao', 'responsavel', 'categoria', 'prioridade', 'status', 'tags', 'criadoPor', 'origemId'];

const DEFAULT_CFG = {
  categorias: 'CSIRT, Engenharia de Detecção, SOC, Projetos, Melhoria Contínua',
  lembreteDiasPadrao: '2',          // janela padrão (dias antes do prazo)
  horaLembrete: '8',                // hora do digest diário (America/Sao_Paulo)
  emailGestor: '',                  // CSV de e-mails: resumo semanal + alerta de bloqueio
  resumoSemanalGestor: 'true',      // segunda-feira
  notificarAtribuicao: 'true',      // e-mail ao novo responsável
  notificarComentario: 'true',      // e-mail a responsável + criador em novo comentário (@menção notifica sempre)
  urlPortal: ''                     // URL /exec do deployment (link nos e-mails)
};

let _ss, _team, _cfg; // memo por execução

/* ─────────────────────────── Setup / operação ─────────────────────────── */

function setup() {
  guardOwner_();
  const props = PropertiesService.getScriptProperties();
  let ss;
  const id = props.getProperty(APP.PROP_SSID);
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create(APP.NAME + ' — DB');
    props.setProperty(APP.PROP_SSID, ss.getId());
  }
  _ss = ss;

  const tasks = ensureSheet_(ss, APP.SH_TASKS, COLS);
  ensureTaskHeader_(tasks);
  ensureSheet_(ss, APP.SH_REQ, RCOLS);
  ensureSheet_(ss, APP.SH_CHECK, KCOLS);
  const rows = Math.max(tasks.getMaxRows() - 1, 1);
  DAY_COLS.forEach(k => tasks.getRange(2, C[k] + 1, rows, 1).setNumberFormat('yyyy-mm-dd'));
  TS_COLS.forEach(k => tasks.getRange(2, C[k] + 1, rows, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'));

  const team = ensureSheet_(ss, APP.SH_TEAM, ['nome', 'email', 'papel', 'ativo', 'tema']);
  if (team.getRange(1, 5).getValue() !== 'tema') team.getRange(1, 5).setValue('tema').setFontWeight('bold');
  if (team.getLastRow() < 2) {
    const me = Session.getEffectiveUser().getEmail().toLowerCase();
    team.appendRow([me.split('@')[0], me, 'admin', true, 'system']);
  }

  const cfg = ensureSheet_(ss, APP.SH_CFG, ['chave', 'valor']);
  const existing = cfg.getLastRow() > 1 ? cfg.getRange(2, 1, cfg.getLastRow() - 1, 1).getValues().flat() : [];
  Object.entries(DEFAULT_CFG).forEach(([k, v]) => { if (!existing.includes(k)) cfg.appendRow([k, "'" + v]); });

  ensureSheet_(ss, APP.SH_LOG, ['timestamp', 'usuario', 'acao', 'id', 'detalhe']);
  const cm = ensureSheet_(ss, APP.SH_COMMENTS, CCOLS);
  ['criadoEm', 'editadoEm'].forEach(k =>
    cm.getRange(2, CC[k] + 1, Math.max(cm.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss'));

  ['Sheet1', 'Página1', 'Planilha1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  _cfg = null;
  installTriggers_(Number(readCfg_().horaLembrete) || 8);
  console.log('DB: ' + ss.getUrl());
}

/** Rode após alterar Config.horaLembrete. */
function reinstallTriggers() {
  guardOwner_();
  installTriggers_(Number(readCfg_().horaLembrete) || 8);
}

function installTriggers_(hour) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runReminders')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runReminders').timeBased().everyDays(1).atHour(hour).inTimezone(APP.TZ).create();
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Bloqueia invocação via google.script.run por visitantes (execute-as-me). */
function guardOwner_() {
  const active = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const eff = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (active && active !== eff) throw new Error('Execução restrita ao owner do script.');
}

/* ─────────────────────────── Web App ─────────────────────────── */

/**
 * Roteamento:
 *  ?acao=decidir&req=R-…&d=aprovar|reprovar  → página de confirmação (somente admin)
 *  membro ativo                              → app
 *  usuário do domínio sem acesso             → página "Solicitar acesso"
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) return simplePage_('Não foi possível identificar o usuário. Acesse com sua conta Google corporativa.');

  if (p.acao === 'decidir') {
    const admin = getTeam_().find(t => t.email === email && t.ativo && t.papel === 'admin');
    if (!admin) return accessPage_('negado', { email });
    let req = null;
    try { req = findRequest_(reqSheet_(), p.req).r; } catch (err) { /* inexistente */ }
    return accessPage_('decidir', {
      req, decisao: p.d === 'reprovar' ? 'reprovar' : 'aprovar', url: portalUrl_(), admin: admin.nome
    });
  }

  if (getTeam_().some(t => t.email === email && t.ativo)) {
    const t = HtmlService.createTemplateFromFile('Index');
    t.appName = APP.NAME;
    return frame_(t.evaluate());
  }
  return accessPage_('solicitar', accessState_(email));
}

function frame_(out) {
  return out.setTitle(APP.NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // embed no Google Sites
}

function simplePage_(msg) {
  return frame_(HtmlService.createHtmlOutput(
    '<p style="font-family:Segoe UI,Arial,sans-serif;padding:24px">' + esc_(msg) + '</p>'));
}

/** JSON seguro para embutir em <script> (neutraliza </script> e afins em dados do usuário). */
function safeJson_(v) {
  return JSON.stringify(v).replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
}

function accessPage_(mode, data) {
  const t = HtmlService.createTemplateFromFile('Acesso');
  t.appName = APP.NAME;
  t.modeJson = safeJson_(mode);
  t.dataJson = safeJson_(data);
  return frame_(t.evaluate());
}

function include(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

/* ─────────────────────────── Acesso a dados ─────────────────────────── */

function ss_() {
  if (_ss) return _ss;
  const id = PropertiesService.getScriptProperties().getProperty(APP.PROP_SSID);
  if (!id) throw new Error('Planner não inicializado: execute setup() no editor.');
  return (_ss = SpreadsheetApp.openById(id));
}
function sheet_(name) { return ss_().getSheetByName(name); }

function readCfg_() {
  if (_cfg) return _cfg;
  const sh = sheet_(APP.SH_CFG);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues() : [];
  _cfg = Object.assign({}, DEFAULT_CFG);
  rows.forEach(([k, v]) => { if (k) _cfg[String(k).trim()] = String(v).trim(); });
  return _cfg;
}

function getTeam_() {
  if (_team) return _team;
  const sh = sheet_(APP.SH_TEAM);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues() : [];
  _team = rows
    .map(([nome, email, papel, ativo, tema], i) => ({
      nome: String(nome || email).trim(),
      email: String(email).trim().toLowerCase(),
      papel: String(papel || 'membro').trim().toLowerCase(),
      ativo: ativo === true || String(ativo).toUpperCase() === 'TRUE',
      tema: THEMES.includes(String(tema).trim()) ? String(tema).trim() : 'system',
      _row: i + 2
    }))
    .filter(m => m.email);
  return _team;
}

const THEMES = ['light', 'dark', 'system'];

/** Preferência de tema por usuário (aba Equipe, coluna "tema"): sincroniza entre dispositivos. */
function apiSetTheme(tema) {
  const user = currentUser_();
  if (!THEMES.includes(tema)) throw new Error('Tema inválido.');
  if (user.tema === tema) return tema;
  withLock_(() => {
    const sh = sheet_(APP.SH_TEAM);
    if (sh.getRange(1, 5).getValue() !== 'tema') sh.getRange(1, 5).setValue('tema').setFontWeight('bold');
    sh.getRange(user._row, 5).setValue(tema);
  });
  return tema;
}

function currentUser_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) throw new Error('Não foi possível identificar o usuário (acesso fora do domínio Workspace?).');
  const m = getTeam_().find(t => t.email === email && t.ativo);
  if (!m) throw new Error('Acesso negado: ' + email + ' não está ativo na aba Equipe.');
  return m;
}

function assertAdmin_(u) {
  if (u.papel !== 'admin') throw new Error('Ação restrita a administradores.');
}

function readTasks_(includeArchived) {
  const sh = sheet_(APP.SH_TASKS);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const out = [];
  const w = Math.min(COLS.length, sh.getMaxColumns()); // tolera planilha antiga/colunas removidas
  sh.getRange(2, 1, n, w).getValues().forEach((r, i) => {
    while (r.length < COLS.length) r.push('');
    if (!r[C.id]) return;
    const t = rowToTask_(r, i + 2);
    if (includeArchived || !t.arquivado) out.push(t);
  });
  return out;
}

function findTask_(sh, id) {
  const cell = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1)
    .createTextFinder(String(id)).matchEntireCell(true).findNext();
  if (!cell) throw new Error('Atividade não encontrada: ' + id);
  const rowIdx = cell.getRow();
  return { rowIdx, task: rowToTask_(sh.getRange(rowIdx, 1, 1, COLS.length).getValues()[0], rowIdx) };
}

function rowToTask_(r, rowIdx) {
  const t = {};
  COLS.forEach((k, i) => {
    const v = r[i];
    if (DAY_COLS.includes(k)) t[k] = fmtDay_(v);
    else if (TS_COLS.includes(k)) t[k] = v instanceof Date ? Utilities.formatDate(v, APP.TZ, "yyyy-MM-dd'T'HH:mm:ss") : String(v || '');
    else t[k] = v;
  });
  TEXT_COLS.forEach(k => { t[k] = t[k] instanceof Date ? fmtDay_(t[k]) : String(t[k] == null ? '' : t[k]); });
  t.versao = Number(t.versao) || 1;
  t.lembreteDias = (t.lembreteDias === '' || t.lembreteDias == null) ? '' : Number(t.lembreteDias);
  t.arquivado = t.arquivado === true || String(t.arquivado).toUpperCase() === 'TRUE';
  if (rowIdx) t._row = rowIdx;
  return t;
}

/** forSheet=true: strings de texto com prefixo "'" (literal; neutraliza formula/CSV injection e autoconversão). */
function taskToRow_(t, forSheet) {
  return COLS.map(k => {
    let v = t[k];
    if (v == null) v = '';
    if ((DAY_COLS.includes(k) || TS_COLS.includes(k)) && v !== '' && !(v instanceof Date)) v = parseDate_(v);
    if (forSheet && TEXT_COLS.includes(k) && typeof v === 'string' && v !== '') v = "'" + v;
    return v;
  });
}

const toClient_ = t => { const c = Object.assign({}, t); delete c._row; delete c._sit; return c; };

function fmtDay_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, APP.TZ, 'yyyy-MM-dd');
  const s = String(v || '');
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : '';
}

function parseDate_(s) {
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (!m) return '';
  return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
}

function daysBetween_(a, b) {
  const p = s => { const [y, m, d] = s.split('-').map(Number); return Date.UTC(y, m - 1, d); };
  return Math.round((p(b) - p(a)) / 864e5);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Planner ocupado (lock timeout). Tente novamente.');
  try { return fn(); } finally { lock.releaseLock(); }
}

function log_(user, acao, id, detalhe) {
  const safe = v => { const s = String(v == null ? '' : v).slice(0, 2000); return s ? "'" + s : ''; };
  sheet_(APP.SH_LOG).appendRow([new Date(), safe(user), safe(acao), safe(id), safe(detalhe)]);
}

function newId_() {
  return 'T-' + Utilities.formatDate(new Date(), APP.TZ, 'yyMMdd') + '-' +
    Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase(); // 8 hex: seguro p/ import em lote
}

/* ─────────────────────────── API (cliente) ─────────────────────────── */

function apiBootstrap() {
  const user = currentUser_();
  const cfg = readCfg_();
  return {
    user: { nome: user.nome, email: user.email, papel: user.papel, tema: user.tema },
    team: getTeam_().filter(t => t.ativo).map(({ nome, email, papel }) => ({ nome, email, papel })),
    statuses: APP.STATUSES,
    done: APP.DONE,
    blocked: APP.BLOCKED,
    priorities: APP.PRIORITIES,
    categorias: String(cfg.categorias || '').split(',').map(s => s.trim()).filter(Boolean),
    lembreteDiasPadrao: Number(cfg.lembreteDiasPadrao) || 2,
    pendingAccess: user.papel === 'admin' ? readRequests_().filter(r => r.status === 'pendente').length : 0,
    tasks: withCounts_(readTasks_())
  };
}

function apiListTasks() {
  currentUser_();
  return withCounts_(readTasks_());
}

function apiSaveTask(input) {
  const user = currentUser_();
  input = input || {};
  const isCreate = !input.id;
  const clean = validate_(input, isCreate);

  const res = withLock_(() => {
    const sh = sheet_(APP.SH_TASKS);
    const now = new Date();
    let task, before = null;

    if (isCreate) {
      task = Object.assign(Object.fromEntries(COLS.map(k => [k, ''])), clean, {
        id: newId_(), criadoPor: user.email, criadoEm: now, atualizadoEm: now, versao: 1, arquivado: false
      });
      if (task.status === APP.DONE) task.concluidoEm = now;
      sh.appendRow(taskToRow_(task, true));
      log_(user.email, 'CREATE', task.id, task.titulo);
    } else {
      const found = findTask_(sh, input.id);
      before = found.task;
      if (before.arquivado) throw new Error('Atividade arquivada.');
      assertVersion_(before, input.versao);
      task = Object.assign({}, before, clean, { atualizadoEm: now, versao: before.versao + 1 });
      if (task.status === APP.DONE && before.status !== APP.DONE) task.concluidoEm = now;
      if (task.status !== APP.DONE) task.concluidoEm = '';
      sh.getRange(found.rowIdx, 1, 1, COLS.length).setValues([taskToRow_(task, true)]);
      log_(user.email, 'UPDATE', task.id, diff_(before, task));
    }
    SpreadsheetApp.flush();
    return { before, task };
  });

  notify_(res.before, res.task, user); // fora do lock: I/O de e-mail não segura o lock
  return toClient_(rowToTask_(taskToRow_(res.task, false)));
}

function apiArchiveTask(id, versao) {
  const user = currentUser_();
  withLock_(() => {
    const sh = sheet_(APP.SH_TASKS);
    const { rowIdx, task } = findTask_(sh, id);
    if (user.papel !== 'admin' && task.criadoPor !== user.email) {
      throw new Error('Somente o criador ou um admin pode arquivar.');
    }
    assertVersion_(task, versao);
    sh.getRange(rowIdx, C.arquivado + 1).setValue(true);
    sh.getRange(rowIdx, C.versao + 1).setValue(task.versao + 1);
    sh.getRange(rowIdx, C.atualizadoEm + 1).setValue(new Date());
    log_(user.email, 'ARCHIVE', id, task.titulo);
  });
  return true;
}

function apiRunRemindersNow() {
  const user = currentUser_();
  assertAdmin_(user);
  const r = runReminders_();
  log_(user.email, 'REMINDERS_MANUAL', '', JSON.stringify(r));
  return r;
}

function assertVersion_(current, versao) {
  if (versao != null && versao !== '' && Number(versao) !== current.versao) {
    throw new Error(`CONFLITO: ${current.id} foi alterada por outro usuário (v${current.versao}). Recarregando.`);
  }
}

function validate_(i, isCreate) {
  const o = {};
  const str = (v, max) => String(v == null ? '' : v).trim().slice(0, max);
  const has = k => Object.prototype.hasOwnProperty.call(i, k);

  if (has('titulo') || isCreate) {
    o.titulo = str(i.titulo, 200);
    if (!o.titulo) throw new Error('Título obrigatório.');
  }
  if (has('descricao')) o.descricao = str(i.descricao, 5000);
  if (has('responsavel') || isCreate) {
    o.responsavel = str(i.responsavel, 254).toLowerCase();
    if (!getTeam_().some(t => t.ativo && t.email === o.responsavel)) throw new Error('Responsável inválido ou inativo.');
  }
  if (has('categoria')) o.categoria = str(i.categoria, 80);
  if (has('prioridade') || isCreate) o.prioridade = APP.PRIORITIES.includes(i.prioridade) ? i.prioridade : 'Média';
  if (has('status') || isCreate) {
    o.status = i.status || 'Backlog';
    if (!APP.STATUSES.includes(o.status)) throw new Error('Status inválido.');
  }
  DAY_COLS.forEach(k => {
    if (!has(k)) return;
    const v = str(i[k], 10);
    if (v && !/^\d{4}-\d{2}-\d{2}$/.test(v)) throw new Error(`Data inválida em ${k}.`);
    o[k] = v;
  });
  if (has('lembreteDias')) {
    o.lembreteDias = (i.lembreteDias === '' || i.lembreteDias == null)
      ? '' : Math.max(0, Math.min(60, parseInt(i.lembreteDias, 10) || 0));
  }
  if (has('tags')) o.tags = str(i.tags, 300);
  if (o.inicio && o.prazo && o.inicio > o.prazo) throw new Error('Início posterior ao prazo.');
  return o;
}

function diff_(a, b) {
  return ['titulo', 'responsavel', 'categoria', 'prioridade', 'status', 'inicio', 'prazo', 'lembreteDias', 'tags', 'descricao']
    .filter(k => String(a[k] == null ? '' : a[k]) !== String(b[k] == null ? '' : b[k]))
    .map(k => k === 'descricao' ? 'descricao alterada' : `${k}: ${a[k]} → ${b[k]}`)
    .join('; ');
}

/* ─────────────────────────── Comentários ─────────────────────────── */
// Aba própria (append-only + soft delete). Não altera versao/atualizadoEm da atividade:
// comentar nunca gera CONFLITO com quem está editando o formulário.

const CCOLS = ['id', 'taskId', 'autor', 'texto', 'criadoEm', 'editadoEm', 'excluido'];
const CC = Object.fromEntries(CCOLS.map((k, i) => [k, i]));
const MAX_COMMENT = 5000;

function commentsSheet_() {
  return sheet_(APP.SH_COMMENTS) || ensureSheet_(ss_(), APP.SH_COMMENTS, CCOLS); // auto-cria em upgrades
}

function tsStr_(v) {
  return v instanceof Date ? Utilities.formatDate(v, APP.TZ, "yyyy-MM-dd'T'HH:mm:ss") : String(v || '');
}

function rowToComment_(r) {
  return {
    id: String(r[CC.id]),
    taskId: String(r[CC.taskId]),
    autor: String(r[CC.autor]).toLowerCase(),
    texto: String(r[CC.texto]),
    criadoEm: tsStr_(r[CC.criadoEm]),
    editadoEm: tsStr_(r[CC.editadoEm]),
    excluido: r[CC.excluido] === true || String(r[CC.excluido]).toUpperCase() === 'TRUE'
  };
}

function readComments_() {
  const sh = sheet_(APP.SH_COMMENTS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, CCOLS.length).getValues()
    .filter(r => r[CC.id]).map(rowToComment_);
}

function findComment_(sh, id) {
  const n = sh.getLastRow() - 1;
  const cell = n > 0 ? sh.getRange(2, 1, n, 1).createTextFinder(String(id)).matchEntireCell(true).findNext() : null;
  if (!cell) throw new Error('Comentário não encontrado.');
  return { rowIdx: cell.getRow(), c: rowToComment_(sh.getRange(cell.getRow(), 1, 1, CCOLS.length).getValues()[0]) };
}

function withCounts_(tasks) {
  const m = {}, ck = {};
  readComments_().forEach(c => { if (!c.excluido) m[c.taskId] = (m[c.taskId] || 0) + 1; });
  readChecklist_().forEach(i => {
    if (i.excluido) return;
    const k = ck[i.taskId] = ck[i.taskId] || { t: 0, d: 0 };
    k.t++; if (i.feito) k.d++;
  });
  return tasks.map(t => Object.assign(toClient_(t), {
    nComentarios: m[t.id] || 0, ckTotal: (ck[t.id] || {}).t || 0, ckDone: (ck[t.id] || {}).d || 0
  }));
}

function commentToClient_(c, user) {
  return {
    id: c.id, taskId: c.taskId, autor: c.autor, texto: c.texto,
    criadoEm: c.criadoEm, editadoEm: c.editadoEm,
    podeEditar: c.autor === user.email,
    podeExcluir: c.autor === user.email || user.papel === 'admin'
  };
}

function cleanComment_(texto) {
  const s = String(texto == null ? '' : texto).replace(/\r\n?/g, '\n').trim().slice(0, MAX_COMMENT);
  if (!s) throw new Error('Comentário vazio.');
  return s;
}

/** @menção por parte local do e-mail (@carlos.quintas) ou e-mail completo. */
function mentionsIn_(texto) {
  const team = getTeam_().filter(m => m.ativo);
  const out = new Set();
  (texto.match(/@[\w.\-]+(?:@[\w.\-]+)?/g) || []).forEach(tok => {
    const k = tok.slice(1).toLowerCase().replace(/[.\-]+$/, '');
    const m = team.find(t => t.email === k || t.email.split('@')[0] === k);
    if (m) out.add(m.email);
  });
  return out;
}

function apiListComments(taskId) {
  const user = currentUser_();
  return readComments_()
    .filter(c => c.taskId === String(taskId) && !c.excluido)
    .sort((a, b) => a.criadoEm.localeCompare(b.criadoEm))
    .map(c => commentToClient_(c, user));
}

function apiAddComment(taskId, texto) {
  const user = currentUser_();
  texto = cleanComment_(texto);
  const res = withLock_(() => {
    const { task } = findTask_(sheet_(APP.SH_TASKS), taskId);
    if (task.arquivado) throw new Error('Atividade arquivada: comentários bloqueados.');
    const now = new Date();
    const c = {
      id: 'C-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase(),
      taskId: task.id, autor: user.email, texto, criadoEm: tsStr_(now), editadoEm: '', excluido: false
    };
    commentsSheet_().appendRow([c.id, c.taskId, c.autor, "'" + c.texto, now, '', false]);
    log_(user.email, 'COMMENT_ADD', task.id, `${c.id}: ${texto.slice(0, 300)}`);
    return { task, c };
  });
  notifyComment_(res.task, res.c, user); // fora do lock
  return commentToClient_(res.c, user);
}

function apiEditComment(id, texto) {
  const user = currentUser_();
  texto = cleanComment_(texto);
  return withLock_(() => {
    const sh = commentsSheet_();
    const { rowIdx, c } = findComment_(sh, id);
    if (c.excluido) throw new Error('Comentário removido.');
    if (c.autor !== user.email) throw new Error('Só o autor pode editar o comentário.');
    const now = new Date();
    sh.getRange(rowIdx, CC.texto + 1).setValue("'" + texto);
    sh.getRange(rowIdx, CC.editadoEm + 1).setValue(now);
    log_(user.email, 'COMMENT_EDIT', c.taskId, `${id}: ${c.texto.slice(0, 150)} → ${texto.slice(0, 150)}`);
    return commentToClient_(Object.assign(c, { texto, editadoEm: tsStr_(now) }), user);
  });
}

function apiDeleteComment(id) {
  const user = currentUser_();
  withLock_(() => {
    const sh = commentsSheet_();
    const { rowIdx, c } = findComment_(sh, id);
    if (c.excluido) return;
    if (c.autor !== user.email && user.papel !== 'admin') throw new Error('Só o autor ou um admin pode excluir.');
    sh.getRange(rowIdx, CC.excluido + 1).setValue(true); // soft delete: texto preservado para auditoria
    log_(user.email, 'COMMENT_DELETE', c.taskId, `${id} (autor ${c.autor})`);
  });
  return true;
}

function notifyComment_(task, c, actor) {
  try {
    const cfg = readCfg_();
    const mentioned = mentionsIn_(c.texto);
    const to = new Set(mentioned);
    if (cfg.notificarComentario !== 'false') [task.responsavel, task.criadoPor].forEach(e => e && to.add(e));
    to.delete(actor.email);
    const active = new Set(getTeam_().filter(m => m.ativo).map(m => m.email));
    const url = portalUrl_();
    [...to].filter(e => active.has(e)).forEach(email => {
      const men = mentioned.has(email);
      sendMail_(email,
        `[${APP.NAME}] ${men ? 'Você foi mencionado' : 'Novo comentário'} em ${task.id}: ${task.titulo}`,
        `<p><b>${esc_(actor.nome)}</b> ${men ? 'mencionou você em' : 'comentou em'} uma atividade:</p>
         <blockquote style="margin:12px 0;padding:10px 14px;border-left:4px solid #024DA1;background:#f8fafc;white-space:pre-wrap;color:#1f2937">${esc_(c.texto)}</blockquote>` +
        tasksTable_([task], false), url);
    });
  } catch (e) {
    console.warn('notifyComment_ falhou: ' + e);
  }
}

/* ─────────────────────────── Notificações ─────────────────────────── */

function notify_(before, t, actor) {
  try {
    const cfg = readCfg_();
    const url = portalUrl_();
    const reassigned = !before || before.responsavel !== t.responsavel;
    if (cfg.notificarAtribuicao === 'true' && reassigned && t.responsavel && t.responsavel !== actor.email) {
      sendMail_(t.responsavel, `[${APP.NAME}] Atividade atribuída: ${t.titulo}`,
        `<p><b>${esc_(actor.nome)}</b> atribuiu a você a atividade abaixo.</p>` + tasksTable_([t], false) +
        (t.descricao ? `<p style="white-space:pre-wrap;color:#334155">${esc_(t.descricao)}</p>` : ''), url);
    }
    if (cfg.emailGestor && t.status === APP.BLOCKED && (!before || before.status !== APP.BLOCKED)) {
      sendMail_(cfg.emailGestor, `[${APP.NAME}] Atividade BLOQUEADA: ${t.titulo}`,
        `<p><b>${esc_(actor.nome)}</b> moveu a atividade para <b>Bloqueado</b>.</p>` + tasksTable_([t], false), url);
    }
  } catch (e) {
    console.warn('notify_ falhou: ' + e);
  }
}

/** Trigger diário. Digest único por responsável (economia de quota + menos ruído). */
function runReminders() {
  guardOwner_();
  return runReminders_();
}

function runReminders_() {
  const cfg = readCfg_();
  const now = new Date();
  const today = Utilities.formatDate(now, APP.TZ, 'yyyy-MM-dd');
  const defDays = Number(cfg.lembreteDiasPadrao) || 2;
  const team = getTeam_();
  const all = readTasks_();
  const byUser = {};

  all.forEach(t => {
    if (t.status === APP.DONE || !t.responsavel) return;
    if (t.ultimoLembrete && t.ultimoLembrete.slice(0, 10) === today) return; // dedupe diário
    let sit = null;
    if (t.prazo) {
      const d = daysBetween_(today, t.prazo);
      const win = t.lembreteDias === '' ? defDays : t.lembreteDias;
      if (d < 0) sit = { tipo: 'atrasada', ord: 0, txt: `Atrasada ${-d}d` };
      else if (d <= win) sit = { tipo: 'vencendo', ord: 1, txt: d === 0 ? 'Vence hoje' : `Vence em ${d}d` };
    }
    if (!sit && t.status === APP.BLOCKED) sit = { tipo: 'bloqueada', ord: 2, txt: 'Bloqueada' };
    if (!sit) return;
    t._sit = sit;
    (byUser[t.responsavel] = byUser[t.responsavel] || []).push(t);
  });

  const url = portalUrl_();
  const sh = sheet_(APP.SH_TASKS);
  let enviados = 0, itens = 0, pulados = 0;

  Object.entries(byUser).forEach(([email, list]) => {
    const member = team.find(m => m.email === email);
    if (member && !member.ativo) return;
    if (MailApp.getRemainingDailyQuota() < 1) { pulados += list.length; return; }
    list.sort((a, b) => a._sit.ord - b._sit.ord || (a.prazo || '9').localeCompare(b.prazo || '9'));
    const late = list.filter(t => t._sit.tipo === 'atrasada').length;
    try {
      sendMail_(email,
        `[${APP.NAME}] ${list.length} atividade(s) requerem atenção` + (late ? ` — ${late} atrasada(s)` : ''),
        `<p>Olá ${esc_(member ? member.nome : email)}, segue o resumo diário das suas atividades:</p>` + tasksTable_(list, true),
        url);
    } catch (e) {
      console.warn(`Falha ao enviar para ${email}: ${e}`);
      pulados += list.length;
      return;
    }
    // Marca somente após envio OK; re-localiza por ID sob lock (sem bump de versão).
    withLock_(() => list.forEach(t => {
      try { sh.getRange(findTask_(sh, t.id).rowIdx, C.ultimoLembrete + 1).setValue(now); } catch (e) { /* removida */ }
    }));
    enviados++; itens += list.length;
  });

  if (cfg.resumoSemanalGestor === 'true' && cfg.emailGestor && Utilities.formatDate(now, APP.TZ, 'u') === '1') {
    try { sendManagerDigest_(all, url); } catch (e) { console.warn('Digest gestor falhou: ' + e); }
  }

  const r = { enviados, itens, pulados, quotaRestante: MailApp.getRemainingDailyQuota() };
  log_('system', 'REMINDERS', '', JSON.stringify(r));
  return r;
}

function sendManagerDigest_(all, url) {
  const cfg = readCfg_();
  const today = Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd');
  const isDone = t => t.status === APP.DONE;
  const isLate = t => !isDone(t) && t.prazo && daysBetween_(today, t.prazo) < 0;
  const done = all.filter(isDone);
  const open = all.filter(t => !isDone(t));
  const late = open.filter(isLate);
  const blocked = open.filter(t => t.status === APP.BLOCKED);
  const done7 = done.filter(t => t.concluidoEm && daysBetween_(t.concluidoEm.slice(0, 10), today) <= 7);
  const pct = all.length ? Math.round(done.length * 100 / all.length) : 0;
  const name = e => (getTeam_().find(m => m.email === e) || {}).nome || e;

  const kpi = (l, v, c) => `<td style="padding:10px 14px;border:1px solid #e5e7eb;border-radius:8px;text-align:center">
      <div style="font-size:22px;font-weight:700;color:${c || '#0f172a'}">${v}</div>
      <div style="font-size:12px;color:#64748b">${l}</div></td>`;

  const people = {};
  all.forEach(t => {
    const p = people[t.responsavel] = people[t.responsavel] || { total: 0, done: 0, open: 0, late: 0, done7: 0 };
    p.total++;
    if (isDone(t)) p.done++; else p.open++;
    if (isLate(t)) p.late++;
    if (done7.includes(t)) p.done7++;
  });
  const th = s => `<th style="text-align:left;padding:6px 8px;border-bottom:2px solid #e5e7eb;font-size:12px;color:#475569">${s}</th>`;
  const td = s => `<td style="padding:6px 8px;border-bottom:1px solid #f1f5f9">${s}</td>`;
  const peopleTbl = `<table style="border-collapse:collapse;width:100%;font-size:13px"><tr>${th('Responsável')}${th('Abertas')}${th('Atrasadas')}${th('Concluídas 7d')}${th('% Conclusão')}</tr>` +
    Object.entries(people).sort((a, b) => b[1].late - a[1].late).map(([e, p]) =>
      `<tr>${td(esc_(name(e)))}${td(p.open)}${td(p.late ? `<b style="color:#dc2626">${p.late}</b>` : 0)}${td(p.done7)}${td(Math.round(p.done * 100 / p.total) + '%')}</tr>`).join('') +
    '</table>';

  late.forEach(t => { t._sit = { tipo: 'atrasada', txt: `Atrasada ${-daysBetween_(today, t.prazo)}d — ${name(t.responsavel)}` }; });

  const html = `<p>Resumo semanal do time.</p>
    <table style="border-collapse:separate;border-spacing:8px"><tr>
      ${kpi('Conclusão', pct + '%')}${kpi('Abertas', open.length)}${kpi('Atrasadas', late.length, '#dc2626')}
      ${kpi('Bloqueadas', blocked.length, '#7c3aed')}${kpi('Concluídas 7d', done7.length, '#16a34a')}
    </tr></table>
    <h3 style="font-size:14px;margin:18px 0 6px">Por responsável</h3>${peopleTbl}
    ${late.length ? `<h3 style="font-size:14px;margin:18px 0 6px">Atrasadas</h3>${tasksTable_(late, true)}` : ''}`;

  sendMail_(cfg.emailGestor, `[${APP.NAME}] Resumo semanal — ${pct}% concluído, ${late.length} atrasada(s)`, html, url);
}

function sendMail_(to, subject, inner, url) {
  MailApp.sendEmail({ to, subject, htmlBody: shell_(inner, url), name: APP.NAME });
}

function shell_(inner, url) {
  return `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#1f2937;max-width:760px">
    <div style="background:#0f172a;color:#fff;padding:12px 16px;border-radius:8px 8px 0 0;font-weight:600">${esc_(APP.NAME)}</div>
    <div style="border:1px solid #e5e7eb;border-top:0;padding:16px;border-radius:0 0 8px 8px">${inner}
      ${url ? `<p style="margin-top:18px"><a href="${esc_(url)}" style="background:#2563eb;color:#fff;padding:9px 16px;border-radius:6px;text-decoration:none;font-weight:600">Abrir o Planner</a></p>` : ''}
    </div>
    <p style="font-size:11px;color:#94a3b8;margin-top:8px">Mensagem automática — não contém anexos nem solicita credenciais.</p>
  </div>`;
}

function tasksTable_(tasks, withSit) {
  const th = s => `<th style="text-align:left;padding:6px 8px;border-bottom:2px solid #e5e7eb;font-size:12px;color:#475569">${s}</th>`;
  const td = (s, st) => `<td style="padding:6px 8px;border-bottom:1px solid #f1f5f9;${st || ''}">${s}</td>`;
  const color = { atrasada: '#dc2626', vencendo: '#d97706', bloqueada: '#7c3aed' };
  return `<table style="border-collapse:collapse;width:100%;font-size:13px"><tr>
      ${th('ID')}${th('Atividade')}${th('Status')}${th('Prioridade')}${th('Prazo')}${withSit ? th('Situação') : ''}</tr>` +
    tasks.map(t => `<tr>
      ${td(esc_(t.id), 'font-family:Consolas,monospace;font-size:12px;color:#64748b;white-space:nowrap')}
      ${td(esc_(t.titulo))}${td(esc_(t.status))}${td(esc_(t.prioridade))}${td(fmtBr_(t.prazo), 'white-space:nowrap')}
      ${withSit && t._sit ? td(esc_(t._sit.txt), `color:${color[t._sit.tipo]};font-weight:600;white-space:nowrap`) : ''}</tr>`).join('') +
    '</table>';
}

function portalUrl_() {
  const cfg = readCfg_();
  if (cfg.urlPortal) return cfg.urlPortal;
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

function fmtBr_(v) { const d = fmtDay_(v); return d ? d.split('-').reverse().join('/') : '—'; }

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/* ─────────────────────────── Gestão de acesso ─────────────────────────── */
// Fluxo: não-membro do domínio abre o portal → solicita (nome + justificativa) → admins recebem
// e-mail com botões Aprovar/Reprovar → o link abre uma página de CONFIRMAÇÃO (o GET não decide nada;
// scanners de link como Safe Links/Gmail fazem prefetch) → admin confirma via google.script.run.
// Autorização = sessão Google do admin (domínio) + papel admin na aba Equipe.

const RCOLS = ['id', 'email', 'nome', 'justificativa', 'status', 'criadoEm', 'decididoPor', 'decididoEm', 'motivo'];
const RC = Object.fromEntries(RCOLS.map((k, i) => [k, i]));
const REQUEST_COOLDOWN_H = 24; // após reprovação

function reqSheet_() { return sheet_(APP.SH_REQ) || ensureSheet_(ss_(), APP.SH_REQ, RCOLS); }

function rowToRequest_(r) {
  return {
    id: String(r[RC.id]), email: String(r[RC.email]).toLowerCase(), nome: String(r[RC.nome]),
    justificativa: String(r[RC.justificativa]), status: String(r[RC.status] || 'pendente'),
    criadoEm: tsStr_(r[RC.criadoEm]), decididoPor: String(r[RC.decididoPor] || ''),
    decididoEm: tsStr_(r[RC.decididoEm]), motivo: String(r[RC.motivo] || '')
  };
}

function readRequests_() {
  const sh = sheet_(APP.SH_REQ);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, RCOLS.length).getValues().filter(r => r[RC.id]).map(rowToRequest_);
}

function findRequest_(sh, id) {
  const n = sh.getLastRow() - 1;
  const cell = n > 0 && id ? sh.getRange(2, 1, n, 1).createTextFinder(String(id)).matchEntireCell(true).findNext() : null;
  if (!cell) throw new Error('Solicitação não encontrada.');
  return { rowIdx: cell.getRow(), r: rowToRequest_(sh.getRange(cell.getRow(), 1, 1, RCOLS.length).getValues()[0]) };
}

/** Estado da solicitação do usuário atual (para a página de acesso). */
function accessState_(email) {
  const last = readRequests_().filter(r => r.email === email).pop();
  const st = { email, status: 'novo', podeSolicitar: true };
  if (!last || last.status === 'aprovado') return st; // aprovado mas não-membro = foi desativado depois
  Object.assign(st, { status: last.status, criadoEm: last.criadoEm, decididoEm: last.decididoEm, motivo: last.motivo });
  if (last.status === 'pendente') st.podeSolicitar = false;
  if (last.status === 'reprovado') {
    const libera = new Date(new Date(last.decididoEm).getTime() + REQUEST_COOLDOWN_H * 3600e3);
    st.podeSolicitar = Date.now() >= libera.getTime();
    st.liberaEm = tsStr_(libera);
  }
  return st;
}

function ownerDomain_() { return Session.getEffectiveUser().getEmail().toLowerCase().split('@')[1]; }

function apiRequestAccess(nome, justificativa) {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) throw new Error('Usuário não identificado.');
  if (email.split('@')[1] !== ownerDomain_()) throw new Error('Somente contas do domínio corporativo.');
  if (getTeam_().some(t => t.email === email && t.ativo)) throw new Error('Você já tem acesso. Recarregue a página.');
  nome = String(nome || '').trim().slice(0, 120);
  justificativa = String(justificativa || '').replace(/\r\n?/g, '\n').trim().slice(0, 500);
  if (nome.length < 3) throw new Error('Informe seu nome.');
  if (justificativa.length < 10) throw new Error('Descreva a justificativa (mín. 10 caracteres).');

  const req = withLock_(() => {
    const st = accessState_(email);
    if (!st.podeSolicitar) {
      throw new Error(st.status === 'pendente' ? 'Já existe uma solicitação pendente.' : 'Aguarde para solicitar novamente.');
    }
    const now = new Date();
    const r = { id: 'R-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase(),
      email, nome, justificativa, status: 'pendente', criadoEm: tsStr_(now) };
    reqSheet_().appendRow([r.id, "'" + email, "'" + nome, "'" + justificativa, 'pendente', now, '', '', '']);
    log_(email, 'ACCESS_REQUEST', r.id, justificativa.slice(0, 300));
    return r;
  });
  notifyAdminsAccess_(req);
  return accessState_(email);
}

function notifyAdminsAccess_(req) {
  try {
    const admins = getTeam_().filter(t => t.ativo && t.papel === 'admin').map(t => t.email);
    if (!admins.length) return;
    const url = portalUrl_();
    const link = d => `${url}?acao=decidir&req=${encodeURIComponent(req.id)}&d=${d}`;
    const btn = (href, label, bg) =>
      `<td style="padding:0 6px"><a href="${esc_(href)}" style="display:inline-block;background:${bg};color:#fff;padding:12px 26px;border-radius:8px;text-decoration:none;font-weight:700;font-size:15px">${label}</a></td>`;
    const html = `<p>Nova solicitação de acesso ao <b>${esc_(APP.NAME)}</b>:</p>
      <table style="border-collapse:collapse;font-size:14px;margin:8px 0 14px">
        <tr><td style="padding:4px 12px 4px 0;color:#64748b">Nome</td><td><b>${esc_(req.nome)}</b></td></tr>
        <tr><td style="padding:4px 12px 4px 0;color:#64748b">E-mail</td><td>${esc_(req.email)}</td></tr>
        <tr><td style="padding:4px 12px 4px 0;color:#64748b;vertical-align:top">Justificativa</td><td style="white-space:pre-wrap">${esc_(req.justificativa)}</td></tr>
        <tr><td style="padding:4px 12px 4px 0;color:#64748b">Solicitado em</td><td>${esc_(fmtTsBr_(req.criadoEm))}</td></tr>
      </table>
      ${url ? `<table role="presentation" style="border-collapse:collapse;margin:6px 0 10px"><tr>
          ${btn(link('aprovar'), '✔ Aprovar', '#16a34a')}${btn(link('reprovar'), '✖ Reprovar', '#dc2626')}
        </tr></table>
        <p style="font-size:12px;color:#64748b">O botão abre o Planner para você confirmar a decisão (proteção contra scanners de link). Solicitação ${esc_(req.id)}.</p>`
      : '<p style="color:#dc2626">Configure <b>Config.urlPortal</b> para habilitar os botões de aprovação.</p>'}`;
    MailApp.sendEmail({ to: admins.join(','), subject: `[${APP.NAME}] Solicitação de acesso: ${req.nome} (${req.email})`,
      htmlBody: shell_(html, ''), name: APP.NAME });
  } catch (e) { console.warn('notifyAdminsAccess_ falhou: ' + e); }
}

function apiListAccessRequests() {
  assertAdmin_(currentUser_());
  const all = readRequests_();
  return {
    pendentes: all.filter(r => r.status === 'pendente').sort((a, b) => a.criadoEm.localeCompare(b.criadoEm)),
    recentes: all.filter(r => r.status !== 'pendente').sort((a, b) => b.decididoEm.localeCompare(a.decididoEm)).slice(0, 20)
  };
}

function apiDecideAccess(id, decisao, motivo, papel) {
  const user = currentUser_();
  assertAdmin_(user);
  if (!['aprovar', 'reprovar'].includes(decisao)) throw new Error('Decisão inválida.');
  papel = papel === 'admin' ? 'admin' : 'membro';
  motivo = String(motivo || '').trim().slice(0, 500);

  const req = withLock_(() => {
    const sh = reqSheet_();
    const { rowIdx, r } = findRequest_(sh, id);
    if (r.status !== 'pendente') {
      throw new Error(`Solicitação já ${r.status === 'aprovado' ? 'aprovada' : 'reprovada'} por ${r.decididoPor} em ${fmtTsBr_(r.decididoEm)}.`);
    }
    const now = new Date();
    r.status = decisao === 'aprovar' ? 'aprovado' : 'reprovado';
    if (r.status === 'aprovado') upsertMember_(r.email, r.nome, papel);
    sh.getRange(rowIdx, RC.status + 1).setValue(r.status);
    sh.getRange(rowIdx, RC.decididoPor + 1, 1, 3).setValues([["'" + user.email, now, motivo ? "'" + motivo : '']]);
    Object.assign(r, { decididoPor: user.email, decididoEm: tsStr_(now), motivo });
    log_(user.email, r.status === 'aprovado' ? 'ACCESS_APPROVE' : 'ACCESS_REJECT', r.id,
      `${r.email}${r.status === 'aprovado' ? ' papel=' + papel : ''}${motivo ? ' motivo=' + motivo : ''}`);
    return r;
  });

  try {
    const url = portalUrl_();
    if (req.status === 'aprovado') {
      sendMail_(req.email, `[${APP.NAME}] Acesso aprovado`,
        `<p>Olá ${esc_(req.nome)}, seu acesso ao <b>${esc_(APP.NAME)}</b> foi <b style="color:#16a34a">aprovado</b>.</p>` +
        (motivo ? `<p style="white-space:pre-wrap;color:#334155">${esc_(motivo)}</p>` : ''), url);
    } else {
      sendMail_(req.email, `[${APP.NAME}] Solicitação de acesso reprovada`,
        `<p>Olá ${esc_(req.nome)}, sua solicitação de acesso ao <b>${esc_(APP.NAME)}</b> foi <b style="color:#dc2626">reprovada</b>.</p>` +
        (motivo ? `<p><b>Motivo:</b></p><p style="white-space:pre-wrap;color:#334155">${esc_(motivo)}</p>` : '') +
        `<p style="color:#64748b">Você poderá enviar uma nova solicitação após ${REQUEST_COOLDOWN_H}h.</p>`, '');
    }
  } catch (e) { console.warn('e-mail ao solicitante falhou: ' + e); }
  return req;
}

function fmtTsBr_(s) { return s ? `${s.slice(8, 10)}/${s.slice(5, 7)}/${s.slice(0, 4)} ${s.slice(11, 16)}` : '—'; }

/** Cria ou reativa membro na aba Equipe. Chamar dentro do lock. */
function upsertMember_(email, nome, papel) {
  _team = null;
  const sh = sheet_(APP.SH_TEAM);
  const m = getTeam_().find(t => t.email === email);
  if (m) {
    sh.getRange(m._row, 3, 1, 2).setValues([[papel || m.papel, true]]);
    if (!m.nome || m.nome === m.email) sh.getRange(m._row, 1).setValue("'" + nome);
  } else {
    sh.appendRow(["'" + nome, "'" + email, papel || 'membro', true, 'system']);
  }
  _team = null;
}

/* ── Membros (painel Admin) ── */

function apiListMembers() {
  const user = currentUser_();
  assertAdmin_(user);
  _team = null;
  return getTeam_().map(({ nome, email, papel, ativo }) => ({ nome, email, papel, ativo, eu: email === user.email }));
}

function apiUpdateMember(email, patch) {
  const user = currentUser_();
  assertAdmin_(user);
  email = String(email || '').toLowerCase();
  if (email === user.email) throw new Error('Você não pode alterar o próprio acesso (evita lockout). Peça a outro admin.');
  return withLock_(() => {
    _team = null;
    const m = getTeam_().find(t => t.email === email);
    if (!m) throw new Error('Membro não encontrado.');
    const papel = patch && 'papel' in patch ? (patch.papel === 'admin' ? 'admin' : 'membro') : m.papel;
    const ativo = patch && 'ativo' in patch ? !!patch.ativo : m.ativo;
    const adminsRestantes = getTeam_().filter(t => t.ativo && t.papel === 'admin' && t.email !== email).length +
      (ativo && papel === 'admin' ? 1 : 0);
    if (!adminsRestantes) throw new Error('Precisa existir ao menos um admin ativo.');
    sheet_(APP.SH_TEAM).getRange(m._row, 3, 1, 2).setValues([[papel, ativo]]);
    log_(user.email, 'MEMBER_UPDATE', email, `papel=${m.papel}→${papel} ativo=${m.ativo}→${ativo}`);
    _team = null;
    return { email, papel, ativo };
  });
}

function apiAddMember(nome, email, papel) {
  const user = currentUser_();
  assertAdmin_(user);
  email = String(email || '').trim().toLowerCase();
  nome = String(nome || '').trim().slice(0, 120);
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('E-mail inválido.');
  if (email.split('@')[1] !== ownerDomain_()) throw new Error(`Somente e-mails @${ownerDomain_()}.`);
  if (nome.length < 3) throw new Error('Informe o nome.');
  withLock_(() => {
    if (getTeam_().some(t => t.email === email && t.ativo)) throw new Error('Já é membro ativo.');
    upsertMember_(email, nome, papel === 'admin' ? 'admin' : 'membro');
    log_(user.email, 'MEMBER_ADD', email, `papel=${papel}`);
  });
  return apiListMembers();
}

/* ─────────────────────────── Importação (MS Planner) ─────────────────────────── */
// O cliente lê o .xlsx/.csv exportado do Planner (SheetJS), mapeia colunas/pessoas/status e envia
// lotes já normalizados. Aqui: valida tudo de novo (não confia no cliente), deduplica por origemId
// (idempotente: reimportar o mesmo arquivo não duplica) e grava o lote num único setValues.

function ensureTaskHeader_(sh) {
  const n = COLS.length;
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());
  const hdr = sh.getRange(1, 1, 1, n).getValues()[0];
  if (hdr[n - 1] !== COLS[n - 1]) sh.getRange(1, n).setValue(COLS[n - 1]).setFontWeight('bold');
}

function apiImportTasks(rows) {
  const user = currentUser_();
  assertAdmin_(user);
  if (!Array.isArray(rows) || !rows.length) throw new Error('Nada para importar.');
  if (rows.length > 300) throw new Error('Máximo de 300 linhas por lote.');

  return withLock_(() => {
    const sh = sheet_(APP.SH_TASKS);
    ensureTaskHeader_(sh);
    const all = readTasks_(true);
    const seen = new Set(all.map(t => t.origemId).filter(Boolean));
    const ids = new Set(all.map(t => t.id));
    const team = new Set(getTeam_().filter(t => t.ativo).map(t => t.email));
    const now = new Date();
    const out = [];
    const res = { inseridas: 0, duplicadas: 0, erros: [] };
    const ckRows = [];

    rows.forEach((r, i) => {
      try {
        const origemId = String(r.origemId || '').trim().slice(0, 120);
        if (origemId && seen.has(origemId)) { res.duplicadas++; return; }
        const input = {};
        ['titulo', 'descricao', 'responsavel', 'categoria', 'prioridade', 'status', 'inicio', 'prazo', 'tags']
          .forEach(k => { if (r[k] != null) input[k] = r[k]; });
        if (input.inicio && input.prazo && input.inicio > input.prazo) delete input.inicio;
        const clean = validate_(input, true);
        let id; do { id = newId_(); } while (ids.has(id));
        const created = /^\d{4}-\d{2}-\d{2}/.test(String(r.criadoEm || '')) ? parseDate_(r.criadoEm) : now;
        const t = Object.assign(Object.fromEntries(COLS.map(k => [k, ''])), clean, {
          id, criadoPor: team.has(String(r.criadoPor || '').toLowerCase()) ? String(r.criadoPor).toLowerCase() : user.email,
          criadoEm: created, atualizadoEm: now, versao: 1, arquivado: false, origemId
        });
        if (t.status === APP.DONE) {
          t.concluidoEm = /^\d{4}-\d{2}-\d{2}/.test(String(r.concluidoEm || '')) ? parseDate_(r.concluidoEm) : now;
        }
        out.push(taskToRow_(t, true));
        if (Array.isArray(r.checklist)) {
          r.checklist.slice(0, MAX_CK_ITEMS).forEach((it, j) => {
            const texto = String((it && it.texto) || '').replace(/\s+/g, ' ').trim().slice(0, MAX_CK_TEXT);
            if (!texto) return;
            const feito = !!(it && it.feito);
            ckRows.push([newCkId_(), id, "'" + texto, feito, (j + 1) * 10, "'" + user.email, now,
              feito ? "'" + t.criadoPor : '', feito ? now : '', false]);
          });
        }
        ids.add(id); if (origemId) seen.add(origemId);
        res.inseridas++;
      } catch (e) {
        res.erros.push({ linha: r._linha || i + 1, titulo: String(r.titulo || '').slice(0, 80), erro: e.message });
      }
    });

    if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, COLS.length).setValues(out);
    if (ckRows.length) {
      const ks = ckSheet_();
      ks.getRange(ks.getLastRow() + 1, 1, ckRows.length, KCOLS.length).setValues(ckRows);
    }
    res.itensChecklist = ckRows.length;
    log_(user.email, 'IMPORT', '', `MS Planner: ${res.inseridas} inseridas, ${res.duplicadas} duplicadas, ${res.erros.length} erros, ${ckRows.length} itens de checklist`);
    return res;
  });
}


/* ─────────────────────────── Diagnóstico (rodar no editor) ─────────────────────────── */
/** Executa cada etapa da carga do portal e registra tempo/erro. Rode no editor e veja o Registro de execução. */
function diagnostico() {
  guardOwner_();
  const step = (nome, fn) => {
    const t = Date.now();
    try { const r = fn(); console.log(`OK   ${nome} (${Date.now() - t} ms)${r !== undefined ? ' → ' + r : ''}`); return true; }
    catch (e) { console.error(`FALHA ${nome}: ${e && e.stack || e}`); return false; }
  };
  step('Propriedade PLANNER_SSID', () => PropertiesService.getScriptProperties().getProperty(APP.PROP_SSID) || (() => { throw new Error('ausente — rode setup()'); })());
  step('Abrir planilha', () => ss_().getName());
  [APP.SH_TASKS, APP.SH_TEAM, APP.SH_CFG, APP.SH_LOG, APP.SH_COMMENTS, APP.SH_REQ, APP.SH_CHECK].forEach(n =>
    step(`Aba ${n}`, () => { const sh = sheet_(n); if (!sh) throw new Error('não existe — rode setup()'); return `${sh.getLastRow()} linhas × ${sh.getMaxColumns()} colunas`; }));
  step('Cabeçalho Atividades', () => {
    const h = sheet_(APP.SH_TASKS).getRange(1, 1, 1, COLS.length).getValues()[0];
    const diff = COLS.filter((c, i) => h[i] !== c);
    if (diff.length) throw new Error('divergente em: ' + diff.join(', ') + ' — rode setup()');
    return 'ok';
  });
  step('Equipe', () => { _team = null; const t = getTeam_(); return `${t.length} membros, ${t.filter(m => m.ativo && m.papel === 'admin').length} admin(s) ativo(s)`; });
  step('Seu e-mail na Equipe', () => {
    const me = Session.getEffectiveUser().getEmail().toLowerCase();
    const m = getTeam_().find(t => t.email === me);
    if (!m) throw new Error(me + ' não está na aba Equipe');
    return `${me} papel=${m.papel} ativo=${m.ativo}`;
  });
  step('readTasks_', () => readTasks_().length + ' atividades');
  step('readComments_', () => readComments_().length + ' comentários');
  step('readRequests_', () => readRequests_().length + ' solicitações');
  step('readChecklist_', () => readChecklist_().length + ' itens de checklist');
  step('apiBootstrap (como owner)', () => {
    const b = apiBootstrap();
    return `${b.tasks.length} tarefas, payload ${JSON.stringify(b).length} bytes`;
  });
  step('URL do portal', () => portalUrl_() || '(vazia — preencha Config.urlPortal)');
  step('Quota de e-mail restante', () => MailApp.getRemainingDailyQuota());
}


/* ─────────────────────────── Lista de verificação ─────────────────────────── */
// Aba própria, um item por linha, gravação imediata por item (como no MS Planner).
// Desacoplada da atividade: marcar item não altera versao → sem CONFLITO com o formulário aberto;
// registra quem concluiu e quando. Exclusão = soft delete. Qualquer membro ativo pode editar.

const KCOLS = ['id', 'taskId', 'texto', 'feito', 'ordem', 'criadoPor', 'criadoEm', 'feitoPor', 'feitoEm', 'excluido'];
const KC = Object.fromEntries(KCOLS.map((k, i) => [k, i]));
const MAX_CK_ITEMS = 50;
const MAX_CK_TEXT = 300;

function ckSheet_() { return sheet_(APP.SH_CHECK) || ensureSheet_(ss_(), APP.SH_CHECK, KCOLS); }
function newCkId_() { return 'K-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase(); }
const isTrue_ = v => v === true || String(v).toUpperCase() === 'TRUE';

function rowToItem_(r) {
  return {
    id: String(r[KC.id]), taskId: String(r[KC.taskId]), texto: String(r[KC.texto]),
    feito: isTrue_(r[KC.feito]), ordem: Number(r[KC.ordem]) || 0,
    criadoPor: String(r[KC.criadoPor] || '').toLowerCase(), criadoEm: tsStr_(r[KC.criadoEm]),
    feitoPor: String(r[KC.feitoPor] || '').toLowerCase(), feitoEm: tsStr_(r[KC.feitoEm]),
    excluido: isTrue_(r[KC.excluido])
  };
}

function readChecklist_() {
  const sh = sheet_(APP.SH_CHECK);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, KCOLS.length).getValues().filter(r => r[KC.id]).map(rowToItem_);
}

function findItem_(sh, id) {
  const n = sh.getLastRow() - 1;
  const cell = n > 0 && id ? sh.getRange(2, 1, n, 1).createTextFinder(String(id)).matchEntireCell(true).findNext() : null;
  if (!cell) throw new Error('Item da lista não encontrado.');
  return { rowIdx: cell.getRow(), it: rowToItem_(sh.getRange(cell.getRow(), 1, 1, KCOLS.length).getValues()[0]) };
}

const itemToClient_ = i => ({ id: i.id, taskId: i.taskId, texto: i.texto, feito: i.feito, ordem: i.ordem, feitoPor: i.feitoPor, feitoEm: i.feitoEm });

function ckText_(t) {
  const s = String(t == null ? '' : t).replace(/\s+/g, ' ').trim().slice(0, MAX_CK_TEXT);
  if (!s) throw new Error('Item vazio.');
  return s;
}

function openTask_(taskId) {
  const { task } = findTask_(sheet_(APP.SH_TASKS), taskId);
  if (task.arquivado) throw new Error('Atividade arquivada: lista de verificação bloqueada.');
  return task;
}

function apiListChecklist(taskId) {
  currentUser_();
  return readChecklist_().filter(i => i.taskId === String(taskId) && !i.excluido)
    .sort((a, b) => a.ordem - b.ordem).map(itemToClient_);
}

/** Aceita string, [string] ou [{texto, feito}]. Usado no "Adicionar item" e ao criar atividade com itens. */
function apiAddChecklistItems(taskId, itens) {
  const user = currentUser_();
  const list = (Array.isArray(itens) ? itens : [itens]).map(x =>
    typeof x === 'object' && x ? { texto: ckText_(x.texto), feito: !!x.feito } : { texto: ckText_(x), feito: false });
  if (!list.length) throw new Error('Nada para adicionar.');
  return withLock_(() => {
    const task = openTask_(taskId);
    const cur = readChecklist_().filter(i => i.taskId === task.id && !i.excluido);
    if (cur.length + list.length > MAX_CK_ITEMS) throw new Error(`Máximo de ${MAX_CK_ITEMS} itens por atividade.`);
    let ordem = cur.reduce((m, i) => Math.max(m, i.ordem), 0);
    const now = new Date();
    const items = list.map(x => {
      ordem += 10;
      return { id: newCkId_(), taskId: task.id, texto: x.texto, feito: x.feito, ordem,
        feitoPor: x.feito ? user.email : '', feitoEm: x.feito ? tsStr_(now) : '' };
    });
    const sh = ckSheet_();
    sh.getRange(sh.getLastRow() + 1, 1, items.length, KCOLS.length).setValues(items.map(i =>
      [i.id, i.taskId, "'" + i.texto, i.feito, i.ordem, "'" + user.email, now, i.feito ? "'" + user.email : '', i.feito ? now : '', false]));
    log_(user.email, 'CHECK_ADD', task.id, items.map(i => i.texto).join(' | ').slice(0, 500));
    return items;
  });
}

function apiUpdateChecklistItem(id, patch) {
  const user = currentUser_();
  patch = patch || {};
  return withLock_(() => {
    const sh = ckSheet_();
    const { rowIdx, it } = findItem_(sh, id);
    if (it.excluido) throw new Error('Item removido por outro usuário.');
    openTask_(it.taskId);
    if ('texto' in patch) {
      const texto = ckText_(patch.texto);
      if (texto !== it.texto) {
        sh.getRange(rowIdx, KC.texto + 1).setValue("'" + texto);
        log_(user.email, 'CHECK_EDIT', it.taskId, `${id}: ${it.texto} → ${texto}`.slice(0, 500));
        it.texto = texto;
      }
    }
    if ('feito' in patch && !!patch.feito !== it.feito) {
      const now = new Date();
      it.feito = !!patch.feito;
      it.feitoPor = it.feito ? user.email : '';
      it.feitoEm = it.feito ? tsStr_(now) : '';
      sh.getRange(rowIdx, KC.feito + 1).setValue(it.feito);
      sh.getRange(rowIdx, KC.feitoPor + 1, 1, 2).setValues([[it.feito ? "'" + user.email : '', it.feito ? now : '']]);
      log_(user.email, it.feito ? 'CHECK_DONE' : 'CHECK_UNDONE', it.taskId, `${id}: ${it.texto}`.slice(0, 300));
    }
    return itemToClient_(it);
  });
}

function apiDeleteChecklistItem(id) {
  const user = currentUser_();
  withLock_(() => {
    const sh = ckSheet_();
    const { rowIdx, it } = findItem_(sh, id);
    if (it.excluido) return;
    openTask_(it.taskId);
    sh.getRange(rowIdx, KC.excluido + 1).setValue(true);
    log_(user.email, 'CHECK_DELETE', it.taskId, `${id}: ${it.texto}`.slice(0, 300));
  });
  return true;
}

/** Reordena pela lista de IDs recebida (drag & drop). IDs desconhecidos são ignorados. */
function apiReorderChecklist(taskId, ids) {
  const user = currentUser_();
  if (!Array.isArray(ids)) throw new Error('Ordem inválida.');
  return withLock_(() => {
    const task = openTask_(taskId);
    const sh = ckSheet_();
    const n = sh.getLastRow() - 1;
    if (n < 1) return [];
    const vals = sh.getRange(2, 1, n, KCOLS.length).getValues();
    const rowOf = {};
    vals.forEach((r, i) => { if (String(r[KC.taskId]) === task.id && !isTrue_(r[KC.excluido])) rowOf[String(r[KC.id])] = i + 2; });
    const order = ids.map(String).filter(id => rowOf[id]);
    Object.keys(rowOf).forEach(id => { if (!order.includes(id)) order.push(id); }); // itens novos de outros usuários vão ao fim
    order.forEach((id, i) => sh.getRange(rowOf[id], KC.ordem + 1).setValue((i + 1) * 10));
    log_(user.email, 'CHECK_REORDER', task.id, `${order.length} itens`);
    return apiListChecklist(task.id);
  });
}
