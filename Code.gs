/**
 * PLATEFORME DE FORMATION (type MOOC) — API Google Apps Script
 * - Inscription libre : pseudonyme + mot de passe.
 * - L'administrateur crée des cours ; chaque cours a des documents, un pré-test, un post-test et une discussion.
 * - Tout est dans le Google Sheet (un onglet par table) : modifiable à la main.
 * - L'utilisateur peut modifier son pseudo / mot de passe et ses propres messages.
 * Déploiement : Déployer > Application Web > Exécuter en tant que moi, accès : tout le monde.
 * Exécuter setup() une fois pour créer les onglets et autoriser Drive.
 */
const ADMIN_PASSWORD = 'Silo@2000'; // mot de passe de l'espace administrateur
const SHEET_ID = '1hZ7v3_5K2qhD5N4YRmt3C96OsxOIk0iuQdQUwYSrwQc';
const TABS = {
  // Users : l'administrateur voit et modifie pseudo et motdepasse directement ici (id = identifiant stable, ne pas modifier)
  Users: ['id', 'pseudo', 'motdepasse', 'jeton', 'cree'],
  Cours: ['id', 'titre', 'description', 'actif'],
  Documents: ['id', 'cours', 'titre', 'url'],
  // options = JSON ["A","B"] ; bonnes = JSON [0] (index à partir de 0, plusieurs = choix multiple)
  // type : single | multi | chart | rpt | dates | match ; donnees = JSON (contexte, tableaux, dates, appariements)
  Questions: ['id', 'cours', 'phase', 'type', 'question', 'options', 'bonnes', 'donnees', 'explication'],
  Blocs: ['id', 'cours', 'titre', 'texte'],
  Inscriptions: ['uid', 'pseudo', 'cours', 'date'],
  Resultats: ['date', 'uid', 'pseudo', 'cours', 'phase', 'score', 'total'],
  Messages: ['id', 'date', 'cours', 'uid', 'pseudo', 'texte']
};
const PHASES = ['Pré-test', 'Post-test'];

const API = { login, me, profil, updateProfile, enroll, unenroll, course, startTest, submitTest,
  getMessages, postMessage, editMessage, deleteMessage,
  adminLogin, adminOverview, adminSaveCourse, adminAddDoc, adminAddQuestion,
  adminCourseDetail, adminDelete, adminAddBloc, adminImporterMPE };

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    if (!API.hasOwnProperty(b.action)) return out_({ ok: false, message: 'Action inconnue.' });
    return out_(API[b.action].apply(null, b.args || []));
  } catch (err) { return out_({ ok: false, message: err.message }); }
}
function doGet() { return out_({ ok: true, service: 'plateforme-formation' }); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function setup() { Object.keys(TABS).forEach(sh_); txt_(); DriveApp.getRootFolder(); }
/** Colonnes pseudo / mot de passe en texte brut : « 0123 » ne perd pas son zéro. */
function txt_() { sh_('Users').getRange('A:D').setNumberFormat('@'); }

/* ---------- Accès au Sheet ---------- */
function sh_(n) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(TABS[n]); s.setFrozenRows(1); if (n === 'Users') s.getRange('A:D').setNumberFormat('@'); }
  else { const h = hdr_(s); TABS[n].forEach(k => { if (h.indexOf(k) < 0) { s.getRange(1, h.length + 1).setValue(k); h.push(k); } }); }
  return s;
}
function hdr_(s) { return s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0]; }
function rows_(n) {
  const v = sh_(n).getDataRange().getValues(), h = v.shift();
  return v.map((r, i) => { const o = { _r: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; });
}
function add_(n, o) { const s = sh_(n); s.appendRow(hdr_(s).map(k => o[k] === undefined ? '' : o[k])); }
function set_(n, r, k, v) { const s = sh_(n); s.getRange(r, hdr_(s).indexOf(k) + 1).setValue(v); }
const S = v => String(v);
const id_ = p => p + Utilities.getUuid().slice(0, 8);

/* ---------- Comptes ---------- */
function findUser_(p) { return rows_('Users').filter(u => S(u.pseudo).toLowerCase() === S(p).trim().toLowerCase())[0]; }
function auth_(tok) {
  const u = tok && rows_('Users').filter(x => S(x.jeton) === S(tok))[0];
  if (!u) throw new Error('Session expirée : reconnectez-vous.');
  return u;
}
// Critères usuels : pseudonyme de 3 à 30 caractères (lettres, chiffres, . _ -), sans espace ; mot de passe de 6 caractères minimum
const PSEUDO_RE = /^[\p{L}\p{N}._-]{3,30}$/u;
function checkPseudo_(p) { return PSEUDO_RE.test(p) ? '' : 'Pseudonyme : 3 à 30 caractères (lettres, chiffres, point, tiret, souligné), sans espace.'; }
function checkPw_(pw) { return S(pw).length >= 6 ? '' : 'Mot de passe : 6 caractères minimum.'; }
/** Première utilisation : un pseudonyme inconnu crée le compte tout de suite (aucune validation de l'administrateur). */
function login(pseudo, pw) {
  pseudo = S(pseudo).trim(); pw = S(pw);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const u = findUser_(pseudo);
    if (u) {
      if (S(u.motdepasse) !== pw) return { ok: false, message: 'Mot de passe incorrect.' };
      if (!u.jeton) { u.jeton = 't' + Utilities.getUuid(); set_('Users', u._r, 'jeton', u.jeton); }
      return { ok: true, token: S(u.jeton), pseudo: S(u.pseudo) };
    }
    const err = checkPseudo_(pseudo) || checkPw_(pw);
    if (err) return { ok: false, message: err };
    const jeton = 't' + Utilities.getUuid();
    add_('Users', { id: id_('u'), pseudo, motdepasse: pw, jeton, cree: new Date() });
    return { ok: true, token: jeton, pseudo, created: true };
  } finally { lock.releaseLock(); }
}
/** Profil : infos du compte, cours suivis et résultats. */
function profil(tok) {
  const u = auth_(tok), res = rows_('Resultats').filter(r => S(r.uid) === S(u.id)), cs = rows_('Cours');
  const cours = rows_('Inscriptions').filter(i => S(i.uid) === S(u.id)).map(i => {
    const c = cs.filter(x => S(x.id) === S(i.cours))[0], r = ph => { const x = res.filter(y => S(y.cours) === S(i.cours) && y.phase === ph)[0]; return x ? x.score + '/' + x.total : null; };
    return { id: S(i.cours), titre: c ? c.titre : '(cours supprimé)', pre: r('Pré-test'), post: r('Post-test') };
  });
  return { ok: true, pseudo: S(u.pseudo), cree: u.cree, cours };
}
/** Modifier son pseudonyme et/ou son mot de passe (laisser vide pour ne pas changer). */
function updateProfile(tok, newPseudo, newPw) {
  const u = auth_(tok); newPseudo = S(newPseudo || '').trim();
  if (newPseudo && newPseudo !== S(u.pseudo)) {
    const err = checkPseudo_(newPseudo); if (err) return { ok: false, message: err };
    const ex = findUser_(newPseudo); if (ex && ex._r !== u._r) return { ok: false, message: 'Ce pseudonyme est déjà pris.' };
    set_('Users', u._r, 'pseudo', newPseudo);
  }
  if (newPw) { const err = checkPw_(newPw); if (err) return { ok: false, message: err }; set_('Users', u._r, 'motdepasse', S(newPw)); }
  return { ok: true, pseudo: newPseudo || S(u.pseudo) };
}

/* ---------- Cours et inscriptions ---------- */
function activeCourse_(cid) {
  const c = rows_('Cours').filter(x => S(x.id) === S(cid))[0];
  if (!c || c.actif === false || S(c.actif).toUpperCase() === 'FALSE') throw new Error('Cours introuvable.');
  return c;
}
function enrolled_(uid, cid) { return rows_('Inscriptions').some(i => S(i.uid) === S(uid) && S(i.cours) === S(cid)); }
function needEnrolled_(u, cid) { if (!enrolled_(u.id, cid)) throw new Error('Inscrivez-vous d’abord à ce cours.'); }

function me(tok) {
  const u = auth_(tok), mine = rows_('Inscriptions').filter(i => S(i.uid) === S(u.id)).map(i => S(i.cours));
  const cours = rows_('Cours').filter(c => c.actif !== false && S(c.actif).toUpperCase() !== 'FALSE')
    .map(c => ({ id: S(c.id), titre: c.titre, description: c.description, inscrit: mine.indexOf(S(c.id)) >= 0 }));
  return { ok: true, pseudo: S(u.pseudo), cours };
}
function enroll(tok, cid) {
  const u = auth_(tok); activeCourse_(cid);
  if (!enrolled_(u.id, cid)) add_('Inscriptions', { uid: u.id, pseudo: u.pseudo, cours: S(cid), date: new Date() });
  return { ok: true };
}
function unenroll(tok, cid) {
  const u = auth_(tok);
  rows_('Inscriptions').filter(i => S(i.uid) === S(u.id) && S(i.cours) === S(cid)).reverse()
    .forEach(i => sh_('Inscriptions').deleteRow(i._r));
  return { ok: true };
}
function course(tok, cid) {
  const u = auth_(tok), c = activeCourse_(cid); needEnrolled_(u, cid);
  const qs = rows_('Questions').filter(q => S(q.cours) === S(cid));
  const res = {};
  rows_('Resultats').filter(r => S(r.uid) === S(u.id) && S(r.cours) === S(cid)).forEach(r => res[r.phase] = { score: r.score, total: r.total });
  return { ok: true, id: S(c.id), titre: c.titre, description: c.description,
    blocs: rows_('Blocs').filter(b => S(b.cours) === S(cid)).map(b => ({ titre: b.titre, texte: b.texte })),
    documents: rows_('Documents').filter(d => S(d.cours) === S(cid)).map(d => ({ titre: d.titre, url: d.url })),
    nbQuestions: { 'Pré-test': qs.filter(q => q.phase === 'Pré-test').length, 'Post-test': qs.filter(q => q.phase === 'Post-test').length },
    resultats: res };
}

/* ---------- Tests ---------- */
function shuffle_(a) { for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; }
function questions_(cid, phase) { return rows_('Questions').filter(q => S(q.cours) === S(cid) && q.phase === phase); }
function done_(uid, cid, phase) { return rows_('Resultats').some(r => S(r.uid) === S(uid) && S(r.cours) === S(cid) && r.phase === phase); }

function typeOf_(q) { return q.type || (JSON.parse(q.bonnes).length > 1 ? 'multi' : 'single'); }
function dat_(q) { return q.donnees ? JSON.parse(q.donnees) : {}; }
function qOut_(q) {
  const t = typeOf_(q), d = dat_(q);
  const o = { id: S(q.id), type: t, q: q.question, ctx: d.ctx, ctxTable: d.ctxTable, chartData: d.chartData, rptDisease: d.rptDisease };
  if (t === 'dates') o.dates = d.dates.map(x => ({ k: x.k, label: x.label, o: x.o }));
  else if (t === 'match') { o.pre = d.pre; o.tubes = d.tubes; o.rows = d.rows.map(r => r.label); }
  else {
    const opts = JSON.parse(q.options).map((x, i) => ({ i, t: typeof x === 'string' ? x : x.t, chart: x.chart, rpt: x.rpt }));
    o.multi = t === 'multi'; o.o = d.fixe ? opts : shuffle_(opts);
  }
  return o;
}
function grade_(q, rep) {
  const t = typeOf_(q), d = dat_(q);
  if (t === 'dates') return !!rep && d.dates.every(x => S(rep[x.k]) === S(x.a));
  if (t === 'match') return Array.isArray(rep) && d.rows.every((r, i) => rep[i] && S(rep[i][0]) === S(r.p) && S(rep[i][1]) === S(r.t));
  const good = JSON.parse(q.bonnes).map(Number).sort(), got = (Array.isArray(rep) ? rep : []).map(Number).sort();
  return good.length === got.length && good.every((v, i) => v === got[i]);
}
function bonnes_(q) {
  const t = typeOf_(q), d = dat_(q);
  if (t === 'dates') return d.dates.map(x => x.label + ' : ' + x.o[x.a]);
  if (t === 'match') return d.rows.map(r => r.label + ' : ' + d.pre[r.p] + ' → ' + d.tubes[r.t].t);
  if (t === 'rpt') return [];
  return JSON.parse(q.bonnes).map(i => { const x = JSON.parse(q.options)[i]; return typeof x === 'string' ? x : x.t; });
}
function startTest(tok, cid, phase) {
  const u = auth_(tok); activeCourse_(cid); needEnrolled_(u, cid);
  if (PHASES.indexOf(phase) < 0) return { ok: false, message: 'Phase inconnue.' };
  if (done_(u.id, cid, phase)) return { ok: false, message: 'Vous avez déjà passé ce test.' };
  if (phase === 'Post-test' && !done_(u.id, cid, 'Pré-test')) return { ok: false, message: 'Passez d’abord le pré-test.' };
  const qs = questions_(cid, phase);
  if (!qs.length) return { ok: false, message: 'Aucune question n’a encore été ajoutée.' };
  return { ok: true, pseudo: S(u.pseudo), limitSec: 25 * 60, questions: shuffle_(qs.map(qOut_)) };
}
/** reponses : { idQuestion: réponse } — les questions non renseignées comptent comme fausses. */
function submitTest(tok, cid, phase, reponses) {
  const u = auth_(tok); needEnrolled_(u, cid);
  if (done_(u.id, cid, phase)) return { ok: false, message: 'Vous avez déjà passé ce test.' };
  if (phase === 'Post-test' && !done_(u.id, cid, 'Pré-test')) return { ok: false, message: 'Passez d’abord le pré-test.' };
  const qs = questions_(cid, phase); let score = 0;
  const corr = qs.map(q => {
    const ok = grade_(q, (reponses || {})[S(q.id)]); if (ok) score++;
    return { q: q.question, ok, bonnes: bonnes_(q), explication: q.explication };
  });
  add_('Resultats', { date: new Date(), uid: u.id, pseudo: u.pseudo, cours: S(cid), phase, score, total: qs.length });
  return { ok: true, pseudo: S(u.pseudo), score, total: qs.length, corrige: corr };
}

/* ---------- Discussion ---------- */
function getMessages(tok, cid) {
  const u = auth_(tok); needEnrolled_(u, cid);
  const names = {}; rows_('Users').forEach(x => names[S(x.id)] = S(x.pseudo));
  const m = rows_('Messages').filter(x => S(x.cours) === S(cid)).slice(-100)
    .map(x => ({ id: S(x.id), date: x.date, pseudo: names[S(x.uid)] || S(x.pseudo), texte: x.texte, mine: S(x.uid) === S(u.id) }));
  return { ok: true, messages: m };
}
function postMessage(tok, cid, texte) {
  const u = auth_(tok); needEnrolled_(u, cid);
  texte = S(texte || '').trim().slice(0, 1000);
  if (!texte) return { ok: false, message: 'Message vide.' };
  add_('Messages', { id: id_('m'), date: new Date(), cours: S(cid), uid: u.id, pseudo: u.pseudo, texte });
  return { ok: true };
}
function ownMessage_(u, mid) {
  const m = rows_('Messages').filter(x => S(x.id) === S(mid) && S(x.uid) === S(u.id))[0];
  if (!m) throw new Error('Message introuvable.'); return m;
}
function editMessage(tok, mid, texte) { const m = ownMessage_(auth_(tok), mid); set_('Messages', m._r, 'texte', S(texte).trim().slice(0, 1000)); return { ok: true }; }
function deleteMessage(tok, mid) { const m = ownMessage_(auth_(tok), mid); sh_('Messages').deleteRow(m._r); return { ok: true }; }

/* ---------- Administration ---------- */
function admin_(pw) { if (ADMIN_PASSWORD === 'CHANGEZ-MOI') throw new Error('Définissez ADMIN_PASSWORD dans Code.gs puis redéployez.'); if (S(pw) !== ADMIN_PASSWORD) throw new Error('Mot de passe administrateur incorrect.'); }
function adminLogin(pw) { admin_(pw); return { ok: true }; }
function adminOverview(pw) {
  admin_(pw);
  const ins = rows_('Inscriptions'), res = rows_('Resultats'), qs = rows_('Questions'), docs = rows_('Documents');
  const moy = (cid, ph) => { const r = res.filter(x => S(x.cours) === cid && x.phase === ph && x.total > 0); return r.length ? Math.round(r.reduce((s, x) => s + 100 * x.score / x.total, 0) / r.length) : null; };
  return { ok: true, cours: rows_('Cours').map(c => { const id = S(c.id); return {
    id, titre: c.titre, actif: !(c.actif === false || S(c.actif).toUpperCase() === 'FALSE'),
    inscrits: ins.filter(i => S(i.cours) === id).length,
    documents: docs.filter(d => S(d.cours) === id).length,
    pre: qs.filter(q => S(q.cours) === id && q.phase === 'Pré-test').length,
    post: qs.filter(q => S(q.cours) === id && q.phase === 'Post-test').length,
    moyPre: moy(id, 'Pré-test'), moyPost: moy(id, 'Post-test') }; }) };
}
function adminSaveCourse(pw, c) {
  admin_(pw);
  const titre = S(c && c.titre).trim().slice(0, 120); if (!titre) return { ok: false, message: 'Titre requis.' };
  const desc = S(c.description || '').slice(0, 500);
  const ex = c.id && rows_('Cours').filter(x => S(x.id) === S(c.id))[0];
  if (ex) { set_('Cours', ex._r, 'titre', titre); set_('Cours', ex._r, 'description', desc); set_('Cours', ex._r, 'actif', c.actif !== false); return { ok: true, id: S(c.id) }; }
  const id = id_('c'); add_('Cours', { id, titre, description: desc, actif: true }); return { ok: true, id };
}
/** doc : { titre, url } ou { titre, fichier: { nom, mime, b64 } } (déposé dans Drive, lisible par lien). */
function adminAddDoc(pw, cid, doc) {
  admin_(pw); activeCourse_(cid);
  let url = S(doc.url || '').trim();
  if (doc.fichier) {
    const f = DriveApp.createFile(Utilities.newBlob(Utilities.base64Decode(doc.fichier.b64), doc.fichier.mime, doc.fichier.nom));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); url = f.getUrl();
  }
  if (!url) return { ok: false, message: 'Indiquez un lien ou un fichier.' };
  add_('Documents', { id: id_('d'), cours: S(cid), titre: S(doc.titre || (doc.fichier && doc.fichier.nom) || 'Document').slice(0, 120), url });
  return { ok: true };
}
function adminAddQuestion(pw, cid, q) {
  admin_(pw); activeCourse_(cid);
  const opts = (q.options || []).map(x => S(x).trim()).filter(Boolean);
  const bonnes = (q.bonnes || []).map(Number).filter(i => i >= 0 && i < opts.length);
  if (PHASES.indexOf(q.phase) < 0 || !S(q.question).trim() || opts.length < 2 || !bonnes.length) return { ok: false, message: 'Question, au moins 2 choix et une bonne réponse requis.' };
  add_('Questions', { id: id_('q'), cours: S(cid), phase: q.phase, type: bonnes.length > 1 ? 'multi' : 'single', question: S(q.question).trim(), options: JSON.stringify(opts), bonnes: JSON.stringify(bonnes), donnees: '{}', explication: S(q.explication || '') });
  return { ok: true };
}
/** Page d'administration d'un cours : tout son contenu et les résultats (pseudonyme = code de travail). */
function adminCourseDetail(pw, cid) {
  admin_(pw);
  const c = rows_('Cours').filter(x => S(x.id) === S(cid))[0]; if (!c) return { ok: false, message: 'Cours introuvable.' };
  const mine = r => S(r.cours) === S(cid);
  return { ok: true,
    cours: { id: S(c.id), titre: c.titre, description: c.description, actif: !(c.actif === false || S(c.actif).toUpperCase() === 'FALSE') },
    documents: rows_('Documents').filter(mine).map(d => ({ id: S(d.id), titre: d.titre, url: d.url })),
    blocs: rows_('Blocs').filter(mine).map(b => ({ id: S(b.id), titre: b.titre, texte: b.texte })),
    questions: rows_('Questions').filter(mine).map(q => ({ id: S(q.id), phase: q.phase, type: typeOf_(q), question: q.question })),
    resultats: rows_('Resultats').filter(mine).map(r => ({ pseudo: r.pseudo, phase: r.phase, score: r.score, total: r.total, date: r.date })) };
}
function adminDelete(pw, table, id) {
  admin_(pw);
  if (['Documents', 'Questions', 'Blocs'].indexOf(table) < 0) return { ok: false, message: 'Table non autorisée.' };
  const r = rows_(table).filter(x => S(x.id) === S(id))[0]; if (r) sh_(table).deleteRow(r._r);
  return { ok: true };
}
/** Élément de texte affiché sur la page du cours (consigne, résumé, objectifs…). */
function adminAddBloc(pw, cid, titre, texte) {
  admin_(pw); activeCourse_(cid);
  if (!S(texte).trim()) return { ok: false, message: 'Texte requis.' };
  add_('Blocs', { id: id_('b'), cours: S(cid), titre: S(titre || '').slice(0, 120), texte: S(texte).slice(0, 5000) });
  return { ok: true };
}
