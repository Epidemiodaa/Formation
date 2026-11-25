/**
 * PLATEFORME DE FORMATION (type MOOC) — API Google Apps Script
 * - Inscription libre : pseudonyme + mot de passe.
 * - L'administrateur crée des cours ; chaque cours a des documents, un pré-test, un post-test et une discussion.
 * - Tout est dans le Google Sheet (un onglet par table) : modifiable à la main.
 * - L'utilisateur peut modifier son pseudo / mot de passe et ses propres messages.
 * Déploiement : Déployer > Application Web > Exécuter en tant que moi, accès : tout le monde.
 * Exécuter setup() une fois pour créer les onglets et autoriser Drive.
 */
const ADMIN_PASSWORD = 'CHANGEZ-MOI'; // <<< à modifier avant déploiement
const SHEET_ID = '1hZ7v3_5K2qhD5N4YRmt3C96OsxOIk0iuQdQUwYSrwQc';
const TABS = {
  Users: ['pseudo', 'sel', 'hash', 'jeton', 'cree'],
  Cours: ['id', 'titre', 'description', 'actif'],
  Documents: ['id', 'cours', 'titre', 'url'],
  // options = JSON ["A","B"] ; bonnes = JSON [0] (index à partir de 0, plusieurs = choix multiple)
  Questions: ['id', 'cours', 'phase', 'question', 'options', 'bonnes', 'explication'],
  Inscriptions: ['pseudo', 'cours', 'date'],
  Resultats: ['date', 'pseudo', 'cours', 'phase', 'score', 'total'],
  Messages: ['id', 'date', 'cours', 'pseudo', 'texte']
};
const PHASES = ['Pré-test', 'Post-test'];

const API = { register, login, me, updateProfile, enroll, unenroll, course, startTest, submitTest,
  getMessages, postMessage, editMessage, deleteMessage,
  adminLogin, adminOverview, adminSaveCourse, adminAddDoc, adminAddQuestion };

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    if (!API.hasOwnProperty(b.action)) return out_({ ok: false, message: 'Action inconnue.' });
    return out_(API[b.action].apply(null, b.args || []));
  } catch (err) { return out_({ ok: false, message: err.message }); }
}
function doGet() { return out_({ ok: true, service: 'plateforme-formation' }); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function setup() { Object.keys(TABS).forEach(sh_); DriveApp.getRootFolder(); }

/* ---------- Accès au Sheet ---------- */
function sh_(n) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(TABS[n]); s.setFrozenRows(1); }
  return s;
}
function rows_(n) {
  const v = sh_(n).getDataRange().getValues(), h = v.shift();
  return v.map((r, i) => { const o = { _r: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; });
}
function add_(n, o) { sh_(n).appendRow(TABS[n].map(k => o[k] === undefined ? '' : o[k])); }
function set_(n, r, k, v) { sh_(n).getRange(r, TABS[n].indexOf(k) + 1).setValue(v); }
const S = v => String(v);
const id_ = p => p + Utilities.getUuid().slice(0, 8);

/* ---------- Comptes ---------- */
function hash_(sel, pw) {
  return 'h:' + Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sel + pw)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function findUser_(p) { return rows_('Users').filter(u => S(u.pseudo).toLowerCase() === S(p).trim().toLowerCase())[0]; }
function auth_(tok) {
  const u = tok && rows_('Users').filter(x => S(x.jeton) === S(tok))[0];
  if (!u) throw new Error('Session expirée : reconnectez-vous.');
  return u;
}
function validPseudo_(p) {
  p = S(p).trim();
  if (p.length < 3 || p.length > 30) throw new Error('Le pseudonyme doit avoir 3 à 30 caractères.');
  return p;
}
function session_(u) { return { ok: true, token: u.jeton, pseudo: u.pseudo }; }
function register(pseudo, pw) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    pseudo = validPseudo_(pseudo);
    if (S(pw).length < 4) return { ok: false, message: 'Mot de passe : 4 caractères minimum.' };
    if (findUser_(pseudo)) return { ok: false, message: 'Ce pseudonyme est déjà pris.' };
    const sel = Utilities.getUuid(), jeton = 't' + Utilities.getUuid();
    add_('Users', { pseudo, sel, hash: hash_(sel, pw), jeton, cree: new Date() });
    return { ok: true, token: jeton, pseudo };
  } finally { lock.releaseLock(); }
}
function login(pseudo, pw) {
  const u = findUser_(pseudo);
  if (!u || u.hash !== hash_(u.sel, pw)) return { ok: false, message: 'Pseudonyme ou mot de passe incorrect.' };
  return session_(u);
}
/** Modifier son pseudonyme et/ou son mot de passe (laisser vide pour ne pas changer). */
function updateProfile(tok, newPseudo, newPw) {
  const u = auth_(tok);
  if (newPseudo && S(newPseudo).trim() !== S(u.pseudo)) {
    newPseudo = validPseudo_(newPseudo);
    const ex = findUser_(newPseudo);
    if (ex && ex._r !== u._r) return { ok: false, message: 'Ce pseudonyme est déjà pris.' };
    ['Inscriptions', 'Resultats', 'Messages'].forEach(n => rows_(n).forEach(r => { if (S(r.pseudo) === S(u.pseudo)) set_(n, r._r, 'pseudo', newPseudo); }));
    set_('Users', u._r, 'pseudo', newPseudo);
  }
  if (newPw) {
    if (S(newPw).length < 4) return { ok: false, message: 'Mot de passe : 4 caractères minimum.' };
    set_('Users', u._r, 'hash', hash_(u.sel, newPw));
  }
  return { ok: true, pseudo: newPseudo || u.pseudo };
}

/* ---------- Cours et inscriptions ---------- */
function activeCourse_(cid) {
  const c = rows_('Cours').filter(x => S(x.id) === S(cid))[0];
  if (!c || c.actif === false || S(c.actif).toUpperCase() === 'FALSE') throw new Error('Cours introuvable.');
  return c;
}
function enrolled_(pseudo, cid) { return rows_('Inscriptions').some(i => S(i.pseudo) === S(pseudo) && S(i.cours) === S(cid)); }
function needEnrolled_(u, cid) { if (!enrolled_(u.pseudo, cid)) throw new Error('Inscrivez-vous d’abord à ce cours.'); }

function me(tok) {
  const u = auth_(tok), mine = rows_('Inscriptions').filter(i => S(i.pseudo) === S(u.pseudo)).map(i => S(i.cours));
  const cours = rows_('Cours').filter(c => c.actif !== false && S(c.actif).toUpperCase() !== 'FALSE')
    .map(c => ({ id: S(c.id), titre: c.titre, description: c.description, inscrit: mine.indexOf(S(c.id)) >= 0 }));
  return { ok: true, pseudo: u.pseudo, cours };
}
function enroll(tok, cid) {
  const u = auth_(tok); activeCourse_(cid);
  if (!enrolled_(u.pseudo, cid)) add_('Inscriptions', { pseudo: u.pseudo, cours: S(cid), date: new Date() });
  return { ok: true };
}
function unenroll(tok, cid) {
  const u = auth_(tok);
  rows_('Inscriptions').filter(i => S(i.pseudo) === S(u.pseudo) && S(i.cours) === S(cid)).reverse()
    .forEach(i => sh_('Inscriptions').deleteRow(i._r));
  return { ok: true };
}
function course(tok, cid) {
  const u = auth_(tok), c = activeCourse_(cid); needEnrolled_(u, cid);
  const qs = rows_('Questions').filter(q => S(q.cours) === S(cid));
  const res = {};
  rows_('Resultats').filter(r => S(r.pseudo) === S(u.pseudo) && S(r.cours) === S(cid)).forEach(r => res[r.phase] = { score: r.score, total: r.total });
  return { ok: true, id: S(c.id), titre: c.titre, description: c.description,
    documents: rows_('Documents').filter(d => S(d.cours) === S(cid)).map(d => ({ titre: d.titre, url: d.url })),
    nbQuestions: { 'Pré-test': qs.filter(q => q.phase === 'Pré-test').length, 'Post-test': qs.filter(q => q.phase === 'Post-test').length },
    resultats: res };
}

/* ---------- Tests ---------- */
function shuffle_(a) { for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; }
function questions_(cid, phase) { return rows_('Questions').filter(q => S(q.cours) === S(cid) && q.phase === phase); }
function done_(pseudo, cid, phase) { return rows_('Resultats').some(r => S(r.pseudo) === S(pseudo) && S(r.cours) === S(cid) && r.phase === phase); }

function startTest(tok, cid, phase) {
  const u = auth_(tok); activeCourse_(cid); needEnrolled_(u, cid);
  if (PHASES.indexOf(phase) < 0) return { ok: false, message: 'Phase inconnue.' };
  if (done_(u.pseudo, cid, phase)) return { ok: false, message: 'Vous avez déjà passé ce test.' };
  if (phase === 'Post-test' && !done_(u.pseudo, cid, 'Pré-test')) return { ok: false, message: 'Passez d’abord le pré-test.' };
  const qs = questions_(cid, phase);
  if (!qs.length) return { ok: false, message: 'Aucune question n’a encore été ajoutée.' };
  return { ok: true, questions: shuffle_(qs.map(q => ({
    id: S(q.id), q: q.question, multi: JSON.parse(q.bonnes).length > 1,
    o: shuffle_(JSON.parse(q.options).map((t, i) => ({ i, t }))) }))) };
}
/** reponses : { idQuestion: [index choisis] } — les non renseignées comptent comme fausses. */
function submitTest(tok, cid, phase, reponses) {
  const u = auth_(tok); needEnrolled_(u, cid);
  if (done_(u.pseudo, cid, phase)) return { ok: false, message: 'Vous avez déjà passé ce test.' };
  if (phase === 'Post-test' && !done_(u.pseudo, cid, 'Pré-test')) return { ok: false, message: 'Passez d’abord le pré-test.' };
  const qs = questions_(cid, phase); let score = 0;
  const corr = qs.map(q => {
    const good = JSON.parse(q.bonnes).map(Number).sort(), got = ((reponses || {})[S(q.id)] || []).map(Number).sort();
    const ok = good.length === got.length && good.every((v, i) => v === got[i]); if (ok) score++;
    const opts = JSON.parse(q.options);
    return { q: q.question, ok, bonnes: good.map(i => opts[i]), explication: q.explication };
  });
  add_('Resultats', { date: new Date(), pseudo: u.pseudo, cours: S(cid), phase, score, total: qs.length });
  return { ok: true, score, total: qs.length, corrige: corr };
}

/* ---------- Discussion ---------- */
function getMessages(tok, cid) {
  const u = auth_(tok); needEnrolled_(u, cid);
  const m = rows_('Messages').filter(x => S(x.cours) === S(cid)).slice(-100)
    .map(x => ({ id: S(x.id), date: x.date, pseudo: x.pseudo, texte: x.texte, mine: S(x.pseudo) === S(u.pseudo) }));
  return { ok: true, messages: m };
}
function postMessage(tok, cid, texte) {
  const u = auth_(tok); needEnrolled_(u, cid);
  texte = S(texte || '').trim().slice(0, 1000);
  if (!texte) return { ok: false, message: 'Message vide.' };
  add_('Messages', { id: id_('m'), date: new Date(), cours: S(cid), pseudo: u.pseudo, texte });
  return { ok: true };
}
function ownMessage_(u, mid) {
  const m = rows_('Messages').filter(x => S(x.id) === S(mid) && S(x.pseudo) === S(u.pseudo))[0];
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
  add_('Questions', { id: id_('q'), cours: S(cid), phase: q.phase, question: S(q.question).trim(), options: JSON.stringify(opts), bonnes: JSON.stringify(bonnes), explication: S(q.explication || '') });
  return { ok: true };
}
